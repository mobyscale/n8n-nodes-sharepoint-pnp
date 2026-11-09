import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { runInBatches, type BatchResult } from '../../shared/batch';
import { buildViewXml, undefinedIfNotFound } from '../../shared/filesFolders';
import { getSharePointErrorMessage } from '../../shared/utils';
import { buildFormValues, type FieldFormValue, type FieldInfo } from '../listItem/helpers';
import {
	describeFieldErrors,
	isEmptyValue,
	selectNameFor,
	unknownFields,
	type FieldResult,
	type FieldValues,
	type ListFields,
} from './fieldValues';

export type IncludeItems = 'files' | 'folders' | 'both';

export interface MetadataOptions {
	include?: IncludeItems;
	includeFolderItself?: boolean;
	maxDepth?: number;
	onlyFieldsInContentType?: boolean;
	ignoreEmptyValues?: boolean;
	skipUnchanged?: boolean;
	systemUpdate?: boolean;
	batchSize?: number;
	continueOnItemError?: boolean;
}

export interface TargetItem {
	id: number;
	type: 'file' | 'folder';
	serverRelativeUrl: string;
	contentTypeId?: string;
	modifiedUtc?: string;
	editorId?: number;
	/** The current values of the fields to set, as SharePoint returned them. */
	values: Record<string, unknown>;
}

export interface PageResult {
	rows: TargetItem[];
	next?: string;
}

/** Reads one page of the direct children of a folder. `next` is the paging token of the previous page. */
export type PageReader = (folderPath: string, next?: string) => Promise<PageResult>;

export interface ItemPlan {
	item: TargetItem;
	values: FieldValues;
}

export interface ItemWriter {
	/** Writes one batch. The results have the same order as the plans. */
	writeChunk(plans: ItemPlan[]): Promise<Array<BatchResult<unknown>>>;
}

export const MAX_BATCH_SIZE = 100;

/** Field types whose current value in a view row can be compared as plain text. */
const ROW_COMPARABLE_TYPES = new Set(['Text', 'Choice']);

export async function readAllPages(reader: PageReader, folderPath: string): Promise<TargetItem[]> {
	const rows: TargetItem[] = [];
	let next: string | undefined;
	do {
		const page = await reader(folderPath, next);
		rows.push(...page.rows);
		next = page.next;
	} while (next);
	return rows;
}

/**
 * Walks the folder tree one folder at a time. Each request is scoped to one folder and paged, so a
 * large library does not hit the list view threshold. A max depth of 0 has no limit, 1 reads only
 * the direct children.
 */
export async function collectTargets(
	rootPath: string,
	reader: PageReader,
	options: Pick<MetadataOptions, 'include' | 'maxDepth'>,
): Promise<TargetItem[]> {
	const include = options.include ?? 'both';
	const maxDepth = Math.max(0, Math.floor(options.maxDepth ?? 0));
	const targets: TargetItem[] = [];
	const queue: Array<{ path: string; depth: number }> = [{ path: rootPath, depth: 1 }];
	while (queue.length) {
		const { path, depth } = queue.shift() as { path: string; depth: number };
		for (const row of await readAllPages(reader, path)) {
			if (row.type === 'folder') {
				if (include !== 'files') targets.push(row);
				if (maxDepth === 0 || depth < maxDepth) {
					queue.push({ path: row.serverRelativeUrl, depth: depth + 1 });
				}
			} else if (include !== 'folders') {
				targets.push(row);
			}
		}
	}
	return targets;
}

/** Removes empty values when "Ignore Empty Values" is on. */
export function prepareFields(fields: FieldValues, ignoreEmptyValues: boolean): FieldValues {
	if (!ignoreEmptyValues) return { ...fields };
	return Object.fromEntries(Object.entries(fields).filter(([, value]) => !isEmptyValue(value)));
}

/** Decides which values to write to one item. An empty result means the item is skipped. */
export function planItemValues(
	item: TargetItem,
	fields: FieldValues,
	types: Map<string, string>,
	options: {
		contentTypeFields?: Map<string, Set<string>>;
		skipUnchanged?: boolean;
	},
): FieldValues {
	const allowed = item.contentTypeId
		? options.contentTypeFields?.get(item.contentTypeId)
		: undefined;
	const values: FieldValues = {};
	for (const [name, value] of Object.entries(fields)) {
		if (allowed && !allowed.has(name)) continue;
		if (
			options.skipUnchanged &&
			ROW_COMPARABLE_TYPES.has(types.get(name) ?? '') &&
			String(item.values[name] ?? '') === String(value ?? '')
		) {
			continue;
		}
		values[name] = value;
	}
	return values;
}

function errorMessage(error: unknown): string {
	const value = error as { message?: string; isHttpRequestError?: boolean };
	if (value?.isHttpRequestError && typeof value.message === 'string') {
		return getSharePointErrorMessage(value as { message: string });
	}
	return value?.message ?? String(error);
}

export interface ApplyResult {
	updated: number;
	failed: number;
	errors: IDataObject[];
}

/** Writes the plans in batches. Without "Continue On Item Error" the first failed batch stops the run. */
export async function applyPlans(
	plans: ItemPlan[],
	writer: ItemWriter,
	options: Pick<MetadataOptions, 'batchSize' | 'continueOnItemError'>,
): Promise<ApplyResult> {
	const size = Math.min(
		MAX_BATCH_SIZE,
		Math.max(1, Math.floor(options.batchSize ?? MAX_BATCH_SIZE)),
	);
	const result: ApplyResult = { updated: 0, failed: 0, errors: [] };
	for (let start = 0; start < plans.length; start += size) {
		const chunk = plans.slice(start, start + size);
		const outcomes = await writer.writeChunk(chunk);
		outcomes.forEach((outcome, index) => {
			if (outcome.ok) {
				result.updated++;
				return;
			}
			const { item } = chunk[index];
			result.failed++;
			result.errors.push({
				listItemId: item.id,
				serverRelativeUrl: item.serverRelativeUrl,
				message: errorMessage(outcome.error),
			});
		});
		if (result.failed && !options.continueOnItemError) {
			const first = result.errors[0];
			throw new Error(
				`The update of "${String(first.serverRelativeUrl)}" failed: ${String(first.message)}. ${result.updated} item(s) were updated before the run stopped. Turn on "Continue On Item Error" to update the other items.`,
			);
		}
	}
	return result;
}

export interface SetMetadataSource {
	listFields: ListFields;
	readPage: PageReader;
	readSelf(): Promise<TargetItem | undefined>;
	readContentTypeFields(): Promise<Map<string, Set<string>>>;
	writer: ItemWriter;
}

export async function setMetadataRecursive(
	folderPath: string,
	fields: FieldValues,
	options: MetadataOptions,
	source: SetMetadataSource,
): Promise<IDataObject> {
	const values = prepareFields(fields, options.ignoreEmptyValues === true);
	if (!Object.keys(values).length) throw new Error('No fields to set were given');
	const { types } = source.listFields;
	const missing = unknownFields(Object.keys(values), types);
	if (missing.length) {
		throw new Error(
			`These fields do not exist in the library: ${missing.join(', ')}. Use the internal field names.`,
		);
	}
	const contentTypeFields = options.onlyFieldsInContentType
		? await source.readContentTypeFields()
		: undefined;

	const targets: TargetItem[] = [];
	if (options.includeFolderItself) {
		const self = await source.readSelf();
		if (self) targets.push(self);
	}
	targets.push(...(await collectTargets(folderPath, source.readPage, options)));

	const plans: ItemPlan[] = [];
	let skipped = 0;
	for (const item of targets) {
		const itemValues = planItemValues(item, values, types, {
			contentTypeFields,
			skipUnchanged: options.skipUnchanged ?? true,
		});
		if (Object.keys(itemValues).length) plans.push({ item, values: itemValues });
		else skipped++;
	}
	const applied = await applyPlans(plans, source.writer, options);
	return {
		folderPath,
		matched: targets.length,
		updated: applied.updated,
		skipped,
		failed: applied.failed,
		errors: applied.errors,
	};
}

interface ViewRow {
	ID?: number | string;
	FSObjType?: number | string;
	FileRef?: string;
	ContentTypeId?: string;
	Modified?: string;
	'Modified.'?: string;
	Editor?: Array<{ id?: number | string }>;
	[key: string]: unknown;
}

function isIsoDate(value: unknown): value is string {
	return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(value);
}

export function rowToTarget(row: ViewRow, fieldNames: string[]): TargetItem {
	const modified = isIsoDate(row['Modified.']) ? row['Modified.'] : row.Modified;
	const editorId = Number(row.Editor?.[0]?.id);
	return {
		id: Number(row.ID),
		type: String(row.FSObjType) === '1' ? 'folder' : 'file',
		serverRelativeUrl: String(row.FileRef ?? ''),
		contentTypeId: row.ContentTypeId,
		modifiedUtc: isIsoDate(modified) ? modified : undefined,
		editorId: Number.isFinite(editorId) && editorId > 0 ? editorId : undefined,
		values: Object.fromEntries(fieldNames.map((name) => [name, row[name]])),
	};
}

/** Reads folder pages with RenderListDataAsStream. Each page holds up to 500 rows. */
export function createPageReader(sp: SPFI, listId: string, fieldNames: string[]): PageReader {
	const viewXml = buildViewXml({ include: 'all' }, 500, ['ContentTypeId', ...fieldNames]);
	return async (folderPath, next) => {
		const query = next
			? new Map(new URLSearchParams(next.replace(/^\?/, '')).entries())
			: new Map<string, string>();
		const page = await sp.web.lists.getById(listId).renderListDataAsStream(
			{
				ViewXml: viewXml,
				FolderServerRelativeUrl: folderPath,
				RenderOptions: 2,
				DatesInUtc: true,
			},
			null,
			query,
		);
		return {
			rows: (page.Row as unknown as ViewRow[]).map((row) => rowToTarget(row, fieldNames)),
			next: page.NextHref || undefined,
		};
	};
}

export async function readFolderItself(
	sp: SPFI,
	folderPath: string,
	fieldNames: string[],
	types: Map<string, string>,
): Promise<TargetItem | undefined> {
	const selects = fieldNames.map((name) => selectNameFor(name, types.get(name)));
	const item = (await undefinedIfNotFound(
		sp.web
			.getFolderByServerRelativePath(folderPath)
			.listItemAllFields.select(
				'Id',
				'FileRef',
				'ContentTypeId',
				'Modified',
				'EditorId',
				...selects,
			)(),
	)) as Record<string, unknown> | undefined;
	const id = Number(item?.Id);
	if (!item || !Number.isFinite(id) || id <= 0) return undefined;
	const editorId = Number(item.EditorId);
	return {
		id,
		type: 'folder',
		serverRelativeUrl: String(item.FileRef ?? folderPath),
		contentTypeId: item.ContentTypeId as string | undefined,
		modifiedUtc: isIsoDate(item.Modified) ? item.Modified : undefined,
		editorId: Number.isFinite(editorId) && editorId > 0 ? editorId : undefined,
		values: Object.fromEntries(fieldNames.map((name) => [name, item[name]])),
	};
}

export async function readContentTypeFields(
	sp: SPFI,
	listId: string,
): Promise<Map<string, Set<string>>> {
	const contentTypes = (await sp.web.lists
		.getById(listId)
		.contentTypes.select('StringId', 'Fields/InternalName')
		.expand('Fields')()) as unknown as Array<{
		StringId: string;
		Fields?: Array<{ InternalName: string }>;
	}>;
	return new Map(
		contentTypes.map((entry) => [
			entry.StringId,
			new Set((entry.Fields ?? []).map((field) => field.InternalName)),
		]),
	);
}

/** "2026-01-01T10:00:00.000" from utcToLocalTime becomes "2026-01-01 10:00:00". */
export function toFormDate(local: string): string {
	return local
		.replace('T', ' ')
		.replace(/\.\d+$/, '')
		.replace(/Z$/, '');
}

/**
 * Writes items with validateUpdateListItem in SharePoint batches. A system update uses
 * bNewDocumentUpdate (no new version) and sends the current Modified and Editor values back, so
 * they do not change. That needs one more batch for each chunk to convert the dates to site time.
 */
export function createItemWriter(
	sp: SPFI,
	listId: string,
	infos: FieldInfo[],
	systemUpdate: boolean,
): ItemWriter {
	const logins = new Map<number, string | undefined>();

	async function systemValues(chunk: ItemPlan[]): Promise<FieldFormValue[][]> {
		const editors = [...new Set(chunk.map((plan) => plan.item.editorId))].filter(
			(id): id is number => id !== undefined && !logins.has(id),
		);
		if (editors.length) {
			const users = await runInBatches(
				sp,
				editors,
				(batched, id) => batched.web.siteUsers.getById(id).select('LoginName')(),
				editors.length,
			);
			users.forEach((user, index) =>
				logins.set(
					editors[index],
					user.ok ? (user.value as { LoginName?: string }).LoginName : undefined,
				),
			);
		}
		const times = await runInBatches(
			sp,
			chunk,
			(batched, plan) =>
				plan.item.modifiedUtc
					? batched.web.regionalSettings.timeZone.utcToLocalTime(plan.item.modifiedUtc)
					: Promise.resolve(undefined),
			chunk.length,
		);
		return chunk.map((plan, index) => {
			const extra: FieldFormValue[] = [];
			const local = times[index];
			if (local.ok && typeof local.value === 'string') {
				extra.push({ FieldName: 'Modified', FieldValue: toFormDate(local.value) });
			}
			const login = plan.item.editorId === undefined ? undefined : logins.get(plan.item.editorId);
			if (login?.startsWith('i:0#')) {
				extra.push({ FieldName: 'Editor', FieldValue: JSON.stringify([{ Key: login }]) });
			}
			return extra;
		});
	}

	return {
		async writeChunk(chunk) {
			if (!chunk.length) return [];
			const extras = systemUpdate ? await systemValues(chunk) : chunk.map(() => []);
			return await runInBatches(
				sp,
				chunk,
				(batched, plan, index) =>
					batched.web.lists
						.getById(listId)
						.items.getById(plan.item.id)
						.validateUpdateListItem(
							[...buildFormValues(plan.values, infos), ...extras[index]],
							systemUpdate,
						)
						.then((results: FieldResult[]) => {
							const message = describeFieldErrors(results);
							if (message) throw new Error(message);
							return results;
						}),
				chunk.length,
			);
		},
	};
}
