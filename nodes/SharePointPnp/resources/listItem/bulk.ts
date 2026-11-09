import type {
	IDataObject,
	IDisplayOptions,
	IExecuteFunctions,
	INodeProperties,
} from 'n8n-workflow';
import '@pnp/sp/regional-settings/index.js';
import type { IList } from '@pnp/sp/lists/index.js';
import type { SPFI } from '@pnp/sp/index.js';

import { runInBatches } from '../../shared/batch';
import { getList, listRLC } from '../../shared/descriptions';
import { toNodeError, toServerRelativePath } from '../../shared/utils';
import type { BulkItemOutcome, BulkOperationContext, BulkOperationHandler } from '../types';
import {
	MAX_BATCH_SIZE,
	assertFieldResults,
	buildKeyFilter,
	buildKeyViewXml,
	clampBatchSize,
	comparisonSelect,
	isUnchanged,
	newItemId,
	parseListItemId,
	shouldSetKeyOnCreate,
	systemUpdateValues,
	type FieldResult,
	type SystemUpdateSource,
} from './bulkHelpers';
import { buildFormValues, type FieldFormValue, type FieldInfo } from './helpers';

const RESOURCE = 'listItem';
export const BULK_OPERATIONS = ['createMany', 'updateMany', 'deleteMany', 'upsert'];
const MAX_MATCHES = 5000;

const show = (operation: string[]): IDisplayOptions => ({
	show: { resource: [RESOURCE], operation },
});

const batchSizeOption: INodeProperties = {
	displayName: 'Batch Size',
	name: 'batchSize',
	type: 'number',
	typeOptions: { minValue: 1, maxValue: MAX_BATCH_SIZE },
	default: MAX_BATCH_SIZE,
	description: 'The number of requests in one SharePoint $batch request. The maximum is 100.',
};

const continueOnItemErrorOption: INodeProperties = {
	displayName: 'Continue on Item Error',
	name: 'continueOnItemError',
	type: 'boolean',
	default: false,
	description:
		'Whether to output a failed item with "success" set to false and the error, and continue. If off, a failed item stops the node after the batch ran, so other items can already be changed.',
};

const createMissingTermsOption: INodeProperties = {
	displayName: 'Create Missing Terms',
	name: 'createMissingTerms',
	type: 'boolean',
	default: false,
	description:
		'Whether to create a term when a managed metadata (tag) value is a label that does not exist in the term set of the field. The term set must be open for new terms. If off, an unknown label fails.',
};

const folderPathOption: INodeProperties = {
	displayName: 'Folder Path',
	name: 'folderPath',
	type: 'string',
	default: '',
	placeholder: 'Lists/Example/Folder',
	description:
		'Create new items in this folder. A server-relative path, a site-relative path or an absolute URL. Empty creates them in the list root.',
};

const systemUpdateOption: INodeProperties = {
	displayName: 'System Update',
	name: 'systemUpdate',
	type: 'boolean',
	default: false,
	description:
		'Whether to update items without a change to Modified, Modified By or the version number. This uses validateUpdateListItem with bNewDocumentUpdate set to true and reads the current values first.',
};

/** The list and item parameters of the bulk operations. index.ts adds the field input after them. */
export const bulkDescription: INodeProperties[] = [
	{
		...listRLC,
		displayOptions: show(BULK_OPERATIONS),
	},
	{
		displayName: 'Item ID',
		name: 'bulkItemId',
		type: 'string',
		default: '',
		required: true,
		placeholder: 'e.g. {{ $json.ID }}',
		description:
			'The numeric ID of the list item. Use an expression to read the ID from each input item.',
		displayOptions: show(['updateMany', 'deleteMany']),
	},
	{
		displayName: 'Key Field Name or ID',
		name: 'keyField',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getListFieldOptions',
			loadOptionsDependsOn: ['list.value', 'site.value'],
		},
		default: '',
		required: true,
		description:
			'The field that identifies the item. Supported types: single line of text, choice, file name (FileLeafRef), number, currency, ID, yes/no, and lookup or person by ID. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show(['upsert']),
	},
	{
		displayName: 'Key Value',
		name: 'keyValue',
		type: 'string',
		default: '',
		required: true,
		description:
			'The value of the key field to find the item by. For a lookup or person field, use the numeric ID.',
		displayOptions: show(['upsert']),
	},
];

/** The options of the bulk operations. index.ts places them after the field input. */
export const bulkOptionsDescription: INodeProperties[] = [
	{
		displayName: 'Options',
		name: 'bulkCreateOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createMany']),
		options: [
			batchSizeOption,
			continueOnItemErrorOption,
			createMissingTermsOption,
			folderPathOption,
		],
	},
	{
		displayName: 'Options',
		name: 'bulkUpdateOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['updateMany']),
		options: [
			batchSizeOption,
			continueOnItemErrorOption,
			createMissingTermsOption,
			systemUpdateOption,
		],
	},
	{
		displayName: 'Options',
		name: 'bulkDeleteOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['deleteMany']),
		options: [
			batchSizeOption,
			continueOnItemErrorOption,
			{
				displayName: 'Delete Permanently',
				name: 'permanent',
				type: 'boolean',
				default: false,
				description:
					'Whether to delete the items for good. If off, the items go to the recycle bin.',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'upsertOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['upsert']),
		options: [
			batchSizeOption,
			continueOnItemErrorOption,
			createMissingTermsOption,
			{
				displayName: 'Find in Folder',
				name: 'findInFolder',
				type: 'string',
				default: '',
				placeholder: 'Shared Documents/Folder',
				description:
					'Only look for the item directly in this folder, with a CAML query. Use it for lists above the list view threshold (5000 items) when the key field has no index. New items are created in this folder unless Folder Path is set.',
			},
			folderPathOption,
			{
				displayName: 'If Several Match',
				name: 'ifSeveralMatch',
				type: 'options',
				options: [
					{
						name: 'Update First',
						value: 'first',
						description: 'Update the item with the lowest ID',
					},
					{
						name: 'Update All',
						value: 'all',
						description: 'Update all items that match (up to 5000)',
					},
					{ name: 'Fail', value: 'fail', description: 'Fail the input item' },
				],
				default: 'first',
				description: 'What to do when more than one item has the key value',
			},
			{
				displayName: 'Skip Update If Unchanged',
				name: 'skipUnchanged',
				type: 'boolean',
				default: false,
				description:
					'Whether to read the current values of the given fields and skip the update when they are all equal. Text, choice, number, yes/no, date, lookup and multi-choice fields can be compared. Other field types always count as changed.',
			},
			systemUpdateOption,
		],
	},
];

/** Functions of the List Item resource that the bulk operations reuse. */
export interface ListItemFieldTools {
	readFieldValues(
		this: IExecuteFunctions,
		itemIndex: number,
		operation: 'create' | 'update',
	): Record<string, unknown>;
	resolveTags(
		this: IExecuteFunctions,
		siteUrl: string,
		values: Record<string, unknown>,
		fieldInfos: FieldInfo[],
		createMissing: boolean,
	): Promise<Record<string, unknown>>;
}

interface BulkOptions {
	batchSize?: number;
	continueOnItemError?: boolean;
	createMissingTerms?: boolean;
	folderPath?: string;
	findInFolder?: string;
	systemUpdate?: boolean;
	permanent?: boolean;
	ifSeveralMatch?: 'first' | 'all' | 'fail';
	skipUnchanged?: boolean;
}

const OPTIONS_PARAMETER: Record<string, string> = {
	createMany: 'bulkCreateOptions',
	updateMany: 'bulkUpdateOptions',
	deleteMany: 'bulkDeleteOptions',
	upsert: 'upsertOptions',
};

async function getFieldInfos(list: IList): Promise<FieldInfo[]> {
	return (await list.fields.select(
		'InternalName',
		'TypeAsString',
		'TermSetId',
		'SspId',
		'AnchorId',
	)()) as FieldInfo[];
}

/** The input items of one list, with helpers to read options and to record outcomes. */
interface ListScope {
	list: IList;
	/** Builds the list on another SPFI, for example a batched one, from the list parameter. */
	listOn: (sp: SPFI) => IList;
	itemIndexes: number[];
	batchSize: number;
	options: (itemIndex: number) => BulkOptions;
	succeed: (itemIndex: number, result: IDataObject) => void;
	fail: (itemIndex: number, error: unknown, extra?: IDataObject) => void;
}

/**
 * Groups the input items by list and calls `handle` once for each list. A list-level error, for
 * example a list that does not exist, fails the input items of that list that have no outcome yet.
 */
async function forEachList(
	this: IExecuteFunctions,
	operation: string,
	context: BulkOperationContext,
	handle: (scope: ListScope) => Promise<void>,
): Promise<BulkItemOutcome[]> {
	const outcomes = new Map<number, BulkItemOutcome>();
	const optionCache = new Map<number, BulkOptions>();
	const options = (itemIndex: number): BulkOptions => {
		let value = optionCache.get(itemIndex);
		if (!value) {
			value = this.getNodeParameter(OPTIONS_PARAMETER[operation], itemIndex, {}) as BulkOptions;
			optionCache.set(itemIndex, value);
		}
		return value;
	};
	const succeed = (itemIndex: number, result: IDataObject) => {
		outcomes.set(itemIndex, { itemIndex, result });
	};
	const fail = (itemIndex: number, error: unknown, extra: IDataObject = {}) => {
		if (!options(itemIndex).continueOnItemError) {
			outcomes.set(itemIndex, { itemIndex, error });
			return;
		}
		const message = toNodeError(this.getNode(), error, itemIndex).message;
		outcomes.set(itemIndex, { itemIndex, result: { ...extra, success: false, error: message } });
	};

	const groups = new Map<string, number[]>();
	for (const itemIndex of context.itemIndexes) {
		try {
			const locator = this.getNodeParameter('list', itemIndex) as {
				mode?: string;
				value?: unknown;
			};
			const value = String(locator?.value ?? '').trim();
			if (!value) throw new Error('The list must not be empty');
			const key = `${locator.mode === 'title' ? 'title' : 'id'}:${value.toLowerCase()}`;
			groups.set(key, [...(groups.get(key) ?? []), itemIndex]);
		} catch (error) {
			fail(itemIndex, error);
		}
	}

	for (const itemIndexes of groups.values()) {
		const first = itemIndexes[0];
		try {
			await handle({
				list: getList.call(this, context.sp, first),
				listOn: (sp) => getList.call(this, sp, first),
				itemIndexes,
				batchSize: clampBatchSize(options(first).batchSize),
				options,
				succeed,
				fail,
			});
		} catch (error) {
			for (const itemIndex of itemIndexes) {
				if (!outcomes.has(itemIndex)) fail(itemIndex, error);
			}
		}
	}
	return [...outcomes.values()];
}

function toLocalTimes(sp: SPFI, utcTimes: string[], batchSize: number) {
	return runInBatches(
		sp,
		utcTimes,
		(batchedSp, utc) => batchedSp.web.regionalSettings.timeZone.utcToLocalTime(utc),
		batchSize,
	);
}

type SystemUpdateInfo = { source: SystemUpdateSource; localModified?: string } | { error: unknown };

/**
 * Gets what a system update needs for each item: the current Modified and Editor values and Modified
 * in the time zone of the site. Items without a known source are read first. Both steps use $batch.
 */
async function getSystemUpdateInfo(
	listOn: (sp: SPFI) => IList,
	sp: SPFI,
	ids: number[],
	batchSize: number,
	known: Map<number, SystemUpdateSource> = new Map(),
): Promise<Map<number, SystemUpdateInfo>> {
	const info = new Map<number, SystemUpdateInfo>();
	const sources = new Map(known);
	const unknownIds = [...new Set(ids.filter((id) => !known.has(id)))];
	const reads = await runInBatches(
		sp,
		unknownIds,
		(batchedSp, id) =>
			listOn(batchedSp)
				.items.getById(id)
				.select('Modified', 'Editor/Name')
				.expand('Editor')() as Promise<SystemUpdateSource>,
		batchSize,
	);
	for (const read of reads) {
		if (read.ok) sources.set(unknownIds[read.index], read.value as SystemUpdateSource);
		else info.set(unknownIds[read.index], { error: read.error });
	}

	const withModified = [...sources.keys()].filter((id) => sources.get(id)?.Modified);
	const times = await toLocalTimes(
		sp,
		withModified.map((id) => sources.get(id)!.Modified!),
		batchSize,
	);
	for (const [id, source] of sources) info.set(id, { source });
	for (const time of times) {
		const id = withModified[time.index];
		info.set(
			id,
			time.ok ? { source: sources.get(id)!, localModified: time.value } : { error: time.error },
		);
	}
	return info;
}

export function createBulkOperations(
	tools: ListItemFieldTools,
): Record<string, BulkOperationHandler> {
	async function readValues(
		this: IExecuteFunctions,
		siteUrl: string,
		itemIndex: number,
		operation: 'create' | 'update',
		fieldInfos: FieldInfo[],
		options: BulkOptions,
	): Promise<Record<string, unknown>> {
		return await tools.resolveTags.call(
			this,
			siteUrl,
			tools.readFieldValues.call(this, itemIndex, operation),
			fieldInfos,
			options.createMissingTerms === true,
		);
	}

	/** The folder for new items: Folder Path, else Find in Folder (Create or Update), else the list root. */
	function folderPathFor(siteUrl: string, options: BulkOptions, rootFolder: () => Promise<string>) {
		const path = options.folderPath?.trim() || options.findInFolder?.trim();
		return path ? Promise.resolve(toServerRelativePath(siteUrl, path)) : rootFolder();
	}

	function rootFolderReader(list: IList): () => Promise<string> {
		let rootFolder: Promise<string> | undefined;
		return async () =>
			await (rootFolder ??= list.rootFolder
				.select('ServerRelativeUrl')()
				.then((folder) => folder.ServerRelativeUrl));
	}

	async function createMany(this: IExecuteFunctions, context: BulkOperationContext) {
		const { siteUrl } = context;
		return await forEachList.call(this, 'createMany', context, async (scope) => {
			const fieldInfos = await getFieldInfos(scope.list);
			const rootFolder = rootFolderReader(scope.list);
			const prepared: Array<{
				itemIndex: number;
				formValues: FieldFormValue[];
				folderPath: string;
			}> = [];
			for (const itemIndex of scope.itemIndexes) {
				const options = scope.options(itemIndex);
				try {
					const values = await readValues.call(
						this,
						siteUrl,
						itemIndex,
						'create',
						fieldInfos,
						options,
					);
					prepared.push({
						itemIndex,
						formValues: buildFormValues(values, fieldInfos),
						folderPath: await folderPathFor(siteUrl, options, rootFolder),
					});
				} catch (error) {
					scope.fail(itemIndex, error);
				}
			}

			const results = await runInBatches(
				context.sp,
				prepared,
				(batchedSp, entry) =>
					scope
						.listOn(batchedSp)
						.addValidateUpdateItemUsingPath(entry.formValues, entry.folderPath, false)
						.then((fieldResults) => newItemId(fieldResults as FieldResult[])),
				scope.batchSize,
			);
			for (const result of results) {
				const { itemIndex } = prepared[result.index];
				if (result.ok) scope.succeed(itemIndex, { id: result.value, success: true });
				else scope.fail(itemIndex, result.error);
			}
		});
	}

	async function updateMany(this: IExecuteFunctions, context: BulkOperationContext) {
		const { siteUrl } = context;
		return await forEachList.call(this, 'updateMany', context, async (scope) => {
			const fieldInfos = await getFieldInfos(scope.list);
			const prepared: Array<{
				itemIndex: number;
				id: number;
				formValues: FieldFormValue[];
				systemUpdate: boolean;
			}> = [];
			for (const itemIndex of scope.itemIndexes) {
				const options = scope.options(itemIndex);
				let id: number | undefined;
				try {
					id = parseListItemId(this.getNodeParameter('bulkItemId', itemIndex, ''));
					const values = await readValues.call(
						this,
						siteUrl,
						itemIndex,
						'update',
						fieldInfos,
						options,
					);
					const formValues = buildFormValues(values, fieldInfos);
					if (!formValues.length) throw new Error('No fields to update were given');
					prepared.push({ itemIndex, id, formValues, systemUpdate: options.systemUpdate === true });
				} catch (error) {
					scope.fail(itemIndex, error, id ? { id } : {});
				}
			}

			const systemIds = prepared.filter((entry) => entry.systemUpdate).map((entry) => entry.id);
			const systemInfo = systemIds.length
				? await getSystemUpdateInfo(scope.listOn, context.sp, systemIds, scope.batchSize)
				: new Map<number, SystemUpdateInfo>();
			const writable = prepared.filter((entry) => {
				if (!entry.systemUpdate) return true;
				const info = systemInfo.get(entry.id);
				if (!info || 'error' in info) {
					scope.fail(entry.itemIndex, info?.error ?? new Error('The item was not found'), {
						id: entry.id,
					});
					return false;
				}
				entry.formValues.push(
					...systemUpdateValues(info.source, info.localModified, entry.formValues),
				);
				return true;
			});

			const results = await runInBatches(
				context.sp,
				writable,
				(batchedSp, entry) =>
					scope
						.listOn(batchedSp)
						.items.getById(entry.id)
						.validateUpdateListItem(entry.formValues, entry.systemUpdate)
						.then((fieldResults) => assertFieldResults(fieldResults as FieldResult[])),
				scope.batchSize,
			);
			for (const result of results) {
				const { itemIndex, id } = writable[result.index];
				if (result.ok) scope.succeed(itemIndex, { id, success: true });
				else scope.fail(itemIndex, result.error, { id });
			}
		});
	}

	async function deleteMany(this: IExecuteFunctions, context: BulkOperationContext) {
		return await forEachList.call(this, 'deleteMany', context, async (scope) => {
			const prepared: Array<{ itemIndex: number; id: number; permanent: boolean }> = [];
			for (const itemIndex of scope.itemIndexes) {
				try {
					const id = parseListItemId(this.getNodeParameter('bulkItemId', itemIndex, ''));
					prepared.push({ itemIndex, id, permanent: scope.options(itemIndex).permanent === true });
				} catch (error) {
					scope.fail(itemIndex, error);
				}
			}
			const results = await runInBatches(
				context.sp,
				prepared,
				(batchedSp, entry): Promise<unknown> => {
					const item = scope.listOn(batchedSp).items.getById(entry.id);
					return entry.permanent ? item.delete() : item.recycle();
				},
				scope.batchSize,
			);
			for (const result of results) {
				const { itemIndex, id, permanent } = prepared[result.index];
				if (result.ok) scope.succeed(itemIndex, { id, success: true, recycled: !permanent });
				else scope.fail(itemIndex, result.error, { id });
			}
		});
	}

	type Match = Record<string, unknown> & SystemUpdateSource & { Id: number };

	/** The number of matches to read: one for Update First, two to detect several for Fail. */
	function matchLimit(options: BulkOptions): number {
		const mode = options.ifSeveralMatch ?? 'first';
		return mode === 'all' ? MAX_MATCHES : mode === 'fail' ? 2 : 1;
	}

	interface UpsertInput {
		itemIndex: number;
		options: BulkOptions;
		keyField: string;
		keyValue: unknown;
		/** An OData filter, or a CAML query scoped to a folder (Find in Folder). */
		lookup: { filter: string } | { viewXml: string; folder: string };
		values: Record<string, unknown>;
		select: string[];
	}

	interface WriteTask {
		itemIndex: number;
		/** The item to update. Undefined creates an item. */
		id?: number;
		formValues: FieldFormValue[];
		systemUpdate: boolean;
		folderPath?: string;
	}

	async function upsert(this: IExecuteFunctions, context: BulkOperationContext) {
		const { siteUrl } = context;
		return await forEachList.call(this, 'upsert', context, async (scope) => {
			const fieldInfos = await getFieldInfos(scope.list);
			const types = new Map(fieldInfos.map((field) => [field.InternalName, field.TypeAsString]));
			const rootFolder = rootFolderReader(scope.list);

			const inputs: UpsertInput[] = [];
			for (const itemIndex of scope.itemIndexes) {
				const options = scope.options(itemIndex);
				try {
					const keyField = String(this.getNodeParameter('keyField', itemIndex, '') ?? '').trim();
					if (!keyField) throw new Error('The key field must not be empty');
					const keyType = types.get(keyField);
					if (!keyType)
						throw new Error(`The list has no field with the internal name "${keyField}"`);
					const keyValue = this.getNodeParameter('keyValue', itemIndex, '');
					const findInFolder = options.findInFolder?.trim();
					const lookup = findInFolder
						? {
								viewXml: buildKeyViewXml(keyField, keyType, keyValue, matchLimit(options)),
								folder: toServerRelativePath(siteUrl, findInFolder),
							}
						: { filter: buildKeyFilter(keyField, keyType, keyValue) };
					const values = await readValues.call(
						this,
						siteUrl,
						itemIndex,
						'update',
						fieldInfos,
						options,
					);
					const select = new Set(['Id']);
					if (options.skipUnchanged) {
						for (const name of Object.keys(values)) {
							const property = comparisonSelect(name, types.get(name));
							if (property) select.add(property);
						}
					}
					if (options.systemUpdate) {
						select.add('Modified');
						select.add('Editor/Name');
					}
					inputs.push({
						itemIndex,
						options,
						keyField,
						keyValue,
						lookup,
						values,
						select: [...select],
					});
				} catch (error) {
					scope.fail(itemIndex, error);
				}
			}

			const found = await runInBatches(
				context.sp,
				inputs,
				(batchedSp, input) => {
					const expand = input.select.includes('Editor/Name') ? ['Editor'] : [];
					const { lookup } = input;
					if ('viewXml' in lookup) {
						return scope
							.listOn(batchedSp)
							.getItemsByCAMLQuery(
								{ ViewXml: lookup.viewXml, FolderServerRelativeUrl: lookup.folder },
								...expand,
							) as Promise<Match[]>;
					}
					let query = scope.listOn(batchedSp).items.select(...input.select);
					if (expand.length) query = query.expand(...expand);
					return query
						.filter(lookup.filter)
						.orderBy('Id', true)
						.top(matchLimit(input.options))() as Promise<Match[]>;
				},
				scope.batchSize,
			);

			const tasks: WriteTask[] = [];
			const summary = new Map<
				number,
				{ action: 'created' | 'updated' | 'unchanged'; ids: number[] }
			>();
			const sources = new Map<number, SystemUpdateSource>();
			for (const result of found) {
				const input = inputs[result.index];
				const { itemIndex, options } = input;
				try {
					if (!result.ok) throw result.error;
					const matches = result.value ?? [];
					if (!matches.length) {
						const values = { ...input.values };
						if (
							!(input.keyField in values) &&
							shouldSetKeyOnCreate(types.get(input.keyField) ?? '')
						) {
							values[input.keyField] = input.keyValue;
						}
						tasks.push({
							itemIndex,
							formValues: buildFormValues(values, fieldInfos),
							systemUpdate: false,
							folderPath: await folderPathFor(siteUrl, options, rootFolder),
						});
						summary.set(itemIndex, { action: 'created', ids: [] });
						continue;
					}
					if (matches.length > 1 && options.ifSeveralMatch === 'fail') {
						throw new Error(
							`Several items have the value "${String(input.keyValue)}" in the field "${input.keyField}"`,
						);
					}
					const formValues = buildFormValues(input.values, fieldInfos);
					const toUpdate = formValues.length
						? matches.filter(
								(match) => !(options.skipUnchanged && isUnchanged(input.values, types, match)),
							)
						: [];
					summary.set(itemIndex, {
						action: toUpdate.length ? 'updated' : 'unchanged',
						ids: matches.map((match) => match.Id),
					});
					for (const match of toUpdate) {
						if (options.systemUpdate) sources.set(match.Id, match);
						tasks.push({
							itemIndex,
							id: match.Id,
							formValues: [...formValues],
							systemUpdate: options.systemUpdate === true,
						});
					}
				} catch (error) {
					scope.fail(itemIndex, error);
				}
			}

			const failedInputs = new Map<number, unknown>();
			if (sources.size) {
				const systemInfo = await getSystemUpdateInfo(
					scope.listOn,
					context.sp,
					[...sources.keys()],
					scope.batchSize,
					sources,
				);
				for (const task of tasks) {
					if (!task.systemUpdate || task.id === undefined) continue;
					const info = systemInfo.get(task.id);
					if (!info || 'error' in info) {
						if (!failedInputs.has(task.itemIndex)) {
							failedInputs.set(task.itemIndex, info?.error ?? new Error('The item was not found'));
						}
						continue;
					}
					task.formValues.push(
						...systemUpdateValues(info.source, info.localModified, task.formValues),
					);
				}
			}

			const writable = tasks.filter((task) => !failedInputs.has(task.itemIndex));
			const writes = await runInBatches(
				context.sp,
				writable,
				(batchedSp, task) => {
					const batchedList = scope.listOn(batchedSp);
					if (task.id === undefined) {
						return batchedList
							.addValidateUpdateItemUsingPath(task.formValues, task.folderPath ?? '', false)
							.then((fieldResults) => newItemId(fieldResults as FieldResult[]));
					}
					const id = task.id;
					return batchedList.items
						.getById(id)
						.validateUpdateListItem(task.formValues, task.systemUpdate)
						.then((fieldResults) => {
							assertFieldResults(fieldResults as FieldResult[]);
							return id;
						});
				},
				scope.batchSize,
			);
			for (const write of writes) {
				const task = writable[write.index];
				if (!write.ok) {
					if (!failedInputs.has(task.itemIndex)) failedInputs.set(task.itemIndex, write.error);
				} else if (task.id === undefined) {
					summary.get(task.itemIndex)?.ids.push(write.value as number);
				}
			}

			for (const [itemIndex, { action, ids }] of summary) {
				if (failedInputs.has(itemIndex)) {
					scope.fail(itemIndex, failedInputs.get(itemIndex), ids.length ? { id: ids[0] } : {});
					continue;
				}
				const result: IDataObject = { id: ids[0], success: true, action };
				if (ids.length > 1) result.ids = ids;
				scope.succeed(itemIndex, result);
			}
		});
	}

	return { createMany, updateMany, deleteMany, upsert };
}
