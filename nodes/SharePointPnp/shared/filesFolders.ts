import type { IDataObject, INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError, sleep } from 'n8n-workflow';
import type { IMoveCopyOptions, SPFI } from '@pnp/sp/index.js';
import { SPQueryable, spPost } from '@pnp/sp/index.js';

import { pathRLC } from './descriptions';
import { getSiteServerRelativeUrl, isNotFoundError, splitPath, trimTrailingSlash } from './utils';

export type IfExists = 'fail' | 'replace' | 'keepBoth';

export interface MoveCopyFields {
	newName?: string;
	ifExists?: IfExists;
	bypassSharedLock?: boolean;
	retainEditorAndModified?: boolean;
	resetAuthorAndCreated?: boolean;
}

/** A file path picker. Only the "By Path" mode exists, because the folder search lists folders only. */
export function filePathRLC(
	overrides: Partial<INodeProperties> & Pick<INodeProperties, 'displayName' | 'name'>,
): INodeProperties {
	const base = pathRLC(overrides);
	return {
		...base,
		default: { mode: 'path', value: '' },
		modes: (base.modes ?? []).filter((mode) => mode.name === 'path'),
		placeholder: undefined,
	};
}

export function buildMoveCopyOptions(
	kind: 'move' | 'copy',
	fields: MoveCopyFields,
): { overwrite: boolean; options: Partial<IMoveCopyOptions> } {
	const ifExists = fields.ifExists ?? 'fail';
	const options: Partial<IMoveCopyOptions> = {
		KeepBoth: ifExists === 'keepBoth',
		ShouldBypassSharedLocks: fields.bypassSharedLock ?? false,
	};
	if (kind === 'move')
		options.RetainEditorAndModifiedOnMove = fields.retainEditorAndModified ?? false;
	else options.ResetAuthorAndCreatedOnCopy = fields.resetAuthorAndCreated ?? false;
	return { overwrite: ifExists === 'replace', options };
}

export function getDestinationPath(
	sourcePath: string,
	destinationFolder: string,
	newName?: string,
): string {
	const name = newName?.trim() || splitPath(sourcePath).name;
	return trimTrailingSlash(`${trimTrailingSlash(destinationFolder)}/${name}`);
}

/** Returns the paths to check when looking for the closest existing parent, deepest first. */
export function getAncestorPaths(path: string, siteUrl: string): string[] {
	const root = getSiteServerRelativeUrl(siteUrl);
	const result: string[] = [];
	let current = trimTrailingSlash(path);
	while (current && current !== '/' && current.length > root.length && current.startsWith(root)) {
		result.push(current);
		current = splitPath(current).parent;
	}
	return result;
}

/** Returns the missing folder paths, shallowest first, below the given existing ancestor. */
export function getMissingSegments(path: string, existingAncestor: string): string[] {
	const trimmed = trimTrailingSlash(path);
	const base = trimTrailingSlash(existingAncestor);
	const rest = trimmed.slice(base.length).split('/').filter(Boolean);
	return rest.map(
		(_, index) => `${base === '/' ? '' : base}/${rest.slice(0, index + 1).join('/')}`,
	);
}

/** Picks the list whose root folder is the longest parent of the path (or the path itself). */
export function findOwningList<T extends { rootUrl: string }>(
	path: string,
	lists: T[],
): T | undefined {
	const target = trimTrailingSlash(path).toLowerCase();
	let best: T | undefined;
	for (const list of lists) {
		const root = trimTrailingSlash(list.rootUrl).toLowerCase();
		if (target === root || target.startsWith(`${root}/`)) {
			if (!best || root.length > trimTrailingSlash(best.rootUrl).length) best = list;
		}
	}
	return best;
}

export type IncludeKind = 'all' | 'files' | 'folders';

export interface FileFolderFilters {
	include: IncludeKind;
	nameContains?: string;
	modifiedAfter?: string;
	modifiedBefore?: string;
}

export function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function combineConditions(operator: 'And' | 'Or', conditions: string[]): string {
	if (conditions.length === 1) return conditions[0];
	const [first, ...rest] = conditions;
	return rest.reduce((left, right) => `<${operator}>${left}${right}</${operator}>`, first);
}

function toIsoDate(value: string): string {
	const date = new Date(value);
	if (Number.isNaN(date.getTime())) throw new Error(`"${value}" is not a valid date`);
	return date.toISOString();
}

const VIEW_FIELDS = [
	'FileLeafRef',
	'FileRef',
	'FSObjType',
	'UniqueId',
	'ID',
	'File_x0020_Size',
	'Created',
	'Modified',
	'Author',
	'Editor',
];

/** Builds the CAML for one folder level. Recursion is done by the caller, one request per folder. */
export function buildViewXml(filters: FileFolderFilters, pageSize = 500): string {
	const conditions: string[] = [];
	if (filters.include === 'files') {
		conditions.push('<Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq>');
	} else if (filters.include === 'folders') {
		conditions.push('<Eq><FieldRef Name="FSObjType"/><Value Type="Integer">1</Value></Eq>');
	}
	if (filters.nameContains) {
		conditions.push(
			`<Contains><FieldRef Name="FileLeafRef"/><Value Type="Text">${escapeXml(filters.nameContains)}</Value></Contains>`,
		);
	}
	if (filters.modifiedAfter) {
		conditions.push(
			`<Gt><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">${toIsoDate(filters.modifiedAfter)}</Value></Gt>`,
		);
	}
	if (filters.modifiedBefore) {
		conditions.push(
			`<Lt><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">${toIsoDate(filters.modifiedBefore)}</Value></Lt>`,
		);
	}
	const where = conditions.length
		? `<Query><Where>${combineConditions('And', conditions)}</Where></Query>`
		: '';
	const viewFields = `<ViewFields>${VIEW_FIELDS.map((name) => `<FieldRef Name="${name}"/>`).join('')}</ViewFields>`;
	return `<View Scope="All">${viewFields}${where}<RowLimit Paged="TRUE">${pageSize}</RowLimit></View>`;
}

export function parseExtensions(value: string | undefined): string[] {
	return (value ?? '')
		.split(/[,;\s]+/)
		.map((entry) =>
			entry
				.trim()
				.replace(/^\*?\./, '')
				.toLowerCase(),
		)
		.filter(Boolean);
}

export function hasExtension(name: string, extensions: string[]): boolean {
	if (!extensions.length) return true;
	const index = name.lastIndexOf('.');
	if (index < 0) return false;
	return extensions.includes(name.slice(index + 1).toLowerCase());
}

interface PersonField {
	title?: string;
	email?: string;
	id?: string;
}

function toPerson(value: unknown): IDataObject | undefined {
	const person = (Array.isArray(value) ? value[0] : value) as PersonField | undefined;
	if (!person || typeof person !== 'object') return undefined;
	return { id: person.id ?? null, name: person.title ?? null, email: person.email ?? null };
}

function asNumber(value: unknown): number | undefined {
	if (value === undefined || value === null || value === '') return undefined;
	const number = Number(value);
	return Number.isFinite(number) ? number : undefined;
}

function pickRaw(row: IDataObject, field: string): unknown {
	return row[`${field}.`] ?? row[field];
}

export function normalizeRow(row: IDataObject, includeFields: boolean): IDataObject {
	const isFolder = String(row.FSObjType) === '1';
	const entry: IDataObject = {
		name: String(row.FileLeafRef ?? ''),
		type: isFolder ? 'folder' : 'file',
		serverRelativeUrl: String(row.FileRef ?? ''),
		uniqueId: String(row.UniqueId ?? '').replace(/[{}]/g, ''),
		listItemId: asNumber(row.ID) ?? null,
		size: isFolder ? null : (asNumber(row.File_x0020_Size) ?? null),
		created: (pickRaw(row, 'Created') as string | undefined) ?? null,
		modified: (pickRaw(row, 'Modified') as string | undefined) ?? null,
		author: toPerson(row.Author) ?? null,
		editor: toPerson(row.Editor) ?? null,
	};
	if (includeFields) entry.fields = row;
	return entry;
}

export interface CopyJobOptions {
	sourceUrl: string;
	destinationFolderUrl: string;
	move: boolean;
	ifExists: IfExists;
	bypassSharedLock: boolean;
	ignoreVersionHistory: boolean;
	newName?: string;
}

const NAME_CONFLICT_BEHAVIOR: Record<IfExists, number> = { fail: 0, replace: 1, keepBoth: 2 };

export function buildCopyJobPayload(options: CopyJobOptions): IDataObject {
	const jobOptions: IDataObject = {
		AllowSchemaMismatch: true,
		BypassSharedLock: options.bypassSharedLock,
		IgnoreVersionHistory: options.ignoreVersionHistory,
		IsMoveMode: options.move,
		NameConflictBehavior: NAME_CONFLICT_BEHAVIOR[options.ifExists],
	};
	if (options.newName) jobOptions.CustomizedItemName = [options.newName];
	return {
		exportObjectUris: [options.sourceUrl],
		destinationUri: options.destinationFolderUrl,
		options: jobOptions,
	};
}

export function toAbsoluteUrl(siteUrl: string, serverRelativePath: string): string {
	return `${new URL(siteUrl).origin}${encodeURI(serverRelativePath)}`;
}

interface CopyJobInfo {
	EncryptionKey?: string;
	JobId?: string;
	JobQueueUri?: string;
}

function unwrapCollection<T>(response: unknown): T[] {
	if (Array.isArray(response)) return response as T[];
	const data = response as { value?: T[]; d?: { CreateCopyJobs?: { results?: T[] } | T[] } };
	if (Array.isArray(data.value)) return data.value;
	const wrapped = data.d?.CreateCopyJobs;
	if (Array.isArray(wrapped)) return wrapped;
	return wrapped?.results ?? [];
}

const JOB_ERROR_EVENTS = new Set(['JobError', 'JobFatalError']);

export interface CopyJobResult extends IDataObject {
	jobId: string | null;
	completed: boolean;
	errors: IDataObject[];
}

function parseLogs(logs: unknown): IDataObject[] {
	if (!Array.isArray(logs)) return [];
	return logs.map((entry) => {
		if (typeof entry !== 'string') return entry as IDataObject;
		try {
			return JSON.parse(entry) as IDataObject;
		} catch {
			return { Message: entry };
		}
	});
}

export async function runCopyJob(
	node: INode,
	sp: SPFI,
	siteUrl: string,
	payload: IDataObject,
	wait: boolean,
	maxWaitSeconds: number,
	itemIndex: number,
): Promise<CopyJobResult> {
	const post = (path: string, body: IDataObject) =>
		spPost(SPQueryable([sp.web, `${siteUrl}/_api/site/${path}`]), {
			body: JSON.stringify(body),
			headers: { 'Content-Type': 'application/json;odata=nometadata' },
		});

	const jobs = unwrapCollection<CopyJobInfo>(await post('CreateCopyJobs', payload));
	const job = jobs[0];
	if (!job?.JobId) {
		throw new NodeOperationError(node, 'SharePoint did not return a copy job', { itemIndex });
	}
	if (!wait) return { jobId: job.JobId, completed: false, errors: [] };

	const deadline = Date.now() + maxWaitSeconds * 1000;
	let sawLogs = false;
	for (;;) {
		const progress = (await post('GetCopyJobProgress', { copyJobInfo: job })) as {
			JobState?: number;
			Logs?: unknown;
		};
		const logs = parseLogs(progress.Logs);
		sawLogs = sawLogs || logs.length > 0;
		const errors = logs.filter(
			(log) =>
				JOB_ERROR_EVENTS.has(String(log.Event)) ||
				(log.Event === 'JobEnd' && Number(log.TotalErrors ?? 0) > 0),
		);
		if (errors.length) return { jobId: job.JobId, completed: false, errors };
		const importEnded = logs.some(
			(log) => log.Event === 'JobEnd' && log.MigrationDirection !== 'Export',
		);
		if (importEnded || (progress.JobState === 0 && sawLogs)) {
			return { jobId: job.JobId, completed: true, errors: [] };
		}
		if (Date.now() > deadline) {
			throw new NodeOperationError(
				node,
				`The copy job ${job.JobId} did not finish within ${maxWaitSeconds} seconds`,
				{ itemIndex },
			);
		}
		await sleep(2000);
	}
}

/** Resolves to undefined when SharePoint answers 404. Other errors reach the node router. */
export async function undefinedIfNotFound<T>(request: Promise<T>): Promise<T | undefined> {
	return await request.catch((error: unknown) =>
		isNotFoundError(error) ? undefined : Promise.reject(error),
	);
}
