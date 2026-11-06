import type { IDataObject, INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { body } from '@pnp/queryable/index.js';
import { Folder } from '@pnp/sp/folders/index.js';
import { spGet, SPQueryable, spPost, type SPFI } from '@pnp/sp/index.js';
import type { IChangeQuery } from '@pnp/sp/types.js';
import type { IWeb } from '@pnp/sp/webs/index.js';

import { getListFromLocator } from './descriptions';
import { isNotFoundError, toServerRelativePath, trimTrailingSlash } from './utils';

export type ChangeScope = 'siteCollection' | 'web' | 'list';

/** The object type flags of SP.ChangeQuery (PnP Core ChangeQueryOptions). */
export const OBJECT_TYPE_FLAGS = [
	'Item',
	'File',
	'Folder',
	'List',
	'Web',
	'Site',
	'User',
	'Group',
	'SecurityPolicy',
	'Field',
	'ContentType',
	'View',
	'Alert',
	'Navigation',
] as const;

/** The change type flags of SP.ChangeQuery (PnP Core ChangeQueryOptions). */
export const CHANGE_TYPE_FLAGS = [
	'Add',
	'Update',
	'DeleteObject',
	'Rename',
	'Move',
	'Restore',
	'SystemUpdate',
	'RoleAssignmentAdd',
	'RoleAssignmentDelete',
	'RoleDefinitionAdd',
	'RoleDefinitionDelete',
	'RoleDefinitionUpdate',
	'GroupMembershipAdd',
	'GroupMembershipDelete',
	'Activity',
] as const;

export type ObjectTypeFlag = (typeof OBJECT_TYPE_FLAGS)[number];
export type ChangeTypeFlag = (typeof CHANGE_TYPE_FLAGS)[number];

/** SP.ChangeType values, in order from 0. */
export const CHANGE_TYPE_NAMES = [
	'NoChange',
	'Add',
	'Update',
	'DeleteObject',
	'Rename',
	'MoveAway',
	'MoveInto',
	'Restore',
	'RoleAdd',
	'RoleDelete',
	'RoleUpdate',
	'AssignmentAdd',
	'AssignmentDelete',
	'MemberAdd',
	'MemberDelete',
	'SystemUpdate',
	'Navigation',
	'ScopeAdd',
	'ScopeDelete',
	'ListContentTypeAdd',
	'ListContentTypeDelete',
	'Dirty',
	'Activity',
] as const;

export const ChangeType = {
	Add: 1,
	Update: 2,
	DeleteObject: 3,
	Rename: 4,
	MoveAway: 5,
	MoveInto: 6,
	Restore: 7,
	AssignmentAdd: 11,
	AssignmentDelete: 12,
	SystemUpdate: 15,
	ScopeAdd: 17,
	ScopeDelete: 18,
} as const;

const SCOPE_CODES: Record<ChangeScope, number> = { siteCollection: 1, web: 2, list: 3 };

const TICKS_AT_UNIX_EPOCH = BigInt('621355968000000000');
const TICKS_PER_MILLISECOND = BigInt(10000);

export function getChangeTypeName(changeType: unknown): string {
	const value = Number(changeType);
	return CHANGE_TYPE_NAMES[value] ?? `Unknown (${String(changeType)})`;
}

/**
 * Builds a change token that starts at a date: "1;<scope>;<id>;<ticks>;-1". The scope is 1 for a
 * site collection, 2 for a site and 3 for a list. The ID is the GUID of that site, web or list.
 */
export function createChangeTokenFromDate(scope: ChangeScope, id: string, date: Date): string {
	const time = date.getTime();
	if (!Number.isFinite(time)) throw new Error('The date is not valid');
	const ticks = BigInt(time) * TICKS_PER_MILLISECOND + TICKS_AT_UNIX_EPOCH;
	return `1;${SCOPE_CODES[scope]};${id.replace(/[{}]/g, '').toLowerCase()};${ticks.toString()};-1`;
}

const OBJECT_TYPE_BY_PROPERTY: Array<[string, string]> = [
	['AlertId', 'Alert'],
	['FieldId', 'Field'],
	['ViewId', 'View'],
	['ContentTypeId', 'ContentType'],
	['GroupId', 'Group'],
	['UserId', 'User'],
	['ItemId', 'Item'],
	['ListId', 'List'],
	['WebId', 'Web'],
	['SiteId', 'Site'],
];

/** Returns "Item", "File", "Folder", "List", "Web" and so on, from the OData type or from the properties. */
export function getObjectType(change: IDataObject): string {
	const odataType =
		(change['odata.type'] as string | undefined) ??
		((change.__metadata as IDataObject | undefined)?.type as string | undefined);
	if (odataType?.startsWith('SP.Change')) return odataType.slice('SP.Change'.length) || 'Unknown';
	for (const [property, objectType] of OBJECT_TYPE_BY_PROPERTY) {
		if (change[property] !== undefined && change[property] !== null) return objectType;
	}
	return 'Unknown';
}

export function getChangeTokenString(change: IDataObject): string | undefined {
	const token = change.ChangeToken as { StringValue?: string } | string | undefined;
	if (typeof token === 'string') return token;
	return token?.StringValue;
}

export interface QueryOptions {
	endPosition?: string;
	fetchLimit?: number;
	latestFirst?: boolean;
	recursiveAll?: boolean;
	requireSecurityTrim?: boolean;
	ignoreStartTokenNotFoundError?: boolean;
}

export interface ChangeOutputOptions {
	includeItemDetails?: boolean;
	itemFields?: string;
	resolveDeletedItems?: boolean;
	folderPathFilter?: string;
}

export interface ChangeQueryInput {
	objectTypes: string[];
	changeTypes: string[];
	startToken?: string;
	queryOptions?: QueryOptions;
}

export type ChangeQuery = IChangeQuery & {
	/** An Edm.Int64 value. SharePoint only accepts it as a string. */
	FetchLimit?: string;
	LatestFirst?: boolean;
	RecursiveAll?: boolean;
	RequireSecurityTrim?: boolean;
	IgnoreStartTokenNotFoundError?: boolean;
	Activity?: boolean;
};

/** Builds the SP.ChangeQuery body from the selected flags, the start token and the query options. */
export function buildChangeQuery(input: ChangeQueryInput): ChangeQuery {
	const query: Record<string, unknown> = {};
	for (const flag of input.objectTypes) {
		if ((OBJECT_TYPE_FLAGS as readonly string[]).includes(flag)) query[flag] = true;
	}
	for (const flag of input.changeTypes) {
		if ((CHANGE_TYPE_FLAGS as readonly string[]).includes(flag)) query[flag] = true;
	}
	const startToken = input.startToken?.trim();
	if (startToken) query.ChangeTokenStart = { StringValue: startToken };

	const options = input.queryOptions ?? {};
	const endPosition = options.endPosition?.trim();
	if (endPosition) query.ChangeTokenEnd = { StringValue: endPosition };
	if (options.fetchLimit !== undefined && options.fetchLimit > 0) {
		query.FetchLimit = String(Math.floor(options.fetchLimit));
	}
	if (options.latestFirst) query.LatestFirst = true;
	if (options.recursiveAll !== undefined) query.RecursiveAll = options.recursiveAll;
	if (options.requireSecurityTrim) query.RequireSecurityTrim = true;
	if (options.ignoreStartTokenNotFoundError) query.IgnoreStartTokenNotFoundError = true;
	return query as ChangeQuery;
}

export type ChangeKind = 'file' | 'folder' | 'item' | 'other';

export interface ChangeEvent {
	name: string;
	value: string;
	description: string;
	objectTypes: ObjectTypeFlag[];
	changeTypeFlags: ChangeTypeFlag[];
	changeTypes: number[];
	kinds: ChangeKind[];
}

const CONTENT: ObjectTypeFlag[] = ['Item', 'File', 'Folder'];

export const CHANGE_EVENTS: ChangeEvent[] = [
	{
		name: 'File Created',
		value: 'fileCreated',
		description: 'A file is added',
		objectTypes: CONTENT,
		changeTypeFlags: ['Add'],
		changeTypes: [ChangeType.Add],
		kinds: ['file'],
	},
	{
		name: 'File Updated',
		value: 'fileUpdated',
		description: 'A file or its properties change',
		objectTypes: CONTENT,
		changeTypeFlags: ['Update'],
		changeTypes: [ChangeType.Update],
		kinds: ['file'],
	},
	{
		name: 'File Deleted',
		value: 'fileDeleted',
		description: 'A file is deleted or moved to the recycle bin',
		objectTypes: CONTENT,
		changeTypeFlags: ['DeleteObject'],
		changeTypes: [ChangeType.DeleteObject],
		kinds: ['file'],
	},
	{
		name: 'Folder Created',
		value: 'folderCreated',
		description: 'A folder is added',
		objectTypes: CONTENT,
		changeTypeFlags: ['Add'],
		changeTypes: [ChangeType.Add],
		kinds: ['folder'],
	},
	{
		name: 'Folder Deleted',
		value: 'folderDeleted',
		description: 'A folder is deleted or moved to the recycle bin',
		objectTypes: CONTENT,
		changeTypeFlags: ['DeleteObject'],
		changeTypes: [ChangeType.DeleteObject],
		kinds: ['folder'],
	},
	{
		name: 'Item Created',
		value: 'itemCreated',
		description: 'A list item is added (not a file or folder)',
		objectTypes: ['Item'],
		changeTypeFlags: ['Add'],
		changeTypes: [ChangeType.Add],
		kinds: ['item'],
	},
	{
		name: 'Item Updated',
		value: 'itemUpdated',
		description: 'A list item changes (not a file or folder)',
		objectTypes: ['Item'],
		changeTypeFlags: ['Update'],
		changeTypes: [ChangeType.Update],
		kinds: ['item'],
	},
	{
		name: 'Item Deleted',
		value: 'itemDeleted',
		description: 'A list item is deleted (not a file or folder)',
		objectTypes: ['Item'],
		changeTypeFlags: ['DeleteObject'],
		changeTypes: [ChangeType.DeleteObject],
		kinds: ['item'],
	},
	{
		name: 'Moved',
		value: 'moved',
		description: 'A file, folder or item is moved',
		objectTypes: CONTENT,
		changeTypeFlags: ['Move'],
		changeTypes: [ChangeType.MoveAway, ChangeType.MoveInto],
		kinds: ['file', 'folder', 'item'],
	},
	{
		name: 'Permissions Changed',
		value: 'permissionsChanged',
		description: 'Role assignments are added or removed, or inheritance is broken or reset',
		objectTypes: [...CONTENT, 'List', 'Web', 'SecurityPolicy'],
		changeTypeFlags: ['RoleAssignmentAdd', 'RoleAssignmentDelete'],
		changeTypes: [
			ChangeType.AssignmentAdd,
			ChangeType.AssignmentDelete,
			ChangeType.ScopeAdd,
			ChangeType.ScopeDelete,
		],
		kinds: ['file', 'folder', 'item', 'other'],
	},
	{
		name: 'Renamed',
		value: 'renamed',
		description: 'A file, folder or item is renamed',
		objectTypes: CONTENT,
		changeTypeFlags: ['Rename'],
		changeTypes: [ChangeType.Rename],
		kinds: ['file', 'folder', 'item'],
	},
	{
		name: 'Restored',
		value: 'restored',
		description: 'A file, folder or item is restored from the recycle bin',
		objectTypes: CONTENT,
		changeTypeFlags: ['Restore'],
		changeTypes: [ChangeType.Restore],
		kinds: ['file', 'folder', 'item'],
	},
];

const EVENT_BY_VALUE = new Map(CHANGE_EVENTS.map((event) => [event.value, event]));

/** Returns the object and change type flags that a set of trigger events needs. */
export function getEventFlags(events: string[]): {
	objectTypes: ObjectTypeFlag[];
	changeTypes: ChangeTypeFlag[];
} {
	const objectTypes = new Set<ObjectTypeFlag>();
	const changeTypes = new Set<ChangeTypeFlag>();
	for (const value of events) {
		const event = EVENT_BY_VALUE.get(value);
		if (!event) continue;
		event.objectTypes.forEach((flag) => objectTypes.add(flag));
		event.changeTypeFlags.forEach((flag) => changeTypes.add(flag));
	}
	return { objectTypes: [...objectTypes], changeTypes: [...changeTypes] };
}

/** Returns the trigger events that a classified change matches. */
export function getMatchingEvents(
	changeType: number,
	kind: ChangeKind,
	events: string[],
): string[] {
	return events.filter((value) => {
		const event = EVENT_BY_VALUE.get(value);
		return !!event && event.changeTypes.includes(changeType) && event.kinds.includes(kind);
	});
}

/** SP.RecycleBinItemType values that this node maps to a kind. */
const RECYCLE_BIN_KIND: Record<number, ChangeKind> = {
	1: 'file',
	3: 'item',
	5: 'folder',
	6: 'folder',
};

const DOCUMENT_LIBRARY_BASE_TYPE = 1;

export interface ItemDetails extends IDataObject {
	exists: boolean;
	kind?: ChangeKind;
	name?: string;
	serverRelativeUrl?: string;
	isFolder?: boolean;
	isFile?: boolean;
	deleted?: boolean;
	recycleBinItemId?: string;
	deletedDate?: string;
	fields?: IDataObject;
}

export interface RecycleBinEntry {
	Id: string;
	UniqueId?: string;
	DeletedDate: string;
	DirName: string;
	LeafName: string;
	ItemType: number;
}

export interface RecycleBinSource {
	/** Finds an entry by the UniqueId of the deleted item. Large recycle bins can reject this query. */
	findByUniqueId(uniqueId: string): Promise<RecycleBinEntry | undefined>;
	/** Reads one page, newest first. */
	getPage(after?: RecycleBinEntry): Promise<RecycleBinEntry[]>;
}

const RECYCLE_BIN_SELECT = ['Id', 'UniqueId', 'DeletedDate', 'DirName', 'LeafName', 'ItemType'];
const RECYCLE_BIN_PAGE_SIZE = 200;
const RECYCLE_BIN_MAX_PAGES = 25;
const RECYCLE_BIN_MARGIN_MS = 5 * 60 * 1000;

function isAccessDenied(error: unknown): boolean {
	const status = (error as { status?: number })?.status;
	return status === 401 || status === 403;
}

/** Use with Promise.catch: returns undefined for a 404 and rethrows other errors. */
function undefinedIfNotFound(error: unknown): undefined {
	if (isNotFoundError(error)) return undefined;
	throw error;
}

/**
 * Reads the recycle bin of the site collection (all sites, both stages). If the user cannot read
 * it, it reads the recycle bin of the site.
 */
export function createRecycleBinSource(sp: SPFI): RecycleBinSource {
	let useWeb = false;
	const withFallback = async <T>(run: (useWebBin: boolean) => Promise<T>): Promise<T> => {
		if (!useWeb) {
			const result = await run(false).then(
				(value) => ({ value }),
				(error: unknown) => {
					if (!isAccessDenied(error)) throw error;
					useWeb = true;
					return undefined;
				},
			);
			if (result) return result.value;
		}
		return await run(true);
	};
	return {
		findByUniqueId: async (uniqueId) =>
			await withFallback(async (useWebBin) => {
				const bin = useWebBin ? sp.web.recycleBin : sp.site.recycleBin;
				const entries = await bin
					.filter(`UniqueId eq guid'${uniqueId}'`)
					.select(...RECYCLE_BIN_SELECT)();
				return entries[0] as unknown as RecycleBinEntry | undefined;
			}),
		getPage: async (after) =>
			await withFallback(async (useWebBin) => {
				const paging = after
					? `pagingInfo='${encodeURIComponent(
							`id=${encodeURIComponent(after.Id)}&title=${encodeURIComponent(after.LeafName)}&searchValue=${encodeURIComponent(after.DeletedDate)}`,
						).replace(/'/g, "''")}',`
					: '';
				const query = SPQueryable(
					useWebBin ? sp.web : sp.site,
					`getrecyclebinitems(${paging}rowLimit='${RECYCLE_BIN_PAGE_SIZE}',isAscending=false,itemState=0,orderBy=3)`,
				).select(...RECYCLE_BIN_SELECT);
				return (await spGet(query)) as RecycleBinEntry[];
			}),
	};
}

/**
 * Finds deleted items in the recycle bin by their UniqueId. It first asks for the UniqueId. When the
 * recycle bin is too large for that query, it reads the recycle bin newest first, until it passes
 * the time of the change. It keeps what it read, so later lookups in the same run are fast.
 */
export class RecycleBinLookup {
	private readonly entries = new Map<string, RecycleBinEntry>();
	private exactQuery = true;
	private last: RecycleBinEntry | undefined;
	private oldest = Infinity;
	private pages = 0;
	private done = false;

	constructor(private readonly source: RecycleBinSource) {}

	async find(uniqueId: string, changeTime?: string): Promise<RecycleBinEntry | undefined> {
		const key = uniqueId.toLowerCase();
		const known = this.entries.get(key);
		if (known) return known;

		if (this.exactQuery) {
			const result = await this.source.findByUniqueId(uniqueId).then(
				(entry) => ({ entry }),
				(error: unknown) => {
					if (isAccessDenied(error)) throw error;
					this.exactQuery = false;
					return undefined;
				},
			);
			if (result) {
				if (result.entry) this.entries.set(key, result.entry);
				return result.entry;
			}
		}

		const time = changeTime ? Date.parse(changeTime) : NaN;
		const stopBefore = Number.isFinite(time) ? time - RECYCLE_BIN_MARGIN_MS : -Infinity;
		while (!this.done && this.pages < RECYCLE_BIN_MAX_PAGES && this.oldest >= stopBefore) {
			const page = await this.source.getPage(this.last);
			this.pages++;
			if (!page.length) {
				this.done = true;
				break;
			}
			for (const entry of page) {
				if (entry.UniqueId) this.entries.set(entry.UniqueId.toLowerCase(), entry);
				const deleted = Date.parse(entry.DeletedDate);
				if (Number.isFinite(deleted)) this.oldest = Math.min(this.oldest, deleted);
			}
			this.last = page[page.length - 1];
			if (page.length < RECYCLE_BIN_PAGE_SIZE) this.done = true;
			if (this.entries.has(key)) break;
		}
		return this.entries.get(key);
	}
}

interface ListInfo {
	baseType: number;
}

function getFieldValue(item: IDataObject, path: string): unknown {
	let value: unknown = item;
	for (const part of path.split('/')) {
		if (value === null || value === undefined || typeof value !== 'object') return null;
		value = (value as IDataObject)[part];
	}
	return value ?? null;
}

/** Resolves item details for changes. It caches webs, lists and items, so each one is read once. */
export class ChangeResolver {
	private readonly webs = new Map<string, Promise<IWeb>>();
	private readonly lists = new Map<string, Promise<ListInfo | undefined>>();
	private readonly items = new Map<string, Promise<ItemDetails>>();
	private readonly recycleBin: RecycleBinLookup | undefined;
	private currentWebId: Promise<string> | undefined;

	constructor(
		private readonly sp: SPFI,
		private readonly options: {
			itemFields: string[];
			resolveDeletedItems: boolean;
			recycleBinSource?: RecycleBinSource;
		},
	) {
		if (options.resolveDeletedItems) {
			this.recycleBin = new RecycleBinLookup(
				options.recycleBinSource ?? createRecycleBinSource(sp),
			);
		}
	}

	private async getWeb(webId: unknown): Promise<IWeb> {
		const id = String(webId ?? '').toLowerCase();
		if (!id) return this.sp.web;
		this.currentWebId ??= this.sp.web
			.select('Id')()
			.then((web) => String(web.Id).toLowerCase());
		if (id === (await this.currentWebId)) return this.sp.web;
		let web = this.webs.get(id);
		if (!web) {
			web = this.sp.site.openWebById(id).then((result) => result.web);
			this.webs.set(id, web);
		}
		return await web;
	}

	async getListInfo(webId: unknown, listId: unknown): Promise<ListInfo | undefined> {
		const key = String(listId ?? '').toLowerCase();
		if (!key) return undefined;
		let info = this.lists.get(key);
		if (!info) {
			info = this.getWeb(webId)
				.then(async (web) => {
					const list = await web.lists.getById(key).select('BaseType')();
					return { baseType: Number(list.BaseType) };
				})
				.catch(undefinedIfNotFound);
			this.lists.set(key, info);
		}
		return await info;
	}

	async getDetails(change: IDataObject, objectType: string): Promise<ItemDetails | undefined> {
		const uniqueId = change.UniqueId as string | undefined;
		let key: string;
		let load: () => Promise<ItemDetails>;
		if (objectType === 'Item' && change.ListId && change.ItemId !== undefined) {
			key = `item:${String(change.ListId)}:${String(change.ItemId)}`;
			load = async () => await this.loadItem(change);
		} else if ((objectType === 'File' || objectType === 'Folder') && uniqueId) {
			key = `${objectType}:${uniqueId}`;
			load = async () => await this.loadFileOrFolder(change, objectType);
		} else {
			return undefined;
		}
		let details = this.items.get(key);
		if (!details) {
			details = load();
			this.items.set(key, details);
		}
		return await details;
	}

	private async loadItem(change: IDataObject): Promise<ItemDetails> {
		const web = await this.getWeb(change.WebId);
		const fields = this.options.itemFields;
		const expand = [...new Set(fields.filter((f) => f.includes('/')).map((f) => f.split('/')[0]))];
		let query = web.lists
			.getById(String(change.ListId))
			.items.getById(Number(change.ItemId))
			.select('Id', 'FileSystemObjectType', 'FileRef', 'FileLeafRef', 'Title', ...fields);
		if (expand.length) query = query.expand(...expand);
		const item = (await query().catch(undefinedIfNotFound)) as IDataObject | undefined;
		if (!item) return await this.loadDeleted(change);

		const listInfo = await this.getListInfo(change.WebId, change.ListId);
		const isFolder = Number(item.FileSystemObjectType) === 1;
		const isFile = !isFolder && listInfo?.baseType === DOCUMENT_LIBRARY_BASE_TYPE;
		const details: ItemDetails = {
			exists: true,
			kind: isFolder ? 'folder' : isFile ? 'file' : 'item',
			name: (item.FileLeafRef as string | undefined) ?? (item.Title as string | undefined),
			serverRelativeUrl: item.FileRef as string | undefined,
			isFolder,
			isFile,
		};
		if (fields.length) {
			details.fields = Object.fromEntries(
				fields.map((field) => [field, getFieldValue(item, field)]),
			) as IDataObject;
		}
		return details;
	}

	private async loadFileOrFolder(change: IDataObject, objectType: string): Promise<ItemDetails> {
		const web = await this.getWeb(change.WebId);
		const uniqueId = String(change.UniqueId);
		const isFolder = objectType === 'Folder';
		const target = isFolder ? web.getFolderById(uniqueId) : web.getFileById(uniqueId);
		const data = (await target.select('Name', 'ServerRelativeUrl')().catch(undefinedIfNotFound)) as
			| IDataObject
			| undefined;
		if (!data) return await this.loadDeleted(change);
		return {
			exists: true,
			kind: isFolder ? 'folder' : 'file',
			name: data.Name as string,
			serverRelativeUrl: data.ServerRelativeUrl as string,
			isFolder,
			isFile: !isFolder,
		};
	}

	/** Changes do not carry a name or path. The recycle bin keeps the UniqueId of a deleted item. */
	private async loadDeleted(change: IDataObject): Promise<ItemDetails> {
		const uniqueId = change.UniqueId as string | undefined;
		if (!this.recycleBin || !uniqueId) return { exists: false };
		const entry = await this.recycleBin.find(uniqueId, change.Time as string | undefined);
		if (!entry) return { exists: false };
		const kind = RECYCLE_BIN_KIND[Number(entry.ItemType)] ?? 'other';
		return {
			exists: false,
			deleted: true,
			kind,
			name: entry.LeafName,
			serverRelativeUrl: `/${String(entry.DirName ?? '').replace(/^\/+/, '')}/${entry.LeafName}`,
			isFolder: kind === 'folder',
			isFile: kind === 'file',
			recycleBinItemId: entry.Id,
			deletedDate: entry.DeletedDate,
		};
	}
}

/**
 * Returns true when a server-relative path is inside a folder (not the folder itself). Without
 * recursive, only the direct children of the folder match.
 */
export function isUnderFolder(path: string | undefined, folder: string, recursive = true): boolean {
	if (!path) return false;
	const root = trimTrailingSlash(folder).toLowerCase();
	const prefix = `${root === '/' ? '' : root}/`;
	const value = trimTrailingSlash(path).toLowerCase();
	if (!value.startsWith(prefix)) return false;
	return recursive || !value.slice(prefix.length).includes('/');
}

export interface ChangeTarget {
	scope: ChangeScope;
	/** Set when the target reads the change log of one folder and its subfolders. */
	folderPath?: string;
	getChanges(query: ChangeQuery): Promise<IDataObject[]>;
	getCurrentToken(): Promise<string>;
	getId(): Promise<string>;
}

export interface LocatorValue {
	mode: string;
	value: string;
}

async function getCurrentToken(target: {
	select: (...fields: string[]) => () => Promise<unknown>;
}): Promise<string> {
	const data = (await target.select('CurrentChangeToken')()) as {
		CurrentChangeToken: { StringValue: string };
	};
	return data.CurrentChangeToken.StringValue;
}

/**
 * Returns the change log of a site collection, site or list. With a list and a folder path, it reads
 * the change log of that folder and its subfolders (Folder.GetChanges with RecursiveAll). The change
 * tokens are the tokens of the list.
 */
export function getChangeTarget(
	sp: SPFI,
	scope: ChangeScope,
	list?: LocatorValue,
	folderPath?: string,
): ChangeTarget {
	if (scope === 'list') {
		if (!list || !String(list.value ?? '').trim()) throw new Error('Select a list');
		const target = getListFromLocator(sp, list);
		const getId = async () => String((await target.select('Id')()).Id);
		if (folderPath) {
			const folder = sp.web.getFolderByServerRelativePath(folderPath);
			return {
				scope,
				folderPath,
				getChanges: async (query) =>
					(await spPost(
						Folder(folder, 'getchanges'),
						body({ query: { ...query, RecursiveAll: query.RecursiveAll ?? true } }),
					)) as IDataObject[],
				getCurrentToken: async () => await getCurrentToken(target),
				getId,
			};
		}
		return {
			scope,
			getChanges: async (query) => (await target.getChanges(query)) as IDataObject[],
			getCurrentToken: async () => await getCurrentToken(target),
			getId,
		};
	}
	const target = scope === 'siteCollection' ? sp.site : sp.web;
	return {
		scope,
		getChanges: async (query) => (await target.getChanges(query)) as IDataObject[],
		getCurrentToken: async () => await getCurrentToken(target),
		getId: async () => String(((await target.select('Id')()) as { Id: string }).Id),
	};
}

export interface ChangeRunOptions extends ChangeOutputOptions {
	/** Keep reading with the last change token until SharePoint returns no more changes. */
	returnAll: boolean;
	/** The maximum number of output changes when returnAll is false. */
	limit?: number;
	/** The trigger events to keep. When empty or undefined, all changes are kept. */
	events?: string[];
	/** Stop reading more pages after this time (Date.now() value). */
	deadline?: number;
	recycleBinSource?: RecycleBinSource;
}

export interface ChangeRunResult {
	changes: IDataObject[];
	/** The token of the last change that was read, or undefined when no change was read. */
	nextChangeToken?: string;
	/** Whether SharePoint can have more changes after nextChangeToken. */
	hasMore: boolean;
}

const MAX_PAGES = 1000;

function stripMetadata(change: IDataObject): IDataObject {
	const result: IDataObject = {};
	for (const [key, value] of Object.entries(change)) {
		if (key.startsWith('odata.') || key === '__metadata') continue;
		result[key] = value;
	}
	return result;
}

/** Converts a raw SP.Change to the output shape. */
export function formatChange(change: IDataObject): IDataObject {
	return {
		changeToken: getChangeTokenString(change),
		changeType: getChangeTypeName(change.ChangeType),
		objectType: getObjectType(change),
		time: change.Time,
		...stripMetadata(change),
	};
}

export function classifyChange(
	objectType: string,
	details: ItemDetails | undefined,
	listBaseType?: number,
): ChangeKind {
	if (details?.kind) return details.kind;
	if (objectType === 'File') return 'file';
	if (objectType === 'Folder') return 'folder';
	if (objectType === 'Item') {
		return listBaseType === DOCUMENT_LIBRARY_BASE_TYPE ? 'file' : 'item';
	}
	return 'other';
}

function parseFields(value: string | undefined): string[] {
	return (value ?? '')
		.split(',')
		.map((field) => field.trim())
		.filter(Boolean);
}

/**
 * Reads changes page by page and applies the folder path and event filters. The next change token
 * is the token of the last change that was read, also when a filter skipped that change.
 */
export async function runChangeQuery(
	sp: SPFI,
	target: ChangeTarget,
	query: ChangeQuery,
	options: ChangeRunOptions,
): Promise<ChangeRunResult> {
	const folderFilter = options.folderPathFilter?.trim() || target.folderPath;
	const events = options.events?.length ? options.events : undefined;
	const needDetails = !!options.includeItemDetails || !!folderFilter || !!events;
	const resolver = needDetails
		? new ChangeResolver(sp, {
				itemFields: parseFields(options.itemFields),
				resolveDeletedItems: !!options.resolveDeletedItems,
				recycleBinSource: options.recycleBinSource,
			})
		: undefined;
	const limit = options.returnAll ? Infinity : Math.max(1, options.limit ?? 50);
	const changes: IDataObject[] = [];
	let nextChangeToken: string | undefined;
	let currentQuery = query;
	let hasMore = false;

	for (let page = 0; page < MAX_PAGES; page++) {
		const batch = await target.getChanges(currentQuery);
		hasMore = batch.length > 0;
		if (!batch.length) break;

		for (const change of batch) {
			const objectType = getObjectType(change);
			const details = resolver ? await resolver.getDetails(change, objectType) : undefined;

			let keep = true;
			if (folderFilter) {
				const path = details?.serverRelativeUrl;
				// The folder change log only has changes inside the folder, so an unknown path is kept.
				keep =
					path || !target.folderPath
						? isUnderFolder(path, folderFilter, query.RecursiveAll !== false)
						: true;
			}
			let matchedEvents: string[] | undefined;
			if (keep && events) {
				const listInfo =
					objectType === 'Item' && !details?.kind && resolver
						? await resolver.getListInfo(change.WebId, change.ListId)
						: undefined;
				const kind = classifyChange(objectType, details, listInfo?.baseType);
				matchedEvents = getMatchingEvents(Number(change.ChangeType), kind, events);
				keep = matchedEvents.length > 0;
			}

			nextChangeToken = getChangeTokenString(change) ?? nextChangeToken;
			if (!keep) continue;

			const output = formatChange(change);
			if (matchedEvents) output.events = matchedEvents;
			if (details && (options.includeItemDetails || folderFilter)) output.item = details;
			changes.push(output);
			if (changes.length >= limit) break;
		}

		if (changes.length >= limit) break;
		if (query.LatestFirst) break;
		if (options.deadline !== undefined && Date.now() >= options.deadline) break;
		const last = getChangeTokenString(batch[batch.length - 1]);
		if (!last) break;
		currentQuery = { ...query, ChangeTokenStart: { StringValue: last } };
	}

	return { changes, nextChangeToken, hasMore };
}

type DisplayOptions = INodeProperties['displayOptions'];

function withDisplay(property: INodeProperties, show: IDataObject | undefined): INodeProperties {
	if (!show) return property;
	const current = (property.displayOptions?.show ?? {}) as IDataObject;
	return {
		...property,
		displayOptions: { ...property.displayOptions, show: { ...show, ...current } } as DisplayOptions,
	};
}

export const changeScopeProperty: INodeProperties = {
	displayName: 'Scope',
	name: 'scope',
	type: 'options',
	options: [
		{
			name: 'Site Collection',
			value: 'siteCollection',
			description: 'Changes in the site collection of the selected site, including subsites',
		},
		{ name: 'Site', value: 'web', description: 'Changes in the selected site (web)' },
		{ name: 'List', value: 'list', description: 'Changes in one list or library' },
	],
	default: 'list',
	description: 'Where to read changes from',
};

const objectTypeOptions = [
	{ name: 'Alert', value: 'Alert' },
	{ name: 'Content Type', value: 'ContentType' },
	{ name: 'Field', value: 'Field' },
	{ name: 'File', value: 'File' },
	{ name: 'Folder', value: 'Folder' },
	{ name: 'Group', value: 'Group' },
	{
		name: 'Item',
		value: 'Item',
		description: 'List items, including files and folders in libraries',
	},
	{ name: 'List', value: 'List' },
	{ name: 'Navigation', value: 'Navigation' },
	{
		name: 'Security Policy',
		value: 'SecurityPolicy',
		description: 'Role assignments and permissions',
	},
	{ name: 'Site', value: 'Site', description: 'The site collection' },
	{ name: 'User', value: 'User' },
	{ name: 'View', value: 'View' },
	{ name: 'Web', value: 'Web', description: 'Sites and subsites' },
];

const changeTypeOptions = [
	{ name: 'Activity', value: 'Activity' },
	{ name: 'Add', value: 'Add' },
	{ name: 'Delete', value: 'DeleteObject' },
	{ name: 'Group Membership Add', value: 'GroupMembershipAdd' },
	{ name: 'Group Membership Delete', value: 'GroupMembershipDelete' },
	{ name: 'Move', value: 'Move' },
	{ name: 'Rename', value: 'Rename' },
	{ name: 'Restore', value: 'Restore' },
	{ name: 'Role Assignment Add', value: 'RoleAssignmentAdd' },
	{ name: 'Role Assignment Delete', value: 'RoleAssignmentDelete' },
	{ name: 'Role Definition Add', value: 'RoleDefinitionAdd' },
	{ name: 'Role Definition Delete', value: 'RoleDefinitionDelete' },
	{ name: 'Role Definition Update', value: 'RoleDefinitionUpdate' },
	{ name: 'System Update', value: 'SystemUpdate' },
	{ name: 'Update', value: 'Update' },
];

export const objectTypesProperty: INodeProperties = {
	displayName: 'Object Types',
	name: 'objectTypes',
	type: 'multiOptions',
	options: objectTypeOptions,
	default: ['Item', 'File', 'Folder'],
	description: 'The kinds of objects to include changes for',
};

export const changeTypesProperty: INodeProperties = {
	displayName: 'Change Types',
	name: 'changeTypes',
	type: 'multiOptions',
	options: changeTypeOptions,
	default: ['Add', 'Update', 'DeleteObject', 'Rename', 'Move', 'Restore'],
	description: 'The kinds of changes to include',
};

const fetchLimitOption: INodeProperties = {
	displayName: 'Fetch Limit',
	name: 'fetchLimit',
	type: 'number',
	typeOptions: { minValue: 1 },
	default: 1000,
	description: 'The maximum number of changes that SharePoint returns in one request',
};

export const queryOptionsProperty: INodeProperties = {
	displayName: 'Query Options',
	name: 'queryOptions',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	options: [
		{
			displayName: 'End Token',
			name: 'endPosition',
			type: 'string',
			default: '',
			description: 'Only return changes up to this change token',
		},
		fetchLimitOption,
		{
			displayName: 'Ignore Start Token Not Found Error',
			name: 'ignoreStartTokenNotFoundError',
			type: 'boolean',
			default: false,
			description:
				'Whether to ignore the error when the start token is not in the change log, or when there are no changes',
		},
		{
			displayName: 'Latest First',
			name: 'latestFirst',
			type: 'boolean',
			default: false,
			description:
				'Whether to return the most recent changes first. With this option, only one request is sent.',
		},
		{
			displayName: 'Recursive All',
			name: 'recursiveAll',
			type: 'boolean',
			default: true,
			description:
				'Whether the folder path filter includes changes in subfolders. If not set, subfolders are included. With the List scope, SharePoint applies it to the folder change log.',
		},
		{
			displayName: 'Require Security Trim',
			name: 'requireSecurityTrim',
			type: 'boolean',
			default: false,
			description: 'Whether to return only changes that the current user has permission to see',
		},
	],
};

const itemDetailOptions: INodeProperties[] = [
	{
		displayName: 'Folder Path Filter',
		name: 'folderPathFilter',
		type: 'string',
		default: '',
		placeholder: 'Shared Documents/Folder',
		description:
			'Only keep changes to files, folders and items inside this folder. With the List scope, SharePoint reads the change log of the folder. With other scopes, the node reads the item details of each change to find its path.',
	},
	{
		displayName: 'Include Item Details',
		name: 'includeItemDetails',
		type: 'boolean',
		default: false,
		description:
			'Whether to read the name, server-relative URL and type of each changed file, folder or item',
	},
	{
		displayName: 'Item Fields',
		name: 'itemFields',
		type: 'string',
		default: '',
		placeholder: 'Title,Modified,Editor/Title',
		description:
			'Comma-separated internal names of item fields to include in the item details. Only used for list items.',
	},
	{
		displayName: 'Resolve Deleted Items From Recycle Bin',
		name: 'resolveDeletedItems',
		type: 'boolean',
		default: true,
		description:
			'Whether to find the name and path of deleted items in the recycle bin. Delete changes do not include them.',
	},
];

export const changeOptionsProperty: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	options: [
		...itemDetailOptions,
		{
			displayName: 'Output',
			name: 'output',
			type: 'options',
			options: [
				{ name: 'One Item per Change', value: 'perChange' },
				{
					name: 'Single Item',
					value: 'single',
					description: 'One item with the "changes" array and the "nextChangeToken"',
				},
			],
			default: 'perChange',
		},
	],
};

export const triggerOptionsProperty: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	options: [
		fetchLimitOption,
		...itemDetailOptions.filter((option) => option.name !== 'folderPathFilter'),
		{
			displayName: 'Start From Date',
			name: 'startFromDate',
			type: 'dateTime',
			default: '',
			description:
				'On the first poll, emit the changes since this date. If not set, the first poll only saves the current change token.',
		},
	],
};

/** Applies "show" display options to a group of shared change properties. */
export function showFor(properties: INodeProperties[], show?: IDataObject): INodeProperties[] {
	return properties.map((property) => withDisplay(property, show));
}

export function resolveFolderFilter(
	siteUrl: string,
	value: string | undefined,
): string | undefined {
	const trimmed = value?.trim();
	return trimmed ? toServerRelativePath(siteUrl, trimmed) : undefined;
}

export function validateFlags(
	node: INode,
	objectTypes: string[],
	changeTypes: string[],
	itemIndex?: number,
): void {
	if (!objectTypes.length) {
		throw new NodeOperationError(node, 'Select at least one object type', { itemIndex });
	}
	if (!changeTypes.length) {
		throw new NodeOperationError(node, 'Select at least one change type', { itemIndex });
	}
}
