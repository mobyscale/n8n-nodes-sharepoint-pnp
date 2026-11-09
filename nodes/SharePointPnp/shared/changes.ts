import type { IDataObject, INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { body } from '@pnp/queryable/index.js';
import { Folder } from '@pnp/sp/folders/index.js';
import { List } from '@pnp/sp/lists/index.js';
import { spGet, SPQueryable, spPost, type SPFI } from '@pnp/sp/index.js';
import type { IChangeQuery } from '@pnp/sp/types.js';
import type { IWeb } from '@pnp/sp/webs/index.js';

import { runInBatches } from './batch';
import type { EditorSource, ItemKind } from './changeFilters';
import {
	createEditorSource,
	EditorDirectory,
	getEditorEmail,
	isEditorExcluded,
	isWithinDepth,
	matchesFieldFilters,
	parseEditorList,
	parseFieldFilters,
	parseItemKinds,
	parseMaxDepth,
} from './changeFilters';
import { getListFromLocator, pathRLC } from './descriptions';
import { isNotFoundError, toServerRelativePath, trimTrailingSlash } from './utils';
import { locatorValue } from './pickers';

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
	itemFields?: string | string[];
	resolveDeletedItems?: boolean;
	folderPathFilter?: string;
	/** Item Kind: "file", "folder" and "item". Changes of other objects are removed. */
	itemKinds?: string[];
	/** Max Depth: folder levels below the folder path filter or the scope root. */
	maxDepth?: number;
	/** Exclude Editors: a comma-separated list of display names, emails or logins. */
	excludeEditors?: string;
	/** Field Filters: the fixedCollection value ({ filter: [{ field, operator, value }] }). */
	fieldFilters?: unknown;
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
	/** Finds the entries of a few UniqueIds in one request. Large recycle bins can reject this query. */
	findByUniqueIds?(uniqueIds: string[]): Promise<RecycleBinEntry[]>;
	/** Reads one page, newest first. */
	getPage(after?: RecycleBinEntry): Promise<RecycleBinEntry[]>;
}

const RECYCLE_BIN_SELECT = ['Id', 'UniqueId', 'DeletedDate', 'DirName', 'LeafName', 'ItemType'];
const RECYCLE_BIN_PAGE_SIZE = 200;
const RECYCLE_BIN_MAX_PAGES = 25;
const RECYCLE_BIN_MARGIN_MS = 5 * 60 * 1000;
/** UniqueIds in one recycle bin filter. It keeps the URL short. */
const RECYCLE_BIN_IDS_PER_QUERY = 20;

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
	const findByUniqueIds = async (uniqueIds: string[]) =>
		await withFallback(async (useWebBin) => {
			const bin = useWebBin ? sp.web.recycleBin : sp.site.recycleBin;
			const filter = uniqueIds.map((id) => `UniqueId eq guid'${id}'`).join(' or ');
			return (await bin
				.filter(filter)
				.select(...RECYCLE_BIN_SELECT)()) as unknown as RecycleBinEntry[];
		});
	return {
		findByUniqueId: async (uniqueId) => (await findByUniqueIds([uniqueId]))[0],
		findByUniqueIds,
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

export interface RecycleBinRequest {
	uniqueId: string;
	changeTime?: string;
}

/**
 * Finds deleted items in the recycle bin by their UniqueId. It first asks for the UniqueIds, a few in
 * each request. When the recycle bin is too large for that query, it reads the recycle bin newest
 * first, until it passes the time of the oldest change. It keeps what it read, so later lookups in
 * the same run send no more requests.
 */
export class RecycleBinLookup {
	private readonly entries = new Map<string, RecycleBinEntry>();
	private readonly missing = new Set<string>();
	private exactQuery = true;
	private last: RecycleBinEntry | undefined;
	private oldest = Infinity;
	private pages = 0;
	private done = false;

	constructor(private readonly source: RecycleBinSource) {}

	async find(uniqueId: string, changeTime?: string): Promise<RecycleBinEntry | undefined> {
		await this.findMany([{ uniqueId, changeTime }]);
		return this.entries.get(uniqueId.toLowerCase());
	}

	/** Loads the entries of many deleted items. Use find to read them after. */
	async findMany(requests: RecycleBinRequest[]): Promise<void> {
		const pending = new Map<string, RecycleBinRequest>();
		for (const request of requests) {
			const key = request.uniqueId.toLowerCase();
			if (!this.entries.has(key) && !this.missing.has(key)) pending.set(key, request);
		}
		if (!pending.size) return;
		if (this.exactQuery && (await this.findExact([...pending.values()]))) return;

		const times = [...pending.values()].map((request) =>
			request.changeTime ? Date.parse(request.changeTime) : NaN,
		);
		const stopBefore = times.every(Number.isFinite)
			? Math.min(...times) - RECYCLE_BIN_MARGIN_MS
			: -Infinity;
		const keys = [...pending.keys()];
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
			if (keys.every((key) => this.entries.has(key))) break;
		}
	}

	/** Returns false when the recycle bin rejects the UniqueId query. */
	private async findExact(requests: RecycleBinRequest[]): Promise<boolean> {
		for (let start = 0; start < requests.length; start += RECYCLE_BIN_IDS_PER_QUERY) {
			const ids = requests.slice(start, start + RECYCLE_BIN_IDS_PER_QUERY).map((r) => r.uniqueId);
			const query = this.source.findByUniqueIds
				? this.source.findByUniqueIds(ids)
				: this.findOneByOne(ids);
			const result = await query.then(
				(found) => ({ found }),
				(error: unknown) => {
					if (isAccessDenied(error)) throw error;
					this.exactQuery = false;
					return undefined;
				},
			);
			if (!result) return false;
			for (const entry of result.found) {
				if (entry.UniqueId) this.entries.set(entry.UniqueId.toLowerCase(), entry);
			}
			for (const id of ids) {
				const key = id.toLowerCase();
				if (!this.entries.has(key)) this.missing.add(key);
			}
		}
		return true;
	}

	private async findOneByOne(ids: string[]): Promise<RecycleBinEntry[]> {
		const found: RecycleBinEntry[] = [];
		for (const id of ids) {
			const entry = await this.source.findByUniqueId(id);
			if (entry) found.push({ ...entry, UniqueId: entry.UniqueId ?? id });
		}
		return found;
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

/** The maximum number of values in a CAML In element. */
export const ITEM_IDS_PER_QUERY = 500;
/** The number of requests in one $batch request. */
export const REQUESTS_PER_BATCH = 100;

const BASE_ITEM_SELECT = ['Id', 'FileSystemObjectType', 'FileRef', 'FileLeafRef', 'Title'];

/** A CAML query for list items by ID, in all folders. ID is always indexed, so large lists allow it. */
export function buildItemIdQuery(ids: number[]): string {
	const values = ids.map((id) => `<Value Type="Counter">${id}</Value>`).join('');
	return `<View Scope="RecursiveAll"><Query><Where><In><FieldRef Name="ID"/><Values>${values}</Values></In></Where></Query><RowLimit>${ids.length}</RowLimit></View>`;
}

function chunk<T>(values: T[], size: number): T[][] {
	const chunks: T[][] = [];
	for (let start = 0; start < values.length; start += size) {
		chunks.push(values.slice(start, start + size));
	}
	return chunks;
}

export interface DetailRequest {
	change: IDataObject;
	objectType: string;
}

interface PendingDetail {
	key: string;
	change: IDataObject;
}

interface ItemGroup {
	webId: unknown;
	listId: string;
	entries: Array<PendingDetail & { id: number }>;
}

interface FileGroup {
	webId: unknown;
	entries: Array<PendingDetail & { isFolder: boolean }>;
}

/** Returns the cache key of the item of a change, or undefined when the change has no item. */
export function getDetailKey(change: IDataObject, objectType: string): string | undefined {
	if (objectType === 'Item' && change.ListId && change.ItemId !== undefined) {
		return `item:${String(change.ListId).toLowerCase()}:${String(change.ItemId)}`;
	}
	const uniqueId = change.UniqueId as string | undefined;
	if ((objectType === 'File' || objectType === 'Folder') && uniqueId) {
		return `${objectType}:${uniqueId.toLowerCase()}`;
	}
	return undefined;
}

/**
 * Resolves item details for changes. It caches webs, lists and items, so each one is read once.
 * `prefetch` loads the details of many changes in bulk: list items with one CAML query per list for
 * each 500 IDs, and files and folders without a list in $batch requests of 100.
 */
export class ChangeResolver {
	private readonly webs = new Map<string, Promise<IWeb>>();
	private readonly lists = new Map<string, Promise<ListInfo | undefined>>();
	private readonly items = new Map<string, Promise<ItemDetails>>();
	private readonly fieldValues = new WeakMap<ItemDetails, IDataObject>();
	private readonly recycleBin: RecycleBinLookup | undefined;
	private readonly selectFields: string[];
	private readonly expandFields: string[];
	private currentWebId: Promise<string> | undefined;

	constructor(
		private readonly sp: SPFI,
		private readonly options: {
			itemFields: string[];
			/** More fields to read, for example for field filters. They are not in the output. */
			extraFields?: string[];
			resolveDeletedItems: boolean;
			recycleBinSource?: RecycleBinSource;
		},
	) {
		if (options.resolveDeletedItems) {
			this.recycleBin = new RecycleBinLookup(
				options.recycleBinSource ?? createRecycleBinSource(sp),
			);
		}
		this.selectFields = [...new Set([...options.itemFields, ...(options.extraFields ?? [])])];
		this.expandFields = [
			...new Set(this.selectFields.filter((f) => f.includes('/')).map((f) => f.split('/')[0])),
		];
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

	/** The values of the item fields and the extra fields. Undefined when the item has no fields. */
	getFieldValues(details: ItemDetails | undefined): IDataObject | undefined {
		return details ? this.fieldValues.get(details) : undefined;
	}

	async getDetails(change: IDataObject, objectType: string): Promise<ItemDetails | undefined> {
		const key = getDetailKey(change, objectType);
		if (!key) return undefined;
		if (!this.items.has(key)) await this.prefetch([{ change, objectType }]);
		return await this.items.get(key);
	}

	/** Loads the details of many changes in bulk. Changes that are already cached are skipped. */
	async prefetch(requests: DetailRequest[]): Promise<void> {
		const itemGroups = new Map<string, ItemGroup>();
		const fileGroups = new Map<string, FileGroup>();
		const keys = new Set<string>();
		for (const { change, objectType } of requests) {
			const key = getDetailKey(change, objectType);
			if (!key || this.items.has(key) || keys.has(key)) continue;
			keys.add(key);
			const webKey = String(change.WebId ?? '').toLowerCase();
			if (objectType === 'Item') {
				const listId = String(change.ListId).toLowerCase();
				const groupKey = `${webKey}|${listId}`;
				let group = itemGroups.get(groupKey);
				if (!group) {
					group = { webId: change.WebId, listId, entries: [] };
					itemGroups.set(groupKey, group);
				}
				group.entries.push({ key, change, id: Number(change.ItemId) });
			} else {
				let group = fileGroups.get(webKey);
				if (!group) {
					group = { webId: change.WebId, entries: [] };
					fileGroups.set(webKey, group);
				}
				group.entries.push({ key, change, isFolder: objectType === 'Folder' });
			}
		}
		if (!keys.size) return;

		const run = this.loadAll([...itemGroups.values()], [...fileGroups.values()]);
		for (const key of keys) {
			const details = run.then((results) => results.get(key) ?? { exists: false });
			// The error is thrown by `await run` below. This only marks the copies as handled.
			details.catch(() => undefined);
			this.items.set(key, details);
		}
		await run;
	}

	private async loadAll(
		itemGroups: ItemGroup[],
		fileGroups: FileGroup[],
	): Promise<Map<string, ItemDetails>> {
		const results = new Map<string, ItemDetails>();
		const deleted: PendingDetail[] = [];
		for (const group of itemGroups) await this.loadItemGroup(group, results, deleted);
		for (const group of fileGroups) await this.loadFileGroup(group, results, deleted);
		if (deleted.length && this.recycleBin) {
			await this.recycleBin.findMany(
				deleted
					.filter(({ change }) => change.UniqueId)
					.map(({ change }) => ({
						uniqueId: String(change.UniqueId),
						changeTime: change.Time as string | undefined,
					})),
			);
		}
		for (const { key, change } of deleted) results.set(key, await this.loadDeleted(change));
		return results;
	}

	private async loadItemGroup(
		group: ItemGroup,
		results: Map<string, ItemDetails>,
		deleted: PendingDetail[],
	): Promise<void> {
		const web = await this.getWeb(group.webId);
		const ids = [...new Set(group.entries.map((entry) => entry.id))];
		const [rows, listInfo] = await Promise.all([
			this.expandFields.length
				? this.readItemsInBatches(web, group.listId, ids)
				: this.readItemsByQuery(web, group.listId, ids),
			this.getListInfo(group.webId, group.listId),
		]);
		for (const entry of group.entries) {
			const item = rows?.get(entry.id);
			if (item) results.set(entry.key, this.toItemDetails(item, listInfo));
			else deleted.push(entry);
		}
	}

	/** One CAML query for each 500 IDs. Returns undefined when the list does not exist. */
	private async readItemsByQuery(
		web: IWeb,
		listId: string,
		ids: number[],
	): Promise<Map<number, IDataObject> | undefined> {
		const rows = new Map<number, IDataObject>();
		const list = web.lists.getById(listId);
		for (const part of chunk(ids, ITEM_IDS_PER_QUERY)) {
			const query = List(list, 'getitems').select(...BASE_ITEM_SELECT, ...this.selectFields);
			const page = (await spPost(query, body({ query: { ViewXml: buildItemIdQuery(part) } })).catch(
				undefinedIfNotFound,
			)) as IDataObject[] | undefined;
			if (!page) return undefined;
			for (const row of page) rows.set(Number(row.Id ?? row.ID), row);
		}
		return rows;
	}

	/**
	 * The CAML query cannot expand lookup and person fields (for example Editor/Title). Then the items
	 * are read by ID in $batch requests of 100.
	 */
	private async readItemsInBatches(
		web: IWeb,
		listId: string,
		ids: number[],
	): Promise<Map<number, IDataObject>> {
		const results = await runInBatches(
			web,
			ids,
			async (batchedWeb, id) =>
				(await batchedWeb.lists
					.getById(listId)
					.items.getById(id)
					.select(...BASE_ITEM_SELECT, ...this.selectFields)
					.expand(...this.expandFields)()) as IDataObject,
			REQUESTS_PER_BATCH,
			REQUESTS_PER_BATCH,
		);
		const rows = new Map<number, IDataObject>();
		for (const result of results) {
			if (result.ok && result.value) rows.set(ids[result.index], result.value);
			else if (!isNotFoundError(result.error)) throw result.error;
		}
		return rows;
	}

	private toItemDetails(item: IDataObject, listInfo: ListInfo | undefined): ItemDetails {
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
		const fields = this.options.itemFields;
		if (fields.length) {
			details.fields = Object.fromEntries(
				fields.map((field) => [field, getFieldValue(item, field)]),
			) as IDataObject;
		}
		this.fieldValues.set(
			details,
			Object.fromEntries(
				this.selectFields.map((field) => [field, getFieldValue(item, field)]),
			) as IDataObject,
		);
		return details;
	}

	/** Files and folders by UniqueId, in $batch requests of 100. */
	private async loadFileGroup(
		group: FileGroup,
		results: Map<string, ItemDetails>,
		deleted: PendingDetail[],
	): Promise<void> {
		const web = await this.getWeb(group.webId);
		const reads = await runInBatches(
			web,
			group.entries,
			async (batchedWeb, { change, isFolder }) => {
				const uniqueId = String(change.UniqueId);
				const target = isFolder
					? batchedWeb.getFolderById(uniqueId)
					: batchedWeb.getFileById(uniqueId);
				return (await target.select('Name', 'ServerRelativeUrl')()) as IDataObject;
			},
			REQUESTS_PER_BATCH,
			REQUESTS_PER_BATCH,
		);
		for (const read of reads) {
			const entry = group.entries[read.index];
			if (read.ok && read.value) {
				results.set(entry.key, {
					exists: true,
					kind: entry.isFolder ? 'folder' : 'file',
					name: read.value.Name as string,
					serverRelativeUrl: read.value.ServerRelativeUrl as string,
					isFolder: entry.isFolder,
					isFile: !entry.isFolder,
				});
			} else if (isNotFoundError(read.error)) {
				deleted.push(entry);
			} else {
				throw read.error;
			}
		}
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
	/** The server-relative URL of the scope root: the folder, the list root folder, the site. */
	getRootPath?(): Promise<string>;
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
				getRootPath: async () => folderPath,
			};
		}
		return {
			scope,
			getChanges: async (query) => (await target.getChanges(query)) as IDataObject[],
			getCurrentToken: async () => await getCurrentToken(target),
			getId,
			getRootPath: async () =>
				String((await target.rootFolder.select('ServerRelativeUrl')()).ServerRelativeUrl),
		};
	}
	const target = scope === 'siteCollection' ? sp.site : sp.web;
	return {
		scope,
		getChanges: async (query) => (await target.getChanges(query)) as IDataObject[],
		getCurrentToken: async () => await getCurrentToken(target),
		getId: async () => String(((await target.select('Id')()) as { Id: string }).Id),
		getRootPath: async () =>
			String(
				((await target.select('ServerRelativeUrl')()) as { ServerRelativeUrl: string })
					.ServerRelativeUrl,
			),
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
	editorSource?: EditorSource;
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
	fileSystemObjectType?: unknown,
): ChangeKind {
	if (details?.kind) return details.kind;
	if (objectType === 'File') return 'file';
	if (objectType === 'Folder') return 'folder';
	if (objectType === 'Item') {
		if (fileSystemObjectType !== undefined && Number(fileSystemObjectType) === 1) return 'folder';
		return listBaseType === DOCUMENT_LIBRARY_BASE_TYPE ? 'file' : 'item';
	}
	return 'other';
}

function parseFields(value: string | string[] | undefined): string[] {
	const parts = Array.isArray(value) ? value : (value ?? '').split(',');
	return [...new Set(parts.map((field) => String(field).trim()).filter(Boolean))];
}

function hasValue(value: unknown): boolean {
	return value !== undefined && value !== null && value !== '';
}

function isContentChange(objectType: string): boolean {
	return objectType === 'Item' || objectType === 'File' || objectType === 'Folder';
}

/**
 * Reads changes page by page and applies the filters. The next change token is the token of the
 * last change that was read, also when a filter skipped that change.
 *
 * For each page it loads the item details of all changes that need them in bulk (see
 * ChangeResolver.prefetch), so a page of 1000 changes in one list needs about two requests.
 */
export async function runChangeQuery(
	sp: SPFI,
	target: ChangeTarget,
	query: ChangeQuery,
	options: ChangeRunOptions,
): Promise<ChangeRunResult> {
	const folderFilter = options.folderPathFilter?.trim() || target.folderPath;
	const events = options.events?.length ? options.events : undefined;
	const itemKinds = parseItemKinds(options.itemKinds);
	const maxDepth = parseMaxDepth(options.maxDepth);
	const excludedEditors = parseEditorList(options.excludeEditors);
	const fieldFilters = parseFieldFilters(options.fieldFilters);
	const itemFields = parseFields(options.itemFields);

	// These options need the item details of every change.
	const detailsForAll =
		!!options.includeItemDetails || !!folderFilter || !!events || fieldFilters.length > 0;
	// Item Kind and Max Depth only need the details when the change has no kind or path.
	const needsDetails = (change: IDataObject, objectType: string): boolean =>
		detailsForAll ||
		(itemKinds.length > 0 && objectType === 'Item' && !hasValue(change.FileSystemObjectType)) ||
		(maxDepth !== undefined && isContentChange(objectType) && !hasValue(change.ServerRelativeUrl));
	// Without filters that need details, only the changes up to the limit need details.
	const onlyLimitNeedsDetails =
		!folderFilter && !events && !fieldFilters.length && !itemKinds.length && maxDepth === undefined;

	const resolver =
		detailsForAll || itemKinds.length || maxDepth !== undefined
			? new ChangeResolver(sp, {
					itemFields,
					extraFields: fieldFilters.map((filter) => filter.field),
					resolveDeletedItems: !!options.resolveDeletedItems,
					recycleBinSource: options.recycleBinSource,
				})
			: undefined;
	const editors =
		excludedEditors.length || options.includeItemDetails
			? new EditorDirectory(options.editorSource ?? createEditorSource(sp))
			: undefined;
	let depthRoot: Promise<string | undefined> | undefined;
	const getDepthRoot = async () => {
		depthRoot ??= folderFilter
			? Promise.resolve(folderFilter)
			: (target.getRootPath?.() ?? Promise.resolve(undefined));
		return await depthRoot;
	};

	const limit = options.returnAll ? Infinity : Math.max(1, options.limit ?? 50);
	const changes: IDataObject[] = [];
	let nextChangeToken: string | undefined;
	let currentQuery = query;
	let hasMore = false;

	for (let page = 0; page < MAX_PAGES; page++) {
		const batch = await target.getChanges(currentQuery);
		hasMore = batch.length > 0;
		if (!batch.length) break;

		const excluded = new Set<IDataObject>();
		for (const change of batch) {
			if (await isEditorExcluded(change, excludedEditors, editors)) excluded.add(change);
		}
		if (resolver) {
			let requests = batch
				.filter((change) => !excluded.has(change))
				.map((change) => ({ change, objectType: getObjectType(change) }))
				.filter(({ change, objectType }) => needsDetails(change, objectType));
			if (onlyLimitNeedsDetails) requests = requests.slice(0, limit - changes.length);
			await resolver.prefetch(requests);
		}

		for (const change of batch) {
			nextChangeToken = getChangeTokenString(change) ?? nextChangeToken;
			if (excluded.has(change)) continue;

			const objectType = getObjectType(change);
			const details =
				resolver && needsDetails(change, objectType)
					? await resolver.getDetails(change, objectType)
					: undefined;

			if (folderFilter) {
				const path = details?.serverRelativeUrl;
				// The folder change log only has changes inside the folder, so an unknown path is kept.
				const inside =
					path || !target.folderPath
						? isUnderFolder(path, folderFilter, query.RecursiveAll !== false)
						: true;
				if (!inside) continue;
			}

			let kind: ChangeKind | undefined;
			if (events || itemKinds.length) {
				const listInfo =
					objectType === 'Item' && !details?.kind && resolver
						? await resolver.getListInfo(change.WebId, change.ListId)
						: undefined;
				kind = classifyChange(objectType, details, listInfo?.baseType, change.FileSystemObjectType);
			}
			if (itemKinds.length && !itemKinds.includes(kind as ItemKind)) continue;

			if (maxDepth !== undefined) {
				const path = details?.serverRelativeUrl ?? (change.ServerRelativeUrl as string | undefined);
				if (path && !isWithinDepth(path, await getDepthRoot(), maxDepth)) continue;
			}

			if (
				fieldFilters.length &&
				!matchesFieldFilters(resolver?.getFieldValues(details), fieldFilters)
			) {
				continue;
			}

			let matchedEvents: string[] | undefined;
			if (events) {
				matchedEvents = getMatchingEvents(Number(change.ChangeType), kind ?? 'other', events);
				if (!matchedEvents.length) continue;
			}

			const output = formatChange(change);
			if (matchedEvents) output.events = matchedEvents;
			if (details && (options.includeItemDetails || folderFilter)) output.item = details;
			if (options.includeItemDetails) {
				const editorEmail = await getEditorEmail(change, editors);
				if (editorEmail) output.editorEmail = editorEmail;
			}
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

/** What to do when the start token is older than the change log. */
export type ExpiredTokenMode = 'oldest' | 'now' | 'fail';

/** SharePoint Online keeps the change log for 60 days. */
export const CHANGE_LOG_RETENTION_MS = 60 * 24 * 60 * 60 * 1000;

/** Returns the time of a change token (Date.now() value), or undefined when the token has no time. */
export function getChangeTokenTime(token: string | undefined): number | undefined {
	const ticks = token?.split(';')[3];
	if (!ticks || !/^\d+$/.test(ticks)) return undefined;
	return Number((BigInt(ticks) - TICKS_AT_UNIX_EPOCH) / TICKS_PER_MILLISECOND);
}

export function isChangeTokenOutsideRetention(
	token: string | undefined,
	now = Date.now(),
): boolean {
	const time = getChangeTokenTime(token);
	return time !== undefined && time < now - CHANGE_LOG_RETENTION_MS;
}

const EXPIRED_TOKEN_ERROR_PATTERNS = [
	/before the start of the current change log/i,
	/(start|change) ?token[^.]*not (be )?found/i,
	/SPInvalidChangeToken/i,
];

/** Returns true for the SharePoint error about a start token that is not in the change log. */
export function isExpiredChangeTokenError(error: unknown): boolean {
	const message = String((error as { message?: unknown })?.message ?? '');
	return EXPIRED_TOKEN_ERROR_PATTERNS.some((pattern) => pattern.test(message));
}

export class ChangeTokenExpiredError extends Error {
	constructor() {
		super(
			'The change token has expired. SharePoint keeps changes for about 60 days. Set "On Expired Token" to continue from the oldest available change or from now.',
		);
		this.name = 'ChangeTokenExpiredError';
	}
}

export interface ExpiredTokenRunOptions extends ChangeRunOptions {
	onExpiredToken?: ExpiredTokenMode;
	/** The current time (Date.now() value). Only for tests. */
	now?: number;
}

export interface ExpiredTokenRunResult extends ChangeRunResult {
	/** True when the start token was older than the change log. */
	tokenExpired: boolean;
}

/**
 * Runs a change query and handles a start token that is older than the change log.
 *
 * Contract: when the token has expired, the changes between the token and the start of the change
 * log are lost. "oldest" returns all changes that are still in the log. They are all newer than the
 * token, so there are no duplicates and the gap is as small as possible. Each of these changes has
 * `tokenExpired: true`. "now" returns no changes and the current token. "fail" throws.
 *
 * SharePoint Online does not return an error for an old token. It returns the changes that are
 * still in the log, so the age of the token is also checked. Other SharePoint versions return a
 * "start token not found" error. Then the query runs again without a start token.
 */
export async function runChangeQueryWithExpiry(
	sp: SPFI,
	target: ChangeTarget,
	query: ChangeQuery,
	options: ExpiredTokenRunOptions,
): Promise<ExpiredTokenRunResult> {
	const mode = options.onExpiredToken ?? 'oldest';
	const startToken = query.ChangeTokenStart?.StringValue;
	if (!startToken)
		return { ...(await runChangeQuery(sp, target, query, options)), tokenExpired: false };

	// The current token is read before the query, so that no change is lost when nothing is read.
	const getFallbackToken = async (): Promise<string> => {
		if (mode === 'fail') throw new ChangeTokenExpiredError();
		return await target.getCurrentToken();
	};
	const markExpired = (result: ChangeRunResult, fallbackToken: string): ExpiredTokenRunResult => {
		for (const change of result.changes) change.tokenExpired = true;
		return {
			...result,
			nextChangeToken: result.nextChangeToken ?? fallbackToken,
			tokenExpired: true,
		};
	};
	const startFromNow = (currentToken: string): ExpiredTokenRunResult => ({
		changes: [],
		nextChangeToken: currentToken,
		hasMore: false,
		tokenExpired: true,
	});

	let fallbackToken: string | undefined;
	if (isChangeTokenOutsideRetention(startToken, options.now)) {
		fallbackToken = await getFallbackToken();
		if (mode === 'now') return startFromNow(fallbackToken);
	}

	const firstQuery = mode === 'oldest' ? { ...query, IgnoreStartTokenNotFoundError: true } : query;
	const result = await runChangeQuery(sp, target, firstQuery, options).catch((error: unknown) => {
		if (isExpiredChangeTokenError(error)) return undefined;
		throw error;
	});
	if (result) {
		return fallbackToken ? markExpired(result, fallbackToken) : { ...result, tokenExpired: false };
	}

	fallbackToken ??= await getFallbackToken();
	if (mode === 'now') return startFromNow(fallbackToken);
	const withoutStart: ChangeQuery = { ...query };
	delete withoutStart.ChangeTokenStart;
	return markExpired(await runChangeQuery(sp, target, withoutStart, options), fallbackToken);
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

export const onExpiredTokenOption: INodeProperties = {
	displayName: 'On Expired Token',
	name: 'onExpiredToken',
	type: 'options',
	options: [
		{
			name: 'Fail',
			value: 'fail',
			description: 'Stop with an error. Change this option to continue.',
		},
		{
			name: 'Start From Now',
			value: 'now',
			description: 'Save the current token and return no changes. The changes in the gap are lost.',
		},
		{
			name: 'Start From Oldest Available Change',
			value: 'oldest',
			description:
				'Return all changes that are still in the change log. They are all newer than the token, so there are no duplicates.',
		},
	],
	default: 'oldest',
	description:
		'What to do when the start token is older than the change log. SharePoint Online keeps changes for 60 days, so the changes between the token and the start of the change log are lost. Changes that are returned after an expired token have "tokenExpired": true.',
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

function sortByDisplayName(properties: INodeProperties[]): INodeProperties[] {
	return [...properties].sort((a, b) => a.displayName.localeCompare(b.displayName));
}

const itemDetailOptions: INodeProperties[] = [
	pathRLC({
		displayName: 'Folder Path Filter',
		name: 'folderPathFilter',
		required: false,
		default: { mode: 'path', value: '' },
		description:
			'Only keep changes to files, folders and items inside this folder. With the List scope, SharePoint reads the change log of the folder. With other scopes, the node reads the item details of each change to find its path.',
	}),
	{
		displayName: 'Include Item Details',
		name: 'includeItemDetails',
		type: 'boolean',
		default: false,
		description:
			'Whether to read the name, server-relative URL and type of each changed file, folder or item, and the editor email when it is known. The node reads the details in bulk: about one request for each 500 changed items in a list.',
	},
	{
		displayName: 'Item Field Names or IDs',
		name: 'itemFields',
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getChangeItemFieldOptions',
			loadOptionsDependsOn: ['list.value'],
		},
		default: [],
		description:
			'Fields to include in the item details. Only used for list items. The list shows the fields of the selected list, so it needs the List scope. An expression can return an array or a comma-separated list, for example "Title,Editor/Title". Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
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

const itemKindOption: INodeProperties = {
	displayName: 'Item Kind',
	name: 'changeItemKinds',
	type: 'multiOptions',
	options: [
		{ name: 'Files', value: 'file' },
		{ name: 'Folders', value: 'folder' },
		{ name: 'List Items', value: 'item', description: 'Items that are not files or folders' },
	],
	default: [],
	description:
		'Only keep changes to these kinds of objects. Changes to lists, sites and other objects are removed. The node uses the item details, or the FileSystemObjectType of the change when it has one.',
};

const changeFilterOptions: INodeProperties[] = [
	{
		displayName: 'Exclude Editors',
		name: 'excludeEditors',
		type: 'string',
		default: '',
		placeholder: 'e.g. SharePoint App, a@contoso.com',
		description:
			'A comma-separated list of editors to ignore. An entry matches the display name of the editor, the editor email hint of the change, or the email or login of the site user with that display name. Not case-sensitive.',
	},
	{
		displayName: 'Field Filters',
		name: 'changeFieldFilters',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Field Filter',
		default: {},
		description:
			'Only keep changes to list items whose field values match all filters. The node compares the values after it reads the item details, so this option reads the details of every change. Deleted items and changes without item fields do not match.',
		options: [
			{
				displayName: 'Filter',
				name: 'filter',
				values: [
					{
						displayName: 'Field Name or ID',
						name: 'field',
						type: 'options',
						typeOptions: {
							loadOptionsMethod: 'getChangeItemFieldOptions',
							loadOptionsDependsOn: ['list.value'],
						},
						default: '',
						description:
							'The internal name of the field. The list shows the fields of the selected list (List scope). For a lookup or person field, use an expression such as "Editor/Title". Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
					},
					{
						displayName: 'Operator',
						name: 'operator',
						type: 'options',
						options: [
							{ name: 'Contains', value: 'contains' },
							{ name: 'Equals', value: 'equals' },
							{ name: 'Is Empty', value: 'isEmpty' },
							{ name: 'Is Not Empty', value: 'isNotEmpty' },
							{ name: 'Not Equals', value: 'notEquals' },
						],
						default: 'equals',
						description: 'Text compares are not case-sensitive',
					},
					{
						displayName: 'Value',
						name: 'value',
						type: 'string',
						default: '',
						displayOptions: { hide: { operator: ['isEmpty', 'isNotEmpty'] } },
					},
				],
			},
		],
	},
	{
		displayName: 'Max Depth',
		name: 'changeMaxDepth',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 1,
		description:
			'Only keep changes up to this many folder levels below the folder path filter, or below the scope root (the library or list, or the site). 1 keeps only the direct children. Changes without a path are kept.',
	},
];

export const changeOptionsProperty: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	options: sortByDisplayName([
		...itemDetailOptions,
		...changeFilterOptions,
		itemKindOption,
		{
			...onExpiredTokenOption,
			displayOptions: { show: { '/startFrom': ['token', 'date'] } },
		},
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
	]),
};

export const triggerOptionsProperty: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	options: sortByDisplayName([
		fetchLimitOption,
		...itemDetailOptions.filter((option) => option.name !== 'folderPathFilter'),
		...changeFilterOptions,
		onExpiredTokenOption,
		{
			displayName: 'Start From Date',
			name: 'startFromDate',
			type: 'dateTime',
			default: '',
			description:
				'On the first poll, emit the changes since this date. If not set, the first poll only saves the current change token.',
		},
	]),
};

/** Applies "show" display options to a group of shared change properties. */
export function showFor(properties: INodeProperties[], show?: IDataObject): INodeProperties[] {
	return properties.map((property) => withDisplay(property, show));
}

export function resolveFolderFilter(siteUrl: string, value: unknown): string | undefined {
	const trimmed = locatorValue(value);
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
