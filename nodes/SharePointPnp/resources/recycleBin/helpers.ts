import type { IDataObject } from 'n8n-workflow';

export const ITEM_TYPE_NAMES: Record<number, string> = {
	1: 'File',
	2: 'FileVersion',
	3: 'ListItem',
	4: 'List',
	5: 'Folder',
	6: 'FolderWithLists',
	7: 'Attachment',
	8: 'ListItemVersion',
	9: 'CascadeParent',
	10: 'Web',
};

export const ITEM_STATE_NAMES: Record<number, string> = {
	1: 'FirstStage',
	2: 'SecondStage',
};

export const ITEM_STATE_FIRST = 1;
export const ITEM_STATE_SECOND = 2;

const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export interface RecycleBinRawItem {
	Id: string;
	Title?: string;
	ItemType: number;
	ItemState: number;
	DirName?: string;
	LeafName?: string;
	DeletedDate: string;
	DeletedByName?: string;
	DeletedByEmail?: string;
	AuthorName?: string;
	AuthorEmail?: string;
	Size?: number | string;
}

export interface RecycleBinFilters {
	itemType?: number;
	stage?: number;
	deletedAfter?: string;
	deletedBefore?: string;
	deletedBy?: string;
	text?: string;
}

export function isGuid(value: string): boolean {
	return GUID_PATTERN.test(value);
}

/** Accepts a comma or line separated string, or an array, and returns the trimmed unique IDs. */
export function parseIds(value: unknown): string[] {
	const parts = Array.isArray(value) ? value : String(value ?? '').split(/[,\r\n]+/);
	const ids = parts.map((part) => String(part).trim()).filter((part) => part !== '');
	return [...new Set(ids)];
}

export function findInvalidIds(ids: string[]): string[] {
	return ids.filter((id) => !isGuid(id));
}

function includesText(value: string | undefined, search: string): boolean {
	return (value ?? '').toLowerCase().includes(search.toLowerCase());
}

export function mapItem(item: RecycleBinRawItem): IDataObject {
	return {
		id: item.Id,
		title: item.Title ?? '',
		itemType: ITEM_TYPE_NAMES[item.ItemType] ?? String(item.ItemType),
		itemTypeId: item.ItemType,
		itemState: ITEM_STATE_NAMES[item.ItemState] ?? String(item.ItemState),
		itemStateId: item.ItemState,
		dirName: item.DirName ?? '',
		leafName: item.LeafName ?? '',
		deletedDate: item.DeletedDate,
		deletedByName: item.DeletedByName ?? '',
		deletedByEmail: item.DeletedByEmail ?? '',
		size: Number(item.Size ?? 0),
		authorName: item.AuthorName ?? '',
		authorEmail: item.AuthorEmail ?? '',
	};
}

/** Returns true when the item passes every filter that is set. */
export function matchesFilters(item: RecycleBinRawItem, filters: RecycleBinFilters): boolean {
	if (filters.itemType !== undefined && filters.itemType !== -1) {
		if (item.ItemType !== filters.itemType) return false;
	}
	if (filters.stage !== undefined && filters.stage !== -1) {
		if (item.ItemState !== filters.stage) return false;
	}
	const deleted = Date.parse(item.DeletedDate);
	if (filters.deletedAfter && !(deleted > Date.parse(filters.deletedAfter))) return false;
	if (filters.deletedBefore && !(deleted < Date.parse(filters.deletedBefore))) return false;
	const deletedBy = filters.deletedBy?.trim();
	if (deletedBy && !includesText(item.DeletedByEmail, deletedBy)) return false;
	const text = filters.text?.trim();
	if (text) {
		const inTitle = includesText(item.Title, text);
		const inLeaf = includesText(item.LeafName, text);
		const inDir = includesText(item.DirName, text);
		if (!inTitle && !inLeaf && !inDir) return false;
	}
	return true;
}

/**
 * Items come in descending DeletedDate order. When the last item of a page is older than the
 * "deleted after" filter, later pages cannot contain a match.
 */
export function isPastDeletedAfter(item: RecycleBinRawItem, deletedAfter?: string): boolean {
	return Boolean(deletedAfter) && Date.parse(item.DeletedDate) <= Date.parse(deletedAfter!);
}

/** True when the original location of the item is the given site or below it. */
export function isInSite(item: RecycleBinRawItem, siteServerRelativeUrl: string): boolean {
	const prefix = siteServerRelativeUrl.replace(/^\/+|\/+$/g, '').toLowerCase();
	if (!prefix) return true;
	const dir = (item.DirName ?? '').replace(/^\/+/, '').toLowerCase();
	return dir === prefix || dir.startsWith(`${prefix}/`);
}
