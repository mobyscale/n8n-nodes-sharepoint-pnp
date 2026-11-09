import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';
import type { IList } from '@pnp/sp/lists/index.js';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { createSp } from '../transport';

const PAGE_SIZE = 50;
const RECENT_LIMIT = 300;

interface ItemRow {
	Id: number;
	Title?: string | null;
	FileLeafRef?: string | null;
}

function toResult(row: ItemRow): INodeListSearchItems {
	return {
		name: `${row.Title || row.FileLeafRef || '(no title)'} (ID ${row.Id})`,
		value: String(row.Id),
	};
}

function newestItems(list: IList, top: number, conditions: string[]) {
	const query = list.items.select('Id', 'Title', 'FileLeafRef').orderBy('Id', false).top(top);
	return (conditions.length ? query.filter(conditions.join(' and ')) : query)() as Promise<
		ItemRow[]
	>;
}

const SCAN_PREFIX = 'scan:';

/**
 * Lists above the list view threshold reject a text filter on columns without an index, and also a sort
 * on Modified or Created. Then the items are scanned by ID (always indexed, in creation order), newest
 * first. Each page scans the next RECENT_LIMIT older items, matches them here and sorts the matches by
 * name. The pagination token "scan:<last ID>" continues the scan below that ID.
 */
async function matchRecentItems(
	list: IList,
	text: string,
	scanBeforeId?: number,
): Promise<INodeListSearchResult> {
	const needle = text.toLowerCase();
	const rows = await newestItems(list, RECENT_LIMIT, scanBeforeId ? [`Id lt ${scanBeforeId}`] : []);
	const results = rows
		.filter((row) => `${row.Title ?? ''} ${row.FileLeafRef ?? ''}`.toLowerCase().includes(needle))
		.map(toResult)
		.sort((a, b) => a.name.localeCompare(b.name));
	const lastId = rows.length === RECENT_LIMIT ? rows[rows.length - 1].Id : 0;
	return { results, paginationToken: lastId > 1 ? `${SCAN_PREFIX}${lastId}` : undefined };
}

/**
 * Items of the selected list ("list" parameter), newest first. A number finds the item by ID.
 * Above the list view threshold, a text search scans older items page by page and sorts each page A-Z.
 */
export async function searchListItems(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const list = getListFromLocator(
		sp,
		this.getNodeParameter('list') as { mode: string; value: string },
	);
	const text = (filter ?? '').trim();
	const isId = /^\d+$/.test(text);
	if (paginationToken?.startsWith(SCAN_PREFIX)) {
		return await matchRecentItems(list, text, Number(paginationToken.slice(SCAN_PREFIX.length)));
	}
	// SharePoint ignores $skip for list items, so the next page starts below the last ID.
	const beforeId = Number(paginationToken ?? 0);

	const conditions: string[] = [];
	if (isId) {
		conditions.push(`Id eq ${text}`);
	} else if (text) {
		const escaped = text.replace(/'/g, "''");
		conditions.push(`(substringof('${escaped}', Title) or substringof('${escaped}', FileLeafRef))`);
	}
	if (beforeId) conditions.push(`Id lt ${beforeId}`);

	let rows: ItemRow[];
	try {
		rows = await newestItems(list, PAGE_SIZE + 1, conditions);
	} catch (error) {
		if (!text || isId) throw new NodeApiError(this.getNode(), error as JsonObject);
		return await matchRecentItems(list, text);
	}
	return {
		results: rows.slice(0, PAGE_SIZE).map(toResult),
		paginationToken: rows.length > PAGE_SIZE ? String(rows[PAGE_SIZE - 1].Id) : undefined,
	};
}
