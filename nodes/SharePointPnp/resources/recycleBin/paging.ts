import type { SPFI } from '@pnp/sp/index.js';
import { SPQueryable, spGet } from '@pnp/sp/spqueryable.js';

import type { RecycleBinRawItem } from './helpers';

export type RecycleBinScope = 'siteCollection' | 'site';

export const PAGE_SIZE = 200;

/** Sort order value of the recycle bin query for "DeletedDate". */
const ORDER_BY_DELETED_DATE = 3;

export interface RecycleBinPage {
	items: RecycleBinRawItem[];
	/** Paging info for the next page. Not set when this page is the last one. */
	nextPagingInfo?: string;
}

/**
 * Builds the paging info that makes SharePoint return the page after the given item.
 * SharePoint reads "searchValue" in the time zone of the site, so the UTC "Z" must stay off.
 */
export function buildPagingInfo(last: RecycleBinRawItem): string {
	const searchValue = last.DeletedDate.replace(/Z$/i, '');
	return `id=${last.Id}&title=${encodeURIComponent(last.Title ?? '')}&searchValue=${encodeURIComponent(searchValue)}`;
}

export function buildPageUrlPart(
	stage: number,
	pagingInfo: string | undefined,
	rowLimit: number,
): string {
	const paging = (pagingInfo ?? '').replace(/'/g, "''");
	return `GetRecycleBinItems(pagingInfo='${paging}',rowLimit=${rowLimit},isAscending=false,orderBy=${ORDER_BY_DELETED_DATE},itemState=${stage})`;
}

/**
 * Gets one page of site collection recycle bin items, newest first. The web endpoint of the same
 * name ignores paging info, so the Site scope filters these items by path instead.
 */
export async function fetchRecycleBinPage(
	sp: SPFI,
	stage: number,
	pagingInfo: string | undefined,
	rowLimit: number,
): Promise<RecycleBinPage> {
	const query = SPQueryable(sp.site, buildPageUrlPart(stage, pagingInfo, rowLimit));
	const response = await spGet<RecycleBinRawItem[]>(query);
	const items = response ?? [];
	const last = items[items.length - 1];
	return {
		items,
		nextPagingInfo: items.length >= rowLimit && last ? buildPagingInfo(last) : undefined,
	};
}
