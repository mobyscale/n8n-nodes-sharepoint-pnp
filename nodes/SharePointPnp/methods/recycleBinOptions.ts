import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import {
	isInSite,
	ITEM_STATE_FIRST,
	ITEM_STATE_NAMES,
	ITEM_STATE_SECOND,
	toRecycleBinOptionName,
} from '../resources/recycleBin/helpers';
import type { RecycleBinRawItem } from '../resources/recycleBin/helpers';
import { fetchRecycleBinPage, type RecycleBinScope } from '../resources/recycleBin/paging';
import { getSiteUrl } from '../shared/descriptions';
import { MAX_OPTIONS, sortAndCap } from '../shared/paging';
import { getSiteServerRelativeUrl } from '../shared/utils';
import { createSp } from '../transport';

const PAGE_SIZE = 200;
const MAX_PAGES = 5;

/**
 * Reads up to MAX_OPTIONS recycle bin items of the selected scope (the first stage before the second
 * stage) and sorts them A-Z by name. The description has the stage and the deletion date.
 */
export async function getRecycleBinItemOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const scope = this.getNodeParameter('scope', 'siteCollection') as RecycleBinScope;
	const sitePath = scope === 'site' ? getSiteServerRelativeUrl(siteUrl) : undefined;

	const items: RecycleBinRawItem[] = [];
	for (const stage of [ITEM_STATE_FIRST, ITEM_STATE_SECOND]) {
		let pagingInfo: string | undefined;
		for (let pageNumber = 0; pageNumber < MAX_PAGES && items.length < MAX_OPTIONS; pageNumber++) {
			const page = await fetchRecycleBinPage(sp, stage, pagingInfo, PAGE_SIZE);
			for (const item of page.items) {
				if (!sitePath || isInSite(item, sitePath)) items.push(item);
			}
			if (!page.nextPagingInfo) break;
			pagingInfo = page.nextPagingInfo;
		}
	}
	return sortAndCap(
		items.slice(0, MAX_OPTIONS).map((item) => ({
			name: toRecycleBinOptionName(item),
			value: item.Id,
			description: `${ITEM_STATE_NAMES[item.ItemState] ?? ''} ${item.DeletedDate}`.trim(),
		})),
	);
}
