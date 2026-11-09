import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';
import '@pnp/sp/attachments/index.js';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { parseItemId } from '../shared/pickers';
import { pageByOffset, sortAndCap } from '../shared/paging';
import { joinPath, splitPath } from '../shared/utils';
import { createSp } from '../transport';

function matches(value: string, filter?: string): boolean {
	return !filter || value.toLowerCase().includes(filter.toLowerCase());
}

/**
 * Folders of the selected list ("list" parameter). The search text is a path below the list root:
 * "Clients/Ac" lists the subfolders of "Clients" that start with "Ac". Empty text includes the root on the first page. The list pages by offset.
 */
export async function searchListFolders(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const tenantUrl = new URL(siteUrl).origin;
	const list = () =>
		getListFromLocator(sp, this.getNodeParameter('list') as { mode: string; value: string });
	const root = (await list().rootFolder.select('ServerRelativeUrl')()) as {
		ServerRelativeUrl: string;
	};
	const text = (filter ?? '').replace(/^\/+/, '');
	const { parent, name } = text.includes('/') ? splitPath(text) : { parent: '', name: text };

	const folders = (await sp.web
		.getFolderByServerRelativePath(joinPath(root.ServerRelativeUrl, parent))
		.folders.select('Name', 'ServerRelativeUrl')()) as Array<{
		Name: string;
		ServerRelativeUrl: string;
	}>;

	const results: INodeListSearchItems[] = folders
		.filter(
			(folder) =>
				folder.Name !== 'Forms' && folder.Name.toLowerCase().startsWith(name.toLowerCase()),
		)
		.sort((a, b) => a.Name.localeCompare(b.Name))
		.map((folder) => ({
			name: `${parent ? `${parent}/` : ''}${folder.Name}/`,
			value: folder.ServerRelativeUrl,
			url: `${tenantUrl}${folder.ServerRelativeUrl}`,
		}));
	const page = pageByOffset(results, paginationToken);
	if (!text && !paginationToken) {
		page.results.unshift({
			name: '(list root)',
			value: root.ServerRelativeUrl,
			url: `${tenantUrl}${root.ServerRelativeUrl}`,
		});
	}
	return page;
}

/** File names of the attachments of the selected item ("list" and "itemId" parameters). */
export async function searchItemAttachments(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const list = getListFromLocator(
		sp,
		this.getNodeParameter('list') as { mode: string; value: string },
	);
	const itemId = parseItemId(this.getNodeParameter('itemId'));
	const files = (await list.items
		.getById(itemId)
		.attachmentFiles.select('FileName', 'ServerRelativeUrl')()) as Array<{
		FileName: string;
		ServerRelativeUrl: string;
	}>;
	const tenantUrl = new URL(await getSiteUrl(this)).origin;
	return {
		results: sortAndCap(
			files
				.filter((file) => matches(file.FileName, filter))
				.map((file) => ({
					name: file.FileName,
					value: file.FileName,
					url: `${tenantUrl}${file.ServerRelativeUrl}`,
				})),
		),
	};
}

/** The distinct column groups of the site columns. */
export async function searchColumnGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const rows = (await sp.web.fields.select('Group')()) as Array<{ Group?: string | null }>;
	const groups = [...new Set(rows.map((row) => row.Group ?? '').filter(Boolean))];
	return {
		results: sortAndCap(
			groups
				.filter((group) => matches(group, filter))
				.map((group) => ({ name: group, value: group })),
		),
	};
}
