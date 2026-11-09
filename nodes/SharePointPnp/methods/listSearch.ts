import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';

import { getSiteUrl } from '../shared/descriptions';
import { pageByOffset, sortAndCap, sortByName } from '../shared/paging';
import { getSiteServerRelativeUrl, joinPath, splitPath } from '../shared/utils';
import { createSp, getTenantUrl } from '../transport';

const PAGE_SIZE = 50;
const DOCUMENT_LIBRARY_TEMPLATE = 101;

function matches(value: string, filter?: string): boolean {
	return !filter || value.toLowerCase().includes(filter.toLowerCase());
}

export async function searchSites(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const tenantUrl = await getTenantUrl(this);
	const sp = createSp(this, tenantUrl);
	const startRow = Number(paginationToken ?? 0);
	const text = filter?.trim() ? `"${filter.trim().replace(/"/g, '')}*" ` : '';

	const response = await sp.search({
		Querytext: `${text}(contentclass:STS_Site OR contentclass:STS_Web)`,
		SelectProperties: ['Title', 'Path'],
		RowLimit: PAGE_SIZE,
		StartRow: startRow,
		TrimDuplicates: false,
	});

	const results: INodeListSearchItems[] = response.PrimarySearchResults.map((row) => ({
		name: String(row.Title ?? row.Path),
		value: String(row.Path),
		url: String(row.Path),
	}));

	const nextStart = startRow + response.RowCount;
	return {
		results,
		paginationToken: nextStart < response.TotalRows ? String(nextStart) : undefined,
	};
}

async function searchListsByTemplate(
	ctx: ILoadOptionsFunctions,
	filter: string | undefined,
	baseTemplate?: number,
): Promise<INodeListSearchResult> {
	const siteUrl = await getSiteUrl(ctx);
	const sp = createSp(ctx, siteUrl);
	const odataFilter =
		baseTemplate === undefined
			? 'Hidden eq false'
			: `Hidden eq false and BaseTemplate eq ${baseTemplate}`;

	const lists = await sp.web.lists
		.select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
		.expand('RootFolder')
		.filter(odataFilter)();

	const tenantUrl = new URL(siteUrl).origin;
	return {
		results: sortAndCap(
			lists
				.filter((list) => matches(list.Title, filter))
				.map((list) => ({
					name: list.Title,
					value: list.Id,
					url: `${tenantUrl}${(list as { RootFolder?: { ServerRelativeUrl?: string } }).RootFolder?.ServerRelativeUrl ?? ''}`,
				})),
		),
	};
}

export async function searchLists(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	return await searchListsByTemplate(this, filter);
}

export async function searchLibraries(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	return await searchListsByTemplate(this, filter, DOCUMENT_LIBRARY_TEMPLATE);
}

/**
 * Treats the search text as a site-relative path. Without a "/", it lists the document libraries.
 * With a "/", it lists the subfolders of the parent path whose names start with the last segment.
 * Both lists are sorted A-Z and page by offset.
 */
export async function searchFolders(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const tenantUrl = new URL(siteUrl).origin;
	const siteRelativeRoot = getSiteServerRelativeUrl(siteUrl);
	const text = (filter ?? '').replace(/^\/+/, '');

	if (!text.includes('/')) {
		const libraries = await sp.web.lists
			.select('Title', 'RootFolder/ServerRelativeUrl', 'RootFolder/Name')
			.expand('RootFolder')
			.filter(`Hidden eq false and BaseTemplate eq ${DOCUMENT_LIBRARY_TEMPLATE}`)();
		const entries = libraries
			.map((library) => {
				const root = (library as { RootFolder: { ServerRelativeUrl: string; Name: string } })
					.RootFolder;
				return { title: library.Title, url: root.ServerRelativeUrl, name: root.Name };
			})
			.filter((library) => matches(library.name, text) || matches(library.title, text))
			.map((library) => ({
				name: `${library.name}/`,
				value: library.url,
				url: `${tenantUrl}${library.url}`,
			}));
		return pageByOffset(sortByName(entries), paginationToken);
	}

	const { parent, name } = splitPath(text);
	const parentPath = joinPath(siteRelativeRoot, parent);
	const folders = await sp.web
		.getFolderByServerRelativePath(parentPath)
		.folders.select('Name', 'ServerRelativeUrl')();

	const entries = folders
		.filter(
			(folder) =>
				folder.Name !== 'Forms' && folder.Name.toLowerCase().startsWith(name.toLowerCase()),
		)
		.sort((a, b) => a.Name.localeCompare(b.Name))
		.map((folder) => ({
			name: `${parent}/${folder.Name}/`,
			value: folder.ServerRelativeUrl,
			url: `${tenantUrl}${folder.ServerRelativeUrl}`,
		}));
	return pageByOffset(entries, paginationToken);
}
