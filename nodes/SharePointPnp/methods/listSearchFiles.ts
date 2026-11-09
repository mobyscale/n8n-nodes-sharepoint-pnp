import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';

import { getSiteUrl } from '../shared/descriptions';
import {
	getSiteServerRelativeUrl,
	joinPath,
	splitPath,
	toServerRelativePath,
} from '../shared/utils';
import { pageByOffset } from '../shared/paging';
import { createSp } from '../transport';
import { searchFolders } from './listSearch';

interface Entry {
	Name: string;
	ServerRelativeUrl: string;
}

/** "Library/Folder/" lists the content of Folder. "Library/Folder/re" filters it by the name start "re". */
export function parseSearchPath(text: string): { parent: string; name: string } {
	if (text.endsWith('/')) return { parent: text.replace(/\/+$/, ''), name: '' };
	return splitPath(text);
}

/** Folders first, then files. Folder names end with "/", so the user can type the path further. */
export function toFileResults(
	parent: string,
	prefix: string,
	folders: Entry[],
	files: Entry[],
	tenantUrl: string,
): INodeListSearchItems[] {
	const wanted = (entry: Entry) => entry.Name.toLowerCase().startsWith(prefix.toLowerCase());
	const byName = (a: Entry, b: Entry) => a.Name.localeCompare(b.Name);
	return [
		...folders
			.filter((folder) => folder.Name !== 'Forms' && wanted(folder))
			.sort(byName)
			.map((folder) => ({
				name: `${parent}/${folder.Name}/`,
				value: folder.ServerRelativeUrl,
				url: `${tenantUrl}${folder.ServerRelativeUrl}`,
			})),
		...files
			.filter(wanted)
			.sort(byName)
			.map((file) => ({
				name: `${parent}/${file.Name}`,
				value: file.ServerRelativeUrl,
				url: `${tenantUrl}${file.ServerRelativeUrl}`,
			})),
	];
}

/**
 * Treats the search text as a site-relative path. Without a "/", it lists the document libraries.
 * With a "/", it lists the folders and files of the parent path whose names start with the last segment.
 * Folders come first, then files, each sorted A-Z. The list pages by offset.
 */
export async function searchFiles(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const text = (filter ?? '').replace(/^\/+/, '');
	if (!text.includes('/')) return await searchFolders.call(this, text, paginationToken);

	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const { parent, name } = parseSearchPath(text);
	const folder = sp.web.getFolderByServerRelativePath(
		joinPath(getSiteServerRelativeUrl(siteUrl), parent),
	);
	const [folders, files] = await Promise.all([
		folder.folders.select('Name', 'ServerRelativeUrl')(),
		folder.files.select('Name', 'ServerRelativeUrl')(),
	]);
	return pageByOffset(
		toFileResults(parent, name, folders, files, new URL(siteUrl).origin),
		paginationToken,
	);
}

export interface VersionEntry {
	VersionLabel: string;
	Created?: string;
	CheckInComment?: string;
}

function compareLabels(a: string, b: string): number {
	const [aMajor, aMinor] = a.split('.').map(Number);
	const [bMajor, bMinor] = b.split('.').map(Number);
	return bMajor - aMajor || bMinor - aMinor;
}

/** Newest version first. The text filters by the start of the version label. */
export function toVersionResults(
	versions: VersionEntry[],
	filter?: string,
): INodeListSearchItems[] {
	const text = (filter ?? '').trim();
	return versions
		.filter((version) => version.VersionLabel.startsWith(text))
		.sort((a, b) => compareLabels(a.VersionLabel, b.VersionLabel))
		.map((version) => {
			const details = [version.Created?.slice(0, 10), version.CheckInComment?.trim()].filter(
				Boolean,
			);
			return {
				name: details.length
					? `${version.VersionLabel} (${details.join(', ')})`
					: version.VersionLabel,
				value: version.VersionLabel,
			};
		});
}

/** The earlier versions of the file in the "filePath" parameter, newest first. The current version is not in the list. */
export async function searchFileVersions(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const value = String(this.getNodeParameter('filePath', '', { extractValue: true }) ?? '').trim();
	if (!value) return { results: [] };
	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const versions = (await sp.web
		.getFileByServerRelativePath(toServerRelativePath(siteUrl, value))
		.versions.select('VersionLabel', 'Created', 'CheckInComment')()) as VersionEntry[];
	return pageByOffset(toVersionResults(versions, filter), paginationToken);
}
