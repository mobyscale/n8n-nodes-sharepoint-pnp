import type { IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { IList } from '@pnp/sp/lists/index.js';
import type { SPFI } from '@pnp/sp/index.js';

import type { SharePointContext } from '../transport';
import { getTenantUrl } from '../transport';
import { resolveSiteUrl, toServerRelativePath } from './utils';

const GUID_REGEX = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$';

export const siteRLC: INodeProperties = {
	displayName: 'Site',
	name: 'site',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'The SharePoint site (web) to work in',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchSites', searchable: true },
		},
		{
			displayName: 'By URL',
			name: 'url',
			type: 'string',
			placeholder: 'https://contoso.sharepoint.com/sites/Example',
			hint: 'An absolute site URL, or a server-relative path such as /sites/Example',
		},
	],
};

function listLocator(
	displayName: string,
	name: string,
	description: string,
	searchListMethod: string,
): INodeProperties {
	return {
		displayName,
		name,
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod, searchable: true },
			},
			{
				displayName: 'By Title',
				name: 'title',
				type: 'string',
				placeholder: 'Documents',
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: '00000000-0000-0000-0000-000000000000',
				validation: [
					{ type: 'regex', properties: { regex: GUID_REGEX, errorMessage: 'Not a valid list ID' } },
				],
			},
		],
	};
}

export const listRLC = listLocator(
	'List',
	'list',
	'A list or document library in the site',
	'searchLists',
);

export const libraryRLC = listLocator(
	'Library',
	'library',
	'A document library in the site',
	'searchLibraries',
);

/**
 * Builds a folder or file path picker. In "From List" mode the search text works as a path:
 * "Documents/Clients/Ac" lists the children of "Documents/Clients" that start with "Ac".
 */
export function pathRLC(
	overrides: Partial<INodeProperties> & Pick<INodeProperties, 'displayName' | 'name'>,
	searchListMethod = 'searchFolders',
): INodeProperties {
	return {
		type: 'resourceLocator',
		default: { mode: 'path', value: '' },
		required: true,
		description:
			'A server-relative path (/sites/Example/Shared Documents/Folder), a site-relative path (Shared Documents/Folder), or an absolute URL',
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod, searchable: true },
			},
			{
				displayName: 'By Path',
				name: 'path',
				type: 'string',
				placeholder: 'Shared Documents/Folder',
			},
		],
		...overrides,
	};
}

export async function getSiteUrl(ctx: SharePointContext, itemIndex?: number): Promise<string> {
	const value =
		itemIndex === undefined
			? ((ctx as Exclude<SharePointContext, IExecuteFunctions>).getNodeParameter('site', '', {
					extractValue: true,
				}) as string)
			: ((ctx as IExecuteFunctions).getNodeParameter('site', itemIndex, '', {
					extractValue: true,
				}) as string);
	return resolveSiteUrl(await getTenantUrl(ctx), value);
}

interface LocatorValue {
	mode: string;
	value: string;
}

export function getListFromLocator(sp: SPFI, locator: LocatorValue): IList {
	const value = String(locator.value ?? '').trim();
	if (locator.mode === 'title') return sp.web.lists.getByTitle(value);
	return sp.web.lists.getById(value);
}

export function getList(
	this: IExecuteFunctions,
	sp: SPFI,
	itemIndex: number,
	parameterName = 'list',
): IList {
	const locator = this.getNodeParameter(parameterName, itemIndex) as LocatorValue;
	return getListFromLocator(sp, locator);
}

export function getPath(
	this: IExecuteFunctions,
	siteUrl: string,
	itemIndex: number,
	parameterName: string,
): string {
	const value = this.getNodeParameter(parameterName, itemIndex, '', {
		extractValue: true,
	}) as string;
	return toServerRelativePath(siteUrl, value);
}
