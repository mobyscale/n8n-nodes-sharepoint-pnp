import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import { NodeOperationError } from 'n8n-workflow';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { findOwningList, getListRoots } from '../shared/filesFolders';
import { sortAndCap } from '../shared/paging';
import { locatorValue } from '../shared/pickers';
import { toServerRelativePath } from '../shared/utils';
import { createSp } from '../transport';

function readLocator(ctx: ILoadOptionsFunctions, name: string): { mode: string; value: string } {
	return ctx.getNodeParameter(name) as { mode: string; value: string };
}

/** Options for the view field picker: visible fields of the selected list, by internal name. */
export async function getListFieldOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const fields = (await getListFromLocator(sp, readLocator(this, 'list'))
		.fields.filter('Hidden eq false')
		.select('InternalName', 'Title')()) as Array<{ InternalName: string; Title: string }>;
	return sortAndCap(
		fields.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		})),
	);
}

/**
 * Options for a field picker next to a folder ("folderPath" parameter): visible fields of the list or
 * library that contains the folder, by internal name. Empty while no folder is set.
 */
export async function getFolderListFieldOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const value = locatorValue(this.getNodeParameter('folderPath', ''));
	if (!value) return [];
	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const path = toServerRelativePath(siteUrl, value);
	const list = findOwningList(path, await getListRoots(sp));
	if (!list) {
		throw new NodeOperationError(this.getNode(), `No list or library contains "${path}"`);
	}
	const fields = (await sp.web.lists
		.getById(list.id)
		.fields.filter('Hidden eq false')
		.select('InternalName', 'Title')()) as Array<{ InternalName: string; Title: string }>;
	return sortAndCap(
		fields.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		})),
	);
}

/** Options for the lookup field picker: visible fields of the lookup list ("lookupList" parameter). */
export async function getLookupListFieldOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const fields = (await getListFromLocator(sp, readLocator(this, 'lookupList'))
		.fields.filter('Hidden eq false')
		.select('InternalName', 'Title')()) as Array<{ InternalName: string; Title: string }>;
	return sortAndCap(
		fields.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		})),
	);
}

/** Options for the "view to copy" picker: views of the source list. */
export async function getSourceViewOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const views = (await getListFromLocator(sp, readLocator(this, 'sourceList'))
		.views.filter('Hidden eq false')
		.select('Id', 'Title')()) as Array<{ Id: string; Title: string }>;
	return sortAndCap(views.map((view) => ({ name: view.Title, value: view.Id })));
}
