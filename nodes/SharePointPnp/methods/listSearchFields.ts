import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
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
	return fields
		.sort((a, b) => a.Title.localeCompare(b.Title))
		.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		}));
}

/** Options for the "view to copy" picker: views of the source list. */
export async function getSourceViewOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const views = (await getListFromLocator(sp, readLocator(this, 'sourceList'))
		.views.filter('Hidden eq false')
		.select('Id', 'Title')()) as Array<{ Id: string; Title: string }>;
	return views
		.sort((a, b) => a.Title.localeCompare(b.Title))
		.map((view) => ({ name: view.Title, value: view.Id }));
}
