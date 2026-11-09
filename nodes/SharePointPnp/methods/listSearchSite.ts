import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import '@pnp/sp/hubsites/index.js';
import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { sortAndCap } from '../shared/paging';
import { createSp } from '../transport';

type FieldRow = { InternalName: string; Title: string };

function readList(ctx: ILoadOptionsFunctions): { mode: string; value: string } {
	return ctx.getNodeParameter('list') as { mode: string; value: string };
}

/** Fields of the list, or of the site when the "Target" is "Site". The value is the internal name. */
export async function getTargetFieldOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const target = this.getNodeParameter('target', 'list') as string;
	const fields = target === 'site' ? sp.web.fields : getListFromLocator(sp, readList(this)).fields;
	const rows = (await fields
		.filter('Hidden eq false')
		.select('InternalName', 'Title')()) as FieldRow[];
	return sortAndCap(
		rows.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		})),
	);
}

/** Site content types that can be added to a list. */
export async function getSiteContentTypeOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const types = (await sp.web.contentTypes
		.filter('Hidden eq false')
		.select('StringId', 'Name', 'Group')()) as Array<{
		StringId: string;
		Name: string;
		Group: string;
	}>;
	return sortAndCap(
		types.map((type) => ({ name: `${type.Name} (${type.Group})`, value: type.StringId })),
	);
}

/** Content types that are in the selected list. */
export async function getListContentTypeOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const types = (await getListFromLocator(sp, readList(this)).contentTypes.select(
		'StringId',
		'Name',
	)()) as Array<{ StringId: string; Name: string }>;
	return sortAndCap(types.map((type) => ({ name: type.Name, value: type.StringId })));
}

/** Hub sites of the tenant. The value is the ID of the hub site. */
export async function getHubSiteOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const hubs = (await sp.hubSites.select('SiteId', 'Title', 'SiteUrl')()) as Array<{
		SiteId: string;
		Title: string;
		SiteUrl: string;
	}>;
	return sortAndCap(
		hubs.map((hub) => ({ name: hub.Title, value: hub.SiteId, description: hub.SiteUrl })),
	);
}
