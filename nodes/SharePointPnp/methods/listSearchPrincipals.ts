import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';

import { getSiteUrl } from '../shared/descriptions';
import { createSp } from '../transport';

function matches(value: string, filter?: string): boolean {
	return !filter || value.toLowerCase().includes(filter.toLowerCase());
}

export async function searchGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const groups = await sp.web.siteGroups.select('Id', 'Title')();
	const results: INodeListSearchItems[] = groups
		.filter((group) => matches(group.Title, filter))
		.sort((a, b) => a.Title.localeCompare(b.Title))
		.map((group) => ({ name: group.Title, value: String(group.Id) }));
	return { results };
}

export async function searchRoleDefinitions(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const roles = await sp.web.roleDefinitions.select('Id', 'Name', 'Hidden', 'Order')();
	const results: INodeListSearchItems[] = roles
		.filter((role) => !role.Hidden && matches(role.Name, filter))
		.sort((a, b) => a.Order - b.Order)
		.map((role) => ({ name: role.Name, value: String(role.Id) }));
	return { results };
}
