import type { ILoadOptionsFunctions, INodeListSearchResult } from 'n8n-workflow';

import { getSiteUrl } from '../shared/descriptions';
import { getTermGroups, getTermSets } from '../resources/taxonomy/service';

function matches(value: string, filter?: string): boolean {
	return !filter || value.toLowerCase().includes(filter.toLowerCase());
}

export async function searchTermGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const groups = await getTermGroups(this, { siteUrl: await getSiteUrl(this) });
	return {
		results: groups
			.filter((group) => matches(group.name, filter))
			.sort((a, b) => a.name.localeCompare(b.name))
			.map((group) => ({ name: group.name, value: group.id })),
	};
}

export async function searchTermSets(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sets = await getTermSets(this, { siteUrl: await getSiteUrl(this) });
	return {
		results: sets
			.map((set) => ({ name: `${set.groupName ?? ''} / ${set.name}`, value: set.id }))
			.filter((entry) => matches(entry.name, filter))
			.sort((a, b) => a.name.localeCompare(b.name)),
	};
}
