import type { ILoadOptionsFunctions, INodeListSearchResult } from 'n8n-workflow';

import { getSiteUrl } from '../shared/descriptions';
import { sortAndCap } from '../shared/paging';
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
		results: sortAndCap(
			groups
				.filter((group) => matches(group.name, filter))
				.map((group) => ({ name: group.name, value: group.id })),
		),
	};
}

export async function searchTermSets(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sets = await getTermSets(this, { siteUrl: await getSiteUrl(this) });
	return {
		results: sortAndCap(
			sets
				.map((set) => ({ name: `${set.groupName ?? ''} / ${set.name}`, value: set.id }))
				.filter((entry) => matches(entry.name, filter)),
		),
	};
}
