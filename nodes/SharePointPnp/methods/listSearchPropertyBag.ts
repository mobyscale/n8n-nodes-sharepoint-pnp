import type { ILoadOptionsFunctions, INodeListSearchResult } from 'n8n-workflow';

import { sortAndCap } from '../shared/paging';
import { getPropertyBagKeyOptions } from './siteOptions';

export async function searchPropertyBagKeys(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const keys = await getPropertyBagKeyOptions.call(this);
	const text = filter?.toLowerCase();
	return {
		results: sortAndCap(
			keys
				.filter((key) => !text || key.name.toLowerCase().includes(text))
				.map((key) => ({ name: key.name, value: String(key.value) })),
		),
	};
}
