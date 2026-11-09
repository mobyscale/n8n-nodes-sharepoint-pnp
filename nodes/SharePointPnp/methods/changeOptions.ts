import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { sortAndCap } from '../shared/paging';
import { createSp } from '../transport';
import { locatorValue } from '../shared/pickers';

/** Fields of the selected list. Without a list (other scopes), the list is empty. */
export async function getChangeItemFieldOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const list = this.getNodeParameter('list', { mode: 'list', value: '' }) as {
		mode: string;
		value: string;
	};
	if (!locatorValue(list)) return [];
	const sp = createSp(this, await getSiteUrl(this));
	const fields = (await getListFromLocator(sp, list)
		.fields.filter('Hidden eq false')
		.select('InternalName', 'Title')()) as Array<{ InternalName: string; Title: string }>;
	return sortAndCap(
		fields.map((field) => ({
			name: `${field.Title} (${field.InternalName})`,
			value: field.InternalName,
		})),
	);
}
