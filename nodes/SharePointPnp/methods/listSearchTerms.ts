import type { ILoadOptionsFunctions, INodeListSearchResult } from 'n8n-workflow';

import { isGuid, normalizeGuid } from '../resources/taxonomy/csom';
import { getTerms, matchesLabel } from '../resources/taxonomy/service';
import { getSiteUrl } from '../shared/descriptions';
import { pageByOffset, sortByName } from '../shared/paging';
import { locatorValue } from '../shared/pickers';

const OPERATIONS_WITH_HELPER = ['getTerm', 'updateTerm', 'deleteTerm'];

/**
 * Terms of one term set, found by label. The term set comes from "Term Set". The operations that
 * work on a single term have an optional "Term Set (to Pick From)" instead. The service returns all
 * terms, so the list is sorted A-Z and paged by offset here.
 */
export async function searchTerms(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const operation = String(this.getNodeParameter('operation', '') ?? '');
	const parameter = OPERATIONS_WITH_HELPER.includes(operation) ? 'termSetHelper' : 'termSet';
	const termSetId = locatorValue(this.getNodeParameter(parameter, ''));
	if (!isGuid(termSetId)) return { results: [] };

	const terms = await getTerms(
		this,
		{ siteUrl: await getSiteUrl(this) },
		normalizeGuid(termSetId),
		{
			recursive: true,
		},
	);
	const text = filter?.trim();
	const entries = terms
		.filter((term) => !term.isDeprecated)
		.filter((term) => !text || matchesLabel(term, text, 'contains'))
		.map((term) => ({ name: term.path.replace(/;/g, ' > ') || term.name, value: term.id }));
	return pageByOffset(sortByName(entries), paginationToken);
}
