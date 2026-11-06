import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { ISearchQuery, ISearchResponse } from '@pnp/sp/search/index.js';

import type { OperationContext, ResourceModule } from '../types';
import {
	buildQueryText,
	clampRowLimit,
	DEFAULT_SELECT_PROPERTIES,
	flattenCells,
	nextPageSize,
	parseList,
	toSortList,
	type SearchRow,
	type SearchSort,
} from './helpers';

const show = (operation: string[]) => ({ show: { resource: ['search'], operation } });

interface SearchOptions {
	sort?: { values?: SearchSort[] };
	refinementFilters?: string;
	trimDuplicates?: boolean;
	sourceId?: string;
	queryTemplate?: string;
	enableQueryRules?: boolean;
	culture?: number;
	rowLimitPerPage?: number;
	includeRawResult?: boolean;
}

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['search'] } },
		options: [
			{
				name: 'Query',
				value: 'query',
				description: 'Run a KQL search query',
				action: 'Run a search query',
			},
			{
				name: 'Suggest',
				value: 'suggest',
				description: 'Get query suggestions for a text',
				action: 'Get search suggestions',
			},
		],
		default: 'query',
	},
	{
		displayName: 'Query',
		name: 'query',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'contentclass:STS_ListItem_DocumentLibrary',
		description: 'The KQL query text',
		displayOptions: show(['query']),
	},
	{
		displayName: 'Select Properties',
		name: 'selectProperties',
		type: 'string',
		default: DEFAULT_SELECT_PROPERTIES.join(','),
		description:
			'Managed properties to return, separated by commas. Each property becomes a field in the output.',
		displayOptions: show(['query']),
	},
	{
		displayName: 'Scope to Site',
		name: 'scopeToSite',
		type: 'boolean',
		default: false,
		description: 'Whether to return only results below the selected site URL',
		displayOptions: show(['query']),
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['query']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: { show: { resource: ['search'], operation: ['query'], returnAll: [false] } },
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['query']),
		options: [
			{
				displayName: 'Culture',
				name: 'culture',
				type: 'number',
				default: 1033,
				description: 'The locale ID (LCID) for the query, for example 1033 for English (US)',
			},
			{
				displayName: 'Enable Query Rules',
				name: 'enableQueryRules',
				type: 'boolean',
				default: false,
				description: 'Whether to apply query rules to the query',
			},
			{
				displayName: 'Include Raw Result',
				name: 'includeRawResult',
				type: 'boolean',
				default: false,
				description:
					'Whether to add the raw search response to the last output item as "rawResult". It contains the last page only.',
			},
			{
				displayName: 'Query Template',
				name: 'queryTemplate',
				type: 'string',
				default: '',
				placeholder: '{searchTerms} IsDocument:true',
				description: 'A query template. Use {searchTerms} for the query text.',
			},
			{
				displayName: 'Refinement Filters',
				name: 'refinementFilters',
				type: 'string',
				default: '',
				placeholder: 'FileType:equals("docx")',
				description: 'Refinement filters (FQL), separated by commas',
			},
			{
				displayName: 'Result Source ID',
				name: 'sourceId',
				type: 'string',
				default: '',
				description: 'The ID of the result source to use',
			},
			{
				displayName: 'Row Limit per Page',
				name: 'rowLimitPerPage',
				type: 'number',
				typeOptions: { minValue: 1, maxValue: 500 },
				default: 100,
				description: 'How many rows to get in each request. SharePoint allows up to 500.',
			},
			{
				displayName: 'Sort',
				name: 'sort',
				type: 'fixedCollection',
				typeOptions: { multipleValues: true },
				placeholder: 'Add Sort',
				default: {},
				options: [
					{
						displayName: 'Sort',
						name: 'values',
						values: [
							{
								displayName: 'Property',
								name: 'property',
								type: 'string',
								default: '',
								placeholder: 'LastModifiedTime',
								description: 'A sortable managed property',
							},
							{
								displayName: 'Direction',
								name: 'direction',
								type: 'options',
								options: [
									{ name: 'Ascending', value: 'ascending' },
									{ name: 'Descending', value: 'descending' },
								],
								default: 'descending',
							},
						],
					},
				],
			},
			{
				displayName: 'Trim Duplicates',
				name: 'trimDuplicates',
				type: 'boolean',
				default: false,
				description: 'Whether to remove duplicate results',
			},
		],
	},
	{
		displayName: 'Text',
		name: 'suggestText',
		type: 'string',
		required: true,
		default: '',
		description: 'The text to get suggestions for',
		displayOptions: show(['suggest']),
	},
	{
		displayName: 'Count',
		name: 'suggestCount',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 10,
		description: 'Max number of query suggestions',
		displayOptions: show(['suggest']),
	},
];

async function query(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const queryText = (this.getNodeParameter('query', itemIndex, '') as string).trim();
	if (!queryText) {
		throw new NodeOperationError(this.getNode(), 'The query must not be empty', { itemIndex });
	}
	const select = parseList(this.getNodeParameter('selectProperties', itemIndex, ''));
	const scopeToSite = this.getNodeParameter('scopeToSite', itemIndex, false) as boolean;
	const returnAll = this.getNodeParameter('returnAll', itemIndex, false) as boolean;
	const wanted = returnAll ? undefined : (this.getNodeParameter('limit', itemIndex, 50) as number);
	const options = this.getNodeParameter('options', itemIndex, {}) as SearchOptions;

	const pageSize = clampRowLimit(options.rowLimitPerPage);
	const request: ISearchQuery = {
		Querytext: buildQueryText(queryText, scopeToSite ? siteUrl : undefined),
		SelectProperties: select.length ? select : DEFAULT_SELECT_PROPERTIES,
		TrimDuplicates: options.trimDuplicates ?? false,
	};
	const sortList = toSortList(options.sort?.values ?? []);
	if (sortList.length) request.SortList = sortList as unknown as ISearchQuery['SortList'];
	const refinements = parseList(options.refinementFilters);
	if (refinements.length) request.RefinementFilters = refinements;
	if (options.sourceId?.trim()) request.SourceId = options.sourceId.trim();
	if (options.queryTemplate?.trim()) request.QueryTemplate = options.queryTemplate.trim();
	if (options.enableQueryRules !== undefined) request.EnableQueryRules = options.enableQueryRules;
	if (options.culture) request.Culture = options.culture;

	const rows: IDataObject[] = [];
	let startRow = 0;
	let totalRows = Infinity;
	let raw: ISearchResponse | undefined;
	for (;;) {
		const size = nextPageSize(pageSize, startRow, totalRows, wanted);
		if (size === 0) break;
		const response = await sp.search({ ...request, StartRow: startRow, RowLimit: size });
		raw = response.RawSearchResults;
		totalRows = response.TotalRows;
		const pageRows = (raw.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? []) as SearchRow[];
		rows.push(...pageRows.map(flattenCells));
		if (pageRows.length === 0) break;
		startRow += pageRows.length;
	}

	if (options.includeRawResult && raw) {
		if (rows.length === 0) return [{ rawResult: raw as unknown as IDataObject }];
		rows[rows.length - 1] = { ...rows[rows.length - 1], rawResult: raw as unknown as IDataObject };
	}
	return rows;
}

async function suggest(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const querytext = (this.getNodeParameter('suggestText', itemIndex, '') as string).trim();
	if (!querytext) {
		throw new NodeOperationError(this.getNode(), 'The text must not be empty', { itemIndex });
	}
	const count = this.getNodeParameter('suggestCount', itemIndex, 10) as number;
	const result = await sp.searchSuggest({ querytext, count });
	return {
		queries: result.Queries ?? [],
		peopleNames: result.PeopleNames ?? [],
		personalResults: result.PersonalResults ?? [],
	} as unknown as IDataObject;
}

export const searchResource: ResourceModule = {
	option: { name: 'Search', value: 'search', description: 'Run SharePoint search (KQL) queries' },
	description,
	operations: { query, suggest },
};
