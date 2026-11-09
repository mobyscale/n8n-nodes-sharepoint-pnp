import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import type { OperationContext, ResourceModule } from '../types';
import { isGuid, normalizeGuid } from './csom';
import {
	createTerm,
	deleteTerm,
	getTerm,
	getTermGroups,
	getTermSet,
	getTermSets,
	getTerms,
	matchesLabel,
	updateTerm,
	type NewLabel,
	type TermLocation,
	type TermMatchType,
} from './service';
import { locatorValue } from '../../shared/pickers';

const RESOURCE = 'taxonomy';
const GUID_REGEX = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$';

const show = (operation: string[], extra: Record<string, Array<string | boolean>> = {}) => ({
	show: { resource: [RESOURCE], operation, ...extra },
});

function taxonomyLocator(
	displayName: string,
	name: string,
	description: string,
	searchListMethod: string,
	required: boolean,
): INodeProperties {
	return {
		displayName,
		name,
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required,
		description,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod, searchable: true },
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: '00000000-0000-0000-0000-000000000000',
				validation: [
					{ type: 'regex', properties: { regex: GUID_REGEX, errorMessage: 'Not a valid ID' } },
				],
			},
		],
	};
}

function termLocator(
	displayName: string,
	name: string,
	description: string,
	required: boolean,
): INodeProperties {
	return {
		...taxonomyLocator(displayName, name, description, 'searchTerms', required),
		hint: 'The list shows the terms of the selected term set',
	};
}

const languageOption = (description: string): INodeProperties => ({
	displayName: 'Language (LCID)',
	name: 'language',
	type: 'number',
	typeOptions: { minValue: 0 },
	default: 0,
	description,
});

const labelsOption = (description: string): INodeProperties => ({
	displayName: 'Additional Labels',
	name: 'labels',
	type: 'fixedCollection',
	typeOptions: { multipleValues: true },
	placeholder: 'Add Label',
	default: {},
	description,
	options: [
		{
			displayName: 'Label',
			name: 'values',
			values: [
				{ displayName: 'Text', name: 'value', type: 'string', default: '' },
				{
					displayName: 'Language Name or ID',
					name: 'language',
					type: 'options',
					typeOptions: { loadOptionsMethod: 'getTermStoreLanguageOptions' },
					default: 0,
					description:
						'The language of the label. 0 uses the default language of the term store. Choose from the list, or specify an LCID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
				},
			],
		},
	],
});

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Create Term',
				value: 'createTerm',
				description: 'Create a term in a term set or under another term',
				action: 'Create a term',
			},
			{
				name: 'Delete Term',
				value: 'deleteTerm',
				description: 'Delete a term and its child terms',
				action: 'Delete a term',
			},
			{ name: 'Get Term', value: 'getTerm', description: 'Get a term', action: 'Get a term' },
			{
				name: 'Get Term Groups',
				value: 'getTermGroups',
				description: 'Get the term groups of the term store',
				action: 'Get term groups',
			},
			{
				name: 'Get Term Set',
				value: 'getTermSet',
				description: 'Get a term set',
				action: 'Get a term set',
			},
			{
				name: 'Get Term Sets',
				value: 'getTermSets',
				description: 'Get the term sets of a group, or of all groups',
				action: 'Get term sets',
			},
			{
				name: 'Get Terms',
				value: 'getTerms',
				description: 'Get the terms of a term set',
				action: 'Get terms',
			},
			{
				name: 'Search Terms',
				value: 'searchTerms',
				description: 'Find the terms of a term set by label',
				action: 'Search terms by label',
			},
			{
				name: 'Update Term',
				value: 'updateTerm',
				description: 'Change the name, description or labels of a term',
				action: 'Update a term',
			},
		],
		default: 'getTermGroups',
	},

	{
		...taxonomyLocator(
			'Term Group',
			'termGroup',
			'The term group. Empty returns the term sets of all groups.',
			'searchTermGroups',
			false,
		),
		displayOptions: show(['getTermSets']),
	},
	{
		...taxonomyLocator(
			'Term Set',
			'termSet',
			'The term set (a list of terms)',
			'searchTermSets',
			true,
		),
		displayOptions: show(['getTermSet', 'getTerms', 'searchTerms', 'createTerm']),
	},
	{
		...taxonomyLocator(
			'Term Set (to Pick From)',
			'termSetHelper',
			'Optional. Only fills the "Term" list below. The operation does not use it.',
			'searchTermSets',
			false,
		),
		displayOptions: show(['getTerm', 'updateTerm', 'deleteTerm']),
	},
	{
		...termLocator('Term', 'termId', 'The term', true),
		displayOptions: show(['getTerm', 'updateTerm', 'deleteTerm']),
	},

	{
		displayName: 'Term Name',
		name: 'termName',
		type: 'string',
		default: '',
		required: true,
		description: 'The name of the new term. It becomes the default label.',
		displayOptions: show(['createTerm']),
	},
	{
		displayName: 'Options',
		name: 'createOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createTerm']),
		options: [
			{
				...termLocator(
					'Parent Term',
					'parentTermId',
					'Create the term under this term. Empty creates it at the top of the term set.',
					false,
				),
				default: { mode: 'list', value: '' },
			},
			{
				displayName: 'Description',
				name: 'description',
				type: 'string',
				default: '',
				description: 'The description of the term',
			},
			languageOption('The language of the name. 0 uses the default language of the term store.'),
			labelsOption('More labels (synonyms or translations) for the term'),
		],
	},
	{
		displayName: 'Update Fields',
		name: 'updateOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['updateTerm']),
		options: [
			{
				displayName: 'Name',
				name: 'name',
				type: 'string',
				default: '',
				description: 'The new name (default label) of the term',
			},
			{
				displayName: 'Description',
				name: 'description',
				type: 'string',
				default: '',
				description: 'The new description. Use an empty text to clear it.',
			},
			languageOption(
				'The language of the description and of new labels. 0 uses the default language.',
			),
			labelsOption('Labels to add to the term'),
		],
	},

	{
		displayName: 'Text',
		name: 'searchText',
		type: 'string',
		default: '',
		required: true,
		description: 'The label text to look for',
		displayOptions: show(['searchTerms']),
	},
	{
		displayName: 'Match',
		name: 'matchType',
		type: 'options',
		options: [
			{ name: 'Exact', value: 'exact' },
			{ name: 'Starts With', value: 'startsWith' },
			{ name: 'Contains', value: 'contains' },
		],
		default: 'startsWith',
		description: 'How the label must match the text. The case is ignored.',
		displayOptions: show(['searchTerms']),
	},
	{
		displayName: 'Options',
		name: 'searchOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['searchTerms']),
		options: [
			languageOption('Only look at labels of this language. 0 looks at all languages.'),
			{
				displayName: 'Include Deprecated',
				name: 'includeDeprecated',
				type: 'boolean',
				default: false,
				description: 'Whether to include deprecated terms',
			},
		],
	},

	{
		displayName: 'Options',
		name: 'getTermsOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getTerms']),
		options: [
			{
				displayName: 'Include Children (Recursive)',
				name: 'recursive',
				type: 'boolean',
				default: false,
				description:
					'Whether to include all child terms at all levels. If off, only the terms directly under the term set or parent term are returned.',
			},
			{
				...termLocator(
					'Parent Term',
					'parentTermId',
					'Only return the terms under this term. Empty starts at the top of the term set.',
					false,
				),
				default: { mode: 'list', value: '' },
			},
		],
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['getTerms', 'searchTerms', 'getTermSets']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: show(['getTerms', 'searchTerms', 'getTermSets'], { returnAll: [false] }),
	},
];

function locator(
	this: IExecuteFunctions,
	name: string,
	itemIndex: number,
	required: boolean,
): string {
	const value = String(
		this.getNodeParameter(name, itemIndex, '', { extractValue: true }) ?? '',
	).trim();
	if (!value) {
		if (!required) return '';
		throw new NodeOperationError(this.getNode(), `The parameter "${name}" must not be empty`, {
			itemIndex,
		});
	}
	if (!isGuid(value)) {
		throw new NodeOperationError(this.getNode(), `"${value}" is not a valid ID (GUID)`, {
			itemIndex,
		});
	}
	return normalizeGuid(value);
}

function termId(this: IExecuteFunctions, name: string, itemIndex: number, required = true): string {
	const value = locatorValue(this.getNodeParameter(name, itemIndex, ''));
	if (!value) {
		if (!required) return '';
		throw new NodeOperationError(this.getNode(), `The parameter "${name}" must not be empty`, {
			itemIndex,
		});
	}
	if (!isGuid(value)) {
		throw new NodeOperationError(this.getNode(), `"${value}" is not a valid term ID (GUID)`, {
			itemIndex,
		});
	}
	return normalizeGuid(value);
}

function limitOf(this: IExecuteFunctions, itemIndex: number): number {
	return this.getNodeParameter('returnAll', itemIndex, false)
		? Number.POSITIVE_INFINITY
		: (this.getNodeParameter('limit', itemIndex, 50) as number);
}

function readLabels(raw: unknown): NewLabel[] {
	const values = (raw as { values?: Array<{ value?: string; language?: number }> })?.values ?? [];
	return values
		.filter((entry) => entry.value?.trim())
		.map((entry) => ({ value: String(entry.value).trim(), language: Number(entry.language ?? 0) }));
}

const where = ({ siteUrl }: OperationContext): TermLocation => ({ siteUrl });

async function getTermGroupsOperation(this: IExecuteFunctions, context: OperationContext) {
	return (await getTermGroups(this, where(context))) as unknown as IDataObject[];
}

async function getTermSetsOperation(this: IExecuteFunctions, context: OperationContext) {
	const groupId = locator.call(this, 'termGroup', context.itemIndex, false);
	const sets = await getTermSets(this, where(context), groupId || undefined);
	return sets.slice(0, limitOf.call(this, context.itemIndex)) as unknown as IDataObject[];
}

async function getTermSetOperation(this: IExecuteFunctions, context: OperationContext) {
	const id = locator.call(this, 'termSet', context.itemIndex, true);
	return (await getTermSet(this, where(context), id)) as unknown as IDataObject;
}

async function getTermsOperation(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const options = this.getNodeParameter('getTermsOptions', itemIndex, {}) as {
		recursive?: boolean;
		parentTermId?: unknown;
	};
	const parentTermId = locatorValue(options.parentTermId);
	if (parentTermId && !isGuid(parentTermId)) {
		throw new NodeOperationError(
			this.getNode(),
			`"${parentTermId}" is not a valid term ID (GUID)`,
			{
				itemIndex,
			},
		);
	}
	const terms = await getTerms(
		this,
		where(context),
		locator.call(this, 'termSet', itemIndex, true),
		{
			recursive: options.recursive === true,
			parentTermId: parentTermId && normalizeGuid(parentTermId),
		},
	);
	return terms.slice(0, limitOf.call(this, itemIndex)) as unknown as IDataObject[];
}

async function searchTermsOperation(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const text = String(this.getNodeParameter('searchText', itemIndex, '')).trim();
	if (!text) {
		throw new NodeOperationError(this.getNode(), 'The search text must not be empty', {
			itemIndex,
		});
	}
	const matchType = this.getNodeParameter('matchType', itemIndex, 'startsWith') as TermMatchType;
	const options = this.getNodeParameter('searchOptions', itemIndex, {}) as {
		language?: number;
		includeDeprecated?: boolean;
	};
	const terms = await getTerms(
		this,
		where(context),
		locator.call(this, 'termSet', itemIndex, true),
		{
			recursive: true,
		},
	);
	return terms
		.filter((term) => options.includeDeprecated || !term.isDeprecated)
		.filter((term) => matchesLabel(term, text, matchType, options.language))
		.slice(0, limitOf.call(this, itemIndex)) as unknown as IDataObject[];
}

async function getTermOperation(this: IExecuteFunctions, context: OperationContext) {
	return (await getTerm(
		this,
		where(context),
		termId.call(this, 'termId', context.itemIndex),
	)) as unknown as IDataObject;
}

async function createTermOperation(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const name = String(this.getNodeParameter('termName', itemIndex, '')).trim();
	if (!name) {
		throw new NodeOperationError(this.getNode(), 'The term name must not be empty', { itemIndex });
	}
	const options = this.getNodeParameter('createOptions', itemIndex, {}) as {
		parentTermId?: unknown;
		description?: string;
		language?: number;
		labels?: unknown;
	};
	const parent = locatorValue(options.parentTermId);
	if (parent && !isGuid(parent)) {
		throw new NodeOperationError(this.getNode(), `"${parent}" is not a valid term ID (GUID)`, {
			itemIndex,
		});
	}
	return (await createTerm(this, where(context), locator.call(this, 'termSet', itemIndex, true), {
		name,
		language: options.language,
		description: options.description?.trim() || undefined,
		parentTermId: parent ? normalizeGuid(parent) : undefined,
		labels: readLabels(options.labels),
	})) as unknown as IDataObject;
}

async function updateTermOperation(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const options = this.getNodeParameter('updateOptions', itemIndex, {}) as {
		name?: string;
		description?: string;
		language?: number;
		labels?: unknown;
	};
	const labels = readLabels(options.labels);
	if (options.name === undefined && options.description === undefined && !labels.length) {
		throw new NodeOperationError(
			this.getNode(),
			'No changes were given. Add an option to update.',
			{
				itemIndex,
			},
		);
	}
	return (await updateTerm(this, where(context), termId.call(this, 'termId', itemIndex), {
		name: options.name?.trim() || undefined,
		description: options.description,
		language: options.language,
		labels,
	})) as unknown as IDataObject;
}

async function deleteTermOperation(this: IExecuteFunctions, context: OperationContext) {
	const id = termId.call(this, 'termId', context.itemIndex);
	await deleteTerm(this, where(context), id);
	return { success: true, id } as IDataObject;
}

export const taxonomyResource: ResourceModule = {
	option: {
		name: 'Taxonomy',
		value: RESOURCE,
		description: 'Read and change managed metadata terms (tags)',
	},
	description,
	operations: {
		getTermGroups: getTermGroupsOperation,
		getTermSets: getTermSetsOperation,
		getTermSet: getTermSetOperation,
		getTerms: getTermsOperation,
		getTerm: getTermOperation,
		searchTerms: searchTermsOperation,
		createTerm: createTermOperation,
		updateTerm: updateTermOperation,
		deleteTerm: deleteTermOperation,
	},
};
