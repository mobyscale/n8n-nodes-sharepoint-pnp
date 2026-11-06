import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';

import { getList, listRLC } from '../../shared/descriptions';
import type { OperationContext, ResourceModule } from '../types';

const show = (operation: string[]) => ({ show: { resource: ['list'], operation } });

const LIST_SELECT = [
	'Id',
	'Title',
	'Description',
	'BaseTemplate',
	'BaseType',
	'ItemCount',
	'Hidden',
	'Created',
	'LastItemModifiedDate',
	'RootFolder/ServerRelativeUrl',
];

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['list'] } },
		options: [
			{ name: 'Get', value: 'get', description: 'Get a list or library', action: 'Get a list' },
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many lists and libraries in a site',
				action: 'Get many lists',
			},
		],
		default: 'getAll',
	},
	{ ...listRLC, displayOptions: show(['get']) },
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['getAll']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: { show: { resource: ['list'], operation: ['getAll'], returnAll: [false] } },
	},
	{
		displayName: 'Filters',
		name: 'filters',
		type: 'collection',
		placeholder: 'Add Filter',
		default: {},
		displayOptions: show(['getAll']),
		options: [
			{
				displayName: 'Include Hidden',
				name: 'includeHidden',
				type: 'boolean',
				default: false,
				description: 'Whether to include hidden lists',
			},
			{
				displayName: 'Type',
				name: 'baseTemplate',
				type: 'options',
				default: -1,
				options: [
					{ name: 'Any', value: -1 },
					{ name: 'Document Library', value: 101 },
					{ name: 'Generic List', value: 100 },
					{ name: 'Picture Library', value: 109 },
					{ name: 'Site Pages', value: 119 },
				],
			},
		],
	},
];

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	return (await getList
		.call(this, sp, itemIndex)
		.select(...LIST_SELECT)
		.expand('RootFolder')()) as unknown as IDataObject;
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const filters = this.getNodeParameter('filters', itemIndex, {}) as {
		includeHidden?: boolean;
		baseTemplate?: number;
	};

	const conditions: string[] = [];
	if (!filters.includeHidden) conditions.push('Hidden eq false');
	if (filters.baseTemplate !== undefined && filters.baseTemplate !== -1) {
		conditions.push(`BaseTemplate eq ${filters.baseTemplate}`);
	}

	let query = sp.web.lists.select(...LIST_SELECT).expand('RootFolder');
	if (conditions.length) query = query.filter(conditions.join(' and '));
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);

	return (await query()) as unknown as IDataObject[];
}

export const listResource: ResourceModule = {
	option: {
		name: 'List',
		value: 'list',
		description: 'Get, create and manage lists and libraries',
	},
	description,
	operations: { get, getAll },
};
