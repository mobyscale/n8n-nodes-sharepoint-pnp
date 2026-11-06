import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { IRecycleBin } from '@pnp/sp/recycle-bin/index.js';

import type { OperationContext, ResourceModule } from '../types';
import {
	findInvalidIds,
	isInSite,
	isPastDeletedAfter,
	ITEM_STATE_FIRST,
	ITEM_STATE_SECOND,
	mapItem,
	matchesFilters,
	parseIds,
	type RecycleBinFilters,
	type RecycleBinRawItem,
} from './helpers';
import { getSiteServerRelativeUrl } from '../../shared/utils';
import { fetchRecycleBinPage, PAGE_SIZE, type RecycleBinScope } from './paging';

const show = (operation: string[]) => ({ show: { resource: ['recycleBin'], operation } });

const scopeProperty: INodeProperties = {
	displayName: 'Scope',
	name: 'scope',
	type: 'options',
	default: 'siteCollection',
	options: [
		{
			name: 'Site Collection',
			value: 'siteCollection',
			description: 'The recycle bin of the site collection, both stages',
		},
		{
			name: 'Site',
			value: 'site',
			description:
				'Items deleted from the selected site. Items of its subsites are included. Get Many reads the site collection bin and filters by path.',
		},
	],
	description: 'Which recycle bin to use',
	displayOptions: { show: { resource: ['recycleBin'] } },
};

const idsProperty = (operationDescription: string): INodeProperties => ({
	displayName: 'Item IDs',
	name: 'ids',
	type: 'string',
	required: true,
	default: '',
	placeholder: '00000000-0000-0000-0000-000000000000',
	description: `${operationDescription}. Use a comma-separated list, or an expression that returns an array.`,
	displayOptions: show(['restore', 'deletePermanently', 'moveToSecondStage']),
});

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['recycleBin'] } },
		options: [
			{
				name: 'Delete Permanently',
				value: 'deletePermanently',
				description: 'Delete recycle bin items for good',
				action: 'Delete recycle bin items permanently',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many recycle bin items',
				action: 'Get many recycle bin items',
			},
			{
				name: 'Move to Second Stage',
				value: 'moveToSecondStage',
				description: 'Move recycle bin items to the second stage recycle bin',
				action: 'Move recycle bin items to the second stage',
			},
			{
				name: 'Restore',
				value: 'restore',
				description: 'Restore recycle bin items to their original location',
				action: 'Restore recycle bin items',
			},
		],
		default: 'getAll',
	},
	scopeProperty,
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
		displayOptions: {
			show: { resource: ['recycleBin'], operation: ['getAll'], returnAll: [false] },
		},
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
				displayName: 'Deleted After',
				name: 'deletedAfter',
				type: 'dateTime',
				default: '',
				description: 'Only items deleted after this date and time',
			},
			{
				displayName: 'Deleted Before',
				name: 'deletedBefore',
				type: 'dateTime',
				default: '',
				description: 'Only items deleted before this date and time',
			},
			{
				displayName: 'Deleted By (Email Contains)',
				name: 'deletedBy',
				type: 'string',
				default: '',
				description: 'Only items where the email of the user who deleted it contains this text',
			},
			{
				displayName: 'Item Type',
				name: 'itemType',
				type: 'options',
				default: -1,
				options: [
					{ name: 'Any', value: -1 },
					{ name: 'File', value: 1 },
					{ name: 'Folder', value: 5 },
					{ name: 'List', value: 4 },
					{ name: 'List Item', value: 3 },
				],
			},
			{
				displayName: 'Stage',
				name: 'stage',
				type: 'options',
				default: -1,
				options: [
					{ name: 'Any', value: -1 },
					{ name: 'First', value: ITEM_STATE_FIRST },
					{ name: 'Second', value: ITEM_STATE_SECOND },
				],
				description: 'The recycle bin stage',
			},
			{
				displayName: 'Title or Original Location Contains',
				name: 'text',
				type: 'string',
				default: '',
				description:
					'Only items where the title, file name or original location contains this text',
			},
		],
	},
	{ ...idsProperty('The IDs of the items to restore'), displayOptions: show(['restore']) },
	{
		...idsProperty('The IDs of the items to delete for good. This cannot be undone'),
		displayOptions: show(['deletePermanently']),
	},
	{
		...idsProperty('The IDs of the items to move to the second stage'),
		displayOptions: show(['moveToSecondStage']),
	},
];

function getBin({ sp }: OperationContext, scope: RecycleBinScope): IRecycleBin {
	return scope === 'site' ? sp.web.recycleBin : sp.site.recycleBin;
}

function getIds(this: IExecuteFunctions, itemIndex: number): string[] {
	const ids = parseIds(this.getNodeParameter('ids', itemIndex, ''));
	if (ids.length === 0) {
		throw new NodeOperationError(this.getNode(), 'Provide at least one recycle bin item ID', {
			itemIndex,
		});
	}
	const invalid = findInvalidIds(ids);
	if (invalid.length) {
		throw new NodeOperationError(
			this.getNode(),
			`These values are not valid recycle bin item IDs: ${invalid.join(', ')}`,
			{ itemIndex },
		);
	}
	return ids;
}

function getAction(
	run: (bin: IRecycleBin, id: string) => Promise<void>,
	resultKey: string,
): (this: IExecuteFunctions, context: OperationContext) => Promise<IDataObject[]> {
	return async function (this: IExecuteFunctions, context: OperationContext) {
		const scope = this.getNodeParameter(
			'scope',
			context.itemIndex,
			'siteCollection',
		) as RecycleBinScope;
		const bin = getBin(context, scope);
		const results: IDataObject[] = [];
		for (const id of getIds.call(this, context.itemIndex)) {
			await run(bin, id);
			results.push({ id, [resultKey]: true });
		}
		return results;
	};
}

const restore = getAction((bin, id) => bin.getById(id).restore(), 'restored');
const deletePermanently = getAction((bin, id) => bin.getById(id).delete(), 'deleted');
const moveToSecondStage = getAction((bin, id) => bin.getById(id).moveToSecondStage(), 'moved');

async function getAll(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const scope = this.getNodeParameter('scope', itemIndex, 'siteCollection') as RecycleBinScope;
	const returnAll = this.getNodeParameter('returnAll', itemIndex, false) as boolean;
	const limit = returnAll ? Infinity : (this.getNodeParameter('limit', itemIndex, 50) as number);
	const filters = this.getNodeParameter('filters', itemIndex, {}) as RecycleBinFilters;

	const stages =
		filters.stage === ITEM_STATE_FIRST || filters.stage === ITEM_STATE_SECOND
			? [filters.stage]
			: [ITEM_STATE_FIRST, ITEM_STATE_SECOND];
	const sitePath = scope === 'site' ? getSiteServerRelativeUrl(siteUrl) : undefined;

	const results: RecycleBinRawItem[] = [];
	for (const stage of stages) {
		let pagingInfo: string | undefined;
		let done = false;
		while (!done && results.length < limit) {
			const page = await fetchRecycleBinPage(sp, stage, pagingInfo, PAGE_SIZE);
			for (const item of page.items) {
				if (matchesFilters(item, filters) && (!sitePath || isInSite(item, sitePath))) {
					results.push(item);
					if (results.length >= limit) break;
				}
			}
			const last = page.items[page.items.length - 1];
			done = !page.nextPagingInfo || !last || isPastDeletedAfter(last, filters.deletedAfter);
			pagingInfo = page.nextPagingInfo;
		}
		if (results.length >= limit) break;
	}
	return results.map(mapItem);
}

export const recycleBinResource: ResourceModule = {
	option: {
		name: 'Recycle Bin',
		value: 'recycleBin',
		description: 'Find, restore and delete recycled items',
	},
	description,
	operations: { getAll, restore, deletePermanently, moveToSecondStage },
};
