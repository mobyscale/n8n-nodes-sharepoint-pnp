import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import type {
	ChangeOutputOptions,
	ChangeScope,
	ExpiredTokenMode,
	LocatorValue,
	QueryOptions,
} from '../../shared/changes';
import {
	buildChangeQuery,
	changeOptionsProperty,
	changeScopeProperty,
	changeTypesProperty,
	createChangeTokenFromDate,
	getChangeTarget,
	objectTypesProperty,
	queryOptionsProperty,
	resolveFolderFilter,
	runChangeQueryWithExpiry,
	showFor,
	validateFlags,
} from '../../shared/changes';
import { listRLC } from '../../shared/descriptions';
import type { OperationContext, ResourceModule } from '../types';

const show = (operation: string[]) => ({ resource: ['change'], operation });

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['change'] } },
		options: [
			{
				name: 'Get Current Token',
				value: 'getCurrentToken',
				description: 'Get the current change token of a site collection, site or list',
				action: 'Get the current change token',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many changes from the change log',
				action: 'Get many changes',
			},
		],
		default: 'getAll',
	},
	...showFor([changeScopeProperty], show(['getAll', 'getCurrentToken'])),
	...showFor([listRLC], { ...show(['getAll', 'getCurrentToken']), scope: ['list'] }),
	...showFor(
		[
			{
				displayName: 'Start From',
				name: 'startFrom',
				type: 'options',
				options: [
					{
						name: 'Beginning',
						value: 'beginning',
						description: 'All changes that are still in the change log',
					},
					{ name: 'Change Token', value: 'token', description: 'Changes after a change token' },
					{ name: 'Date', value: 'date', description: 'Changes after a date and time' },
				],
				default: 'token',
			},
		],
		show(['getAll']),
	),
	...showFor(
		[
			{
				displayName: 'Change Token',
				name: 'startPosition',
				type: 'string',
				required: true,
				default: '',
				placeholder: '1;3;00000000-0000-0000-0000-000000000000;638000000000000000;12345',
				description:
					'Return the changes after this token. Use "Get Current Token" or the "nextChangeToken" of an earlier run.',
			},
		],
		{ ...show(['getAll']), startFrom: ['token'] },
	),
	...showFor(
		[
			{
				displayName: 'Start Date',
				name: 'startDate',
				type: 'dateTime',
				required: true,
				default: '',
				description: 'Return the changes after this date and time',
			},
		],
		{ ...show(['getAll']), startFrom: ['date'] },
	),
	...showFor([objectTypesProperty, changeTypesProperty], show(['getAll'])),
	...showFor(
		[
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				default: false,
				description: 'Whether to return all results or only up to a given limit',
			},
		],
		show(['getAll']),
	),
	...showFor(
		[
			{
				displayName: 'Limit',
				name: 'limit',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 50,
				description: 'Max number of results to return',
			},
		],
		{ ...show(['getAll']), returnAll: [false] },
	),
	...showFor([queryOptionsProperty, changeOptionsProperty], show(['getAll'])),
];

function getTarget(
	this: IExecuteFunctions,
	{ sp, itemIndex }: OperationContext,
	folderPath?: string,
) {
	const scope = this.getNodeParameter('scope', itemIndex) as ChangeScope;
	const list =
		scope === 'list' ? (this.getNodeParameter('list', itemIndex) as LocatorValue) : undefined;
	if (scope === 'list' && !String(list?.value ?? '').trim()) {
		throw new NodeOperationError(this.getNode(), 'Select a list', { itemIndex });
	}
	return getChangeTarget(sp, scope, list, scope === 'list' ? folderPath : undefined);
}

async function getCurrentToken(this: IExecuteFunctions, context: OperationContext) {
	const target = getTarget.call(this, context);
	return { scope: target.scope, changeToken: await target.getCurrentToken() };
}

async function getAll(this: IExecuteFunctions, context: OperationContext) {
	const { sp, siteUrl, itemIndex } = context;
	const startFrom = this.getNodeParameter('startFrom', itemIndex) as string;
	const objectTypes = this.getNodeParameter('objectTypes', itemIndex, []) as string[];
	const changeTypes = this.getNodeParameter('changeTypes', itemIndex, []) as string[];
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const limit = returnAll ? undefined : (this.getNodeParameter('limit', itemIndex) as number);
	const queryOptions = this.getNodeParameter('queryOptions', itemIndex, {}) as QueryOptions;
	const options = this.getNodeParameter('options', itemIndex, {}) as Pick<
		ChangeOutputOptions,
		'includeItemDetails' | 'itemFields' | 'resolveDeletedItems' | 'excludeEditors'
	> & {
		changeFieldFilters?: unknown;
		changeItemKinds?: string[];
		changeMaxDepth?: number;
		folderPathFilter?: unknown;
		onExpiredToken?: ExpiredTokenMode;
		output?: string;
	};
	validateFlags(this.getNode(), objectTypes, changeTypes, itemIndex);
	const folderPathFilter = resolveFolderFilter(siteUrl, options.folderPathFilter);
	const target = getTarget.call(this, context, folderPathFilter);

	let startToken: string | undefined;
	if (startFrom === 'token') {
		startToken = String(this.getNodeParameter('startPosition', itemIndex, '')).trim();
		if (!startToken) {
			throw new NodeOperationError(this.getNode(), 'The change token must not be empty', {
				itemIndex,
			});
		}
	} else if (startFrom === 'date') {
		const date = new Date(this.getNodeParameter('startDate', itemIndex, '') as string);
		if (!Number.isFinite(date.getTime())) {
			throw new NodeOperationError(this.getNode(), 'The start date is not valid', { itemIndex });
		}
		startToken = createChangeTokenFromDate(target.scope, await target.getId(), date);
	}

	const query = buildChangeQuery({ objectTypes, changeTypes, startToken, queryOptions });
	const result = await runChangeQueryWithExpiry(sp, target, query, {
		includeItemDetails: options.includeItemDetails,
		itemFields: options.itemFields,
		resolveDeletedItems: options.resolveDeletedItems,
		onExpiredToken: options.onExpiredToken,
		itemKinds: options.changeItemKinds,
		maxDepth: options.changeMaxDepth,
		excludeEditors: options.excludeEditors,
		fieldFilters: options.changeFieldFilters,
		folderPathFilter,
		returnAll,
		limit,
	});

	if (options.output === 'single') {
		const nextChangeToken =
			result.nextChangeToken ?? startToken ?? (await target.getCurrentToken());
		const output: IDataObject = {
			changes: result.changes,
			nextChangeToken,
			count: result.changes.length,
		};
		if (result.tokenExpired) output.tokenExpired = true;
		return output;
	}
	return result.changes;
}

export const changeResource: ResourceModule = {
	option: {
		name: 'Change',
		value: 'change',
		description: 'Read the change log with change tokens',
	},
	description,
	operations: { getAll, getCurrentToken },
};
