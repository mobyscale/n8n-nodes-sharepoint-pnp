import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import { isSystemAccount, toLoginName, userRLC } from '../../shared/securables';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import type { OperationContext, ResourceModule } from '../types';
import { locatorValue } from '../../shared/pickers';

const RESOURCE = 'user';
const show = (operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [RESOURCE], operation: operations, ...extra },
});

const USER_SELECT = [
	'Id',
	'Title',
	'LoginName',
	'Email',
	'PrincipalType',
	'IsSiteAdmin',
	'IsHiddenInUI',
	'UserPrincipalName',
];

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Ensure',
				value: 'ensure',
				description: 'Get a user and add the user to the site user list if needed',
				action: 'Ensure a user',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a user of the site by ID, email address or login name',
				action: 'Get a user',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many users of the site',
				action: 'Get many users',
			},
		],
		default: 'get',
	},
	{
		...userRLC,
		description:
			'The user. Pick a user of the site or of the directory, or type an email address or a login name.',
		displayOptions: show(['ensure']),
	},
	{
		displayName: 'Identify By',
		name: 'identifyBy',
		type: 'options',
		default: 'email',
		options: [
			{ name: 'Email', value: 'email' },
			{ name: 'ID', value: 'id' },
			{ name: 'Login Name', value: 'loginName' },
		],
		displayOptions: show(['get']),
	},
	{
		...userRLC,
		displayName: 'Email',
		name: 'email',
		description: 'The email address of a user of the site',
		displayOptions: show(['get'], { identifyBy: ['email'] }),
	},
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'number',
		typeOptions: { minValue: 1 },
		required: true,
		default: 1,
		displayOptions: show(['get'], { identifyBy: ['id'] }),
	},
	{
		...userRLC,
		displayName: 'Login Name',
		name: 'loginName',
		description: 'The login name of a user of the site',
		displayOptions: show(['get'], { identifyBy: ['loginName'] }),
	},
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
		displayOptions: show(['getAll'], { returnAll: [false] }),
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
				displayName: 'Exclude System Accounts',
				name: 'excludeSystemAccounts',
				type: 'boolean',
				default: false,
				description: 'Whether to leave out hidden users and built-in accounts such as Everyone',
			},
			{
				displayName: 'Principal Type',
				name: 'principalType',
				type: 'options',
				default: 1,
				options: [
					{ name: 'Distribution List', value: 2 },
					{ name: 'Security Group', value: 4 },
					{ name: 'SharePoint Group', value: 8 },
					{ name: 'User', value: 1 },
				],
			},
		],
	},
	{
		displayName: 'Options',
		name: 'getOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: selectionOptions(USER_SELECT),
	},
	{
		displayName: 'Options',
		name: 'getAllOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: selectionOptions(USER_SELECT),
	},
];

async function ensure(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const user = locatorValue(this.getNodeParameter('user', itemIndex, ''));
	if (!user) {
		throw new NodeOperationError(this.getNode(), 'The user must not be empty', { itemIndex });
	}
	return (await sp.web.ensureUser(toLoginName(user))) as unknown as IDataObject;
}

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const identifyBy = this.getNodeParameter('identifyBy', itemIndex) as string;
	const options = this.getNodeParameter('getOptions', itemIndex, {}) as SelectionValues;
	const users = sp.web.siteUsers;
	const read = <T extends Parameters<typeof applySelection>[0]>(query: T) =>
		applySelection(query, USER_SELECT, options);
	if (identifyBy === 'id') {
		const id = this.getNodeParameter('userId', itemIndex) as number;
		return (await read(users.getById(id))()) as unknown as IDataObject;
	}
	if (identifyBy === 'loginName') {
		const loginName = locatorValue(this.getNodeParameter('loginName', itemIndex, ''));
		if (!loginName) {
			throw new NodeOperationError(this.getNode(), 'The login name must not be empty', {
				itemIndex,
			});
		}
		return (await read(users.getByLoginName(toLoginName(loginName)))()) as unknown as IDataObject;
	}
	const email = locatorValue(this.getNodeParameter('email', itemIndex, ''));
	if (!email) {
		throw new NodeOperationError(this.getNode(), 'The email must not be empty', { itemIndex });
	}
	// The picker can give a login name when a user has no email address.
	if (email.includes('|') || !email.includes('@')) {
		return (await read(users.getByLoginName(toLoginName(email)))()) as unknown as IDataObject;
	}
	return (await read(users.getByEmail(email))()) as unknown as IDataObject;
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const filters = this.getNodeParameter('filters', itemIndex, {}) as {
		principalType?: number;
		excludeSystemAccounts?: boolean;
	};

	const options = this.getNodeParameter('getAllOptions', itemIndex, {}) as SelectionValues;
	let query = applySelection(sp.web.siteUsers, USER_SELECT, options);
	if (filters.principalType !== undefined) {
		query = query.filter(`PrincipalType eq ${filters.principalType}`);
	}
	let users = (await query()) as unknown as Array<{
		LoginName?: string;
		IsHiddenInUI?: boolean;
	}>;
	if (filters.excludeSystemAccounts) users = users.filter((user) => !isSystemAccount(user));
	if (!returnAll) users = users.slice(0, this.getNodeParameter('limit', itemIndex) as number);
	return users as unknown as IDataObject[];
}

export const userResource: ResourceModule = {
	option: { name: 'User', value: RESOURCE, description: 'Get and ensure site users' },
	description,
	operations: { ensure, get, getAll },
};
