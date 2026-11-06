import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import { isSystemAccount, toLoginName } from '../../shared/securables';
import type { OperationContext, ResourceModule } from '../types';

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
		displayName: 'User',
		name: 'user',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'user@contoso.com',
		description: 'The email address or login name of the user',
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
		displayName: 'Email',
		name: 'email',
		type: 'string',
		placeholder: 'name@email.com',
		required: true,
		default: '',
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
		displayName: 'Login Name',
		name: 'loginName',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'i:0#.f|membership|user@contoso.com',
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
];

async function ensure(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const user = String(this.getNodeParameter('user', itemIndex, '') ?? '').trim();
	if (!user) {
		throw new NodeOperationError(this.getNode(), 'The user must not be empty', { itemIndex });
	}
	return (await sp.web.ensureUser(toLoginName(user))) as unknown as IDataObject;
}

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const identifyBy = this.getNodeParameter('identifyBy', itemIndex) as string;
	const users = sp.web.siteUsers;
	if (identifyBy === 'id') {
		const id = this.getNodeParameter('userId', itemIndex) as number;
		return (await users.getById(id).select(...USER_SELECT)()) as unknown as IDataObject;
	}
	if (identifyBy === 'loginName') {
		const loginName = String(this.getNodeParameter('loginName', itemIndex, '') ?? '').trim();
		if (!loginName) {
			throw new NodeOperationError(this.getNode(), 'The login name must not be empty', {
				itemIndex,
			});
		}
		return (await users
			.getByLoginName(toLoginName(loginName))
			.select(...USER_SELECT)()) as unknown as IDataObject;
	}
	const email = String(this.getNodeParameter('email', itemIndex, '') ?? '').trim();
	if (!email) {
		throw new NodeOperationError(this.getNode(), 'The email must not be empty', { itemIndex });
	}
	return (await users.getByEmail(email).select(...USER_SELECT)()) as unknown as IDataObject;
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const filters = this.getNodeParameter('filters', itemIndex, {}) as {
		principalType?: number;
		excludeSystemAccounts?: boolean;
	};

	let query = sp.web.siteUsers.select(...USER_SELECT);
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
