import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import {
	ensureUserId,
	findUserId,
	getUserParameter,
	groupRLC,
	parsePrincipalOrThrow,
	resolveGroupId,
	resolvePrincipalReference,
	type LocatorValue,
} from '../../shared/securables';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import type { OperationContext, ResourceModule } from '../types';

const RESOURCE = 'group';
const show = (operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [RESOURCE], operation: operations, ...extra },
});

const GROUP_SELECT = [
	'Id',
	'Title',
	'Description',
	'LoginName',
	'OwnerTitle',
	'PrincipalType',
	'AllowMembersEditMembership',
	'AllowRequestToJoinLeave',
	'AutoAcceptRequestToJoinLeave',
	'OnlyAllowMembersViewMembership',
	'RequestToJoinLeaveEmailSetting',
	'IsHiddenInUI',
];

const USER_SELECT = ['Id', 'Title', 'LoginName', 'Email', 'PrincipalType', 'IsSiteAdmin'];

interface GroupFields {
	title?: string;
	description?: string;
	owner?: string;
	allowMembersEditMembership?: boolean;
	allowRequestToJoinLeave?: boolean;
	autoAcceptRequestToJoinLeave?: boolean;
	onlyAllowMembersViewMembership?: boolean;
	requestToJoinLeaveEmailSetting?: string;
}

const settingFields = (): INodeProperties[] => [
	{
		displayName: 'Allow Members to Edit Membership',
		name: 'allowMembersEditMembership',
		type: 'boolean',
		default: false,
		description: 'Whether members of the group can add and remove members',
	},
	{
		displayName: 'Allow Requests to Join or Leave',
		name: 'allowRequestToJoinLeave',
		type: 'boolean',
		default: false,
		description: 'Whether users can ask to join or leave the group',
	},
	{
		displayName: 'Auto-Accept Requests',
		name: 'autoAcceptRequestToJoinLeave',
		type: 'boolean',
		default: false,
		description: 'Whether requests to join or leave are accepted without approval',
	},
	{
		displayName: 'Description',
		name: 'description',
		type: 'string',
		default: '',
		description: 'The description of the group',
	},
	{
		displayName: 'Only Members Can View Membership',
		name: 'onlyAllowMembersViewMembership',
		type: 'boolean',
		default: false,
		description: 'Whether only members of the group can see who the members are',
	},
	{
		displayName: 'Owner',
		name: 'owner',
		type: 'string',
		default: '',
		placeholder: 'owner@contoso.com',
		description:
			'The owner of the group: a user email address or login name, a group name, or a principal ID',
	},
	{
		displayName: 'Request Email',
		name: 'requestToJoinLeaveEmailSetting',
		type: 'string',
		default: '',
		placeholder: 'group-owner@contoso.com',
		description: 'The email address that receives requests to join or leave',
	},
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
				name: 'Add Member',
				value: 'addMember',
				description: 'Add a user to a group',
				action: 'Add a member to a group',
			},
			{
				name: 'Create',
				value: 'create',
				description: 'Create a SharePoint group',
				action: 'Create a group',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a SharePoint group',
				action: 'Delete a group',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a SharePoint group',
				action: 'Get a group',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many SharePoint groups of the site',
				action: 'Get many groups',
			},
			{
				name: 'Get Members',
				value: 'getMembers',
				description: 'Get the members of a group',
				action: 'Get members of a group',
			},
			{
				name: 'Remove Member',
				value: 'removeMember',
				description: 'Remove a user from a group',
				action: 'Remove a member from a group',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update the properties of a group',
				action: 'Update a group',
			},
		],
		default: 'getAll',
	},
	{
		...groupRLC,
		displayOptions: show(['get', 'update', 'delete', 'getMembers', 'addMember', 'removeMember']),
	},
	{
		displayName: 'Title',
		name: 'title',
		type: 'string',
		required: true,
		default: '',
		description: 'The name of the group',
		displayOptions: show(['create']),
	},
	{
		displayName: 'User',
		name: 'user',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'user@contoso.com',
		description: 'The email address or login name of the user',
		displayOptions: show(['addMember', 'removeMember']),
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: show(['create']),
		options: settingFields(),
	},
	{
		displayName: 'Update Fields',
		name: 'updateFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: show(['update']),
		options: [
			{
				displayName: 'Title',
				name: 'title',
				type: 'string',
				default: '',
				description: 'The new name of the group',
			},
			...settingFields(),
		],
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['getAll', 'getMembers']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: show(['getAll', 'getMembers'], { returnAll: [false] }),
	},
	{
		displayName: 'Options',
		name: 'getOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: selectionOptions(GROUP_SELECT),
	},
	{
		displayName: 'Options',
		name: 'getAllOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: selectionOptions(GROUP_SELECT),
	},
	{
		displayName: 'Options',
		name: 'getMembersOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getMembers']),
		options: selectionOptions(USER_SELECT),
	},
];

function buildProperties(fields: GroupFields): IDataObject {
	const properties: IDataObject = {};
	if (fields.title !== undefined && fields.title !== '') properties.Title = fields.title;
	if (fields.description !== undefined) properties.Description = fields.description;
	if (fields.allowMembersEditMembership !== undefined) {
		properties.AllowMembersEditMembership = fields.allowMembersEditMembership;
	}
	if (fields.allowRequestToJoinLeave !== undefined) {
		properties.AllowRequestToJoinLeave = fields.allowRequestToJoinLeave;
	}
	if (fields.autoAcceptRequestToJoinLeave !== undefined) {
		properties.AutoAcceptRequestToJoinLeave = fields.autoAcceptRequestToJoinLeave;
	}
	if (fields.onlyAllowMembersViewMembership !== undefined) {
		properties.OnlyAllowMembersViewMembership = fields.onlyAllowMembersViewMembership;
	}
	if (fields.requestToJoinLeaveEmailSetting) {
		properties.RequestToJoinLeaveEmailSetting = fields.requestToJoinLeaveEmailSetting;
	}
	return properties;
}

async function getGroupId(this: IExecuteFunctions, sp: SPFI, itemIndex: number): Promise<number> {
	return await resolveGroupId(
		this.getNode(),
		sp,
		this.getNodeParameter('group', itemIndex) as LocatorValue,
		itemIndex,
	);
}

async function setOwner(
	context: IExecuteFunctions,
	sp: SPFI,
	groupId: number,
	owner: string | undefined,
	itemIndex: number,
) {
	if (!owner?.trim()) return;
	const ownerId = await resolvePrincipalReference(
		context.getNode(),
		sp,
		parsePrincipalOrThrow(context.getNode(), owner, itemIndex),
		itemIndex,
	);
	await sp.web.siteGroups.getById(groupId).setUserAsOwner(ownerId);
}

async function readGroup(
	sp: SPFI,
	groupId: number,
	options: SelectionValues = {},
): Promise<IDataObject> {
	return (await applySelection(
		sp.web.siteGroups.getById(groupId),
		GROUP_SELECT,
		options,
	)()) as unknown as IDataObject;
}

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('getOptions', itemIndex, {}) as SelectionValues;
	return await readGroup(sp, await getGroupId.call(this, sp, itemIndex), options);
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const options = this.getNodeParameter('getAllOptions', itemIndex, {}) as SelectionValues;
	let query = applySelection(sp.web.siteGroups, GROUP_SELECT, options);
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);
	return (await query()) as unknown as IDataObject[];
}

async function create(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const title = String(this.getNodeParameter('title', itemIndex, '') ?? '').trim();
	if (!title) {
		throw new NodeOperationError(this.getNode(), 'The title must not be empty', { itemIndex });
	}
	const fields = this.getNodeParameter('additionalFields', itemIndex, {}) as GroupFields;
	const created = (await sp.web.siteGroups.add({
		...buildProperties(fields),
		Title: title,
	})) as unknown as { Id: number };
	await setOwner(this, sp, created.Id, fields.owner, itemIndex);
	return await readGroup(sp, created.Id);
}

async function update(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const groupId = await getGroupId.call(this, sp, itemIndex);
	const fields = this.getNodeParameter('updateFields', itemIndex, {}) as GroupFields;
	const properties = buildProperties(fields);
	if (Object.keys(properties).length) {
		await sp.web.siteGroups.getById(groupId).update(properties);
	}
	await setOwner(this, sp, groupId, fields.owner, itemIndex);
	return await readGroup(sp, groupId);
}

async function deleteGroup(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const groupId = await getGroupId.call(this, sp, itemIndex);
	await sp.web.siteGroups.removeById(groupId);
	return { deleted: true, id: groupId };
}

async function getMembers(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const groupId = await getGroupId.call(this, sp, itemIndex);
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const options = this.getNodeParameter('getMembersOptions', itemIndex, {}) as SelectionValues;
	let query = applySelection(sp.web.siteGroups.getById(groupId).users, USER_SELECT, options);
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);
	return (await query()) as unknown as IDataObject[];
}

async function addMember(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const groupId = await getGroupId.call(this, sp, itemIndex);
	const user = getUserParameter.call(this, itemIndex);
	const userId = await ensureUserId(sp, user);
	const siteUser = await sp.web.siteUsers.getById(userId).select('LoginName')();
	await sp.web.siteGroups.getById(groupId).users.add(siteUser.LoginName);
	return { added: true, groupId, userId };
}

async function removeMember(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const groupId = await getGroupId.call(this, sp, itemIndex);
	const userId = await findUserId(sp, getUserParameter.call(this, itemIndex));
	await sp.web.siteGroups.getById(groupId).users.removeById(userId);
	return { removed: true, groupId, userId };
}

export const groupResource: ResourceModule = {
	option: { name: 'Group', value: RESOURCE, description: 'Manage site groups and their members' },
	description,
	operations: {
		get,
		getAll,
		create,
		update,
		delete: deleteGroup,
		getMembers,
		addMember,
		removeMember,
	},
};
