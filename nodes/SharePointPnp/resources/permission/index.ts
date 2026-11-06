import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import type { SecurableQueryable } from '@pnp/sp/security/index.js';

import {
	buildSecurable,
	decodePermissions,
	diffAssignments,
	findUserId,
	flattenAssignments,
	getPrincipalId,
	getRoleId,
	hasUniqueRoleAssignments,
	parsePrincipalOrThrow,
	principalProperties,
	readTarget,
	resolvePrincipalReference,
	resolveRoleId,
	roleRLC,
	targetProperties,
	toLoginName,
	type RoleAssignmentEntry,
	type RoleAssignmentWithBindings,
	type TargetDescriptor,
} from '../../shared/securables';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import type { OperationContext, ResourceModule } from '../types';
import { toAssignmentOutput } from './helpers';

const RESOURCE = 'permission';
const show = (operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [RESOURCE], operation: operations, ...extra },
});

const ASSIGNMENT_SELECT = [
	'PrincipalId',
	'Member/Id',
	'Member/Title',
	'Member/LoginName',
	'Member/PrincipalType',
	'Member/Email',
	'RoleDefinitionBindings/Id',
	'RoleDefinitionBindings/Name',
];

const ASSIGNMENT_EXPAND = ['Member', 'RoleDefinitionBindings'];

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Give a group or user a role on a target',
				action: 'Add a role assignment',
			},
			{
				name: 'Break Inheritance',
				value: 'breakInheritance',
				description: 'Stop a target from inheriting permissions from its parent',
				action: 'Break permission inheritance',
			},
			{
				name: 'Get Effective Permissions',
				value: 'getEffectivePermissions',
				description: 'Get the permissions that a user has on a target',
				action: 'Get effective permissions of a user',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get the role assignments of a target',
				action: 'Get many role assignments',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove a role or all roles of a group or user from a target',
				action: 'Remove a role assignment',
			},
			{
				name: 'Reset Inheritance',
				value: 'resetInheritance',
				description: 'Make a target inherit permissions from its parent again',
				action: 'Reset permission inheritance',
			},
			{
				name: 'Set',
				value: 'set',
				description: 'Replace all role assignments of a target with a given list',
				action: 'Set role assignments',
			},
		],
		default: 'getAll',
	},
	...targetProperties(RESOURCE, [
		'getAll',
		'breakInheritance',
		'resetInheritance',
		'add',
		'remove',
		'set',
		'getEffectivePermissions',
	]),
	{
		displayName: 'Options',
		name: 'breakOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['breakInheritance']),
		options: [
			{
				displayName: 'Clear Subscopes',
				name: 'clearSubscopes',
				type: 'boolean',
				default: true,
				description:
					'Whether to remove unique permissions of child objects, so that they inherit from the target',
			},
			{
				displayName: 'Copy Role Assignments',
				name: 'copyRoleAssignments',
				type: 'boolean',
				default: false,
				description: 'Whether to copy the role assignments of the parent to the target',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'getAllOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: selectionOptions(ASSIGNMENT_SELECT).map((option) =>
			option.name === 'additionalProperties'
				? {
						...option,
						description: `${option.description} The result keeps the default output and adds the SharePoint object in "properties".`,
					}
				: option,
		),
	},
	...principalProperties(RESOURCE, ['add', 'remove']),
	{ ...roleRLC, displayOptions: show(['add']) },
	{
		...roleRLC,
		required: false,
		description: 'The role to remove. Leave empty to remove all roles of the principal.',
		default: { mode: 'list', value: '' },
		displayOptions: show(['remove']),
	},
	{
		displayName: 'Options',
		name: 'addOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['add']),
		options: [
			{
				displayName: 'Break Inheritance First',
				name: 'breakInheritanceFirst',
				type: 'boolean',
				default: true,
				description:
					'Whether to break inheritance first (and copy the current role assignments) if the target inherits permissions',
			},
		],
	},
	{
		displayName: 'Assignments',
		name: 'assignments',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Assignment',
		default: { entries: [] },
		description: 'The role assignments that the target has after the operation',
		displayOptions: show(['set']),
		options: [
			{
				displayName: 'Entries',
				name: 'entries',
				values: [
					{
						displayName: 'Principal',
						name: 'principal',
						type: 'string',
						default: '',
						placeholder: 'Example Members or user@contoso.com',
						description:
							'A group name, a group or user ID, or the email address or login name of a user',
					},
					{
						displayName: 'Role',
						name: 'role',
						type: 'string',
						default: '',
						placeholder: 'Contribute',
						description: 'The name or ID of a role',
					},
				],
			},
		],
	},
	{
		displayName: 'Options',
		name: 'setOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['set']),
		options: [
			{
				displayName: 'Copy Current Assignments First',
				name: 'copyCurrentFirst',
				type: 'boolean',
				default: false,
				description:
					'Whether to copy the current role assignments when this operation has to break inheritance first',
			},
		],
	},
	{
		displayName: 'User',
		name: 'effectiveUser',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'user@contoso.com',
		description: 'The email address or login name of the user',
		displayOptions: show(['getEffectivePermissions']),
	},
];

function describeTarget(target: TargetDescriptor): IDataObject {
	return {
		target: target.kind,
		...(target.path ? { path: target.path } : {}),
		...(target.itemId ? { itemId: target.itemId } : {}),
	};
}

async function getAssignments(
	securable: SecurableQueryable,
	options: SelectionValues = {},
): Promise<RoleAssignmentWithBindings[]> {
	return (await applySelection(
		securable.roleAssignments,
		ASSIGNMENT_SELECT,
		options,
		ASSIGNMENT_EXPAND,
	)()) as unknown as RoleAssignmentWithBindings[];
}

function requireUnique(
	context: IExecuteFunctions,
	unique: boolean,
	itemIndex: number,
	action: string,
): void {
	if (!unique) {
		throw new NodeOperationError(
			context.getNode(),
			`The target inherits its permissions. Break inheritance before you ${action}.`,
			{ itemIndex },
		);
	}
}

async function getAll(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const securable = buildSecurable(sp, readTarget.call(this, siteUrl, itemIndex));
	const options = this.getNodeParameter('getAllOptions', itemIndex, {}) as SelectionValues;
	const [unique, assignments] = await Promise.all([
		hasUniqueRoleAssignments(securable),
		getAssignments(securable, options),
	]);
	return assignments.map((assignment) => toAssignmentOutput(assignment, unique, options));
}

async function breakInheritance(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const options = this.getNodeParameter('breakOptions', itemIndex, {}) as {
		copyRoleAssignments?: boolean;
		clearSubscopes?: boolean;
	};
	const copy = options.copyRoleAssignments ?? false;
	const clear = options.clearSubscopes ?? true;
	const securable = buildSecurable(sp, target);
	await securable.breakRoleInheritance(copy, clear);
	return {
		...describeTarget(target),
		hasUniqueRoleAssignments: true,
		copyRoleAssignments: copy,
		clearSubscopes: clear,
	};
}

async function resetInheritance(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	await buildSecurable(sp, target).resetRoleInheritance();
	return { ...describeTarget(target), hasUniqueRoleAssignments: false };
}

async function add(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const options = this.getNodeParameter('addOptions', itemIndex, {}) as {
		breakInheritanceFirst?: boolean;
	};
	const principalId = await getPrincipalId.call(this, sp, itemIndex);
	const roleDefinitionId = await getRoleId.call(this, sp, itemIndex);

	let unique = await hasUniqueRoleAssignments(securable);
	let inheritanceBroken = false;
	if (!unique && (options.breakInheritanceFirst ?? true)) {
		await securable.breakRoleInheritance(true, false);
		unique = true;
		inheritanceBroken = true;
	}
	requireUnique(this, unique, itemIndex, 'add a role assignment');

	await securable.roleAssignments.add(principalId, roleDefinitionId);
	return { ...describeTarget(target), principalId, roleDefinitionId, inheritanceBroken };
}

async function remove(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const principalId = await getPrincipalId.call(this, sp, itemIndex, false);
	const role = this.getNodeParameter('role', itemIndex, {}) as { mode: string; value: string };
	const hasRole = String(role?.value ?? '').trim() !== '';
	const roleDefinitionId = hasRole ? await getRoleId.call(this, sp, itemIndex) : undefined;

	requireUnique(
		this,
		await hasUniqueRoleAssignments(securable),
		itemIndex,
		'remove a role assignment',
	);

	if (roleDefinitionId === undefined) {
		await securable.roleAssignments.getById(principalId).delete();
	} else {
		await securable.roleAssignments.remove(principalId, roleDefinitionId);
	}
	return {
		...describeTarget(target),
		principalId,
		...(roleDefinitionId === undefined ? { allRolesRemoved: true } : { roleDefinitionId }),
	};
}

async function set(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const entries =
		(
			this.getNodeParameter('assignments', itemIndex, {}) as {
				entries?: Array<{ principal: string; role: string }>;
			}
		).entries ?? [];
	const options = this.getNodeParameter('setOptions', itemIndex, {}) as {
		copyCurrentFirst?: boolean;
	};

	const desired: RoleAssignmentEntry[] = [];
	for (const entry of entries) {
		const principalId = await resolvePrincipalReference(
			this.getNode(),
			sp,
			parsePrincipalOrThrow(this.getNode(), entry.principal ?? '', itemIndex),
			itemIndex,
		);
		const roleDefinitionId = await resolveRoleId(this.getNode(), sp, entry.role ?? '', itemIndex);
		desired.push({ principalId, roleDefinitionId });
	}

	let inheritanceBroken = false;
	if (!(await hasUniqueRoleAssignments(securable))) {
		await securable.breakRoleInheritance(options.copyCurrentFirst ?? false, false);
		inheritanceBroken = true;
	}

	const current = flattenAssignments(await getAssignments(securable));
	const { toAdd, toRemove } = diffAssignments(current, desired);

	const [batchedSp, execute] = sp.batched();
	const batched = buildSecurable(batchedSp as SPFI, target);
	const pending: Array<Promise<void>> = [];
	for (const entry of toAdd) {
		pending.push(batched.roleAssignments.add(entry.principalId, entry.roleDefinitionId));
	}
	for (const entry of toRemove) {
		pending.push(batched.roleAssignments.remove(entry.principalId, entry.roleDefinitionId));
	}
	if (pending.length) {
		await execute();
		await Promise.all(pending);
	}

	return {
		...describeTarget(target),
		inheritanceBroken,
		added: toAdd.length,
		removed: toRemove.length,
		unchanged: current.length - toRemove.length,
	};
}

async function getEffectivePermissions(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const identifier = String(this.getNodeParameter('effectiveUser', itemIndex, '') ?? '').trim();
	if (!identifier) {
		throw new NodeOperationError(this.getNode(), 'The user must not be empty', { itemIndex });
	}
	const userId = await findUserId(sp, identifier);
	const user = await sp.web.siteUsers.getById(userId).select('Id', 'LoginName', 'Title')();
	const mask = await buildSecurable(sp, target).getUserEffectivePermissions(
		user.LoginName ?? toLoginName(identifier),
	);
	return {
		...describeTarget(target),
		user: { id: user.Id, title: user.Title, loginName: user.LoginName },
		low: mask.Low,
		high: mask.High,
		permissions: decodePermissions(Number(mask.Low), Number(mask.High)),
	};
}

export const permissionResource: ResourceModule = {
	option: {
		name: 'Permission',
		value: RESOURCE,
		description: 'Manage role assignments and permission inheritance',
	},
	description,
	operations: {
		getAll,
		breakInheritance,
		resetInheritance,
		add,
		remove,
		set,
		getEffectivePermissions,
	},
};
