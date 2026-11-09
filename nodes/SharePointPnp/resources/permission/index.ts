import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import {
	buildSecurable,
	decodePermissions,
	dedupeEntries,
	diffAssignments,
	findUserId,
	flattenAssignments,
	getPrincipalId,
	getRoleId,
	hasUniqueRoleAssignments,
	parsePrincipalOrThrow,
	principalProperties,
	principalRLC,
	readPrincipalInput,
	readTarget,
	resolvePrincipalReference,
	resolveRoleId,
	roleRLC,
	targetProperties,
	toLoginName,
	userRLC,
	type RoleAssignmentEntry,
	type TargetDescriptor,
} from '../../shared/securables';
import { toNodeError } from '../../shared/utils';
import { selectionOptions, type SelectionValues } from '../../shared/selection';
import type { OperationContext, ResourceModule } from '../types';
import {
	applyAssignmentChanges,
	ASSIGNMENT_SELECT,
	getAssignments,
	toAssignmentOutput,
} from './helpers';
import {
	mergeByPrincipal,
	parseRoleAssignmentsJson,
	type ParseOptions,
	type ResolvedAssignment,
	type RoleAssignmentInput,
} from './json';
import { locatorValue } from '../../shared/pickers';
import { getPath } from '../../shared/descriptions';
import {
	applyPermissionMapping,
	createMappingGateway,
	parseMappingsJson,
	type MappingInput,
	type MappingOptions,
	type MappingResult,
} from './mapping';
import { mappingProperties } from './mappingDescription';

const RESOURCE = 'permission';
const show = (operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [RESOURCE], operation: operations, ...extra },
});

const ROLE_OPERATIONS = ['add', 'remove', 'set'];

/** Shows a property only in Fields mode. Operations without a saved Input Mode use Fields. */
const inFieldsMode = (property: INodeProperties): INodeProperties => ({
	...property,
	displayOptions: {
		...property.displayOptions,
		show: { ...property.displayOptions?.show, inputMode: ['fields'] },
	},
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
				name: 'Add',
				value: 'add',
				description: 'Give a group or user a role on a target',
				action: 'Add a role assignment',
			},
			{
				name: 'Apply Mapping',
				value: 'applyMapping',
				description:
					'Apply a map of paths, principals and roles to many folders below a base folder',
				action: 'Apply a permission mapping',
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
	{
		displayName: 'Input Mode',
		name: 'inputMode',
		type: 'options',
		default: 'fields',
		description: 'How to give the principals and roles',
		options: [
			{ name: 'Fields', value: 'fields', description: 'Use the fields of this node' },
			{
				name: 'JSON',
				value: 'json',
				description: 'Give a list of principals and roles as JSON, to change many at once',
			},
		],
		displayOptions: show([...ROLE_OPERATIONS, 'applyMapping']),
	},
	{
		displayName: 'Role Assignments (JSON)',
		name: 'roleAssignmentsJson',
		type: 'json',
		required: true,
		default: '[]',
		description:
			'An array of objects with "principal" and "role" (or "roles"). "principalType" is optional. See the documentation for an example.',
		displayOptions: show(ROLE_OPERATIONS, { inputMode: ['json'] }),
	},
	...principalProperties(RESOURCE, ['add', 'remove']).map(inFieldsMode),
	inFieldsMode({ ...roleRLC, displayOptions: show(['add']) }),
	inFieldsMode({
		...roleRLC,
		required: false,
		description: 'The role to remove. Leave empty to remove all roles of the principal.',
		default: { mode: 'list', value: '' },
		displayOptions: show(['remove']),
	}),
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
			{
				displayName: 'Clear Subscopes',
				name: 'clearSubscopes',
				type: 'boolean',
				default: false,
				description:
					'Whether to remove unique permissions of child objects when this operation breaks inheritance',
			},
			{
				displayName: 'Skip Unresolved Principals',
				name: 'skipUnresolved',
				type: 'boolean',
				default: false,
				description:
					'Whether to skip an entry whose principal or role cannot be found, and report it in "warnings", instead of failing the operation',
			},
		],
	},
	inFieldsMode({
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
					{ ...principalRLC, required: false },
					{
						...roleRLC,
						required: false,
						description: 'A permission level of the site',
						typeOptions: { loadOptionsDependsOn: ['site.value'] },
					},
				],
			},
		],
	}),
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
			{
				displayName: 'Always Break Inheritance',
				name: 'alwaysBreakInheritance',
				type: 'boolean',
				default: false,
				description:
					'Whether to call break inheritance again when the target already has unique permissions. SharePoint can keep the existing assignments in that case, so the operation still adds and removes entries to match the list.',
			},
			{
				displayName: 'Clear Subscopes',
				name: 'clearSubscopes',
				type: 'boolean',
				default: false,
				description:
					'Whether to remove unique permissions of child objects when this operation breaks inheritance',
			},
			{
				displayName: 'Skip Unresolved Principals',
				name: 'skipUnresolved',
				type: 'boolean',
				default: false,
				description:
					'Whether to skip an entry whose principal or role cannot be found, and report it in "warnings", instead of failing the operation',
			},
		],
	},
	{
		...userRLC,
		name: 'effectiveUser',
		displayOptions: show(['getEffectivePermissions']),
	},
	...mappingProperties,
];

function describeTarget(target: TargetDescriptor): IDataObject {
	return {
		target: target.kind,
		...(target.path ? { path: target.path } : {}),
		...(target.itemId ? { itemId: target.itemId } : {}),
	};
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

const FIELD_MODE = 'fields';

function readInputMode(this: IExecuteFunctions, itemIndex: number): string {
	return this.getNodeParameter('inputMode', itemIndex, FIELD_MODE) as string;
}

function readJsonEntries(
	this: IExecuteFunctions,
	itemIndex: number,
	options: ParseOptions,
): RoleAssignmentInput[] {
	try {
		return parseRoleAssignmentsJson(
			this.getNodeParameter('roleAssignmentsJson', itemIndex, '[]'),
			options,
		);
	} catch (error) {
		throw new NodeOperationError(this.getNode(), (error as Error).message, { itemIndex });
	}
}

function wrapEntryError(
	context: IExecuteFunctions,
	error: unknown,
	position: number,
	subject: string,
	itemIndex: number,
): Error {
	if (error instanceof NodeOperationError) return error;
	const reason = error instanceof Error ? error.message : String(error);
	return new NodeOperationError(
		context.getNode(),
		`Role Assignments (JSON): entry ${position + 1} "${subject}" was not found or cannot be used (${reason})`,
		{ itemIndex },
	);
}

/** Resolves names and emails to IDs. Nothing is changed in SharePoint before every entry resolves. */
export async function resolveEntries(
	context: IExecuteFunctions,
	sp: SPFI,
	entries: RoleAssignmentInput[],
	itemIndex: number,
	ensureUsers: boolean,
	warnings?: string[],
): Promise<ResolvedAssignment[]> {
	const node = context.getNode();
	const resolved: ResolvedAssignment[] = [];
	for (const [position, entry] of entries.entries()) {
		try {
			let principalId: number;
			try {
				principalId = await resolvePrincipalReference(
					node,
					sp,
					parsePrincipalOrThrow(node, entry.principal, itemIndex, entry.principalType),
					itemIndex,
					ensureUsers,
				);
			} catch (error) {
				throw wrapEntryError(context, error, position, entry.principal, itemIndex);
			}
			const roleDefinitionIds: number[] = [];
			for (const role of entry.roles) {
				try {
					roleDefinitionIds.push(await resolveRoleId(node, sp, role, itemIndex));
				} catch (error) {
					throw wrapEntryError(
						context,
						error,
						position,
						`${entry.principal}, role ${role}`,
						itemIndex,
					);
				}
			}
			resolved.push({ principalId, roleDefinitionIds });
		} catch (error) {
			if (!warnings) throw toNodeError(node, error, itemIndex);
			warnings.push(`Entry ${position + 1} was skipped: ${(error as Error).message}`);
		}
	}
	return resolved;
}

const withWarnings = (result: IDataObject, warnings: string[] | undefined): IDataObject =>
	warnings ? { ...result, warnings } : result;

interface AddOptions {
	breakInheritanceFirst?: boolean;
	clearSubscopes?: boolean;
	skipUnresolved?: boolean;
}

function readAddOptions(this: IExecuteFunctions, itemIndex: number): AddOptions {
	return this.getNodeParameter('addOptions', itemIndex, {}) as AddOptions;
}

async function addFromJson(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const options = readAddOptions.call(this, itemIndex);
	const warnings = options.skipUnresolved ? [] : undefined;
	const entries = readJsonEntries.call(this, itemIndex, {
		rolesOptional: false,
		allowEmpty: false,
	});
	const resolved = await resolveEntries(this, sp, entries, itemIndex, true, warnings);
	const desired = dedupeEntries(
		resolved.flatMap((entry) =>
			entry.roleDefinitionIds.map((roleDefinitionId) => ({
				principalId: entry.principalId,
				roleDefinitionId,
			})),
		),
	);

	if (desired.length === 0 && warnings) {
		return withWarnings(
			{ ...describeTarget(target), inheritanceBroken: false, added: 0, assignments: [] },
			warnings,
		);
	}

	let unique = await hasUniqueRoleAssignments(securable);
	let inheritanceBroken = false;
	if (!unique && (options.breakInheritanceFirst ?? true)) {
		await securable.breakRoleInheritance(true, options.clearSubscopes ?? false);
		unique = true;
		inheritanceBroken = true;
	}
	requireUnique(this, unique, itemIndex, 'add a role assignment');

	for (const entry of desired) {
		await securable.roleAssignments.add(entry.principalId, entry.roleDefinitionId);
	}
	return withWarnings(
		{
			...describeTarget(target),
			inheritanceBroken,
			added: desired.length,
			assignments: desired as unknown as IDataObject[],
		},
		warnings,
	);
}

async function removeFromJson(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const entries = readJsonEntries.call(this, itemIndex, { rolesOptional: true, allowEmpty: false });
	const resolved = mergeByPrincipal(await resolveEntries(this, sp, entries, itemIndex, false));
	requireUnique(
		this,
		await hasUniqueRoleAssignments(securable),
		itemIndex,
		'remove a role assignment',
	);

	for (const entry of resolved) {
		if (entry.roleDefinitionIds.length === 0) {
			await securable.roleAssignments.getById(entry.principalId).delete();
		} else {
			for (const roleDefinitionId of entry.roleDefinitionIds) {
				await securable.roleAssignments.remove(entry.principalId, roleDefinitionId);
			}
		}
	}
	return {
		...describeTarget(target),
		removed: resolved.map((entry) =>
			entry.roleDefinitionIds.length
				? { principalId: entry.principalId, roleDefinitionIds: entry.roleDefinitionIds }
				: { principalId: entry.principalId, allRolesRemoved: true },
		) as IDataObject[],
	};
}

async function add(this: IExecuteFunctions, context: OperationContext) {
	if (readInputMode.call(this, context.itemIndex) === 'json') {
		return await addFromJson.call(this, context);
	}
	const { sp, siteUrl, itemIndex } = context;
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const options = readAddOptions.call(this, itemIndex);
	let principalId: number;
	let roleDefinitionId: number;
	try {
		principalId = await getPrincipalId.call(this, sp, itemIndex);
		roleDefinitionId = await getRoleId.call(this, sp, itemIndex);
	} catch (error) {
		if (!options.skipUnresolved) throw toNodeError(this.getNode(), error, itemIndex);
		return {
			...describeTarget(target),
			inheritanceBroken: false,
			added: 0,
			warnings: [`The role assignment was skipped: ${(error as Error).message}`],
		};
	}

	let unique = await hasUniqueRoleAssignments(securable);
	let inheritanceBroken = false;
	if (!unique && (options.breakInheritanceFirst ?? true)) {
		await securable.breakRoleInheritance(true, options.clearSubscopes ?? false);
		unique = true;
		inheritanceBroken = true;
	}
	requireUnique(this, unique, itemIndex, 'add a role assignment');

	await securable.roleAssignments.add(principalId, roleDefinitionId);
	return withWarnings(
		{ ...describeTarget(target), principalId, roleDefinitionId, inheritanceBroken },
		options.skipUnresolved ? [] : undefined,
	);
}

async function remove(this: IExecuteFunctions, context: OperationContext) {
	if (readInputMode.call(this, context.itemIndex) === 'json') {
		return await removeFromJson.call(this, context);
	}
	const { sp, siteUrl, itemIndex } = context;
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

function readSetEntries(this: IExecuteFunctions, itemIndex: number): RoleAssignmentInput[] {
	if (readInputMode.call(this, itemIndex) === 'json') {
		return readJsonEntries.call(this, itemIndex, { rolesOptional: false, allowEmpty: true });
	}
	const entries =
		(
			this.getNodeParameter('assignments', itemIndex, {}) as {
				entries?: Array<{ principal?: unknown; role?: unknown }>;
			}
		).entries ?? [];
	return entries.map((entry) => {
		const { text, type } = readPrincipalInput(entry.principal);
		return {
			principal: text,
			...(type ? { principalType: type } : {}),
			roles: [locatorValue(entry.role)],
		};
	});
}

async function set(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const securable = buildSecurable(sp, target);
	const entries = readSetEntries.call(this, itemIndex);
	const options = this.getNodeParameter('setOptions', itemIndex, {}) as {
		copyCurrentFirst?: boolean;
		alwaysBreakInheritance?: boolean;
		clearSubscopes?: boolean;
		skipUnresolved?: boolean;
	};
	const warnings = options.skipUnresolved ? [] : undefined;

	const desired: RoleAssignmentEntry[] = [];
	const resolved = await resolveEntries(this, sp, entries, itemIndex, true, warnings);
	for (const entry of resolved) {
		for (const roleDefinitionId of entry.roleDefinitionIds) {
			desired.push({ principalId: entry.principalId, roleDefinitionId });
		}
	}

	let inheritanceBroken = false;
	if (options.alwaysBreakInheritance || !(await hasUniqueRoleAssignments(securable))) {
		await securable.breakRoleInheritance(
			options.copyCurrentFirst ?? false,
			options.clearSubscopes ?? false,
		);
		inheritanceBroken = true;
	}

	const current = flattenAssignments(await getAssignments(securable));
	const { toAdd, toRemove } = diffAssignments(current, desired);

	await applyAssignmentChanges(sp, target, toAdd, toRemove);

	return withWarnings(
		{
			...describeTarget(target),
			inheritanceBroken,
			added: toAdd.length,
			removed: toRemove.length,
			unchanged: current.length - toRemove.length,
		},
		warnings,
	);
}

async function getEffectivePermissions(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = readTarget.call(this, siteUrl, itemIndex);
	const identifier = locatorValue(this.getNodeParameter('effectiveUser', itemIndex, ''));
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

/** Reads the mapping rows or the Mappings (JSON) parameter without calling SharePoint. */
export function readMappings(this: IExecuteFunctions, itemIndex: number): MappingInput[] {
	if (readInputMode.call(this, itemIndex) === 'json') {
		try {
			return parseMappingsJson(this.getNodeParameter('mappingsJson', itemIndex, '[]'));
		} catch (error) {
			throw new NodeOperationError(this.getNode(), (error as Error).message, { itemIndex });
		}
	}
	const rows =
		(
			this.getNodeParameter('mappings', itemIndex, {}) as {
				entries?: Array<{ path?: unknown; principal?: unknown; role?: unknown }>;
			}
		).entries ?? [];
	if (!rows.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one mapping', { itemIndex });
	}
	return rows.map((row) => {
		const { text, type } = readPrincipalInput(row.principal);
		return {
			path: String(row.path ?? '').trim(),
			principal: text,
			...(type ? { principalType: type } : {}),
			roles: [locatorValue(row.role)],
		};
	});
}

interface MappingParameters extends Partial<Omit<MappingOptions, 'mode'>> {
	output?: 'single' | 'perFolder';
}

export function readMappingOptions(
	this: IExecuteFunctions,
	itemIndex: number,
): { options: MappingOptions; output: 'single' | 'perFolder' } {
	const values = this.getNodeParameter('mappingOptions', itemIndex, {}) as MappingParameters;
	return {
		options: {
			mode: this.getNodeParameter('mappingMode', itemIndex, 'set') === 'add' ? 'add' : 'set',
			breakInheritance: values.breakInheritance === 'always' ? 'always' : 'ifInheriting',
			copyCurrentFirst: values.copyCurrentFirst ?? false,
			clearSubscopes: values.clearSubscopes ?? false,
			createMissingFolders: values.createMissingFolders ?? false,
			ifFolderMissing: values.ifFolderMissing === 'fail' ? 'fail' : 'skip',
			skipUnresolved: values.skipUnresolved ?? false,
		},
		output: values.output === 'perFolder' ? 'perFolder' : 'single',
	};
}

/** Builds the output of Apply Mapping: one item, or one item for each folder and missing folder. */
export function toMappingOutput(
	result: MappingResult,
	baseFolder: string,
	mode: MappingOptions['mode'],
	output: 'single' | 'perFolder',
): IDataObject | IDataObject[] {
	const summary = { baseFolder, mode };
	if (output === 'single') {
		return { ...summary, ...result } as unknown as IDataObject;
	}
	const items: IDataObject[] = [
		...result.folders.map((folder) => ({ ...summary, ...folder, missing: false })),
		...result.missing.map((folder) => ({ ...summary, ...folder, missing: true })),
	].map((item) => ({ ...item, warnings: result.warnings }));
	return items.length ? items : { ...summary, ...result };
}

async function applyMapping(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const rawBase = String(
		this.getNodeParameter('baseFolder', itemIndex, '', { extractValue: true }) ?? '',
	).trim();
	if (!rawBase) {
		throw new NodeOperationError(this.getNode(), 'The base folder must not be empty', {
			itemIndex,
		});
	}
	const baseFolder = getPath.call(this, siteUrl, itemIndex, 'baseFolder');
	const mappings = readMappings.call(this, itemIndex);
	const { options, output } = readMappingOptions.call(this, itemIndex);
	const gateway = createMappingGateway(sp, siteUrl, this.getNode(), itemIndex);
	try {
		const result = await applyPermissionMapping(gateway, baseFolder, mappings, options);
		return toMappingOutput(result, baseFolder, options.mode, output);
	} catch (error) {
		throw toNodeError(this.getNode(), error, itemIndex);
	}
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
		applyMapping,
	},
};
