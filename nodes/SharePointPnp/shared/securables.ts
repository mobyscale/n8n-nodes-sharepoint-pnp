import type { IDataObject, IExecuteFunctions, INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import type { SecurableQueryable } from '@pnp/sp/security/index.js';
import { PermissionKind } from '@pnp/sp/security/index.js';
import { Item } from '@pnp/sp/items/index.js';

import { getListFromLocator, listRLC, pathRLC } from './descriptions';
import { toServerRelativePath } from './utils';
import { locatorValue } from './pickers';

export type TargetKind = 'site' | 'list' | 'folder' | 'file' | 'listItem';

export type PrincipalReference =
	| { kind: 'id'; id: number }
	| { kind: 'user'; loginName: string }
	| { kind: 'group'; name: string };

export interface RoleAssignmentEntry {
	principalId: number;
	roleDefinitionId: number;
}

export interface LocatorValue {
	mode: string;
	value: string;
}

/** Reads the text of a resource locator value. Plain strings, for example from an expression, work too. */
export interface PrincipalInput {
	text: string;
	type?: 'user' | 'group';
}

/** Reads a principal field. The "By Email or Login" mode means a user and the "By Name" mode means a group. */
export function readPrincipalInput(raw: unknown): PrincipalInput {
	const text = locatorValue(raw);
	const mode =
		raw !== null && typeof raw === 'object' ? (raw as { mode?: string }).mode : undefined;
	if (mode === 'user') return { text, type: 'user' };
	if (mode === 'name') return { text, type: 'group' };
	return { text };
}

const TARGET_KINDS: TargetKind[] = ['site', 'list', 'folder', 'file', 'listItem'];

export function isTargetKind(value: string): value is TargetKind {
	return (TARGET_KINDS as string[]).includes(value);
}

/**
 * Turns an email address or a login name into the claims login name that SharePoint expects.
 * Login names that already contain a "|" stay as they are.
 */
export function toLoginName(identifier: string): string {
	const value = identifier.trim();
	if (!value) throw new Error('The user must not be empty');
	if (value.includes('|')) return value;
	if (value.includes('@')) return `i:0#.f|membership|${value}`;
	return value;
}

/**
 * Reads a principal from free text: digits are a principal ID, text with "@" or "|" is a user,
 * and all other text is a SharePoint group name. A principal type decides what text is when
 * the guess would be wrong, for example a user login name without "@" or "|".
 */
export function parsePrincipal(input: string, type?: 'user' | 'group'): PrincipalReference {
	const value = input.trim();
	if (!value) throw new Error('The principal must not be empty');
	if (/^\d+$/.test(value)) return { kind: 'id', id: Number(value) };
	if (type === 'group') return { kind: 'group', name: value };
	if (type === 'user') return { kind: 'user', loginName: toLoginName(value) };
	if (value.includes('@') || value.includes('|')) {
		return { kind: 'user', loginName: toLoginName(value) };
	}
	return { kind: 'group', name: value };
}

export function parsePrincipalOrThrow(
	node: INode,
	input: string,
	itemIndex: number,
	type?: 'user' | 'group',
): PrincipalReference {
	try {
		return parsePrincipal(input, type);
	} catch (error) {
		throw new NodeOperationError(node, (error as Error).message, { itemIndex });
	}
}

export function parseRoleId(input: string | number): number | string {
	const value = String(input).trim();
	if (!value) throw new Error('The role must not be empty');
	return /^\d+$/.test(value) ? Number(value) : value;
}

const pairKey = (entry: RoleAssignmentEntry) => `${entry.principalId}:${entry.roleDefinitionId}`;

export function dedupeEntries(entries: RoleAssignmentEntry[]): RoleAssignmentEntry[] {
	const seen = new Set<string>();
	return entries.filter((entry) => {
		const key = pairKey(entry);
		if (seen.has(key)) return false;
		seen.add(key);
		return true;
	});
}

/** Works out which role assignments to add and to remove so that "current" becomes "desired". */
export function diffAssignments(
	current: RoleAssignmentEntry[],
	desired: RoleAssignmentEntry[],
): { toAdd: RoleAssignmentEntry[]; toRemove: RoleAssignmentEntry[] } {
	const wanted = dedupeEntries(desired);
	const existing = dedupeEntries(current);
	const currentKeys = new Set(existing.map(pairKey));
	const wantedKeys = new Set(wanted.map(pairKey));
	return {
		toAdd: wanted.filter((entry) => !currentKeys.has(pairKey(entry))),
		toRemove: existing.filter((entry) => !wantedKeys.has(pairKey(entry))),
	};
}

export interface RoleAssignmentWithBindings {
	PrincipalId: number;
	RoleDefinitionBindings?: Array<{ Id: number }>;
}

export function flattenAssignments(
	assignments: RoleAssignmentWithBindings[],
): RoleAssignmentEntry[] {
	return assignments.flatMap((assignment) =>
		(assignment.RoleDefinitionBindings ?? []).map((role) => ({
			principalId: assignment.PrincipalId,
			roleDefinitionId: role.Id,
		})),
	);
}

/** Lists the permission names that are set in a SharePoint base permissions mask. */
export function decodePermissions(low: number, high: number): string[] {
	const names: string[] = [];
	for (const key of Object.keys(PermissionKind)) {
		const kind = Number(key);
		if (!Number.isInteger(kind) || kind <= 0 || kind >= 64) continue;
		const bit = kind - 1;
		const word = bit < 32 ? low : high;
		if (((word >>> bit % 32) & 1) === 1) names.push(PermissionKind[kind]);
	}
	return names;
}

const SYSTEM_LOGIN_PATTERNS = [
	/^SHAREPOINT\\system$/i,
	/^c:0\(\.s\|true$/i,
	/^c:0!\.s\|/i,
	/spo-grid-all-users/i,
	/^i:0#\.w\|nt service\\/i,
	/^i:0#\.w\|nt authority\\/i,
	/^nt authority\\/i,
	/^i:0i\.t\|ms\.sp\.ext\|/i,
	/^i:0\(\.w\|/i,
	/@sharepoint$/i,
];

export function isSystemAccount(user: { LoginName?: string; IsHiddenInUI?: boolean }): boolean {
	if (user.IsHiddenInUI) return true;
	const login = user.LoginName ?? '';
	return SYSTEM_LOGIN_PATTERNS.some((pattern) => pattern.test(login));
}

const show = (resource: string, operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [resource], operation: operations, ...extra },
});

export const roleRLC: INodeProperties = {
	displayName: 'Role',
	name: 'role',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'A permission level of the site, such as Full Control, Edit, Contribute or Read',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchRoleDefinitions', searchable: true },
		},
		{ displayName: 'By Name', name: 'name', type: 'string', placeholder: 'Contribute' },
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: '1073741827',
			validation: [
				{ type: 'regex', properties: { regex: '^\\d+$', errorMessage: 'Not a valid ID' } },
			],
		},
	],
};

export const groupRLC: INodeProperties = {
	displayName: 'Group',
	name: 'group',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'A SharePoint group of the site',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchGroups', searchable: true },
		},
		{ displayName: 'By Name', name: 'name', type: 'string', placeholder: 'Example Members' },
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: '5',
			validation: [
				{ type: 'regex', properties: { regex: '^\\d+$', errorMessage: 'Not a valid ID' } },
			],
		},
	],
};

export const userRLC: INodeProperties = {
	displayName: 'User',
	name: 'user',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'The user. Pick a user of the site, or type an email address or a login name.',
	typeOptions: { loadOptionsDependsOn: ['site.value'] },
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchUsers', searchable: true },
		},
		{
			displayName: 'By Email or Login',
			name: 'user',
			type: 'string',
			placeholder: 'user@contoso.com',
			hint: 'An email address or a login name such as i:0#.f|membership|user@contoso.com',
		},
	],
};

/** A group or a user. In "From List" mode the search lists both. */
export const principalRLC: INodeProperties = {
	displayName: 'Principal',
	name: 'principal',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'A group or a user',
	typeOptions: { loadOptionsDependsOn: ['site.value'] },
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchPrincipals', searchable: true },
		},
		{
			displayName: 'By Email or Login',
			name: 'user',
			type: 'string',
			placeholder: 'user@contoso.com',
			hint: 'An email address or a login name such as i:0#.f|membership|user@contoso.com',
		},
		{ displayName: 'By Group Name', name: 'name', type: 'string', placeholder: 'Example Members' },
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: '5',
			validation: [
				{ type: 'regex', properties: { regex: '^\\d+$', errorMessage: 'Not a valid ID' } },
			],
		},
	],
};

export const itemIdRLC: INodeProperties = {
	displayName: 'Item',
	name: 'itemId',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'The list item',
	typeOptions: { loadOptionsDependsOn: ['site.value', 'list.value'] },
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchListItems', searchable: true },
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: '1',
			validation: [
				{ type: 'regex', properties: { regex: '^\\d+$', errorMessage: 'Not a valid ID' } },
			],
		},
	],
};

/** The "Target" picker of the permission operations. */
export function targetProperties(resource: string, operations: string[]): INodeProperties[] {
	return [
		{
			displayName: 'Target',
			name: 'target',
			type: 'options',
			default: 'site',
			description: 'The object whose permissions to work with',
			options: [
				{ name: 'File', value: 'file' },
				{ name: 'Folder', value: 'folder' },
				{ name: 'List Item', value: 'listItem' },
				{ name: 'List or Library', value: 'list' },
				{ name: 'Site', value: 'site' },
			],
			displayOptions: show(resource, operations),
		},
		{
			...listRLC,
			displayOptions: show(resource, operations, { target: ['list', 'listItem'] }),
		},
		{
			...itemIdRLC,
			displayOptions: show(resource, operations, { target: ['listItem'] }),
		},
		pathRLC({
			displayName: 'Folder',
			name: 'folderPath',
			displayOptions: show(resource, operations, { target: ['folder'] }),
		}),
		pathRLC({
			displayName: 'File',
			name: 'filePath',
			modes: [
				{
					displayName: 'By Path',
					name: 'path',
					type: 'string',
					placeholder: 'Shared Documents/Folder/File.docx',
				},
			],
			default: { mode: 'path', value: '' },
			displayOptions: show(resource, operations, { target: ['file'] }),
		}),
	];
}

/** Principal fields: a group from a picker, or a user by email or login name. */
export function principalProperties(
	resource: string,
	operations: string[],
	userOnly = false,
): INodeProperties[] {
	const userField: INodeProperties = {
		...userRLC,
		displayOptions: show(resource, operations, userOnly ? {} : { principalType: ['user'] }),
	};
	if (userOnly) return [userField];
	return [
		{
			displayName: 'Principal Type',
			name: 'principalType',
			type: 'options',
			default: 'group',
			options: [
				{ name: 'Group', value: 'group' },
				{ name: 'User', value: 'user' },
			],
			displayOptions: show(resource, operations),
		},
		{ ...groupRLC, displayOptions: show(resource, operations, { principalType: ['group'] }) },
		userField,
	];
}

function getLocator(this: IExecuteFunctions, name: string, itemIndex: number): LocatorValue {
	return this.getNodeParameter(name, itemIndex) as LocatorValue;
}

export async function resolveGroupId(
	node: INode,
	sp: SPFI,
	locator: LocatorValue,
	itemIndex: number,
): Promise<number> {
	const value = String(locator.value ?? '').trim();
	if (!value) throw new NodeOperationError(node, 'The group must not be empty', { itemIndex });
	if (locator.mode === 'name') {
		return (await sp.web.siteGroups.getByName(value).select('Id')()).Id;
	}
	if (!/^\d+$/.test(value)) {
		throw new NodeOperationError(node, `"${value}" is not a valid group ID`, { itemIndex });
	}
	return Number(value);
}

export async function resolveRoleId(
	node: INode,
	sp: SPFI,
	input: LocatorValue | string | number,
	itemIndex: number,
): Promise<number> {
	let parsed: number | string;
	try {
		if (typeof input === 'object' && input !== null) {
			parsed = input.mode === 'name' ? String(input.value).trim() : parseRoleId(input.value);
			if (parsed === '') throw new Error('The role must not be empty');
		} else {
			parsed = parseRoleId(input);
		}
	} catch (error) {
		throw new NodeOperationError(node, (error as Error).message, { itemIndex });
	}
	if (typeof parsed === 'number') return parsed;
	return (await sp.web.roleDefinitions.getByName(parsed).select('Id')()).Id;
}

/** Finds a user without changing the site. Used by read and remove operations. */
export async function findUserId(sp: SPFI, identifier: string): Promise<number> {
	const value = identifier.trim();
	if (/^\d+$/.test(value)) return Number(value);
	const users = sp.web.siteUsers;
	const user =
		value.includes('@') && !value.includes('|')
			? await users.getByEmail(value).select('Id')()
			: await users.getByLoginName(toLoginName(value)).select('Id')();
	return user.Id;
}

/** Resolves a user and adds it to the site user list when needed. */
export async function ensureUserId(sp: SPFI, identifier: string): Promise<number> {
	const value = identifier.trim();
	if (/^\d+$/.test(value)) return Number(value);
	return (await sp.web.ensureUser(toLoginName(value))).Id;
}

export async function resolvePrincipalReference(
	node: INode,
	sp: SPFI,
	reference: PrincipalReference,
	itemIndex: number,
	ensure = true,
): Promise<number> {
	if (reference.kind === 'id') return reference.id;
	if (reference.kind === 'group') {
		return await resolveGroupId(node, sp, { mode: 'name', value: reference.name }, itemIndex);
	}
	return ensure
		? await ensureUserId(sp, reference.loginName)
		: await findUserId(sp, reference.loginName);
}

export function getUserParameter(this: IExecuteFunctions, itemIndex: number): string {
	const user = locatorValue(this.getNodeParameter('user', itemIndex, ''));
	if (!user) {
		throw new NodeOperationError(this.getNode(), 'The user must not be empty', { itemIndex });
	}
	return user;
}

/** Resolves the principal fields of an operation to a principal ID. */
export async function getPrincipalId(
	this: IExecuteFunctions,
	sp: SPFI,
	itemIndex: number,
	ensure = true,
): Promise<number> {
	const type = this.getNodeParameter('principalType', itemIndex) as string;
	if (type === 'group') {
		return await resolveGroupId(
			this.getNode(),
			sp,
			getLocator.call(this, 'group', itemIndex),
			itemIndex,
		);
	}
	const user = getUserParameter.call(this, itemIndex);
	return ensure ? await ensureUserId(sp, user) : await findUserId(sp, user);
}

export async function getRoleId(
	this: IExecuteFunctions,
	sp: SPFI,
	itemIndex: number,
	parameterName = 'role',
): Promise<number> {
	return await resolveRoleId(
		this.getNode(),
		sp,
		getLocator.call(this, parameterName, itemIndex),
		itemIndex,
	);
}

export interface TargetDescriptor {
	kind: TargetKind;
	path?: string;
	list?: LocatorValue;
	itemId?: number;
}

function readItemId(this: IExecuteFunctions, itemIndex: number): number {
	const text = locatorValue(this.getNodeParameter('itemId', itemIndex, ''));
	if (!/^\d+$/.test(text) || Number(text) < 1) {
		throw new NodeOperationError(this.getNode(), `"${text}" is not a valid list item ID`, {
			itemIndex,
		});
	}
	return Number(text);
}

/** Reads the target parameters without calling SharePoint. */
export function readTarget(
	this: IExecuteFunctions,
	siteUrl: string,
	itemIndex: number,
): TargetDescriptor {
	const kind = this.getNodeParameter('target', itemIndex, 'site') as string;
	if (!isTargetKind(kind)) {
		throw new NodeOperationError(this.getNode(), `The target "${kind}" is not supported`, {
			itemIndex,
		});
	}
	switch (kind) {
		case 'list':
			return { kind, list: getLocator.call(this, 'list', itemIndex) };
		case 'listItem':
			return {
				kind,
				list: getLocator.call(this, 'list', itemIndex),
				itemId: readItemId.call(this, itemIndex),
			};
		case 'folder':
		case 'file': {
			const name = kind === 'folder' ? 'folderPath' : 'filePath';
			const value = String(
				this.getNodeParameter(name, itemIndex, '', { extractValue: true }) ?? '',
			).trim();
			if (!value) {
				throw new NodeOperationError(this.getNode(), `The ${kind} path must not be empty`, {
					itemIndex,
				});
			}
			return { kind, path: toServerRelativePath(siteUrl, value) };
		}
		default:
			return { kind };
	}
}

/** Builds the PnPjs object that holds the permissions of a target. Folders and files use their list item. */
export function buildSecurable(sp: SPFI, target: TargetDescriptor): SecurableQueryable {
	switch (target.kind) {
		case 'list':
			return getListFromLocator(sp, target.list!) as unknown as SecurableQueryable;
		case 'listItem':
			return getListFromLocator(sp, target.list!).items.getById(
				target.itemId!,
			) as unknown as SecurableQueryable;
		case 'folder':
			return Item(
				sp.web.getFolderByServerRelativePath(target.path!).listItemAllFields,
			) as unknown as SecurableQueryable;
		case 'file':
			return Item(
				sp.web.getFileByServerRelativePath(target.path!).listItemAllFields,
			) as unknown as SecurableQueryable;
		default:
			return sp.web as unknown as SecurableQueryable;
	}
}

type Selectable = { select: (...fields: string[]) => () => Promise<IDataObject> };

export async function hasUniqueRoleAssignments(securable: SecurableQueryable): Promise<boolean> {
	const info = (await (securable as unknown as Selectable).select(
		'HasUniqueRoleAssignments',
	)()) as { HasUniqueRoleAssignments?: boolean };
	return info.HasUniqueRoleAssignments === true;
}
