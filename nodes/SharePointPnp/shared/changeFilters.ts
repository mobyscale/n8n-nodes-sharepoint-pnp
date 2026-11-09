import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { trimTrailingSlash } from './utils';

export type ItemKind = 'file' | 'folder' | 'item';

export const ITEM_KINDS: ItemKind[] = ['file', 'folder', 'item'];

export type FieldFilterOperator = 'equals' | 'notEquals' | 'contains' | 'isEmpty' | 'isNotEmpty';

export interface FieldFilter {
	field: string;
	operator: FieldFilterOperator;
	value: string;
}

const OPERATORS: FieldFilterOperator[] = [
	'equals',
	'notEquals',
	'contains',
	'isEmpty',
	'isNotEmpty',
];

/** Reads the "Field Filters" fixedCollection value. Rows without a field name are ignored. */
export function parseFieldFilters(value: unknown): FieldFilter[] {
	const rows = (value as { filter?: unknown } | undefined)?.filter;
	if (!Array.isArray(rows)) return [];
	return rows
		.map((row: IDataObject) => ({
			field: String(row?.field ?? '').trim(),
			operator: (OPERATORS.includes(row?.operator as FieldFilterOperator)
				? row.operator
				: 'equals') as FieldFilterOperator,
			value: String(row?.value ?? ''),
		}))
		.filter((row) => row.field);
}

export function parseItemKinds(value: unknown): ItemKind[] {
	const parts = Array.isArray(value) ? value : String(value ?? '').split(',');
	return [
		...new Set(
			parts
				.map((part) => String(part).trim().toLowerCase())
				.filter((part): part is ItemKind => (ITEM_KINDS as string[]).includes(part)),
		),
	];
}

/** Splits a comma-separated list into lowercase entries. */
export function parseEditorList(value: unknown): string[] {
	const parts = Array.isArray(value) ? value : String(value ?? '').split(',');
	return [...new Set(parts.map((part) => String(part).trim().toLowerCase()).filter(Boolean))];
}

export function parseMaxDepth(value: unknown): number | undefined {
	if (value === undefined || value === null || value === '') return undefined;
	const depth = Math.floor(Number(value));
	return Number.isFinite(depth) && depth >= 1 ? depth : undefined;
}

function getValue(values: IDataObject, path: string): unknown {
	let value: unknown = values;
	if (path in values) return values[path];
	for (const part of path.split('/')) {
		if (value === null || value === undefined || typeof value !== 'object') return undefined;
		value = (value as IDataObject)[part];
	}
	return value;
}

function isEmptyValue(value: unknown): boolean {
	if (value === undefined || value === null || value === '') return true;
	if (Array.isArray(value)) return value.every(isEmptyValue);
	if (typeof value === 'object') {
		const results = (value as { results?: unknown }).results;
		if (Array.isArray(results)) return results.every(isEmptyValue);
	}
	return false;
}

/** Returns the scalar values of a field value: a multi-value field has one value per entry. */
function toScalars(value: unknown): string[] {
	if (value === undefined || value === null) return [];
	if (Array.isArray(value)) return value.flatMap(toScalars);
	if (typeof value === 'object') {
		const object = value as IDataObject;
		if (Array.isArray(object.results)) return toScalars(object.results);
		const scalar = object.Label ?? object.Title ?? object.Url ?? object.Description;
		return scalar === undefined ? [JSON.stringify(value).toLowerCase()] : toScalars(scalar);
	}
	return [String(value).trim().toLowerCase()];
}

function isEqual(actual: string, expected: string): boolean {
	if (actual === expected) return true;
	const a = Number(actual);
	const b = Number(expected);
	if (actual !== '' && expected !== '' && Number.isFinite(a) && Number.isFinite(b)) return a === b;
	const dateA = Date.parse(actual);
	const dateB = Date.parse(expected);
	return /\d{4}-\d{2}-\d{2}/.test(expected) && Number.isFinite(dateA) && dateA === dateB;
}

/**
 * Compares one field filter with the field values of an item. Text compares are not case-sensitive.
 * A multi-value field matches Equals and Contains when one of its values matches.
 */
export function matchesFieldFilter(values: IDataObject, filter: FieldFilter): boolean {
	const value = getValue(values, filter.field);
	switch (filter.operator) {
		case 'isEmpty':
			return isEmptyValue(value);
		case 'isNotEmpty':
			return !isEmptyValue(value);
		case 'contains': {
			const expected = filter.value.trim().toLowerCase();
			return toScalars(value).some((scalar) => scalar.includes(expected));
		}
		case 'notEquals': {
			const expected = filter.value.trim().toLowerCase();
			return !toScalars(value).some((scalar) => isEqual(scalar, expected));
		}
		default: {
			const expected = filter.value.trim().toLowerCase();
			const scalars = toScalars(value);
			if (!scalars.length) return expected === '';
			return scalars.some((scalar) => isEqual(scalar, expected));
		}
	}
}

/**
 * All filters must match. Without field values (deleted items, File and Folder changes, changes
 * without item details) no filter matches.
 */
export function matchesFieldFilters(
	values: IDataObject | undefined,
	filters: FieldFilter[],
): boolean {
	if (!filters.length) return true;
	if (!values) return false;
	return filters.every((filter) => matchesFieldFilter(values, filter));
}

/**
 * The number of folder levels between a root folder and a path. A child of the root is level 1.
 * Returns undefined when the path is not below the root.
 */
export function getDepthBelow(path: string, root: string): number | undefined {
	const base = trimTrailingSlash(root).toLowerCase();
	const value = trimTrailingSlash(path).toLowerCase();
	if (value === base) return 0;
	const prefix = `${base === '/' ? '' : base}/`;
	if (!value.startsWith(prefix)) return undefined;
	return value.slice(prefix.length).split('/').length;
}

/** Changes without a path, or with a path outside the root, are kept. */
export function isWithinDepth(
	path: string | undefined,
	root: string | undefined,
	maxDepth: number,
): boolean {
	if (!path || !root) return true;
	const depth = getDepthBelow(path, root);
	return depth === undefined || depth <= maxDepth;
}

export interface SiteUserEntry {
	Title?: string;
	Email?: string;
	LoginName?: string;
	UserPrincipalName?: string | null;
}

export interface EditorInfo {
	email?: string;
	login?: string;
	userPrincipalName?: string;
}

export interface EditorSource {
	/** Returns the site users with this display name. */
	findByDisplayName(name: string): Promise<SiteUserEntry[]>;
}

export function createEditorSource(sp: SPFI): EditorSource {
	return {
		findByDisplayName: async (name) =>
			(await sp.web.siteUsers
				.filter(`Title eq '${name.replace(/'/g, "''")}'`)
				.select('Title', 'Email', 'LoginName', 'UserPrincipalName')()) as SiteUserEntry[],
	};
}

/** Removes the claims prefix of a login name: "i:0#.f|membership|a@contoso.com" -> "a@contoso.com". */
export function stripClaims(login: string): string {
	const index = login.lastIndexOf('|');
	return index >= 0 ? login.slice(index + 1) : login;
}

/**
 * Finds the site user of a change editor by display name. It reads each name once per run. It only
 * returns a user when exactly one site user has that display name, because names are not unique.
 */
export class EditorDirectory {
	private readonly cache = new Map<string, Promise<EditorInfo | undefined>>();

	constructor(private readonly source: EditorSource) {}

	async resolve(name: unknown): Promise<EditorInfo | undefined> {
		const displayName = String(name ?? '').trim();
		if (!displayName) return undefined;
		const key = displayName.toLowerCase();
		let info = this.cache.get(key);
		if (!info) {
			info = this.source.findByDisplayName(displayName).then(
				(users) => {
					const matches = users.filter((user) => String(user.Title ?? '').toLowerCase() === key);
					if (matches.length !== 1) return undefined;
					const [user] = matches;
					return {
						email: user.Email || undefined,
						login: user.LoginName || undefined,
						userPrincipalName: user.UserPrincipalName || undefined,
					};
				},
				() => undefined,
			);
			this.cache.set(key, info);
		}
		return await info;
	}
}

function editorHint(change: IDataObject): string | undefined {
	const hint = String(change.EditorEmailHint ?? '').trim();
	return hint || undefined;
}

/** True when an exclude entry can only match a login or email, not a display name. */
function needsLogin(entries: string[]): boolean {
	return entries.some((entry) => entry.includes('@') || entry.includes('|'));
}

/**
 * Exclude Editors: matches the editor display name, the EditorEmailHint of the change, or the
 * email and login of the site user with that display name. Not case-sensitive.
 */
export async function isEditorExcluded(
	change: IDataObject,
	excluded: string[],
	directory: EditorDirectory | undefined,
): Promise<boolean> {
	if (!excluded.length) return false;
	const candidates = [change.Editor, editorHint(change)]
		.filter((value) => value !== undefined && value !== null && value !== '')
		.map((value) => String(value).trim().toLowerCase());
	if (candidates.some((value) => excluded.includes(value))) return true;
	if (!directory || !needsLogin(excluded)) return false;
	const info = await directory.resolve(change.Editor);
	if (!info) return false;
	const values = [
		info.email,
		info.login,
		info.login && stripClaims(info.login),
		info.userPrincipalName,
	]
		.filter((value): value is string => !!value)
		.map((value) => value.toLowerCase());
	return values.some((value) => excluded.includes(value));
}

/** The editor email: the EditorEmailHint of the change, or the email of the matching site user. */
export async function getEditorEmail(
	change: IDataObject,
	directory: EditorDirectory | undefined,
): Promise<string | undefined> {
	const hint = editorHint(change);
	if (hint) return hint;
	if (!directory) return undefined;
	return (await directory.resolve(change.Editor))?.email;
}
