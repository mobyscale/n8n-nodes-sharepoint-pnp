import type { SPFI } from '@pnp/sp/index.js';

import { findOwningList } from '../../shared/filesFolders';
import { formatFieldValue, type FieldInfo } from '../listItem/helpers';

export type FieldValues = Record<string, unknown>;

export interface FieldResult {
	FieldName?: string;
	HasException?: boolean;
	ErrorMessage?: string;
}

export interface KeyValueRow {
	name?: string;
	value?: unknown;
}

const LOOKUP_ID_TYPES = new Set(['Lookup', 'LookupMulti', 'User', 'UserMulti']);
const NUMBER_TYPES = new Set(['Number', 'Currency', 'Integer', 'Counter']);
const MULTI_TYPES = new Set(['MultiChoice', 'LookupMulti', 'UserMulti']);

const INVALID_JSON = Symbol('invalid JSON');

/** Parses JSON text. Returns INVALID_JSON when the text is not valid JSON. */
export function tryParseJson(text: string): unknown {
	try {
		return JSON.parse(text) as unknown;
	} catch {
		return INVALID_JSON;
	}
}

export function isInvalidJson(value: unknown): boolean {
	return value === INVALID_JSON;
}

/** Reads a JSON object parameter. An empty value is an empty object. */
export function parseFieldsObject(raw: unknown, label: string): FieldValues {
	let value = raw;
	if (typeof value === 'string') {
		if (!value.trim()) return {};
		value = tryParseJson(value);
		if (isInvalidJson(value)) throw new Error(`"${label}" is not valid JSON`);
	}
	if (value === undefined || value === null) return {};
	if (typeof value !== 'object' || Array.isArray(value)) {
		throw new Error(`"${label}" must be a JSON object`);
	}
	return value as FieldValues;
}

export function fromKeyValueRows(rows: KeyValueRow[] | undefined): FieldValues {
	const result: FieldValues = {};
	for (const row of rows ?? []) {
		const name = String(row.name ?? '').trim();
		if (name) result[name] = row.value ?? '';
	}
	return result;
}

/** Splits a list of values: a JSON array, or text with commas or line breaks. */
export function toValueList(value: unknown): string[] {
	if (Array.isArray(value)) return value.map(String).filter((entry) => entry !== '');
	if (value === undefined || value === null) return [];
	const text = String(value).trim();
	if (text.startsWith('[')) {
		try {
			const parsed: unknown = JSON.parse(text);
			if (Array.isArray(parsed)) return parsed.map(String).filter((entry) => entry !== '');
		} catch {
			// not JSON, treat as a delimited list
		}
	}
	return text
		.split(/[\r\n,]+/)
		.map((entry) => entry.trim())
		.filter(Boolean);
}

/** The name to put in $select. Lookup and person fields return their IDs as "<Name>Id". */
export function selectNameFor(name: string, type: string | undefined): string {
	return LOOKUP_ID_TYPES.has(type ?? '') ? `${name}Id` : name;
}

export function readCurrentValue(
	item: Record<string, unknown>,
	name: string,
	type: string | undefined,
): unknown {
	const value = item[selectNameFor(name, type)];
	const results = (value as { results?: unknown } | null)?.results;
	return Array.isArray(results) ? results : value;
}

function idOf(value: unknown): string {
	if (value !== null && typeof value === 'object') {
		const entry = value as { Id?: unknown; id?: unknown; LookupId?: unknown };
		return String(entry.Id ?? entry.id ?? entry.LookupId ?? '');
	}
	return String(value);
}

function normalizeForCompare(type: string | undefined, value: unknown): string {
	if (value === undefined || value === null || value === '') return '';
	if (MULTI_TYPES.has(type ?? '')) {
		const list = Array.isArray(value) ? value.map(idOf) : toValueList(value);
		return [...new Set(list)].sort().join(';#');
	}
	if (type === 'Lookup' || type === 'User') return idOf(value);
	if (NUMBER_TYPES.has(type ?? '')) {
		const number = Number(value);
		return Number.isFinite(number) ? String(number) : String(value);
	}
	if (type === 'DateTime') {
		const date = new Date(String(value));
		return Number.isNaN(date.getTime()) ? String(value) : date.toISOString();
	}
	return formatFieldValue(type, value) ?? '';
}

/**
 * Returns true when the current value of a field already has the new value. Text, choice, number,
 * yes/no, date, lookup ID and multi-choice values are compared. Values that cannot be compared are
 * not equal, so they are written.
 */
export function valuesEqual(
	type: string | undefined,
	current: unknown,
	incoming: unknown,
): boolean {
	return normalizeForCompare(type, current) === normalizeForCompare(type, incoming);
}

export function isEmptyValue(value: unknown): boolean {
	if (value === undefined || value === null) return true;
	if (typeof value === 'string') return value.trim() === '';
	if (Array.isArray(value)) return value.length === 0;
	return false;
}

/** Returns a message for the fields that validateUpdateListItem rejected, or undefined. */
export function describeFieldErrors(results: FieldResult[] | undefined): string | undefined {
	const failed = (results ?? []).filter((entry) => entry.HasException);
	if (!failed.length) return undefined;
	return failed
		.map(
			(entry) => `${entry.FieldName ?? 'Unknown field'}: ${entry.ErrorMessage ?? 'invalid value'}`,
		)
		.join('; ');
}

export function unknownFields(names: Iterable<string>, types: Map<string, string>): string[] {
	return [...new Set(names)].filter((name) => !types.has(name));
}

export interface ListFields {
	listId: string;
	infos: FieldInfo[];
	types: Map<string, string>;
	/** The server-relative URL of the root folder of the list. */
	rootUrl?: string;
}

/** Finds the list that contains the path, and reads the internal names and types of its fields. */
export async function loadListFields(sp: SPFI, path: string): Promise<ListFields> {
	const lists = await sp.web.lists
		.select('Id', 'RootFolder/ServerRelativeUrl')
		.expand('RootFolder')();
	const list = findOwningList(
		path,
		lists.map((entry) => ({
			id: entry.Id as string,
			rootUrl: (entry as unknown as { RootFolder: { ServerRelativeUrl: string } }).RootFolder
				.ServerRelativeUrl,
		})),
	);
	if (!list) throw new Error(`No list or library contains "${path}"`);
	const infos = (await sp.web.lists
		.getById(list.id)
		.fields.select('InternalName', 'TypeAsString')()) as FieldInfo[];
	return {
		listId: list.id,
		rootUrl: list.rootUrl,
		infos,
		types: new Map(infos.map((field) => [field.InternalName, field.TypeAsString])),
	};
}
