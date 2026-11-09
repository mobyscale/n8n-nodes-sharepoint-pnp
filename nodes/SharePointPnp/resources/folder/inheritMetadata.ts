import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { undefinedIfNotFound } from '../../shared/filesFolders';
import { splitPath, trimTrailingSlash } from '../../shared/utils';
import {
	isEmptyValue,
	readCurrentValue,
	toValueList,
	unknownFields,
	type FieldValues,
} from './fieldValues';
import {
	prepareFields,
	setMetadataRecursive,
	type MetadataOptions,
	type SetMetadataSource,
} from './setMetadata';

export interface InheritOptions extends MetadataOptions {
	/** Walk up from the parent until a folder has a value for each field. */
	searchUp?: boolean;
}

/** Reads the field values of one folder in the format that the writer accepts. */
export type FolderValueReader = (folderPath: string) => Promise<FieldValues | undefined>;

export interface InheritSource extends SetMetadataSource {
	libraryRoot: string;
	readValues: FolderValueReader;
}

export interface InheritedValues {
	parentFolder: string;
	/** The folder that gave the values. With values from more than one folder, the nearest one. */
	sourceFolder: string | null;
	values: FieldValues;
	valueSources: Record<string, string | null>;
	missingFields: string[];
}

/** Field types whose values the node cannot read back in a format that it can write. */
const UNSUPPORTED_TYPES = new Set([
	'TaxonomyFieldType',
	'TaxonomyFieldTypeMulti',
	'Calculated',
	'Computed',
]);

/**
 * Reads the picked fields and the "Other Fields To Inherit" text: names separated by commas or line
 * breaks, or arrays. Older workflows have only the text.
 */
export function parseFieldNames(...raw: unknown[]): string[] {
	const names = [...new Set(raw.flatMap(toValueList))];
	if (!names.length) throw new Error('Add at least one field to inherit');
	return names;
}

function isBelow(path: string, root: string): boolean {
	return trimTrailingSlash(path).toLowerCase().startsWith(`${root.toLowerCase()}/`);
}

/**
 * Finds the values to inherit. The search starts at the parent of the folder. Without "search up"
 * it reads only the parent. With "search up" it goes up one folder at a time until each field has a
 * value, and stops below the library root (a library root folder has no fields). A field without a
 * value is null and is in `missingFields`.
 */
export async function findInheritedValues(
	folderPath: string,
	libraryRoot: string,
	fieldNames: string[],
	searchUp: boolean,
	read: FolderValueReader,
): Promise<InheritedValues> {
	const root = trimTrailingSlash(libraryRoot);
	if (!isBelow(folderPath, root)) {
		throw new Error(
			`The folder "${folderPath}" is a library root folder. It has no parent folder.`,
		);
	}
	const parentFolder = splitPath(folderPath).parent;
	if (!isBelow(parentFolder, root)) {
		throw new Error(
			`The parent of "${folderPath}" is the library root folder, which has no fields to inherit`,
		);
	}
	const found = new Map<string, { value: unknown; folder: string }>();
	let current = parentFolder;
	do {
		const values = (await read(current)) ?? {};
		for (const name of fieldNames) {
			if (!found.has(name) && !isEmptyValue(values[name])) {
				found.set(name, { value: values[name], folder: current });
			}
		}
		current = splitPath(current).parent;
	} while (searchUp && found.size < fieldNames.length && isBelow(current, root));

	const values: FieldValues = {};
	const valueSources: Record<string, string | null> = {};
	for (const name of fieldNames) {
		const entry = found.get(name);
		values[name] = entry ? entry.value : null;
		valueSources[name] = entry ? entry.folder : null;
	}
	const sources = [...found.values()].map((entry) => entry.folder);
	return {
		parentFolder,
		sourceFolder: sources.length
			? sources.reduce((nearest, folder) => (folder.length > nearest.length ? folder : nearest))
			: null,
		values,
		valueSources,
		missingFields: fieldNames.filter((name) => !found.has(name)),
	};
}

function checkFields(fieldNames: string[], types: Map<string, string>): void {
	const missing = unknownFields(fieldNames, types);
	if (missing.length) {
		throw new Error(
			`These fields do not exist in the library: ${missing.join(', ')}. Use the internal field names.`,
		);
	}
	const unsupported = fieldNames.filter((name) => UNSUPPORTED_TYPES.has(types.get(name) ?? ''));
	if (unsupported.length) {
		throw new Error(
			`These fields cannot be inherited: ${unsupported.join(', ')}. Managed metadata, calculated and computed fields are not supported.`,
		);
	}
}

/**
 * Reads the values of the fields from the parent folder (or the nearest folder above it with values)
 * and sets them on the folder and on all items below it. "Include The Folder Itself" is on unless the
 * options turn it off.
 */
export async function inheritMetadataFromParent(
	folderPath: string,
	fieldNames: string[],
	options: InheritOptions,
	source: InheritSource,
): Promise<IDataObject> {
	checkFields(fieldNames, source.listFields.types);
	const inherited = await findInheritedValues(
		folderPath,
		source.libraryRoot,
		fieldNames,
		options.searchUp === true,
		source.readValues,
	);
	const summary = {
		folderPath,
		parentFolder: inherited.parentFolder,
		sourceFolder: inherited.sourceFolder,
		values: inherited.values,
		valueSources: inherited.valueSources,
		missingFields: inherited.missingFields,
	};
	if (!Object.keys(prepareFields(inherited.values, options.ignoreEmptyValues === true)).length) {
		return { ...summary, matched: 0, updated: 0, skipped: 0, failed: 0, errors: [] };
	}
	const result = await setMetadataRecursive(
		folderPath,
		inherited.values,
		{ ...options, includeFolderItself: options.includeFolderItself ?? true },
		source,
	);
	return { ...summary, ...result };
}

const LOOKUP_TYPES = new Set(['Lookup', 'LookupMulti']);
const USER_TYPES = new Set(['User', 'UserMulti']);

function listOf(value: unknown): unknown[] {
	const results = (value as { results?: unknown } | null)?.results;
	if (Array.isArray(results)) return results;
	if (Array.isArray(value)) return value;
	return value === undefined || value === null ? [] : [value];
}

function userEntry(value: unknown): { Key: string } | undefined {
	const name = (value as { Name?: unknown } | null)?.Name;
	return typeof name === 'string' && name ? { Key: name } : undefined;
}

/**
 * Converts a value that the REST API returned for a list item into the format of the writer:
 * lookup IDs, { Key: login name } for persons, "URL, description" for links.
 */
export function toWritableValue(
	item: Record<string, unknown>,
	name: string,
	type: string | undefined,
): unknown {
	if (type === 'User') return userEntry(item[name]) ?? null;
	if (type === 'UserMulti') {
		return listOf(item[name])
			.map(userEntry)
			.filter((entry) => entry !== undefined);
	}
	if (type === 'URL') {
		const link = item[name] as { Url?: string; Description?: string } | null | undefined;
		if (!link?.Url) return null;
		return link.Description ? `${link.Url}, ${link.Description}` : link.Url;
	}
	const value = readCurrentValue(item, name, type);
	return value === undefined ? null : value;
}

/** Reads the fields of a folder list item. Lookups by ID, persons by login name. */
export function createFolderValueReader(
	sp: SPFI,
	fieldNames: string[],
	types: Map<string, string>,
): FolderValueReader {
	const selects = ['Id'];
	const expands: string[] = [];
	for (const name of fieldNames) {
		const type = types.get(name) ?? '';
		if (LOOKUP_TYPES.has(type)) selects.push(`${name}Id`);
		else if (USER_TYPES.has(type)) {
			selects.push(`${name}/Name`);
			expands.push(name);
		} else selects.push(name);
	}
	return async (folderPath) => {
		let query = sp.web
			.getFolderByServerRelativePath(folderPath)
			.listItemAllFields.select(...selects);
		if (expands.length) query = query.expand(...expands);
		const item = (await undefinedIfNotFound(query())) as Record<string, unknown> | undefined;
		const id = Number(item?.Id);
		if (!item || !Number.isFinite(id) || id <= 0) return undefined;
		return Object.fromEntries(
			fieldNames.map((name) => [name, toWritableValue(item, name, types.get(name))]),
		);
	};
}
