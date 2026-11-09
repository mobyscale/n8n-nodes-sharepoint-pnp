import type { IDataObject, INodeProperties } from 'n8n-workflow';
import type { IMoveCopyOptions, SPFI } from '@pnp/sp/index.js';

import { pathRLC } from './descriptions';
import { toList, type SelectionValues } from './selection';
import { getSiteServerRelativeUrl, isNotFoundError, splitPath, trimTrailingSlash } from './utils';

export type IfExists = 'fail' | 'replace' | 'keepBoth';

export interface MoveCopyFields {
	newName?: string;
	ifExists?: IfExists;
	bypassSharedLock?: boolean;
	retainEditorAndModified?: boolean;
	resetAuthorAndCreated?: boolean;
}

/**
 * A file path picker. In "From List" mode the search text works as a path, like the folder picker,
 * and the results have the folders and files of the parent path.
 */
export function filePathRLC(
	overrides: Partial<INodeProperties> & Pick<INodeProperties, 'displayName' | 'name'>,
): INodeProperties {
	return { ...pathRLC(overrides, 'searchFiles'), default: { mode: 'path', value: '' } };
}

export function buildMoveCopyOptions(
	kind: 'move' | 'copy',
	fields: MoveCopyFields,
): { overwrite: boolean; options: Partial<IMoveCopyOptions> } {
	const ifExists = fields.ifExists ?? 'fail';
	const options: Partial<IMoveCopyOptions> = {
		KeepBoth: ifExists === 'keepBoth',
		ShouldBypassSharedLocks: fields.bypassSharedLock ?? false,
	};
	if (kind === 'move')
		options.RetainEditorAndModifiedOnMove = fields.retainEditorAndModified ?? false;
	else options.ResetAuthorAndCreatedOnCopy = fields.resetAuthorAndCreated ?? false;
	return { overwrite: ifExists === 'replace', options };
}

export function getDestinationPath(
	sourcePath: string,
	destinationFolder: string,
	newName?: string,
): string {
	const name = newName?.trim() || splitPath(sourcePath).name;
	return trimTrailingSlash(`${trimTrailingSlash(destinationFolder)}/${name}`);
}

/** Returns the paths to check when looking for the closest existing parent, deepest first. */
export function getAncestorPaths(path: string, siteUrl: string): string[] {
	const root = getSiteServerRelativeUrl(siteUrl);
	const result: string[] = [];
	let current = trimTrailingSlash(path);
	while (current && current !== '/' && current.length > root.length && current.startsWith(root)) {
		result.push(current);
		current = splitPath(current).parent;
	}
	return result;
}

/** Returns the missing folder paths, shallowest first, below the given existing ancestor. */
export function getMissingSegments(path: string, existingAncestor: string): string[] {
	const trimmed = trimTrailingSlash(path);
	const base = trimTrailingSlash(existingAncestor);
	const rest = trimmed.slice(base.length).split('/').filter(Boolean);
	return rest.map(
		(_, index) => `${base === '/' ? '' : base}/${rest.slice(0, index + 1).join('/')}`,
	);
}

export async function folderExists(sp: SPFI, path: string): Promise<boolean> {
	const folder = await undefinedIfNotFound(
		sp.web.getFolderByServerRelativePath(path).select('Exists')(),
	);
	return folder?.Exists ?? false;
}

/**
 * Creates the folder and the parent folders that are missing. Returns false when the folder exists.
 * The first folder level must be a library that exists.
 */
export async function ensureFolder(sp: SPFI, siteUrl: string, path: string): Promise<boolean> {
	if (await folderExists(sp, path)) return false;
	let existingAncestor: string | undefined;
	for (const candidate of getAncestorPaths(path, siteUrl).slice(1)) {
		if (await folderExists(sp, candidate)) {
			existingAncestor = candidate;
			break;
		}
	}
	if (!existingAncestor) {
		throw new Error(
			`No parent of "${path}" exists. The first folder level must be a library that exists.`,
		);
	}
	for (const missing of getMissingSegments(path, existingAncestor)) {
		await sp.web.folders.addUsingPath(missing, false);
	}
	return true;
}

/** Picks the list whose root folder is the longest parent of the path (or the path itself). */
export function findOwningList<T extends { rootUrl: string }>(
	path: string,
	lists: T[],
): T | undefined {
	const target = trimTrailingSlash(path).toLowerCase();
	let best: T | undefined;
	for (const list of lists) {
		const root = trimTrailingSlash(list.rootUrl).toLowerCase();
		if (target === root || target.startsWith(`${root}/`)) {
			if (!best || root.length > trimTrailingSlash(best.rootUrl).length) best = list;
		}
	}
	return best;
}

export interface ListRoot {
	id: string;
	rootUrl: string;
}

/** Reads the ID and the root folder URL of each list of the web. */
export async function getListRoots(sp: SPFI): Promise<ListRoot[]> {
	const lists = await sp.web.lists
		.select('Id', 'RootFolder/ServerRelativeUrl')
		.expand('RootFolder')();
	return lists.map((list) => ({
		id: list.Id as string,
		rootUrl: (list as unknown as { RootFolder: { ServerRelativeUrl: string } }).RootFolder
			.ServerRelativeUrl,
	}));
}

export type IncludeKind = 'all' | 'files' | 'folders';

export interface FileFolderFilters {
	include: IncludeKind;
	nameContains?: string;
	modifiedAfter?: string;
	modifiedBefore?: string;
	/** A CAML condition (for example from Field Filters) that the entries must also match. */
	fieldCondition?: string;
}

export function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

/** Joins CAML conditions with nested binary And or Or elements. */
export function combineConditions(operator: 'And' | 'Or', conditions: string[]): string {
	if (conditions.length === 1) return conditions[0];
	const [first, ...rest] = conditions;
	return rest.reduce((left, right) => `<${operator}>${left}${right}</${operator}>`, first);
}

function toIsoDate(value: string): string {
	const date = new Date(value);
	if (Number.isNaN(date.getTime())) throw new Error(`"${value}" is not a valid date`);
	return date.toISOString();
}

/** The list item fields that every Get Many entry loads. They become name, type, size and so on. */
export const VIEW_FIELDS = [
	'FileLeafRef',
	'FileRef',
	'FSObjType',
	'UniqueId',
	'ID',
	'File_x0020_Size',
	'Created',
	'Modified',
	'Author',
	'Editor',
];

/** Builds the CAML for one folder level. Recursion is done by the caller, one request per folder. */
export function buildViewXml(
	filters: FileFolderFilters,
	pageSize = 500,
	extraFields: string[] = [],
): string {
	const conditions: string[] = [];
	if (filters.include === 'files') {
		conditions.push('<Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq>');
	} else if (filters.include === 'folders') {
		conditions.push('<Eq><FieldRef Name="FSObjType"/><Value Type="Integer">1</Value></Eq>');
	}
	if (filters.nameContains) {
		conditions.push(
			`<Contains><FieldRef Name="FileLeafRef"/><Value Type="Text">${escapeXml(filters.nameContains)}</Value></Contains>`,
		);
	}
	if (filters.modifiedAfter) {
		conditions.push(
			`<Gt><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">${toIsoDate(filters.modifiedAfter)}</Value></Gt>`,
		);
	}
	if (filters.modifiedBefore) {
		conditions.push(
			`<Lt><FieldRef Name="Modified"/><Value Type="DateTime" IncludeTimeValue="TRUE">${toIsoDate(filters.modifiedBefore)}</Value></Lt>`,
		);
	}
	if (filters.fieldCondition) conditions.push(filters.fieldCondition);
	const where = conditions.length
		? `<Query><Where>${combineConditions('And', conditions)}</Where></Query>`
		: '';
	const viewFields = `<ViewFields>${toList([...VIEW_FIELDS, ...extraFields])
		.map((name) => `<FieldRef Name="${name}"/>`)
		.join('')}</ViewFields>`;
	return `<View Scope="All">${viewFields}${where}<RowLimit Paged="TRUE">${pageSize}</RowLimit></View>`;
}

export function parseExtensions(value: string | undefined): string[] {
	return (value ?? '')
		.split(/[,;\s]+/)
		.map((entry) =>
			entry
				.trim()
				.replace(/^\*?\./, '')
				.toLowerCase(),
		)
		.filter(Boolean);
}

export function hasExtension(name: string, extensions: string[]): boolean {
	if (!extensions.length) return true;
	const index = name.lastIndexOf('.');
	if (index < 0) return false;
	return extensions.includes(name.slice(index + 1).toLowerCase());
}

interface PersonField {
	title?: string;
	email?: string;
	id?: string;
}

function toPerson(value: unknown): IDataObject | undefined {
	const person = (Array.isArray(value) ? value[0] : value) as PersonField | undefined;
	if (!person || typeof person !== 'object') return undefined;
	return { id: person.id ?? null, name: person.title ?? null, email: person.email ?? null };
}

function asNumber(value: unknown): number | undefined {
	if (value === undefined || value === null || value === '') return undefined;
	const number = Number(value);
	return Number.isFinite(number) ? number : undefined;
}

function pickRaw(row: IDataObject, field: string): unknown {
	return row[`${field}.`] ?? row[field];
}

/** Returns the entries of a row that belong to the given fields. A date field also has a raw "Field." entry. */
export function pickFields(row: IDataObject, fields: string[]): IDataObject {
	const wanted = new Set(fields);
	return Object.fromEntries(
		Object.entries(row).filter(([key]) => wanted.has(key.replace(/\.$/, ''))),
	);
}

export function normalizeRow(
	row: IDataObject,
	includeFields: boolean,
	extraFields: string[] = [],
): IDataObject {
	const isFolder = String(row.FSObjType) === '1';
	const entry: IDataObject = {
		name: String(row.FileLeafRef ?? ''),
		type: isFolder ? 'folder' : 'file',
		serverRelativeUrl: String(row.FileRef ?? ''),
		uniqueId: String(row.UniqueId ?? '').replace(/[{}]/g, ''),
		listItemId: asNumber(row.ID) ?? null,
		size: isFolder ? null : (asNumber(row.File_x0020_Size) ?? null),
		created: (pickRaw(row, 'Created') as string | undefined) ?? null,
		modified: (pickRaw(row, 'Modified') as string | undefined) ?? null,
		author: toPerson(row.Author) ?? null,
		editor: toPerson(row.Editor) ?? null,
	};
	if (includeFields) entry.fields = row;
	else if (extraFields.length) entry.fields = pickFields(row, extraFields);
	return entry;
}

/** Resolves to undefined when SharePoint answers 404. Other errors reach the node router. */
export async function undefinedIfNotFound<T>(request: Promise<T>): Promise<T | undefined> {
	return await request.catch((error: unknown) =>
		isNotFoundError(error) ? undefined : Promise.reject(error),
	);
}

/**
 * "Include List Item Fields" is the same as adding ListItemAllFields to Expand. It is kept for
 * workflows that were made before the selection options existed.
 */
export function withListItemFields(
	options: SelectionValues & { includeListItem?: boolean },
): SelectionValues {
	const { includeListItem, ...values } = options;
	if (!includeListItem) return values;
	return {
		...values,
		expand: [...toList(values.expand), 'ListItemAllFields'],
	};
}
