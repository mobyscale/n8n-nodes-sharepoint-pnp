import type { IDataObject } from 'n8n-workflow';

import { getSiteServerRelativeUrl, joinPath } from '../../shared/utils';
import {
	ensureOneFolder,
	loadFieldTypes,
	requireBaseFolder,
	toAppend,
	type EnsurePathGateway,
	type FolderFieldSpec,
} from './ensurePath';
import { isInvalidJson, parseFieldsObject, tryParseJson } from './fieldValues';

export interface FolderEntrySpec extends FolderFieldSpec {
	key: string;
	/** The folder names below the base folder, from the top level down. */
	segments: string[];
}

/** The folder rows of the UI. JSON fields arrive as text or as objects. */
export interface FolderEntryRow {
	key?: string;
	path?: string;
	fieldsWhenCreated?: unknown;
	fieldsAlways?: unknown;
	appendField?: string;
	appendValues?: string;
}

export interface EnsuredFolderOutput extends IDataObject {
	path: string;
	serverRelativeUrl: string;
	siteRelativeUrl: string;
	absoluteUrl: string;
	listItemId: number | null;
	uniqueId: string | null;
	created: boolean;
	fieldsUpdated: string[];
}

export interface EnsureFoldersResult extends IDataObject {
	folders: Record<string, EnsuredFolderOutput>;
	created: boolean;
	changed: boolean;
	baseFolder: string;
}

/** Splits a relative path into folder names. "\" works as "/", and empty or "." parts are left out. */
export function splitRelativePath(raw: unknown, label: string): string[] {
	const segments = String(raw ?? '')
		.replace(/\\/g, '/')
		.split('/')
		.map((part) => part.trim())
		.filter((part) => part !== '' && part !== '.');
	if (segments.includes('..')) throw new Error(`${label}: the path must not have ".."`);
	if (!segments.length) throw new Error(`${label}: the path is empty`);
	return segments;
}

function checkKey(raw: unknown, index: number): string {
	const key = String(raw ?? '').trim();
	if (!key) throw new Error(`Folder ${index + 1} has no key`);
	return key;
}

/** Checks that the keys are unique and that two keys do not point to the same folder. */
function checkUnique(entries: FolderEntrySpec[]): FolderEntrySpec[] {
	const keys = new Set<string>();
	const paths = new Map<string, string>();
	for (const entry of entries) {
		if (keys.has(entry.key)) throw new Error(`The key "${entry.key}" is used more than once`);
		keys.add(entry.key);
		const path = entry.segments.join('/').toLowerCase();
		const other = paths.get(path);
		if (other !== undefined) {
			throw new Error(`The keys "${other}" and "${entry.key}" have the same path`);
		}
		paths.set(path, entry.key);
	}
	return entries;
}

export function foldersFromRows(rows: FolderEntryRow[] | undefined): FolderEntrySpec[] {
	return checkUnique(
		(rows ?? []).map((row, index) => {
			const key = checkKey(row.key, index);
			return {
				key,
				segments: splitRelativePath(row.path, `Folder "${key}"`),
				fieldsWhenCreated: parseFieldsObject(
					row.fieldsWhenCreated,
					`Folder "${key}": Fields When Created`,
				),
				fieldsAlways: parseFieldsObject(row.fieldsAlways, `Folder "${key}": Fields Always`),
				appendToMultiChoice: toAppend(row.appendField, row.appendValues),
			};
		}),
	);
}

/**
 * Reads the JSON folders: an array of objects with "key", "path", "fieldsWhenCreated",
 * "fieldsAlways" and "appendToMultiChoice" ({ "field": "...", "values": [...] }).
 */
export function foldersFromJson(raw: unknown): FolderEntrySpec[] {
	let value = raw;
	if (typeof value === 'string') {
		value = tryParseJson(value);
		if (isInvalidJson(value)) throw new Error('"Folders (JSON)" is not valid JSON');
	}
	if (!Array.isArray(value)) throw new Error('"Folders (JSON)" must be an array');
	return checkUnique(
		value.map((entry: unknown, index) => {
			if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
				throw new Error(`Folder ${index + 1} must be an object with "key" and "path"`);
			}
			const data = entry as IDataObject;
			const key = checkKey(data.key, index);
			const append = (data.appendToMultiChoice ?? {}) as IDataObject;
			return {
				key,
				segments: splitRelativePath(data.path, `Folder "${key}"`),
				fieldsWhenCreated: parseFieldsObject(
					data.fieldsWhenCreated,
					`Folder "${key}": fieldsWhenCreated`,
				),
				fieldsAlways: parseFieldsObject(data.fieldsAlways, `Folder "${key}": fieldsAlways`),
				appendToMultiChoice: toAppend(append.field, append.values),
			};
		}),
	);
}

export interface FolderNode {
	/** The lower-case relative path. Folder names are not case-sensitive in SharePoint. */
	id: string;
	parentId: string | null;
	name: string;
	/** The relative path with the casing of the first entry that has this folder. */
	path: string;
	entry?: FolderEntrySpec;
}

/**
 * Turns the entries into one list of folders, parents before children. A parent that more than one
 * entry shares is in the list once. The order follows the order of the entries.
 */
export function planFolderTree(entries: FolderEntrySpec[]): FolderNode[] {
	const nodes = new Map<string, FolderNode>();
	for (const entry of entries) {
		let parentId: string | null = null;
		let path = '';
		entry.segments.forEach((name, index) => {
			path = path ? `${path}/${name}` : name;
			const id = (parentId ? `${parentId}/${name}` : name).toLowerCase();
			let node = nodes.get(id);
			if (!node) {
				node = { id, parentId, name, path };
				nodes.set(id, node);
			}
			if (index === entry.segments.length - 1) node.entry = entry;
			parentId = id;
			path = node.path;
		});
	}
	return [...nodes.values()];
}

function encodePath(serverRelativeUrl: string): string {
	return serverRelativeUrl.split('/').map(encodeURIComponent).join('/');
}

/** Returns the site-relative and absolute URLs of a server-relative URL. */
export function folderUrls(
	siteUrl: string,
	serverRelativeUrl: string,
): { siteRelativeUrl: string; absoluteUrl: string } {
	const site = getSiteServerRelativeUrl(siteUrl);
	const prefix = site === '/' ? '' : site;
	const lower = serverRelativeUrl.toLowerCase();
	let siteRelativeUrl = serverRelativeUrl.replace(/^\/+/, '');
	if (lower === prefix.toLowerCase()) siteRelativeUrl = '';
	else if (lower.startsWith(`${prefix.toLowerCase()}/`)) {
		siteRelativeUrl = serverRelativeUrl.slice(prefix.length + 1);
	}
	return {
		siteRelativeUrl,
		absoluteUrl: `${new URL(siteUrl).origin}${encodePath(serverRelativeUrl)}`,
	};
}

/** Creates the folders of all entries below the base folder, parents first, and sets their fields. */
export async function ensureFolders(
	basePath: string,
	entries: FolderEntrySpec[],
	gateway: EnsurePathGateway,
	siteUrl: string,
): Promise<EnsureFoldersResult> {
	if (!entries.length) throw new Error('Add at least one folder');
	const types = await loadFieldTypes(entries, gateway);
	const base = await requireBaseFolder(basePath, gateway);

	const done = new Map<string, { serverRelativeUrl: string; created: boolean }>();
	const folders: Record<string, EnsuredFolderOutput> = {};
	let changed = false;
	for (const node of planFolderTree(entries)) {
		const parent = node.parentId === null ? undefined : done.get(node.parentId);
		const path = joinPath(parent?.serverRelativeUrl ?? base, node.name);
		const { state, created, fieldsUpdated } = await ensureOneFolder(
			path,
			node.entry,
			parent?.created ?? false,
			types,
			gateway,
		);
		done.set(node.id, { serverRelativeUrl: state.serverRelativeUrl, created });
		if (created || fieldsUpdated.length) changed = true;
		if (!node.entry) continue;
		folders[node.entry.key] = {
			path: node.entry.segments.join('/'),
			serverRelativeUrl: state.serverRelativeUrl,
			...folderUrls(siteUrl, state.serverRelativeUrl),
			listItemId: state.listItemId,
			uniqueId: state.uniqueId ?? null,
			created,
			fieldsUpdated,
		};
	}

	// The output keeps the order of the entries, not the order of creation.
	const ordered: Record<string, EnsuredFolderOutput> = {};
	for (const entry of entries) ordered[entry.key] = folders[entry.key];
	return {
		folders: ordered,
		created: [...done.values()].some((entry) => entry.created),
		changed,
		baseFolder: base,
	};
}

/** "One Item Per Folder": one item for each key, with the key and the folder fields. */
export function toPerFolderItems(result: EnsureFoldersResult): IDataObject[] {
	return Object.entries(result.folders).map(([key, folder]) => ({
		key,
		...folder,
		baseFolder: result.baseFolder,
	}));
}
