import type { INode } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { ensureFolder, undefinedIfNotFound } from '../../shared/filesFolders';
import {
	buildSecurable,
	dedupeEntries,
	diffAssignments,
	flattenAssignments,
	hasUniqueRoleAssignments,
	parsePrincipalOrThrow,
	resolvePrincipalReference,
	resolveRoleId,
	type RoleAssignmentEntry,
	type TargetDescriptor,
} from '../../shared/securables';
import { applyAssignmentChanges, getAssignments } from './helpers';
import {
	parseJsonArray,
	parseRoleAssignmentEntry,
	type ParseOptions,
	type PrincipalType,
	type RoleAssignmentInput,
} from './json';

export const WILDCARD = '*';
/** Matches folders at any depth. At the end of a path it matches all folders below. */
export const DEEP_WILDCARD = '**';
export const MAPPINGS_JSON_PARAMETER = 'Mappings (JSON)';

const isWildcard = (segment: string) => segment === WILDCARD || segment === DEEP_WILDCARD;

export interface MappingInput extends RoleAssignmentInput {
	/** The folder path relative to the base folder. A segment can be "*" or "**". */
	path: string;
}

export interface MappingOptions {
	mode: 'add' | 'set';
	breakInheritance: 'always' | 'ifInheriting';
	/** Set mode only. Add mode always copies the current assignments when it breaks inheritance. */
	copyCurrentFirst: boolean;
	clearSubscopes: boolean;
	createMissingFolders: boolean;
	ifFolderMissing: 'skip' | 'fail';
	skipUnresolved: boolean;
}

export interface ChildFolder {
	name: string;
	serverRelativeUrl: string;
}

/** The SharePoint calls of Apply Mapping. Unit tests use a fake. */
export interface MappingGateway {
	/** Returns the folder, or undefined when it does not exist. */
	getFolder(path: string): Promise<{ serverRelativeUrl: string } | undefined>;
	/** Returns the child folders without system folders, or undefined when the folder does not exist. */
	listChildFolders(path: string): Promise<ChildFolder[] | undefined>;
	/** Creates the folder and its missing parents. Returns false when it exists. */
	ensureFolder(path: string): Promise<boolean>;
	resolvePrincipal(principal: string, type?: PrincipalType): Promise<number>;
	resolveRole(role: string | number): Promise<number>;
	hasUniqueRoleAssignments(path: string): Promise<boolean>;
	breakInheritance(path: string, copy: boolean, clearSubscopes: boolean): Promise<void>;
	getAssignments(path: string): Promise<RoleAssignmentEntry[]>;
	applyChanges(
		path: string,
		toAdd: RoleAssignmentEntry[],
		toRemove: RoleAssignmentEntry[],
	): Promise<void>;
}

export interface FolderResult {
	/** The path relative to the base folder. "." is the base folder. */
	path: string;
	serverRelativeUrl: string;
	/** The mapping paths that matched this folder. */
	matchedFrom: string[];
	inheritanceBroken: boolean;
	added: number;
	removed: number;
	unchanged: number;
	created: boolean;
}

export interface MissingFolder {
	path: string;
	serverRelativeUrl: string;
	matchedFrom: string[];
}

export interface MappingResult {
	folders: FolderResult[];
	warnings: string[];
	missing: MissingFolder[];
}

const JSON_OPTIONS: ParseOptions = {
	rolesOptional: false,
	allowEmpty: false,
	parameter: MAPPINGS_JSON_PARAMETER,
	extraKeys: ['path'],
};

/** Reads and checks the Mappings (JSON) parameter. Accepts JSON text or an array. */
export function parseMappingsJson(raw: unknown): MappingInput[] {
	return parseJsonArray(raw, JSON_OPTIONS).map((value, index) => {
		const entry = parseRoleAssignmentEntry(value, index, JSON_OPTIONS);
		const path = (value as { path?: unknown }).path;
		if (typeof path !== 'string' || path.trim() === '') {
			throw new Error(
				`${MAPPINGS_JSON_PARAMETER}: entry ${index + 1} needs a "path", for example "Contracts" or "Reports/*/Drafts".`,
			);
		}
		return { ...entry, path: path.trim() };
	});
}

/**
 * Splits a mapping path into segments. "." segments are ignored, so "." is the base folder.
 * "*" and "**" must be whole segments. Repeated "**" segments are read as one.
 */
export function parseMappingPath(path: string, label: string): string[] {
	if (path.trim() === '') throw new Error(`${label} needs a path`);
	const segments = path
		.replace(/\\/g, '/')
		.split('/')
		.map((segment) => segment.trim())
		.filter((segment) => segment !== '' && segment !== '.');
	for (const segment of segments) {
		if (segment === '..') throw new Error(`${label}: the path "${path}" must not contain ".."`);
		if (!isWildcard(segment) && segment.includes(WILDCARD)) {
			throw new Error(
				`${label}: in "${path}", "*" and "**" must be whole segments, for example "Reports/*/Drafts" or "Reports/**/Drafts"`,
			);
		}
	}
	return segments.filter(
		(segment, index) => !(segment === DEEP_WILDCARD && segments[index - 1] === DEEP_WILDCARD),
	);
}

const lower = (value: string) => value.toLowerCase();

function relativePath(base: string, serverRelativeUrl: string, segments: string[]): string {
	const prefix = `${lower(base)}/`;
	if (lower(serverRelativeUrl).startsWith(prefix)) return serverRelativeUrl.slice(prefix.length);
	return segments.length ? segments.join('/') : '.';
}

interface ResolvedMapping {
	index: number;
	path: string;
	segments: string[];
	entries: RoleAssignmentEntry[];
}

interface FolderGroup {
	path: string;
	serverRelativeUrl: string;
	matchedFrom: string[];
	created: boolean;
	desired: RoleAssignmentEntry[];
}

function describeError(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

/** Resolves the principals and roles of all mappings once. Each name is resolved only one time. */
async function resolveMappings(
	gateway: MappingGateway,
	mappings: Array<MappingInput & { segments: string[] }>,
	options: MappingOptions,
	warnings: string[],
): Promise<ResolvedMapping[]> {
	const principals = new Map<string, Promise<number>>();
	const roles = new Map<string, Promise<number>>();
	const cached = (
		cache: Map<string, Promise<number>>,
		key: string,
		load: () => Promise<number>,
	) => {
		let value = cache.get(key);
		if (!value) {
			value = load();
			cache.set(key, value);
		}
		return value;
	};

	const resolved: ResolvedMapping[] = [];
	for (const [index, mapping] of mappings.entries()) {
		let subject = mapping.principal;
		let failure: string | undefined;
		try {
			const principalId = await cached(
				principals,
				`${mapping.principalType ?? ''}|${lower(mapping.principal)}`,
				async () => await gateway.resolvePrincipal(mapping.principal, mapping.principalType),
			);
			const entries: RoleAssignmentEntry[] = [];
			for (const role of mapping.roles) {
				subject = `${mapping.principal}, role ${role}`;
				const roleDefinitionId = await cached(
					roles,
					lower(String(role)),
					async () => await gateway.resolveRole(role),
				);
				entries.push({ principalId, roleDefinitionId });
			}
			resolved.push({ index, path: mapping.path, segments: mapping.segments, entries });
		} catch (error) {
			failure = `Mapping ${index + 1} "${subject}" was not found or cannot be used (${describeError(error)})`;
		}
		if (failure) {
			// The handler turns this into a NodeOperationError with the item index.
			if (!options.skipUnresolved) throw new Error(failure);
			warnings.push(`${failure}. The mapping was skipped.`);
		}
	}
	return resolved;
}

/** Applies a permission mapping to the folders below a base folder. */
export async function applyPermissionMapping(
	gateway: MappingGateway,
	baseFolder: string,
	mappings: MappingInput[],
	options: MappingOptions,
): Promise<MappingResult> {
	const warnings: string[] = [];
	const parsed = mappings.map((mapping, index) => ({
		...mapping,
		segments: parseMappingPath(mapping.path, `Mapping ${index + 1}`),
	}));
	const resolved = await resolveMappings(gateway, parsed, options, warnings);

	const childCache = new Map<string, Promise<ChildFolder[] | undefined>>();
	const listChildren = (path: string) => {
		let children = childCache.get(lower(path));
		if (!children) {
			children = gateway.listChildFolders(path);
			childCache.set(lower(path), children);
		}
		return children;
	};
	const folderCache = new Map<string, { serverRelativeUrl: string } | undefined>();
	const getFolder = async (path: string) => {
		const key = lower(path);
		if (!folderCache.has(key)) folderCache.set(key, await gateway.getFolder(path));
		return folderCache.get(key);
	};

	const groups = new Map<string, FolderGroup>();
	const missing = new Map<string, MissingFolder>();
	const addMatch = (
		serverRelativeUrl: string,
		segments: string[],
		mapping: ResolvedMapping,
		created: boolean,
	) => {
		const key = lower(serverRelativeUrl);
		let group = groups.get(key);
		if (!group) {
			group = {
				path: relativePath(baseFolder, serverRelativeUrl, segments),
				serverRelativeUrl,
				matchedFrom: [],
				created: false,
				desired: [],
			};
			groups.set(key, group);
		}
		if (!group.matchedFrom.includes(mapping.path)) group.matchedFrom.push(mapping.path);
		group.created ||= created;
		group.desired.push(...mapping.entries);
	};

	const listDescendants = async (path: string) => {
		const found: string[] = [];
		let level = [path];
		while (level.length) {
			const next: string[] = [];
			for (const folder of level) {
				for (const child of (await listChildren(folder)) ?? []) next.push(child.serverRelativeUrl);
			}
			found.push(...next);
			level = next;
		}
		return found;
	};

	for (const mapping of resolved) {
		const label = `Mapping ${mapping.index + 1}`;
		if (!mapping.segments.some(isWildcard)) {
			const target = [baseFolder, ...mapping.segments].join('/');
			const folder = await getFolder(target);
			if (folder) {
				addMatch(folder.serverRelativeUrl, mapping.segments, mapping, false);
			} else if (options.createMissingFolders) {
				const created = await gateway.ensureFolder(target);
				folderCache.set(lower(target), { serverRelativeUrl: target });
				addMatch(target, mapping.segments, mapping, created);
			} else if (options.ifFolderMissing === 'fail') {
				throw new Error(
					`${label}: the folder "${mapping.path}" does not exist. Turn on Create Missing Folders, or set If Folder Is Missing to Skip.`,
				);
			} else {
				const key = lower(target);
				const entry = missing.get(key) ?? {
					path: mapping.segments.join('/') || '.',
					serverRelativeUrl: target,
					matchedFrom: [],
				};
				if (!entry.matchedFrom.includes(mapping.path)) entry.matchedFrom.push(mapping.path);
				missing.set(key, entry);
			}
			continue;
		}

		let candidates = [baseFolder];
		let afterWildcard = false;
		for (const [index, segment] of mapping.segments.entries()) {
			const next: string[] = [];
			if (segment === DEEP_WILDCARD) {
				afterWildcard = true;
				// In the middle "**" also matches zero folders, so "A/**/B" includes "A/B".
				const last = index === mapping.segments.length - 1;
				for (const candidate of candidates) {
					if (!last) next.push(candidate);
					next.push(...(await listDescendants(candidate)));
				}
			} else if (segment === WILDCARD) {
				afterWildcard = true;
				for (const candidate of candidates) {
					for (const child of (await listChildren(candidate)) ?? []) {
						next.push(child.serverRelativeUrl);
					}
				}
			} else if (!afterWildcard) {
				next.push(...candidates.map((candidate) => `${candidate}/${segment}`));
			} else {
				for (const candidate of candidates) {
					const match = ((await listChildren(candidate)) ?? []).find(
						(child) => lower(child.name) === lower(segment),
					);
					if (match) next.push(match.serverRelativeUrl);
				}
			}
			candidates = [...new Map(next.map((path) => [lower(path), path])).values()];
			if (!candidates.length) break;
		}
		if (!candidates.length) {
			warnings.push(`${label}: the path "${mapping.path}" matched no folders.`);
		}
		for (const candidate of candidates) addMatch(candidate, mapping.segments, mapping, false);
	}

	const folders: FolderResult[] = [];
	for (const group of groups.values()) {
		folders.push(await applyToFolder(gateway, group, options));
	}
	return { folders, warnings, missing: [...missing.values()] };
}

async function applyToFolder(
	gateway: MappingGateway,
	group: FolderGroup,
	options: MappingOptions,
): Promise<FolderResult> {
	const path = group.serverRelativeUrl;
	const copy = options.mode === 'add' ? true : options.copyCurrentFirst;
	let inheritanceBroken = false;
	let current: RoleAssignmentEntry[];
	if (options.breakInheritance === 'always') {
		await gateway.breakInheritance(path, copy, options.clearSubscopes);
		inheritanceBroken = true;
		current = await gateway.getAssignments(path);
	} else {
		const [unique, assignments] = await Promise.all([
			gateway.hasUniqueRoleAssignments(path),
			gateway.getAssignments(path),
		]);
		if (unique) {
			current = assignments;
		} else {
			await gateway.breakInheritance(path, copy, options.clearSubscopes);
			inheritanceBroken = true;
			current = await gateway.getAssignments(path);
		}
	}

	const desired = dedupeEntries(group.desired);
	const diff = diffAssignments(current, desired);
	const toRemove = options.mode === 'set' ? diff.toRemove : [];
	if (diff.toAdd.length || toRemove.length) {
		await gateway.applyChanges(path, diff.toAdd, toRemove);
	}
	return {
		path: group.path,
		serverRelativeUrl: group.serverRelativeUrl,
		matchedFrom: group.matchedFrom,
		inheritanceBroken,
		added: diff.toAdd.length,
		removed: toRemove.length,
		unchanged: desired.length - diff.toAdd.length,
		created: group.created,
	};
}

interface RawFolder {
	Name?: string;
	ServerRelativeUrl?: string;
	Exists?: boolean;
	ListItemAllFields?: { Id?: number };
}

/** Connects Apply Mapping to SharePoint through PnPjs. */
export function createMappingGateway(
	sp: SPFI,
	siteUrl: string,
	node: INode,
	itemIndex: number,
): MappingGateway {
	const target = (path: string): TargetDescriptor => ({ kind: 'folder', path });
	const securable = (path: string) => buildSecurable(sp, target(path));
	return {
		async getFolder(path) {
			const folder = (await undefinedIfNotFound(
				sp.web.getFolderByServerRelativePath(path).select('Exists', 'ServerRelativeUrl')(),
			)) as RawFolder | undefined;
			return folder?.Exists ? { serverRelativeUrl: String(folder.ServerRelativeUrl) } : undefined;
		},
		async listChildFolders(path) {
			const folders = (await undefinedIfNotFound(
				sp.web
					.getFolderByServerRelativePath(path)
					.folders.select('Name', 'ServerRelativeUrl', 'ListItemAllFields/Id')
					.expand('ListItemAllFields')(),
			)) as RawFolder[] | undefined;
			// Folders without a list item are system folders, such as "Forms" in a library root.
			return folders
				?.filter((folder) => folder.ListItemAllFields?.Id)
				.map((folder) => ({
					name: String(folder.Name ?? ''),
					serverRelativeUrl: String(folder.ServerRelativeUrl ?? ''),
				}));
		},
		async ensureFolder(path) {
			return await ensureFolder(sp, siteUrl, path);
		},
		async resolvePrincipal(principal, type) {
			return await resolvePrincipalReference(
				node,
				sp,
				parsePrincipalOrThrow(node, principal, itemIndex, type),
				itemIndex,
				true,
			);
		},
		async resolveRole(role) {
			return await resolveRoleId(node, sp, role, itemIndex);
		},
		async hasUniqueRoleAssignments(path) {
			return await hasUniqueRoleAssignments(securable(path));
		},
		async breakInheritance(path, copy, clearSubscopes) {
			await securable(path).breakRoleInheritance(copy, clearSubscopes);
		},
		async getAssignments(path) {
			return flattenAssignments(await getAssignments(securable(path)));
		},
		async applyChanges(path, toAdd, toRemove) {
			await applyAssignmentChanges(sp, target(path), toAdd, toRemove);
		},
	};
}
