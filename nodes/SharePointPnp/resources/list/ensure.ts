import type { IDataObject, IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import type { IList } from '@pnp/sp/lists/index.js';

import { getListFromLocator } from '../../shared/descriptions';
import { undefinedIfNotFound } from '../../shared/filesFolders';
import { locatorValue } from '../../shared/pickers';
import {
	diffAssignments,
	type LocatorValue,
	type RoleAssignmentEntry,
} from '../../shared/securables';
import {
	getRequiredString,
	toNodeError,
	toServerRelativePath,
	trimTrailingSlash,
} from '../../shared/utils';
import { resolveEntries } from '../permission';
import { parseRoleAssignmentsJson } from '../permission/json';
import type { OperationContext } from '../types';
import { readContentTypeOrder, writeContentTypeOrder } from './setup';
import { buildContentTypeOrder } from './setupHelpers';
import { copyListView, getViewTitle } from './views';

export interface ListSummary {
	id: string;
	title: string;
	serverRelativeUrl: string;
}

export interface ListContentType {
	id: string;
	name: string;
}

/** The SharePoint calls of Ensure. A fake of this interface tests the logic without SharePoint. */
export interface EnsureListGateway {
	findByUrl(serverRelativeUrl: string): Promise<ListSummary | undefined>;
	findByTitle(title: string): Promise<ListSummary | undefined>;
	getSummary(id: string): Promise<ListSummary>;
	/** Creates the list and returns its ID. */
	create(title: string, description: string, template: number): Promise<string>;
	/** Reads list properties by their SharePoint names. */
	getProperties(id: string, names: string[]): Promise<Record<string, unknown>>;
	update(id: string, properties: Record<string, unknown>): Promise<void>;
	getContentTypes(id: string): Promise<ListContentType[]>;
	/** Adds a site content type and returns the name of the new list content type. */
	addContentType(id: string, siteContentTypeId: string): Promise<string>;
	getContentTypeOrder(id: string): Promise<string[]>;
	setContentTypeOrder(id: string, order: string[]): Promise<void>;
	getViewTitles(id: string): Promise<string[]>;
	getSourceViewTitle(sourceList: LocatorValue, view: string): Promise<string>;
	/** Copies the view and returns the view fields that the target does not have. */
	copyView(sourceList: LocatorValue, view: string, targetId: string): Promise<string[]>;
	hasUniqueRoleAssignments(id: string): Promise<boolean>;
	breakInheritance(
		id: string,
		copyRoleAssignments: boolean,
		clearSubscopes: boolean,
	): Promise<void>;
	getRoleAssignments(id: string): Promise<RoleAssignmentEntry[]>;
	addRoleAssignment(id: string, entry: RoleAssignmentEntry): Promise<void>;
	removeRoleAssignment(id: string, entry: RoleAssignmentEntry): Promise<void>;
}

export interface EnsureListSettings {
	onQuickLaunch?: boolean;
	enableVersioning?: boolean;
	majorVersionLimit?: number;
	enableMinorVersions?: boolean;
	minorVersionLimit?: number;
	enableContentTypes?: boolean;
}

export interface EnsureListPermissions {
	breakInheritance: boolean;
	copyRoleAssignments: boolean;
	clearSubscopes: boolean;
	/** "replace" removes the role assignments that are not in the list. "add" keeps them. */
	mode: 'replace' | 'add';
	/** Resolves the role assignments. Called only when the settings are applied. */
	resolveAssignments?: () => Promise<RoleAssignmentEntry[]>;
}

export interface EnsureListSpec {
	title: string;
	/** Server-relative URL of the list root folder. */
	url?: string;
	template: number;
	description?: string;
	applyToExisting: boolean;
	settings: EnsureListSettings;
	contentTypes: string[];
	defaultContentType?: string;
	copyView?: { sourceList: LocatorValue; view: string };
	permissions?: EnsureListPermissions;
}

export interface EnsureListResult {
	id: string;
	title: string;
	serverRelativeUrl: string;
	created: boolean;
	changes: string[];
	warnings?: string[];
}

const SETTING_PROPERTIES: Array<[keyof EnsureListSettings, string]> = [
	['onQuickLaunch', 'OnQuickLaunch'],
	['enableVersioning', 'EnableVersioning'],
	['majorVersionLimit', 'MajorVersionLimit'],
	['enableMinorVersions', 'EnableMinorVersions'],
	['minorVersionLimit', 'MajorWithMinorVersionsLimit'],
];

/** The list properties that the settings ask for. Only set settings are included. */
export function desiredListProperties(
	spec: EnsureListSpec,
	created: boolean,
): Record<string, unknown> {
	const properties: Record<string, unknown> = {};
	for (const [option, property] of SETTING_PROPERTIES) {
		if (spec.settings[option] !== undefined) properties[property] = spec.settings[option];
	}
	if (
		spec.settings.enableContentTypes !== undefined ||
		spec.contentTypes.length ||
		spec.defaultContentType
	) {
		properties.ContentTypesEnabled =
			spec.settings.enableContentTypes === true ||
			spec.contentTypes.length > 0 ||
			Boolean(spec.defaultContentType);
	}
	if (!created && spec.description) properties.Description = spec.description;
	return properties;
}

/** The changed properties only. */
export function diffProperties(
	current: Record<string, unknown>,
	desired: Record<string, unknown>,
): Record<string, unknown> {
	return Object.fromEntries(
		Object.entries(desired).filter(([key, value]) => current[key] !== value),
	);
}

const sameId = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

/**
 * Finds a content type of the list by its list ID, by the ID of the site content type it was made
 * from (list ID = site ID + "00" + 32 hex digits), or by its name.
 */
export function findListContentType(
	listTypes: ListContentType[],
	reference: string,
): ListContentType | undefined {
	const value = reference.trim();
	const lower = value.toLowerCase();
	return (
		listTypes.find((type) => sameId(type.id, value)) ??
		listTypes.find(
			(type) =>
				type.id.length === value.length + 34 && type.id.toLowerCase().startsWith(`${lower}00`),
		) ??
		listTypes.find((type) => type.name.toLowerCase() === lower)
	);
}

function lastSegment(path: string): string {
	const trimmed = trimTrailingSlash(path);
	return trimmed.slice(trimmed.lastIndexOf('/') + 1);
}

async function findList(
	spec: EnsureListSpec,
	gateway: EnsureListGateway,
): Promise<ListSummary | undefined> {
	if (!spec.url) return await gateway.findByTitle(spec.title);
	const byUrl = await gateway.findByUrl(spec.url);
	if (byUrl) return byUrl;
	const byTitle = await gateway.findByTitle(spec.title);
	if (byTitle) {
		throw new Error(
			`A list with the title "${spec.title}" exists at "${byTitle.serverRelativeUrl}", not at "${spec.url}". Change the title or the URL.`,
		);
	}
	return undefined;
}

async function applyContentTypes(
	spec: EnsureListSpec,
	gateway: EnsureListGateway,
	id: string,
	changes: string[],
): Promise<void> {
	if (!spec.contentTypes.length && !spec.defaultContentType) return;
	let listTypes = await gateway.getContentTypes(id);
	for (const siteTypeId of spec.contentTypes) {
		if (findListContentType(listTypes, siteTypeId)) continue;
		const name = await gateway.addContentType(id, siteTypeId);
		changes.push(`Added the content type "${name || siteTypeId}"`);
		listTypes = await gateway.getContentTypes(id);
	}
	if (!spec.defaultContentType) return;

	const target = findListContentType(listTypes, spec.defaultContentType);
	if (!target) {
		throw new Error(
			`The content type "${spec.defaultContentType}" is not in the list. Add it with "Content Types to Add".`,
		);
	}
	const order = await gateway.getContentTypeOrder(id);
	if (order.length && sameId(order[0], target.id)) return;
	await gateway.setContentTypeOrder(id, buildContentTypeOrder(order, target.id));
	changes.push(`Set the default content type to "${target.name}"`);
}

async function applyView(
	spec: EnsureListSpec,
	gateway: EnsureListGateway,
	id: string,
	changes: string[],
	warnings: string[],
): Promise<void> {
	if (!spec.copyView) return;
	const { sourceList, view } = spec.copyView;
	const title = await gateway.getSourceViewTitle(sourceList, view);
	const existing = await gateway.getViewTitles(id);
	if (existing.some((entry) => entry.toLowerCase() === title.toLowerCase())) {
		warnings.push(`The list has a view "${title}". The view was not copied.`);
		return;
	}
	const skipped = await gateway.copyView(sourceList, view, id);
	changes.push(`Copied the view "${title}"`);
	if (skipped.length) {
		warnings.push(
			`The view "${title}" was copied without these fields that the list does not have: ${skipped.join(', ')}`,
		);
	}
}

async function applyPermissions(
	permissions: EnsureListPermissions | undefined,
	assignments: RoleAssignmentEntry[] | undefined,
	gateway: EnsureListGateway,
	id: string,
	changes: string[],
): Promise<void> {
	if (!permissions || (!permissions.breakInheritance && !assignments)) return;
	if (!(await gateway.hasUniqueRoleAssignments(id))) {
		await gateway.breakInheritance(id, permissions.copyRoleAssignments, permissions.clearSubscopes);
		changes.push(
			`Broke permission inheritance (copy role assignments: ${permissions.copyRoleAssignments}, clear subscopes: ${permissions.clearSubscopes})`,
		);
	}
	if (!assignments) return;

	const { toAdd, toRemove } = diffAssignments(await gateway.getRoleAssignments(id), assignments);
	for (const entry of toAdd) {
		await gateway.addRoleAssignment(id, entry);
		changes.push(`Added role ${entry.roleDefinitionId} for principal ${entry.principalId}`);
	}
	if (permissions.mode !== 'replace') return;
	for (const entry of toRemove) {
		await gateway.removeRoleAssignment(id, entry);
		changes.push(`Removed role ${entry.roleDefinitionId} from principal ${entry.principalId}`);
	}
}

/**
 * Finds the list by URL or title. Creates it when it is missing, then applies the settings. An
 * existing list gets the settings only when "applyToExisting" is set.
 */
export async function ensureList(
	spec: EnsureListSpec,
	gateway: EnsureListGateway,
): Promise<EnsureListResult> {
	const found = await findList(spec, gateway);
	if (found && !spec.applyToExisting) return { ...found, created: false, changes: [] };

	// Resolve the principals before the list is created, so that an unknown principal changes nothing.
	const assignments = await spec.permissions?.resolveAssignments?.();

	const changes: string[] = [];
	const warnings: string[] = [];
	let id: string;
	if (found) {
		id = found.id;
	} else {
		const createTitle = spec.url ? lastSegment(spec.url) : spec.title;
		id = await gateway.create(createTitle, spec.description ?? '', spec.template);
		changes.push(`Created the list "${spec.title}" with template ${spec.template}`);
		if (createTitle !== spec.title) await gateway.update(id, { Title: spec.title });
	}

	const desired = desiredListProperties(spec, !found);
	if (Object.keys(desired).length) {
		const changed = diffProperties(await gateway.getProperties(id, Object.keys(desired)), desired);
		if (Object.keys(changed).length) {
			await gateway.update(id, changed);
			for (const [key, value] of Object.entries(changed)) {
				changes.push(`Set ${key} to ${JSON.stringify(value)}`);
			}
		}
	}

	await applyContentTypes(spec, gateway, id, changes);
	await applyView(spec, gateway, id, changes, warnings);
	await applyPermissions(spec.permissions, assignments, gateway, id, changes);

	const summary = await gateway.getSummary(id);
	if (!found && spec.url && summary.serverRelativeUrl.toLowerCase() !== spec.url.toLowerCase()) {
		warnings.push(
			`SharePoint created the list at "${summary.serverRelativeUrl}", not at "${spec.url}". The template sets the parent folder (lists go to "Lists/").`,
		);
	}
	return {
		...summary,
		created: !found,
		changes,
		...(warnings.length ? { warnings } : {}),
	};
}

interface RawList {
	Id: string;
	Title: string;
	RootFolder?: { ServerRelativeUrl?: string };
}

const toSummary = (raw: RawList): ListSummary => ({
	id: raw.Id,
	title: raw.Title,
	serverRelativeUrl: raw.RootFolder?.ServerRelativeUrl ?? '',
});

const readSummary = async (list: IList) =>
	toSummary(
		(await list
			.select('Id', 'Title', 'RootFolder/ServerRelativeUrl')
			.expand('RootFolder')()) as RawList,
	);

export function createEnsureListGateway(sp: SPFI): EnsureListGateway {
	const byId = (id: string) => () => sp.web.lists.getById(id);
	const source = (locator: LocatorValue) => () => getListFromLocator(sp, locator);
	return {
		findByUrl: async (url) => await undefinedIfNotFound(readSummary(sp.web.getList(url))),
		findByTitle: async (title) =>
			await undefinedIfNotFound(readSummary(sp.web.lists.getByTitle(title))),
		getSummary: async (id) => await readSummary(byId(id)()),
		create: async (title, description, template) =>
			(await sp.web.lists.add(title, description, template, false)).Id as string,
		getProperties: async (id, names) =>
			(await byId(id)().select(...names)()) as Record<string, unknown>,
		update: async (id, properties) => {
			await byId(id)().update(properties);
		},
		getContentTypes: async (id) =>
			(
				(await byId(id)().contentTypes.select('StringId', 'Name')()) as Array<{
					StringId: string;
					Name: string;
				}>
			).map((type) => ({ id: type.StringId, name: type.Name })),
		addContentType: async (id, siteContentTypeId) =>
			String(
				(await byId(id)().contentTypes.addAvailableContentType(siteContentTypeId)).data?.Name ?? '',
			),
		getContentTypeOrder: async (id) => await readContentTypeOrder(byId(id)),
		setContentTypeOrder: async (id, order) => await writeContentTypeOrder(byId(id), order),
		getViewTitles: async (id) =>
			((await byId(id)().views.select('Title')()) as Array<{ Title: string }>).map(
				(view) => view.Title,
			),
		getSourceViewTitle: async (locator, view) => await getViewTitle(source(locator), view),
		copyView: async (locator, view, targetId) =>
			(await copyListView(source(locator), view, byId(targetId))).skipped,
		hasUniqueRoleAssignments: async (id) =>
			(
				(await byId(id)().select('HasUniqueRoleAssignments')()) as {
					HasUniqueRoleAssignments?: boolean;
				}
			).HasUniqueRoleAssignments === true,
		breakInheritance: async (id, copy, clear) => {
			await byId(id)().breakRoleInheritance(copy, clear);
		},
		getRoleAssignments: async (id) =>
			(
				(await byId(id)()
					.roleAssignments.select('PrincipalId', 'RoleDefinitionBindings/Id')
					.expand('RoleDefinitionBindings')()) as Array<{
					PrincipalId: number;
					RoleDefinitionBindings?: Array<{ Id: number }>;
				}>
			).flatMap((assignment) =>
				(assignment.RoleDefinitionBindings ?? []).map((role) => ({
					principalId: assignment.PrincipalId,
					roleDefinitionId: role.Id,
				})),
			),
		addRoleAssignment: async (id, entry) => {
			await byId(id)().roleAssignments.add(entry.principalId, entry.roleDefinitionId);
		},
		removeRoleAssignment: async (id, entry) => {
			await byId(id)().roleAssignments.remove(entry.principalId, entry.roleDefinitionId);
		},
	};
}

interface EnsureOptions extends EnsureListSettings {
	applyToExisting?: boolean;
	contentTypes?: string[];
	defaultContentType?: string;
	copyViewSourceList?: unknown;
	copyViewName?: string;
	permissions?: {
		settings?: {
			breakInheritance?: boolean;
			copyRoleAssignments?: boolean;
			clearSubscopes?: boolean;
			roleAssignmentsJson?: unknown;
			roleAssignmentMode?: 'replace' | 'add';
		};
	};
}

function hasJson(value: unknown): boolean {
	if (value === undefined || value === null) return false;
	return typeof value !== 'string' || value.trim() !== '';
}

function readCopyView(
	ctx: IExecuteFunctions,
	options: EnsureOptions,
	itemIndex: number,
): EnsureListSpec['copyView'] {
	const sourceValue = locatorValue(options.copyViewSourceList);
	const view = options.copyViewName?.trim() ?? '';
	if (!sourceValue && !view) return undefined;
	if (!sourceValue || !view) {
		throw new NodeOperationError(
			ctx.getNode(),
			'To copy a view, set both "Copy View From List" and "View to Copy"',
			{ itemIndex },
		);
	}
	const raw = options.copyViewSourceList as Partial<LocatorValue> | string;
	const mode = typeof raw === 'object' && raw?.mode ? raw.mode : 'title';
	return { sourceList: { mode, value: sourceValue }, view };
}

function readPermissions(
	ctx: IExecuteFunctions,
	sp: SPFI,
	options: EnsureOptions,
	itemIndex: number,
): EnsureListPermissions | undefined {
	const settings = options.permissions?.settings;
	if (!settings) return undefined;
	const permissions: EnsureListPermissions = {
		breakInheritance: settings.breakInheritance ?? true,
		copyRoleAssignments: settings.copyRoleAssignments ?? false,
		clearSubscopes: settings.clearSubscopes ?? false,
		mode: settings.roleAssignmentMode ?? 'replace',
	};
	if (hasJson(settings.roleAssignmentsJson)) {
		let entries;
		try {
			entries = parseRoleAssignmentsJson(settings.roleAssignmentsJson, {
				rolesOptional: false,
				allowEmpty: true,
			});
		} catch (error) {
			throw new NodeOperationError(ctx.getNode(), (error as Error).message, { itemIndex });
		}
		permissions.resolveAssignments = async () =>
			(await resolveEntries(ctx, sp, entries, itemIndex, true)).flatMap((entry) =>
				entry.roleDefinitionIds.map((roleDefinitionId) => ({
					principalId: entry.principalId,
					roleDefinitionId,
				})),
			);
	}
	return permissions;
}

export async function ensure(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const title = getRequiredString.call(this, 'title', itemIndex);
	const url = String(this.getNodeParameter('listUrl', itemIndex, '') ?? '').trim();
	const options = this.getNodeParameter('ensureOptions', itemIndex, {}) as EnsureOptions;

	const spec: EnsureListSpec = {
		title,
		url: url ? toServerRelativePath(siteUrl, url) : undefined,
		template: this.getNodeParameter('template', itemIndex, 101) as number,
		description: String(this.getNodeParameter('listDescription', itemIndex, '') ?? '').trim(),
		applyToExisting: options.applyToExisting === true,
		settings: {
			onQuickLaunch: options.onQuickLaunch,
			enableVersioning: options.enableVersioning,
			majorVersionLimit: options.majorVersionLimit,
			enableMinorVersions: options.enableMinorVersions,
			minorVersionLimit: options.minorVersionLimit,
			enableContentTypes: options.enableContentTypes,
		},
		contentTypes: (options.contentTypes ?? []).map((id) => id.trim()).filter(Boolean),
		defaultContentType: options.defaultContentType?.trim() || undefined,
		copyView: readCopyView(this, options, itemIndex),
		permissions: readPermissions(this, sp, options, itemIndex),
	};

	try {
		return (await ensureList(spec, createEnsureListGateway(sp))) as unknown as IDataObject;
	} catch (error) {
		// Plain errors carry a message for the user. SharePoint and node errors keep their details.
		if (error instanceof Error && error.constructor === Error) {
			throw new NodeOperationError(this.getNode(), error.message, { itemIndex });
		}
		throw toNodeError(this.getNode(), error, itemIndex);
	}
}
