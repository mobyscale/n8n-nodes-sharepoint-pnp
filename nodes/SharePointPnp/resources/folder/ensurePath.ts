import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { undefinedIfNotFound } from '../../shared/filesFolders';
import { joinPath } from '../../shared/utils';
import { buildFormValues, mergeMultiValues } from '../listItem/helpers';
import {
	describeFieldErrors,
	isInvalidJson,
	loadListFields,
	parseFieldsObject,
	readCurrentValue,
	selectNameFor,
	toValueList,
	tryParseJson,
	unknownFields,
	valuesEqual,
	type FieldResult,
	type FieldValues,
	type ListFields,
} from './fieldValues';

/** The field values of one folder that Ensure Path or Ensure Folders creates or keeps. */
export interface FolderFieldSpec {
	fieldsWhenCreated: FieldValues;
	fieldsAlways: FieldValues;
	appendToMultiChoice?: { field: string; values: string[] };
}

export interface SegmentSpec extends FolderFieldSpec {
	name: string;
}

export interface SegmentResult extends IDataObject {
	name: string;
	serverRelativeUrl: string;
	listItemId: number | null;
	created: boolean;
	fieldsUpdated: string[];
}

export interface FolderState {
	serverRelativeUrl: string;
	listItemId: number | null;
	uniqueId?: string;
	/** The current values of the requested fields, keyed by the $select names. */
	item: Record<string, unknown>;
}

/** The SharePoint calls that Ensure Path needs. The node uses PnPjs, the unit tests use a fake. */
export interface EnsurePathGateway {
	/** Returns the server-relative URL of a folder that exists. Works for a library root folder too. */
	getBaseFolder(path: string): Promise<string | undefined>;
	getFolder(path: string, selectNames: string[]): Promise<FolderState | undefined>;
	createFolder(path: string): Promise<void>;
	getFieldTypes(): Promise<Map<string, string>>;
	updateItem(listItemId: number, values: FieldValues): Promise<void>;
}

/** The segment rows of the UI. JSON fields arrive as text or as objects. */
export interface SegmentRow {
	name?: string;
	fieldsWhenCreated?: unknown;
	fieldsAlways?: unknown;
	appendField?: string;
	appendValues?: string;
}

function checkName(raw: unknown, index: number): string {
	const name = String(raw ?? '').trim();
	if (!name) throw new Error(`Segment ${index + 1} has no name`);
	if (/[\\/]/.test(name)) {
		throw new Error(
			`The segment name "${name}" must not have "/" or "\\". Use one segment for each level.`,
		);
	}
	return name;
}

export function toAppend(field: unknown, values: unknown): FolderFieldSpec['appendToMultiChoice'] {
	const name = String(field ?? '').trim();
	const list = toValueList(values);
	return name && list.length ? { field: name, values: list } : undefined;
}

export function segmentsFromRows(rows: SegmentRow[] | undefined): SegmentSpec[] {
	return (rows ?? []).map((row, index) => ({
		name: checkName(row.name, index),
		fieldsWhenCreated: parseFieldsObject(
			row.fieldsWhenCreated,
			`Segment ${index + 1}: Fields When Created`,
		),
		fieldsAlways: parseFieldsObject(row.fieldsAlways, `Segment ${index + 1}: Fields Always`),
		appendToMultiChoice: toAppend(row.appendField, row.appendValues),
	}));
}

/**
 * Reads the JSON segments: an array of names, or of objects with "name", "fieldsWhenCreated",
 * "fieldsAlways" and "appendToMultiChoice" ({ "field": "...", "values": [...] }).
 */
export function segmentsFromJson(raw: unknown): SegmentSpec[] {
	let value = raw;
	if (typeof value === 'string') {
		value = tryParseJson(value);
		if (isInvalidJson(value)) throw new Error('"Segments (JSON)" is not valid JSON');
	}
	if (!Array.isArray(value)) throw new Error('"Segments (JSON)" must be an array');
	return value.map((entry: unknown, index) => {
		if (typeof entry === 'string') {
			return { name: checkName(entry, index), fieldsWhenCreated: {}, fieldsAlways: {} };
		}
		if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
			throw new Error(`Segment ${index + 1} must be a name or an object`);
		}
		const data = entry as IDataObject;
		const append = (data.appendToMultiChoice ?? {}) as IDataObject;
		return {
			name: checkName(data.name, index),
			fieldsWhenCreated: parseFieldsObject(
				data.fieldsWhenCreated,
				`Segment ${index + 1}: fieldsWhenCreated`,
			),
			fieldsAlways: parseFieldsObject(data.fieldsAlways, `Segment ${index + 1}: fieldsAlways`),
			appendToMultiChoice: toAppend(append.field, append.values),
		};
	});
}

function fieldNamesOf(segment: FolderFieldSpec): string[] {
	return [
		...Object.keys(segment.fieldsWhenCreated),
		...Object.keys(segment.fieldsAlways),
		...(segment.appendToMultiChoice ? [segment.appendToMultiChoice.field] : []),
	];
}

/** The fields to read before a run on an existing folder: the "always" fields and the append field. */
export function fieldsToRead(segment: FolderFieldSpec): string[] {
	return [
		...new Set([
			...Object.keys(segment.fieldsAlways),
			...(segment.appendToMultiChoice ? [segment.appendToMultiChoice.field] : []),
		]),
	];
}

/**
 * Decides which values to write to a segment folder. A new folder gets the "when created" values,
 * then the "always" values. An existing folder gets only the "always" values that differ from the
 * current values. Append values are merged with the current values of the multi-choice field.
 */
export function planSegmentValues(
	segment: FolderFieldSpec,
	created: boolean,
	current: Record<string, unknown>,
	types: Map<string, string>,
): FieldValues {
	const values: FieldValues = {};
	if (created) {
		Object.assign(values, segment.fieldsWhenCreated, segment.fieldsAlways);
	} else {
		for (const [name, value] of Object.entries(segment.fieldsAlways)) {
			if (!valuesEqual(types.get(name), readCurrentValue(current, name, types.get(name)), value)) {
				values[name] = value;
			}
		}
	}
	const append = segment.appendToMultiChoice;
	if (append) {
		const type = types.get(append.field);
		const base = created
			? values[append.field]
			: (values[append.field] ?? readCurrentValue(current, append.field, type));
		const merged = mergeMultiValues(base ?? [], append.values);
		const existing = created ? undefined : readCurrentValue(current, append.field, type);
		if (created || !valuesEqual(type, existing, merged)) values[append.field] = merged;
	}
	return values;
}

function validateFields(specs: FolderFieldSpec[], types: Map<string, string>): void {
	const missing = unknownFields(specs.flatMap(fieldNamesOf), types);
	if (missing.length) {
		throw new Error(
			`These fields do not exist in the library: ${missing.join(', ')}. Use the internal field names.`,
		);
	}
	for (const spec of specs) {
		const append = spec.appendToMultiChoice;
		if (append && types.get(append.field) !== 'MultiChoice') {
			throw new Error(
				`The field "${append.field}" is not a multi-choice field, so values cannot be appended`,
			);
		}
	}
}

/** Reads the field types once when a spec has fields, and checks that all fields exist. */
export async function loadFieldTypes(
	specs: FolderFieldSpec[],
	gateway: EnsurePathGateway,
): Promise<Map<string, string>> {
	if (!specs.some((spec) => fieldNamesOf(spec).length > 0)) return new Map();
	const types = await gateway.getFieldTypes();
	validateFields(specs, types);
	return types;
}

export async function requireBaseFolder(
	basePath: string,
	gateway: EnsurePathGateway,
): Promise<string> {
	const base = await gateway.getBaseFolder(basePath);
	if (!base) throw new Error(`The base folder "${basePath}" does not exist`);
	return base;
}

export interface EnsuredFolder {
	state: FolderState;
	created: boolean;
	fieldsUpdated: string[];
}

/**
 * Makes sure that one folder exists and sets its fields. When the parent was created in this run,
 * the folder cannot exist, so the node does not look it up first.
 */
export async function ensureOneFolder(
	path: string,
	spec: FolderFieldSpec | undefined,
	parentCreated: boolean,
	types: Map<string, string>,
	gateway: EnsurePathGateway,
): Promise<EnsuredFolder> {
	const selectNames = spec
		? fieldsToRead(spec).map((name) => selectNameFor(name, types.get(name)))
		: [];
	let state = parentCreated ? undefined : await gateway.getFolder(path, selectNames);
	const created = !state;
	if (!state) {
		await gateway.createFolder(path);
		state = await gateway.getFolder(path, []);
		if (!state) throw new Error(`The folder "${path}" was not found after it was created`);
	}
	const values = spec ? planSegmentValues(spec, created, state.item, types) : {};
	const names = Object.keys(values);
	if (names.length) {
		if (state.listItemId === null) {
			throw new Error(
				`The folder "${state.serverRelativeUrl}" has no list item, so its fields cannot be set`,
			);
		}
		await gateway.updateItem(state.listItemId, values);
	}
	return { state, created, fieldsUpdated: names };
}

/** Walks the segments below the base folder, creates the missing folders and sets their fields. */
export async function ensurePath(
	basePath: string,
	segments: SegmentSpec[],
	gateway: EnsurePathGateway,
): Promise<IDataObject> {
	if (!segments.length) throw new Error('Add at least one segment');
	const types = await loadFieldTypes(segments, gateway);
	const base = await requireBaseFolder(basePath, gateway);

	const results: SegmentResult[] = [];
	let parent = base;
	let parentCreated = false;
	for (const segment of segments) {
		const path = joinPath(parent, segment.name);
		const { state, created, fieldsUpdated } = await ensureOneFolder(
			path,
			segment,
			parentCreated,
			types,
			gateway,
		);
		results.push({
			name: segment.name,
			serverRelativeUrl: state.serverRelativeUrl,
			listItemId: state.listItemId,
			created,
			fieldsUpdated,
		});
		parent = state.serverRelativeUrl;
		parentCreated = created;
	}

	const last = results[results.length - 1];
	return {
		folderPath: last.serverRelativeUrl,
		listItemId: last.listItemId,
		created: results.some((entry) => entry.created),
		changed: results.some((entry) => entry.created || entry.fieldsUpdated.length > 0),
		segments: results,
	};
}

interface RawFolder {
	Exists?: boolean;
	ServerRelativeUrl?: string;
	UniqueId?: string;
	ListItemAllFields?: Record<string, unknown> | null;
}

/** Connects Ensure Path to SharePoint through PnPjs. */
export function createEnsurePathGateway(sp: SPFI, basePath: string): EnsurePathGateway {
	let fields: Promise<ListFields> | undefined;
	const listFields = () => (fields ??= loadListFields(sp, basePath));
	return {
		async getBaseFolder(path) {
			const info = await undefinedIfNotFound(
				sp.web.getFolderByServerRelativePath(path).select('Exists', 'ServerRelativeUrl')(),
			);
			return info?.Exists ? (info.ServerRelativeUrl ?? path) : undefined;
		},
		async getFolder(path, selectNames) {
			const select = ['Exists', 'ServerRelativeUrl', 'UniqueId', 'ListItemAllFields/Id'];
			select.push(...selectNames.map((name) => `ListItemAllFields/${name}`));
			const info = (await undefinedIfNotFound(
				sp.web
					.getFolderByServerRelativePath(path)
					.select(...select)
					.expand('ListItemAllFields')(),
			)) as RawFolder | undefined;
			if (!info?.Exists) return undefined;
			const item = info.ListItemAllFields ?? {};
			const id = Number(item.Id);
			return {
				serverRelativeUrl: info.ServerRelativeUrl ?? path,
				listItemId: Number.isFinite(id) && id > 0 ? id : null,
				uniqueId: info.UniqueId,
				item,
			};
		},
		async createFolder(path) {
			// addUsingPath fails when the folder exists. Another run can create it after the check.
			const error = await sp.web.folders.addUsingPath(path, false).then(
				() => undefined,
				(reason: unknown) => reason,
			);
			if (error === undefined) return;
			const info = await undefinedIfNotFound(
				sp.web.getFolderByServerRelativePath(path).select('Exists')(),
			);
			if (!info?.Exists) throw error;
		},
		async getFieldTypes() {
			return (await listFields()).types;
		},
		async updateItem(listItemId, values) {
			const { listId, infos } = await listFields();
			const formValues = buildFormValues(values, infos);
			const results = (await sp.web.lists
				.getById(listId)
				.items.getById(listItemId)
				.validateUpdateListItem(formValues, false)) as FieldResult[];
			const message = describeFieldErrors(results);
			if (message) throw new Error(`SharePoint rejected the field values. ${message}`);
		},
	};
}
