import { escapeODataString, type FieldFormValue } from './helpers';

export const MAX_BATCH_SIZE = 100;

/** Key field types for Create or Update, by the way the filter compares the value. */
const TEXT_KEY_TYPES = new Set(['Text', 'Choice', 'File']);
const NUMBER_KEY_TYPES = new Set(['Number', 'Integer', 'Counter', 'Currency']);
const ID_KEY_TYPES = new Set(['Lookup', 'User']);

export const SUPPORTED_KEY_TYPES = [
	...TEXT_KEY_TYPES,
	...NUMBER_KEY_TYPES,
	'Boolean',
	...ID_KEY_TYPES,
];

/** Key types that are added to the fields of a new item when the fields do not set the key. */
const AUTO_SET_KEY_TYPES = new Set([
	'Text',
	'Choice',
	'Number',
	'Integer',
	'Currency',
	'Boolean',
	'Lookup',
]);

export function clampBatchSize(value: unknown): number {
	const size = Math.floor(Number(value));
	if (!Number.isFinite(size) || size < 1) return MAX_BATCH_SIZE;
	return Math.min(size, MAX_BATCH_SIZE);
}

export function parseListItemId(value: unknown): number {
	const text = String(value ?? '').trim();
	const id = Number(text);
	if (!text || !Number.isInteger(id) || id < 1) {
		throw new Error(`"${text}" is not a valid list item ID`);
	}
	return id;
}

/** OData names of fields that start with an underscore get the prefix "OData_". */
export function toODataName(internalName: string): string {
	return internalName.startsWith('_') ? `OData_${internalName}` : internalName;
}

function parseBoolean(value: unknown): boolean {
	if (typeof value === 'string') return ['1', 'true', 'yes'].includes(value.trim().toLowerCase());
	return Boolean(value);
}

function parseNumber(value: unknown, what: string): number {
	const text = String(value ?? '').trim();
	const number = Number(text);
	if (!text || !Number.isFinite(number)) throw new Error(`The ${what} "${text}" is not a number`);
	return number;
}

/** Validates a key value for the field type. Text keys stay text, other keys become numbers. */
function normalizeKeyValue(
	internalName: string,
	typeAsString: string,
	value: unknown,
): string | number {
	if (TEXT_KEY_TYPES.has(typeAsString)) {
		const text = String(value ?? '');
		if (!text.trim()) throw new Error('The key value must not be empty');
		return text;
	}
	if (NUMBER_KEY_TYPES.has(typeAsString)) return parseNumber(value, 'key value');
	if (typeAsString === 'Boolean') return parseBoolean(value) ? 1 : 0;
	if (ID_KEY_TYPES.has(typeAsString)) {
		const id = parseNumber(value, 'key value');
		if (!Number.isInteger(id) || id < 1) {
			throw new Error(
				`The key value for the ${typeAsString} field must be a numeric item or user ID`,
			);
		}
		return id;
	}
	throw new Error(
		`The field "${internalName}" has the type "${typeAsString}", which cannot be used as a key. Supported types: ${SUPPORTED_KEY_TYPES.join(', ')}`,
	);
}

/** Builds the OData $filter that finds the items with the key value. */
export function buildKeyFilter(internalName: string, typeAsString: string, value: unknown): string {
	const name = toODataName(internalName);
	const key = normalizeKeyValue(internalName, typeAsString, value);
	if (typeof key === 'string') return `${name} eq '${escapeODataString(key)}'`;
	return ID_KEY_TYPES.has(typeAsString) ? `${name}Id eq ${key}` : `${name} eq ${key}`;
}

function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

/**
 * Builds the CAML ViewXml that finds the items with the key value, ordered by ID. A CAML query can
 * be scoped to one folder, which works in lists above the list view threshold.
 */
export function buildKeyViewXml(
	internalName: string,
	typeAsString: string,
	value: unknown,
	rowLimit: number,
): string {
	const key = normalizeKeyValue(internalName, typeAsString, value);
	const name = escapeXml(internalName);
	const isId = ID_KEY_TYPES.has(typeAsString);
	const field = `<FieldRef Name="${name}"${isId ? ' LookupId="TRUE"' : ''}/>`;
	const valueType = isId ? (typeAsString === 'User' ? 'Integer' : 'Lookup') : typeAsString;
	const where = `<Where><Eq>${field}<Value Type="${valueType}">${escapeXml(String(key))}</Value></Eq></Where>`;
	return `<View><Query>${where}<OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy></Query><RowLimit>${rowLimit}</RowLimit></View>`;
}

/** Whether the key value is added to the fields of a new item when the fields do not set the key. */
export function shouldSetKeyOnCreate(typeAsString: string): boolean {
	return AUTO_SET_KEY_TYPES.has(typeAsString);
}

const COMPARABLE_TEXT = new Set(['Text', 'Note', 'Choice', 'File']);
const COMPARABLE_NUMBER = new Set(['Number', 'Integer', 'Counter', 'Currency']);

/**
 * The property to select to compare the current value of a field, or undefined when
 * "Skip Update If Unchanged" cannot compare the field type. Such fields always count as changed.
 */
export function comparisonSelect(
	internalName: string,
	typeAsString: string | undefined,
): string | undefined {
	const type = typeAsString ?? '';
	if (type === 'Lookup' || type === 'LookupMulti') return `${toODataName(internalName)}Id`;
	if (
		COMPARABLE_TEXT.has(type) ||
		COMPARABLE_NUMBER.has(type) ||
		type === 'Boolean' ||
		type === 'DateTime' ||
		type === 'MultiChoice'
	) {
		return toODataName(internalName);
	}
	return undefined;
}

const isEmpty = (value: unknown) => value === undefined || value === null || value === '';

function toList(value: unknown): unknown[] {
	if (Array.isArray(value)) return value;
	const results = (value as { results?: unknown } | null)?.results;
	if (Array.isArray(results)) return results;
	if (typeof value === 'string') {
		const trimmed = value.trim();
		if (trimmed.startsWith('[')) {
			try {
				const parsed: unknown = JSON.parse(trimmed);
				if (Array.isArray(parsed)) return parsed;
			} catch {
				// not JSON, treat as a delimited list
			}
		}
		return trimmed === '' ? [] : trimmed.split(',').map((part) => part.trim());
	}
	return isEmpty(value) ? [] : [value];
}

function lookupIdOf(value: unknown): string {
	if (value !== null && typeof value === 'object') {
		const entry = value as { Id?: unknown; id?: unknown; LookupId?: unknown };
		return String(entry.Id ?? entry.id ?? entry.LookupId ?? '');
	}
	return String(value);
}

function sameSet(a: string[], b: string[]): boolean {
	if (a.length !== b.length) return false;
	const sortedA = [...a].sort();
	const sortedB = [...b].sort();
	return sortedA.every((entry, index) => entry === sortedB[index]);
}

/** Compares a new field value with the current value. Returns false when the type cannot be compared. */
export function isSameValue(
	typeAsString: string | undefined,
	desired: unknown,
	current: unknown,
): boolean {
	const type = typeAsString ?? '';
	if (desired === undefined) return true;
	if (COMPARABLE_TEXT.has(type)) return String(desired ?? '') === String(current ?? '');
	if (COMPARABLE_NUMBER.has(type)) {
		if (isEmpty(desired) || isEmpty(current)) return isEmpty(desired) && isEmpty(current);
		return Number(desired) === Number(current);
	}
	if (type === 'Boolean') return parseBoolean(desired) === parseBoolean(current);
	if (type === 'DateTime') {
		if (isEmpty(desired) || isEmpty(current)) return isEmpty(desired) && isEmpty(current);
		const desiredTime = desired instanceof Date ? desired.getTime() : Date.parse(String(desired));
		return !Number.isNaN(desiredTime) && desiredTime === Date.parse(String(current));
	}
	if (type === 'Lookup') {
		if (isEmpty(desired) || isEmpty(current)) return isEmpty(desired) && isEmpty(current);
		return lookupIdOf(desired) === String(current);
	}
	if (type === 'MultiChoice')
		return sameSet(toList(desired).map(String), toList(current).map(String));
	if (type === 'LookupMulti') {
		return sameSet(toList(desired).map(lookupIdOf), toList(current).map(String));
	}
	return false;
}

/** Whether every given field already has the given value on the item. */
export function isUnchanged(
	values: Record<string, unknown>,
	types: Map<string, string>,
	current: Record<string, unknown>,
): boolean {
	return Object.entries(values).every(([name, desired]) => {
		const select = comparisonSelect(name, types.get(name));
		if (!select) return false;
		return isSameValue(types.get(name), desired, current[select]);
	});
}

export interface SystemUpdateSource {
	Modified?: string;
	Editor?: { Name?: string };
}

/**
 * validateUpdateListItem always sets Modified and Editor. A system update sends the current values
 * back, so they do not change. `localModified` is the current Modified value in the time zone of
 * the site. A user can only be sent by sign-in name, so an app principal editor may change.
 */
export function systemUpdateValues(
	current: SystemUpdateSource,
	localModified: string | undefined,
	formValues: FieldFormValue[],
): FieldFormValue[] {
	const given = new Set(formValues.map((entry) => entry.FieldName));
	const extra: FieldFormValue[] = [];
	if (localModified && !given.has('Modified')) {
		extra.push({
			FieldName: 'Modified',
			FieldValue: localModified.replace('T', ' ').replace(/\.\d+$/, ''),
		});
	}
	if (current.Editor?.Name?.startsWith('i:0#') && !given.has('Editor')) {
		extra.push({
			FieldName: 'Editor',
			FieldValue: JSON.stringify([{ Key: current.Editor.Name }]),
		});
	}
	return extra;
}

export interface FieldResult {
	FieldName?: string;
	FieldValue?: string;
	HasException?: boolean;
	ErrorMessage?: string;
}

/** Throws when SharePoint rejected one or more field values. */
export function assertFieldResults(results: FieldResult[]): void {
	const failed = results.filter((entry) => entry.HasException);
	if (!failed.length) return;
	const message = failed
		.map(
			(entry) => `${entry.FieldName ?? 'Unknown field'}: ${entry.ErrorMessage ?? 'invalid value'}`,
		)
		.join('; ');
	throw new Error(`SharePoint rejected the field values. ${message}`);
}

/** Returns the ID of a new item from the results of addValidateUpdateItemUsingPath. */
export function newItemId(results: FieldResult[]): number {
	assertFieldResults(results);
	const id = Number(results.find((entry) => entry.FieldName === 'Id')?.FieldValue);
	if (!Number.isFinite(id) || id <= 0)
		throw new Error('SharePoint did not return the ID of the new item');
	return id;
}
