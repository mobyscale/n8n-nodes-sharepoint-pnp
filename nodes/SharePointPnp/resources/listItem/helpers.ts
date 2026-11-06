export interface FieldFormValue {
	FieldName: string;
	FieldValue: string;
}

export interface FieldInfo {
	InternalName: string;
	TypeAsString: string;
}

const MULTI_VALUE_TYPES = new Set([
	'MultiChoice',
	'LookupMulti',
	'UserMulti',
	'TaxonomyFieldTypeMulti',
]);

/** Field types that "Append to Multi-Value Fields" can merge. */
export const APPENDABLE_TYPES = new Set(['MultiChoice', 'LookupMulti']);

export function isMultiValueType(typeAsString: string): boolean {
	return MULTI_VALUE_TYPES.has(typeAsString);
}

function toArray(value: unknown): unknown[] {
	if (Array.isArray(value)) return value;
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
	return value === undefined || value === null ? [] : [value];
}

function lookupId(value: unknown): string {
	if (value !== null && typeof value === 'object') {
		const entry = value as { Id?: unknown; id?: unknown; LookupId?: unknown };
		return String(entry.Id ?? entry.id ?? entry.LookupId ?? '');
	}
	return String(value);
}

function userKey(value: unknown): string {
	if (value !== null && typeof value === 'object') {
		const entry = value as { Key?: unknown; LoginName?: unknown; Email?: unknown; Id?: unknown };
		if (entry.Key !== undefined) return String(entry.Key);
		if (entry.LoginName !== undefined) return String(entry.LoginName);
		if (entry.Email !== undefined) return String(entry.Email);
	}
	return String(value);
}

/** Returns the union of the current and new values. New values keep their order after the current ones. */
export function mergeMultiValues(current: unknown, incoming: unknown): unknown[] {
	const result: unknown[] = [];
	const seen = new Set<string>();
	for (const value of [...toArray(current), ...toArray(incoming)]) {
		const key = typeof value === 'object' && value !== null ? lookupId(value) : String(value);
		if (seen.has(key)) continue;
		seen.add(key);
		result.push(value);
	}
	return result;
}

/**
 * Converts a value to the string format that "validateUpdateListItem" expects for the field type.
 * Returns null when the value must not be sent.
 */
export function formatFieldValue(typeAsString: string | undefined, value: unknown): string | null {
	if (value === undefined) return null;
	if (value === null) return '';

	switch (typeAsString) {
		case 'Boolean':
			if (typeof value === 'string') {
				return ['1', 'true', 'yes'].includes(value.trim().toLowerCase()) ? '1' : '0';
			}
			return value ? '1' : '0';
		case 'DateTime':
			return value instanceof Date ? value.toISOString() : String(value);
		case 'MultiChoice': {
			const choices = toArray(value).map(String);
			return choices.length ? `;#${choices.join(';#')};#` : '';
		}
		case 'Lookup':
			return lookupId(value);
		case 'LookupMulti': {
			const ids = toArray(value).map(lookupId);
			return ids.join(';#;#');
		}
		case 'User':
			return JSON.stringify([{ Key: userKey(value) }]);
		case 'UserMulti':
			return JSON.stringify(toArray(value).map((entry) => ({ Key: userKey(entry) })));
		default:
			if (typeof value === 'object') return JSON.stringify(value);
			return String(value);
	}
}

/** Builds the form values for "validateUpdateListItem" and "addValidateUpdateItemUsingPath". */
export function buildFormValues(
	values: Record<string, unknown>,
	fieldInfos: FieldInfo[],
): FieldFormValue[] {
	const types = new Map(fieldInfos.map((field) => [field.InternalName, field.TypeAsString]));
	const result: FieldFormValue[] = [];
	for (const [name, value] of Object.entries(values)) {
		const formatted = formatFieldValue(types.get(name), value);
		if (formatted !== null) result.push({ FieldName: name, FieldValue: formatted });
	}
	return result;
}

export interface ViewXmlOptions {
	recursive?: boolean;
	rowLimit?: number;
}

/**
 * Builds a ViewXml document. The input can be the content of a Query element (Where, OrderBy),
 * or a full View element. A full View element is only changed to add the Scope attribute for
 * recursive queries and a RowLimit when none exists.
 */
export function buildViewXml(caml: string, options: ViewXmlOptions = {}): string {
	const input = caml.trim();
	const rowLimit = options.rowLimit ? `<RowLimit Paged="TRUE">${options.rowLimit}</RowLimit>` : '';

	if (/^<View[\s>]/i.test(input)) {
		let xml = input;
		if (options.recursive && !/<View[^>]*\sScope=/i.test(xml)) {
			xml = xml.replace(/^<View/i, '<View Scope="RecursiveAll"');
		}
		if (rowLimit && !/<RowLimit/i.test(xml)) {
			xml = xml.replace(/<\/View>\s*$/i, `${rowLimit}</View>`);
		}
		return xml;
	}

	const scope = options.recursive ? ' Scope="RecursiveAll"' : '';
	const query = /^<Query[\s>]/i.test(input) ? input : `<Query>${input}</Query>`;
	return `<View${scope}>${input ? query : ''}${rowLimit}</View>`;
}

/** Returns the navigation properties that a $select list needs in $expand (for example "Author/Title"). */
export function expandFromSelect(select: string[]): string[] {
	const expand = new Set<string>();
	for (const entry of select) {
		const index = entry.indexOf('/');
		if (index > 0) expand.add(entry.slice(0, index));
	}
	return [...expand];
}

export function escapeODataString(value: string): string {
	return value.replace(/'/g, "''");
}

export function parseCommaList(value: string | undefined): string[] {
	return (value ?? '')
		.split(',')
		.map((part) => part.trim())
		.filter(Boolean);
}
