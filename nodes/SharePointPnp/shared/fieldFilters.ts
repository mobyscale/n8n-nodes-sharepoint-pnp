import type { IDisplayOptions, INodeProperties } from 'n8n-workflow';
import type { IList } from '@pnp/sp/lists/index.js';

import { combineConditions, escapeXml } from './filesFolders';

export type FieldFilterOperator =
	| 'eq'
	| 'ne'
	| 'contains'
	| 'beginsWith'
	| 'gt'
	| 'lt'
	| 'isEmpty'
	| 'isNotEmpty';

export type FieldFilterCombine = 'and' | 'or';

export interface FieldFilter {
	field: string;
	operator: FieldFilterOperator;
	value: string;
}

const OPERATOR_NAMES: Record<FieldFilterOperator, string> = {
	eq: 'Equals',
	ne: 'Not Equals',
	contains: 'Contains',
	beginsWith: 'Begins With',
	gt: 'Greater Than',
	lt: 'Less Than',
	isEmpty: 'Is Empty',
	isNotEmpty: 'Is Not Empty',
};

const NO_VALUE: FieldFilterOperator[] = ['isEmpty', 'isNotEmpty'];

/** The "Field Filters" property. The field picker uses the given load options method. */
export function fieldFiltersProperty(settings: {
	loadOptionsMethod: string;
	loadOptionsDependsOn: string[];
	displayOptions: IDisplayOptions;
	description: string;
}): INodeProperties {
	return {
		displayName: 'Field Filters',
		name: 'fieldValueFilters',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Field Filter',
		default: {},
		description: settings.description,
		displayOptions: settings.displayOptions,
		options: [
			{
				displayName: 'Condition',
				name: 'conditions',
				values: [
					{
						displayName: 'Field Name or ID',
						name: 'field',
						type: 'options',
						typeOptions: {
							loadOptionsMethod: settings.loadOptionsMethod,
							loadOptionsDependsOn: settings.loadOptionsDependsOn,
						},
						default: '',
						description:
							'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>',
					},
					{
						displayName: 'Operator',
						name: 'operator',
						type: 'options',
						default: 'eq',
						options: (Object.keys(OPERATOR_NAMES) as FieldFilterOperator[]).map((value) => ({
							name: OPERATOR_NAMES[value],
							value,
						})),
					},
					{
						displayName: 'Value',
						name: 'value',
						type: 'string',
						default: '',
						description:
							'The value to compare with. Use an ISO date for date fields, 1 or 0 (true or false) for yes/no fields, and the item ID for lookup and person fields.',
						displayOptions: { hide: { operator: NO_VALUE } },
					},
				],
			},
		],
	};
}

/** The "Combine Field Filters" option, for an Options collection. */
export const fieldFiltersCombineOption: INodeProperties = {
	displayName: 'Combine Field Filters',
	name: 'fieldFiltersCombine',
	type: 'options',
	default: 'and',
	description: 'How the field filters are combined',
	options: [
		{ name: 'All (AND)', value: 'and', description: 'An entry must match all field filters' },
		{ name: 'Any (OR)', value: 'or', description: 'An entry must match at least one field filter' },
	],
};

/** Reads the "Field Filters" parameter. */
export function readFieldFilters(raw: unknown): FieldFilter[] {
	const rows = (raw as { conditions?: unknown } | null | undefined)?.conditions;
	if (!Array.isArray(rows)) return [];
	return rows.map((row: { field?: unknown; operator?: unknown; value?: unknown }) => {
		const field = String(row.field ?? '').trim();
		const operator = String(row.operator ?? 'eq') as FieldFilterOperator;
		const value = row.value === undefined || row.value === null ? '' : String(row.value).trim();
		if (!field) throw new Error('Choose a field for each field filter');
		if (!(operator in OPERATOR_NAMES)) {
			throw new Error(`"${operator}" is not a valid field filter operator`);
		}
		if (!NO_VALUE.includes(operator) && value === '') {
			throw new Error(
				`Give a value for the field filter on "${field}", or use the operator "Is Empty"`,
			);
		}
		return { field, operator, value };
	});
}

export function readCombine(value: unknown): FieldFilterCombine {
	return value === 'or' ? 'or' : 'and';
}

type ValueKind = 'text' | 'number' | 'date' | 'boolean' | 'lookup' | 'user';

const TEXT_TYPES = new Set(['Text', 'Note', 'Choice', 'MultiChoice']);
const NUMBER_TYPES = new Set(['Number', 'Currency', 'Integer', 'Counter']);

function kindOf(type: string | undefined): ValueKind {
	if (!type) return 'text';
	if (NUMBER_TYPES.has(type)) return 'number';
	if (type === 'DateTime') return 'date';
	if (type === 'Boolean') return 'boolean';
	if (type === 'Lookup' || type === 'LookupMulti') return 'lookup';
	if (type === 'User' || type === 'UserMulti') return 'user';
	return 'text';
}

function toNumber(filter: FieldFilter): string {
	const number = Number(filter.value);
	if (!Number.isFinite(number)) {
		throw new Error(`"${filter.value}" is not a number. "${filter.field}" is a number field.`);
	}
	return String(number);
}

function toBoolean(filter: FieldFilter): '1' | '0' {
	const value = filter.value.toLowerCase();
	if (['1', 'true', 'yes'].includes(value)) return '1';
	if (['0', 'false', 'no'].includes(value)) return '0';
	throw new Error(
		`"${filter.value}" is not a yes/no value. Use 1 or 0 for "${filter.field}" (true or false).`,
	);
}

function toId(filter: FieldFilter): string {
	if (!/^\d+$/.test(filter.value)) {
		throw new Error(
			`"${filter.value}" is not an ID. "${filter.field}" is a lookup or person field: use the item ID or the user ID.`,
		);
	}
	return String(Number(filter.value));
}

/** A date without a time compares only the date. A date with a time compares the full value. */
function toDate(filter: FieldFilter): { value: string; includeTime: boolean } {
	const date = new Date(filter.value);
	if (Number.isNaN(date.getTime())) throw new Error(`"${filter.value}" is not a valid date`);
	if (/^\d{4}-\d{2}-\d{2}$/.test(filter.value)) {
		return { value: filter.value, includeTime: false };
	}
	return { value: date.toISOString(), includeTime: true };
}

function checkTextOperator(filter: FieldFilter, kind: ValueKind, type: string | undefined): void {
	if ((filter.operator === 'contains' || filter.operator === 'beginsWith') && kind !== 'text') {
		throw new Error(
			`The operator "${OPERATOR_NAMES[filter.operator]}" works only with text fields. "${filter.field}" is a ${type} field.`,
		);
	}
}

const CAML_OPERATORS: Record<FieldFilterOperator, string> = {
	eq: 'Eq',
	ne: 'Neq',
	contains: 'Contains',
	beginsWith: 'BeginsWith',
	gt: 'Gt',
	lt: 'Lt',
	isEmpty: 'IsNull',
	isNotEmpty: 'IsNotNull',
};

function camlValue(filter: FieldFilter, kind: ValueKind, type: string | undefined): string {
	switch (kind) {
		case 'number':
			return `<Value Type="${type}">${toNumber(filter)}</Value>`;
		case 'boolean':
			return `<Value Type="Boolean">${toBoolean(filter)}</Value>`;
		case 'lookup':
			return `<Value Type="Lookup">${toId(filter)}</Value>`;
		case 'user':
			return `<Value Type="Integer">${toId(filter)}</Value>`;
		case 'date': {
			const { value, includeTime } = toDate(filter);
			return `<Value Type="DateTime"${includeTime ? ' IncludeTimeValue="TRUE"' : ''}>${value}</Value>`;
		}
		default: {
			const valueType = type && TEXT_TYPES.has(type) ? type : 'Text';
			return `<Value Type="${valueType}">${escapeXml(filter.value)}</Value>`;
		}
	}
}

/** Builds the CAML condition of one field filter. A field type that is not known is used as Text. */
export function buildCamlCondition(filter: FieldFilter, type?: string): string {
	const kind = kindOf(type);
	const tag = CAML_OPERATORS[filter.operator];
	const name = escapeXml(filter.field);
	if (NO_VALUE.includes(filter.operator)) return `<${tag}><FieldRef Name="${name}"/></${tag}>`;
	checkTextOperator(filter, kind, type);
	const lookupId = kind === 'lookup' || kind === 'user' ? ' LookupId="TRUE"' : '';
	return `<${tag}><FieldRef Name="${name}"${lookupId}/>${camlValue(filter, kind, type)}</${tag}>`;
}

/** Builds one CAML condition for all field filters, or undefined when there are none. */
export function buildCamlFilter(
	filters: FieldFilter[],
	types: Map<string, string>,
	combine: FieldFilterCombine = 'and',
): string | undefined {
	if (!filters.length) return undefined;
	return combineConditions(
		combine === 'or' ? 'Or' : 'And',
		filters.map((filter) => buildCamlCondition(filter, types.get(filter.field))),
	);
}

function escapeODataString(value: string): string {
	return value.replace(/'/g, "''");
}

function odataLiteral(filter: FieldFilter, kind: ValueKind): string {
	switch (kind) {
		case 'number':
			return toNumber(filter);
		case 'boolean':
			return toBoolean(filter);
		case 'lookup':
		case 'user':
			return toId(filter);
		case 'date': {
			const { value, includeTime } = toDate(filter);
			return `datetime'${includeTime ? value : `${value}T00:00:00Z`}'`;
		}
		default:
			return `'${escapeODataString(filter.value)}'`;
	}
}

/** Builds the OData condition of one field filter. Lookup and person fields compare the ID. */
export function buildODataCondition(filter: FieldFilter, type?: string): string {
	const kind = kindOf(type);
	if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(filter.field)) {
		throw new Error(`"${filter.field}" is not a valid internal field name`);
	}
	const name = kind === 'lookup' || kind === 'user' ? `${filter.field}Id` : filter.field;
	if (filter.operator === 'isEmpty') return `${name} eq null`;
	if (filter.operator === 'isNotEmpty') return `${name} ne null`;
	checkTextOperator(filter, kind, type);
	const literal = odataLiteral(filter, kind);
	if (filter.operator === 'contains') return `substringof(${literal},${name})`;
	if (filter.operator === 'beginsWith') return `startswith(${name},${literal})`;
	return `${name} ${filter.operator} ${literal}`;
}

/** Builds one OData $filter expression for all field filters, or undefined when there are none. */
export function buildODataFilter(
	filters: FieldFilter[],
	types: Map<string, string>,
	combine: FieldFilterCombine = 'and',
): string | undefined {
	if (!filters.length) return undefined;
	const conditions = filters.map((filter) => buildODataCondition(filter, types.get(filter.field)));
	if (conditions.length === 1) return conditions[0];
	return conditions.map((condition) => `(${condition})`).join(` ${combine} `);
}

/**
 * Adds a condition to a CAML query with And. The query can be the content of a Query element, a
 * Query element or a full View element, like the CAML Query parameter of List Item Get Many.
 */
export function addCamlCondition(caml: string, condition: string): string {
	const input = caml.trim();
	const where = /<Where\s*>([\s\S]*)<\/Where\s*>/i;
	if (where.test(input)) {
		return input.replace(
			where,
			(_match, inner: string) => `<Where><And>${inner.trim()}${condition}</And></Where>`,
		);
	}
	const newWhere = `<Where>${condition}</Where>`;
	if (/<Query(\s[^>]*)?>/i.test(input)) {
		return input.replace(/<Query(\s[^>]*)?>/i, (open) => `${open}${newWhere}`);
	}
	if (/^<View[\s>]/i.test(input)) {
		return input.replace(/^<View(\s[^>]*)?>/i, (open) => `${open}<Query>${newWhere}</Query>`);
	}
	return `${newWhere}${input}`;
}

/**
 * Reads the types of the fields that the filters use. Throws when a field does not exist, so the
 * error names the field instead of a SharePoint query error.
 */
export async function loadFilterFieldTypes(
	list: IList,
	filters: FieldFilter[],
): Promise<Map<string, string>> {
	const fields = (await list.fields.select('InternalName', 'TypeAsString')()) as Array<{
		InternalName: string;
		TypeAsString: string;
	}>;
	const types = new Map(fields.map((field) => [field.InternalName, field.TypeAsString]));
	const missing = [...new Set(filters.map((filter) => filter.field))].filter(
		(name) => !types.has(name),
	);
	if (missing.length) {
		throw new Error(
			`These filter fields do not exist in the list: ${missing.join(', ')}. Use the internal field names.`,
		);
	}
	return types;
}
