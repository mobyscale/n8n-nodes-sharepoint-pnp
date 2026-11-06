import type { INodeProperties } from 'n8n-workflow';

/**
 * One contract for property selection in every Get and Get Many operation:
 * - each operation loads a documented default set of properties,
 * - "Additional Properties" adds to that set, "Expand" loads related objects,
 * - "Return All Properties" sends no $select, so SharePoint returns its default entity.
 */
export interface SelectionValues {
	additionalProperties?: string | string[];
	expand?: string | string[];
	returnAllProperties?: boolean;
}

interface Selectable {
	select(...selects: string[]): this;
	expand(...expands: string[]): this;
}

export function toList(value: string | string[] | undefined): string[] {
	const entries = Array.isArray(value) ? value : String(value ?? '').split(',');
	return [...new Set(entries.map((entry) => String(entry).trim()).filter(Boolean))];
}

/** Returns the $select and $expand values. A "Nav/Prop" selection adds "Nav" to $expand, as SharePoint requires. */
export function buildSelection(
	defaults: string[],
	values: SelectionValues = {},
	defaultExpand: string[] = [],
): { select: string[]; expand: string[] } {
	const select = values.returnAllProperties
		? []
		: toList([...defaults, ...toList(values.additionalProperties)]);
	const navigation = select
		.filter((entry) => entry.includes('/'))
		.map((entry) => entry.split('/')[0]);
	const expand = toList([...defaultExpand, ...toList(values.expand), ...navigation]);
	return { select, expand };
}

export function applySelection<T extends Selectable>(
	query: T,
	defaults: string[],
	values: SelectionValues = {},
	defaultExpand: string[] = [],
): T {
	const { select, expand } = buildSelection(defaults, values, defaultExpand);
	let result = query;
	if (select.length) result = result.select(...select);
	if (expand.length) result = result.expand(...expand);
	return result;
}

/**
 * The options to put in an operation's "Options" collection. Read them back with the same names
 * (additionalProperties, expand, returnAllProperties) and pass them to applySelection.
 */
export function selectionOptions(defaults: string[]): INodeProperties[] {
	return [
		{
			displayName: 'Additional Properties',
			name: 'additionalProperties',
			type: 'string',
			default: '',
			placeholder: 'e.g. TimeCreated, Author/Title',
			description: `Comma-separated properties to load in addition to the defaults (${defaults.join(', ')}). Use "Navigation/Property" for related objects.`,
		},
		{
			displayName: 'Expand',
			name: 'expand',
			type: 'string',
			default: '',
			placeholder: 'e.g. ListItemAllFields',
			description: 'Comma-separated related objects to load',
		},
		{
			displayName: 'Return All Properties',
			name: 'returnAllProperties',
			type: 'boolean',
			default: false,
			description:
				'Whether to return every property that SharePoint returns by default, instead of the default set',
		},
	];
}
