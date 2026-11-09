import type { INodeProperties } from 'n8n-workflow';

import { pathRLC } from './descriptions';

/** Reads the value of a resource locator, or of a plain value that an older workflow saved. */
export function locatorValue(input: unknown): string {
	if (input !== null && typeof input === 'object') {
		return String((input as { value?: unknown }).value ?? '').trim();
	}
	return String(input ?? '').trim();
}

export function parseItemId(input: unknown): number {
	const value = locatorValue(input);
	if (!/^\d+$/.test(value) || Number(value) < 1) {
		throw new Error(`"${value}" is not a valid item ID`);
	}
	return Number(value);
}

/** An item picker that searches the items of the selected "list" parameter, or takes an ID. */
export function itemIdRLC(
	overrides: Partial<INodeProperties> = {},
	description = 'The list item',
): INodeProperties {
	return {
		displayName: 'Item',
		name: 'itemId',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod: 'searchListItems', searchable: true },
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: '12',
				validation: [
					{
						type: 'regex',
						properties: { regex: '^[1-9][0-9]*$', errorMessage: 'Not a valid item ID' },
					},
				],
			},
		],
		...overrides,
	};
}

/** A folder picker that lists the folders of the selected "list" parameter. */
export function listFolderRLC(
	overrides: Partial<INodeProperties> & Pick<INodeProperties, 'displayName' | 'name'>,
): INodeProperties {
	return pathRLC(
		{
			default: { mode: 'list', value: '' },
			required: false,
			...overrides,
		},
		'searchListFolders',
	);
}

/** Merges the picked field names with the free-text names, without duplicates. */
export function mergeFieldNames(picked: unknown, text: unknown): string[] {
	const toArray = (value: unknown): string[] =>
		Array.isArray(value) ? value.map(String) : String(value ?? '').split(',');
	return [...new Set([...toArray(picked), ...toArray(text)].map((v) => v.trim()).filter(Boolean))];
}

/** A column group picker: an existing group, or a typed name. */
export const groupRLC: INodeProperties = {
	displayName: 'Group',
	name: 'group',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	description: 'The column group of the field',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchColumnGroups', searchable: true },
		},
		{ displayName: 'By Name', name: 'name', type: 'string', placeholder: 'Custom Columns' },
	],
};
