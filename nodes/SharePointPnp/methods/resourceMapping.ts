import type {
	FieldType,
	ILoadOptionsFunctions,
	ResourceMapperField,
	ResourceMapperFields,
} from 'n8n-workflow';

import { getListFromLocator, getSiteUrl } from '../shared/descriptions';
import { createSp } from '../transport';

interface RawField {
	InternalName: string;
	Title: string;
	TypeAsString: string;
	Required: boolean;
	Choices?: string[];
}

const SKIPPED_FIELDS = new Set([
	'ContentType',
	'ContentTypeId',
	'Attachments',
	'Edit',
	'LinkTitleNoMenu',
]);

export function toMapperType(typeAsString: string): FieldType {
	switch (typeAsString) {
		case 'Integer':
		case 'Number':
		case 'Currency':
		case 'Counter':
			return 'number';
		case 'Boolean':
			return 'boolean';
		case 'DateTime':
			return 'dateTime';
		case 'Choice':
			return 'options';
		case 'MultiChoice':
		case 'LookupMulti':
		case 'UserMulti':
		case 'TaxonomyFieldTypeMulti':
			return 'array';
		default:
			return 'string';
	}
}

export async function getListItemFields(
	this: ILoadOptionsFunctions,
): Promise<ResourceMapperFields> {
	const sp = createSp(this, await getSiteUrl(this));
	const locator = this.getNodeParameter('list') as { mode: string; value: string };
	const fields = (await getListFromLocator(sp, locator)
		.fields.filter('Hidden eq false and ReadOnlyField eq false')
		.select('InternalName', 'Title', 'TypeAsString', 'Required', 'Choices')()) as RawField[];

	const result: ResourceMapperField[] = fields
		.filter((field) => !SKIPPED_FIELDS.has(field.InternalName))
		.sort((a, b) => a.Title.localeCompare(b.Title))
		.map((field) => {
			const type = toMapperType(field.TypeAsString);
			return {
				id: field.InternalName,
				displayName: `${field.Title} (${field.InternalName})`,
				defaultMatch: false,
				canBeUsedToMatch: false,
				required: field.Required,
				display: true,
				type,
				removed: !field.Required,
				...(type === 'options'
					? { options: (field.Choices ?? []).map((choice) => ({ name: choice, value: choice })) }
					: {}),
			};
		});

	return { fields: result };
}
