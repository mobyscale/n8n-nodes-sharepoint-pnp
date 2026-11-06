import type { IDataObject, IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { IFields } from '@pnp/sp/fields/index.js';
import type { IContentTypes } from '@pnp/sp/content-types/index.js';
import type { SPFI } from '@pnp/sp/index.js';

import { getList, getListFromLocator } from '../../shared/descriptions';
import { applySelection, type SelectionValues } from '../../shared/selection';
import { getRequiredString } from '../../shared/utils';
import type { OperationContext } from '../types';
import { CONTENT_TYPE_SELECT } from './setupDescription';
import {
	buildContentTypeOrder,
	buildFieldPlan,
	buildFieldUpdate,
	buildListUpdate,
	restTypeForFieldKind,
	type CommonFieldSettings,
	type FieldType,
	type FieldTypeSettings,
	type FieldUpdateOptions,
	type ListUpdateOptions,
} from './setupHelpers';

const FIELD_RESULT_SELECT = [
	'Id',
	'Title',
	'InternalName',
	'TypeAsString',
	'FieldTypeKind',
	'Required',
	'Indexed',
	'Group',
	'Description',
	'DefaultValue',
];

interface LocatorValue {
	mode: string;
	value: string;
}

function isSiteTarget(ctx: IExecuteFunctions, itemIndex: number): boolean {
	return ctx.getNodeParameter('target', itemIndex, 'list') === 'site';
}

/** The fields collection of the list or of the site. PnPjs keeps query state, so call this for each request. */
function targetFields(ctx: IExecuteFunctions, sp: SPFI, itemIndex: number): IFields {
	return isSiteTarget(ctx, itemIndex)
		? sp.web.fields
		: getList.call(ctx, sp, itemIndex).fields;
}

function targetContentTypes(ctx: IExecuteFunctions, sp: SPFI, itemIndex: number): IContentTypes {
	return isSiteTarget(ctx, itemIndex)
		? sp.web.contentTypes
		: getList.call(ctx, sp, itemIndex).contentTypes;
}

function readFieldType(ctx: IExecuteFunctions, itemIndex: number): FieldTypeSettings {
	const read = <T>(name: string) => ctx.getNodeParameter(name, itemIndex, undefined) as T | undefined;
	const lookupList = read<LocatorValue>('lookupList');
	return {
		maxLength: read('maxLength'),
		numberOfLines: read('numberOfLines'),
		richText: read('richText'),
		appendOnly: read('appendOnly'),
		minimumValue: read('minimumValue'),
		maximumValue: read('maximumValue'),
		currencyLocaleId: read('currencyLocaleId'),
		dateTimeFormat: read('dateTimeFormat'),
		choices: read('choices'),
		choiceFormat: read('choiceFormat'),
		fillInChoice: read('fillInChoice'),
		userSelectionMode: read('userSelectionMode'),
		lookupField: read('lookupField'),
		lookupAllowMultiple: read('lookupAllowMultiple'),
		urlFormat: read('urlFormat'),
		formula: read('formula'),
		calculatedOutputType: read('calculatedOutputType'),
		schemaXml: read('schemaXml'),
		lookupListId: lookupList ? lookupList.value : undefined,
	};
}

export async function update(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('updateOptions', itemIndex, {}) as ListUpdateOptions;
	const title = (this.getNodeParameter('newTitle', itemIndex, '') as string).trim();
	const properties = buildListUpdate({ ...options, ...(title ? { title } : {}) });
	if (!Object.keys(properties).length) {
		throw new NodeOperationError(this.getNode(), 'Set a title or at least one option to update', {
			itemIndex,
		});
	}

	await getList.call(this, sp, itemIndex).update(properties);
	return (await getList
		.call(this, sp, itemIndex)
		.select(
			'Id',
			'Title',
			'Description',
			'EnableVersioning',
			'MajorVersionLimit',
			'EnableMinorVersions',
			'MajorWithMinorVersionsLimit',
			'EnableModeration',
			'ForceCheckout',
			'EnableAttachments',
			'OnQuickLaunch',
			'Hidden',
			'NoCrawl',
			'EnableFolderCreation',
			'ContentTypesEnabled',
		)()) as unknown as IDataObject;
}

export async function createField(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const type = this.getNodeParameter('fieldType', itemIndex) as FieldType;
	const common = this.getNodeParameter('commonSettings', itemIndex, {}) as Omit<
		CommonFieldSettings,
		'displayName'
	> & { addToDefaultView?: boolean };
	const displayName =
		type === 'schemaXml' ? '' : getRequiredString.call(this, 'displayName', itemIndex);
	const plan = buildFieldPlan(type, { ...common, displayName }, readFieldType(this, itemIndex));

	const fields = () => targetFields(this, sp, itemIndex);
	let created: Record<string, unknown>;
	if (plan.method === 'createFieldAsXml') {
		created = (await fields().createFieldAsXml({
			SchemaXml: plan.schemaXml!,
			Options: common.addToDefaultView && !isSiteTarget(this, itemIndex) ? 16 : 0,
		})) as Record<string, unknown>;
	} else {
		const collection = fields();
		const add = collection[plan.method] as (title: string, props: unknown) => Promise<unknown>;
		created = (await add.call(collection, plan.title, plan.properties)) as Record<string, unknown>;
	}

	const id = String(created.Id ?? (created.d as { Id?: string } | undefined)?.Id ?? '');
	const locate = () =>
		id ? fields().getById(id) : fields().getByInternalNameOrTitle(plan.title);

	if (plan.renameTo) await locate().update({ Title: plan.renameTo });
	const result = (await locate().select(...FIELD_RESULT_SELECT)()) as unknown as IDataObject;

	if (common.addToDefaultView && !isSiteTarget(this, itemIndex) && plan.method !== 'createFieldAsXml') {
		await getList.call(this, sp, itemIndex).defaultView.fields.add(String(result.InternalName));
	}
	return result;
}

export async function updateField(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const name = getRequiredString.call(this, 'field', itemIndex);
	const options = this.getNodeParameter('updateFieldOptions', itemIndex, {}) as FieldUpdateOptions;
	const properties = buildFieldUpdate(options);
	if (!Object.keys(properties).length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one option to update', {
			itemIndex,
		});
	}

	const field = () => targetFields(this, sp, itemIndex).getByInternalNameOrTitle(name);
	const info = (await field().select('FieldTypeKind')()) as { FieldTypeKind: number };
	await field().update(properties, restTypeForFieldKind(info.FieldTypeKind));
	return (await field().select(...FIELD_RESULT_SELECT)()) as unknown as IDataObject;
}

export async function deleteField(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const name = getRequiredString.call(this, 'field', itemIndex);
	const field = () => targetFields(this, sp, itemIndex).getByInternalNameOrTitle(name);
	const info = (await field().select('Id', 'Title', 'InternalName')()) as {
		Id: string;
		Title: string;
		InternalName: string;
	};
	await field().delete();
	return {
		success: true,
		id: info.Id,
		title: info.Title,
		internalName: info.InternalName,
	} as IDataObject;
}

export async function getContentTypes(
	this: IExecuteFunctions,
	{ sp, itemIndex }: OperationContext,
) {
	const options = this.getNodeParameter(
		'getContentTypesOptions',
		itemIndex,
		{},
	) as SelectionValues & { includeHidden?: boolean };
	let query = applySelection(targetContentTypes(this, sp, itemIndex), CONTENT_TYPE_SELECT, options);
	if (!options.includeHidden) query = query.filter('Hidden eq false');
	return (await query()) as unknown as IDataObject[];
}

export async function addContentType(
	this: IExecuteFunctions,
	{ sp, itemIndex }: OperationContext,
) {
	const contentTypeId = getRequiredString.call(this, 'siteContentType', itemIndex);
	const options = this.getNodeParameter('addContentTypeOptions', itemIndex, {}) as {
		enableContentTypes?: boolean;
	};
	const list = () => getList.call(this, sp, itemIndex);

	if (options.enableContentTypes !== false) {
		const info = (await list().select('ContentTypesEnabled')()) as { ContentTypesEnabled: boolean };
		if (!info.ContentTypesEnabled) await list().update({ ContentTypesEnabled: true });
	}
	await list().contentTypes.addAvailableContentType(contentTypeId);
	return (await list().contentTypes.select(...CONTENT_TYPE_SELECT)()) as unknown as IDataObject[];
}

export async function removeContentType(
	this: IExecuteFunctions,
	{ sp, itemIndex }: OperationContext,
) {
	const contentTypeId = getRequiredString.call(this, 'listContentType', itemIndex);
	const list = () => getList.call(this, sp, itemIndex);
	const info = (await list()
		.contentTypes.getById(contentTypeId)
		.select('StringId', 'Name')()) as { StringId: string; Name: string };
	await list().contentTypes.getById(contentTypeId).delete();
	return { success: true, id: info.StringId, name: info.Name } as IDataObject;
}

export async function setDefaultContentType(
	this: IExecuteFunctions,
	{ sp, itemIndex }: OperationContext,
) {
	const contentTypeId = getRequiredString.call(this, 'listContentType', itemIndex);
	const locator = this.getNodeParameter('list', itemIndex) as LocatorValue;
	const list = () => getListFromLocator(sp, locator);

	const folder = (await list().rootFolder.select('UniqueContentTypeOrder')()) as {
		UniqueContentTypeOrder?: Array<{ StringId: string }>;
	};
	const current = folder.UniqueContentTypeOrder?.length
		? folder.UniqueContentTypeOrder.map((entry) => entry.StringId)
		: (
				(await list().contentTypes.select('StringId')()) as Array<{ StringId: string }>
			).map((entry) => entry.StringId);
	const order = buildContentTypeOrder(current, contentTypeId);

	await list().rootFolder.update({
		UniqueContentTypeOrder: {
			__metadata: { type: 'Collection(SP.ContentTypeId)' },
			results: order.map((id) => ({ __metadata: { type: 'SP.ContentTypeId' }, StringId: id })),
		},
	} as never);
	return { success: true, defaultContentTypeId: contentTypeId, order } as IDataObject;
}
