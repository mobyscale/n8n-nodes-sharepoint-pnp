import type {
	IDataObject,
	IDisplayOptions,
	IExecuteFunctions,
	INodeProperties,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import '@pnp/sp/regional-settings/index.js';
import type { IItem } from '@pnp/sp/items/index.js';
import type { IList } from '@pnp/sp/lists/index.js';
import type { SPFI } from '@pnp/sp/index.js';

import { getList, getPath, listRLC, pathRLC } from '../../shared/descriptions';
import { toServerRelativePath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
import {
	APPENDABLE_TYPES,
	buildFormValues,
	buildViewXml,
	escapeODataString,
	expandFromSelect,
	mergeMultiValues,
	parseCommaList,
	type FieldFormValue,
	type FieldInfo,
} from './helpers';
import {
	createTagClient,
	hasTaxonomyValues,
	resolveTaxonomyValues,
} from './taxonomy';

const RESOURCE = 'listItem';
const NEEDS_TARGET = ['get', 'update', 'delete'];
const STREAM_PAGE_SIZE = 500;
const MAX_ROW_LIMIT = 5000;
const DEFAULT_SELECT = ['*', 'FileLeafRef', 'FileRef', 'FileDirRef', 'FSObjType'];

const show = (
	operation: string[],
	extra: Record<string, Array<string | boolean>> = {},
): IDisplayOptions => ({
	show: { resource: [RESOURCE], operation, ...extra },
});

const fieldsDescription = (
	operation: 'create' | 'update',
	extra: Record<string, string[]> = {},
) => {
	const verb = operation === 'create' ? 'set on the new item' : 'change';
	return [
		{
			displayName: 'Fields Input',
			name: 'fieldsInputMode',
			type: 'options',
			options: [
				{ name: 'Define Below', value: 'mapper' },
				{ name: 'JSON', value: 'json' },
			],
			default: 'mapper',
			description: `How to give the fields to ${verb}`,
			displayOptions: show([operation], extra),
		},
		{
			displayName: 'Fields',
			name: 'fields',
			type: 'resourceMapper',
			noDataExpression: true,
			default: { mappingMode: 'defineBelow', value: null },
			required: true,
			typeOptions: {
				loadOptionsDependsOn: ['list.value', 'site.value'],
				resourceMapper: {
					resourceMapperMethod: 'getListItemFields',
					mode: 'add',
					fieldWords: { singular: 'field', plural: 'fields' },
					addAllFields: false,
					multiKeyMatch: false,
					supportAutoMap: false,
				},
			},
			displayOptions: show([operation], { ...extra, fieldsInputMode: ['mapper'] }),
		},
		{
			displayName: 'Fields (JSON)',
			name: 'fieldsJson',
			type: 'json',
			default: '{}',
			required: true,
			placeholder: '{"Title": "Example", "Status": "Open"}',
			description: 'An object with the internal field names as keys. Values use the same format as the field mapper: arrays for multi-value fields, ISO text for dates.',
			displayOptions: show([operation], { ...extra, fieldsInputMode: ['json'] }),
		},
	] as INodeProperties[];
};

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a list item or a folder',
				action: 'Create a list item',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a list item or move it to the recycle bin',
				action: 'Delete a list item',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get a list item',
				action: 'Get a list item',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many list items',
				action: 'Get many list items',
			},
			{
				name: 'Update',
				value: 'update',
				description: 'Update the fields of a list item',
				action: 'Update a list item',
			},
		],
		default: 'get',
	},
	{
		displayName: 'Identify Item',
		name: 'identifyBy',
		type: 'options',
		options: [
			{ name: 'By ID', value: 'id', description: 'The list and the item ID' },
			{
				name: 'By File or Folder Path',
				value: 'path',
				description: 'The path of a file or folder in a document library',
			},
		],
		default: 'id',
		displayOptions: show(NEEDS_TARGET),
	},
	{
		...listRLC,
		displayOptions: show(['create', 'getAll']),
	},
	{
		...listRLC,
		displayOptions: show(NEEDS_TARGET, { identifyBy: ['id'] }),
	},
	{
		displayName: 'Item ID',
		name: 'itemId',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 1,
		required: true,
		description: 'The numeric ID of the list item',
		displayOptions: show(NEEDS_TARGET, { identifyBy: ['id'] }),
	},
	pathRLC({
		displayName: 'File or Folder',
		name: 'itemPath',
		displayOptions: show(NEEDS_TARGET, { identifyBy: ['path'] }),
	}),

	...fieldsDescription('create'),
	...fieldsDescription('update', { identifyBy: ['id'] }),
	{
		displayName: 'Fields (JSON)',
		name: 'fieldsJson',
		type: 'json',
		default: '{}',
		required: true,
		placeholder: '{"Title": "Example"}',
		description: 'An object with the internal field names as keys. The field mapper needs a list, so use JSON when you identify the item by path.',
		displayOptions: show(['update'], { identifyBy: ['path'] }),
	},

	{
		displayName: 'Query Mode',
		name: 'queryMode',
		type: 'options',
		options: [
			{ name: 'Filter', value: 'filter', description: 'Use OData filter and sort options' },
			{ name: 'CAML Query', value: 'caml', description: 'Use a CAML query (ViewXml)' },
		],
		default: 'filter',
		displayOptions: show(['getAll']),
	},
	{
		displayName: 'Filter',
		name: 'odataFilter',
		type: 'string',
		default: '',
		placeholder: "Status eq 'Open'",
		description: 'An OData $filter expression. Use internal field names.',
		displayOptions: show(['getAll'], { queryMode: ['filter'] }),
	},
	{
		displayName: 'Order By',
		name: 'orderBy',
		type: 'string',
		default: '',
		placeholder: 'Modified desc',
		description: 'An OData $orderby expression, for example "Modified desc"',
		displayOptions: show(['getAll'], { queryMode: ['filter'] }),
	},
	{
		displayName: 'Select Fields',
		name: 'selectFields',
		type: 'string',
		default: '',
		placeholder: 'ID,Title,Author/Title',
		description:
			'A comma-separated list of internal field names. Empty returns all fields. A name with a slash, such as "Author/Title", is expanded automatically',
		displayOptions: show(['getAll'], { queryMode: ['filter'] }),
	},
	{
		displayName: 'CAML Query',
		name: 'camlQuery',
		type: 'string',
		typeOptions: { rows: 5 },
		default: '',
		required: true,
		placeholder: '<Where><Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq></Where>',
		description: 'The content of the Query element (Where, OrderBy), or a full View element. Large results are read in pages.',
		displayOptions: show(['getAll'], { queryMode: ['caml'] }),
	},
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['getAll']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: show(['getAll'], { returnAll: [false] }),
	},

	{
		displayName: 'Options',
		name: 'getOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: [
			{
				displayName: 'Fields to Select',
				name: 'selectFields',
				type: 'string',
				default: '',
				placeholder: 'ID,Title,Author/Title',
				description:
					'A comma-separated list of internal field names. Empty returns all fields. A name with a slash, such as "Author/Title", is expanded automatically',
			},
			{
				displayName: 'Return Values As',
				name: 'returnValuesAs',
				type: 'options',
				options: [
					{ name: 'Raw', value: 'raw', description: 'The stored values' },
					{
						name: 'Text',
						value: 'text',
						description: 'The values as display text (FieldValuesAsText), for example user names',
					},
				],
				default: 'raw',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'getAllOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: [
			{
				displayName: 'Folder Path',
				name: 'folderPath',
				type: 'string',
				default: '',
				placeholder: 'Shared Documents/Folder',
				description: 'Only return items in this folder. A server-relative path, a site-relative path or an absolute URL.',
			},
			{
				displayName: 'Recursive',
				name: 'recursive',
				type: 'boolean',
				default: false,
				description:
					'Whether to include items in subfolders of the folder path. In filter mode this uses the FileDirRef field, which can fail on lists above the list view threshold. Use CAML mode for those lists',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'createOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['create']),
		options: [
			{
				displayName: 'Create as Folder',
				name: 'asFolder',
				type: 'boolean',
				default: false,
				description:
					'Whether to create a folder instead of an item. Set the folder name with the field FileLeafRef. A document library cannot create files here, only folders',
			},
			{
				displayName: 'Create Missing Terms',
				name: 'createMissingTerms',
				type: 'boolean',
				default: false,
				description:
					'Whether to create a term when a managed metadata (tag) value is a label that does not exist in the term set of the field. The term set must be open for new terms. If off, an unknown label fails.',
			},
			{
				displayName: 'Folder Path',
				name: 'folderPath',
				type: 'string',
				default: '',
				placeholder: 'Shared Documents/Folder',
				description:
					'Create the item in this folder. A server-relative path, a site-relative path or an absolute URL. Empty creates it in the list root',
			},
		],
	},
	{
		displayName: 'Update Fields',
		name: 'updateOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['update']),
		options: [
			{
				displayName: 'Append to Multi-Value Fields',
				name: 'appendToMultiValue',
				type: 'boolean',
				default: false,
				description: 'Whether to add the new values to the current values of multi-choice, multi-lookup and multi-value managed metadata (tag) fields instead of replacing them. Other field types are replaced.',
			},
			{
				displayName: 'Create Missing Terms',
				name: 'createMissingTerms',
				type: 'boolean',
				default: false,
				description:
					'Whether to create a term when a managed metadata (tag) value is a label that does not exist in the term set of the field. The term set must be open for new terms. If off, an unknown label fails.',
			},
			{
				displayName: 'System Update',
				name: 'systemUpdate',
				type: 'boolean',
				default: false,
				description:
					'Whether to update the item without a change to Modified, Modified By or the version number. This uses validateUpdateListItem with bNewDocumentUpdate set to true. The item is not checked out or checked in, and the update does not start "item updated" workflows or alerts',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'deleteOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['delete']),
		options: [
			{
				displayName: 'Delete Permanently',
				name: 'permanent',
				type: 'boolean',
				default: false,
				description: 'Whether to delete the item for good. If off, the item goes to the recycle bin.',
			},
		],
	},
];

async function getFieldInfos(list: IList): Promise<FieldInfo[]> {
	return (await list.fields.select(
		'InternalName',
		'TypeAsString',
		'TermSetId',
		'SspId',
		'AnchorId',
	)()) as FieldInfo[];
}

async function resolveItemByPath(sp: SPFI, path: string): Promise<{ list: IList; itemId: number }> {
	const folder = await sp.web.getFolderByServerRelativePath(path).select('Exists')();
	const itemFields = folder.Exists
		? sp.web.getFolderByServerRelativePath(path).listItemAllFields
		: sp.web.getFileByServerRelativePath(path).listItemAllFields;
	const info = (await itemFields.select('Id', 'ParentList/Id').expand('ParentList')()) as {
		Id?: number;
		ParentList?: { Id: string };
	};
	if (!info.Id || !info.ParentList) {
		throw new Error(`No list item found for the path "${path}"`);
	}
	return { list: sp.web.lists.getById(info.ParentList.Id), itemId: info.Id };
}

async function resolveTarget(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
): Promise<{ list: IList; itemId: number }> {
	const identifyBy = this.getNodeParameter('identifyBy', itemIndex, 'id') as string;
	if (identifyBy === 'path') {
		return await resolveItemByPath(sp, getPath.call(this, siteUrl, itemIndex, 'itemPath'));
	}
	return {
		list: getList.call(this, sp, itemIndex),
		itemId: this.getNodeParameter('itemId', itemIndex) as number,
	};
}

function readFieldValues(
	this: IExecuteFunctions,
	itemIndex: number,
	operation: 'create' | 'update',
): Record<string, unknown> {
	const identifyBy = this.getNodeParameter('identifyBy', itemIndex, 'id') as string;
	const inputMode =
		operation === 'update' && identifyBy === 'path'
			? 'json'
			: (this.getNodeParameter('fieldsInputMode', itemIndex, 'mapper') as string);

	if (inputMode === 'json') {
		const raw = this.getNodeParameter('fieldsJson', itemIndex, '{}') as string | IDataObject;
		let parsed: unknown = raw;
		if (typeof raw === 'string') {
			try {
				parsed = JSON.parse(raw);
			} catch {
				throw new NodeOperationError(
					this.getNode(),
					'The "Fields (JSON)" value is not valid JSON',
					{
						itemIndex,
					},
				);
			}
		}
		if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
			throw new NodeOperationError(this.getNode(), 'The "Fields (JSON)" value must be an object', {
				itemIndex,
			});
		}
		return parsed as Record<string, unknown>;
	}

	const mapper = this.getNodeParameter('fields', itemIndex, {}) as {
		value?: Record<string, unknown> | null;
	};
	return mapper.value ?? {};
}

function assertNoFieldErrors(
	this: IExecuteFunctions,
	results: Array<{ FieldName?: string; HasException?: boolean; ErrorMessage?: string }>,
	itemIndex: number,
): void {
	const failed = results.filter((entry) => entry.HasException);
	if (failed.length) {
		const message = failed
			.map(
				(entry) =>
					`${entry.FieldName ?? 'Unknown field'}: ${entry.ErrorMessage ?? 'invalid value'}`,
			)
			.join('; ');
		throw new NodeOperationError(
			this.getNode(),
			`SharePoint rejected the field values. ${message}`,
			{
				itemIndex,
			},
		);
	}
}

async function resolveTags(
	this: IExecuteFunctions,
	siteUrl: string,
	values: Record<string, unknown>,
	fieldInfos: FieldInfo[],
	createMissing: boolean,
	current?: Record<string, unknown>,
): Promise<Record<string, unknown>> {
	if (!hasTaxonomyValues(values, fieldInfos)) return values;
	return await resolveTaxonomyValues(values, fieldInfos, createTagClient(this, siteUrl), {
		createMissing,
		current,
	});
}

async function create(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const list = getList.call(this, sp, itemIndex);
	const options = this.getNodeParameter('createOptions', itemIndex, {}) as {
		asFolder?: boolean;
		folderPath?: string;
		createMissingTerms?: boolean;
	};
	const fieldInfos = await getFieldInfos(list);
	const values = await resolveTags.call(
		this,
		siteUrl,
		readFieldValues.call(this, itemIndex, 'create'),
		fieldInfos,
		options.createMissingTerms === true,
	);
	const formValues = buildFormValues(values, fieldInfos);

	const rootFolder = await list.rootFolder.select('ServerRelativeUrl')();
	const folderPath = options.folderPath?.trim()
		? toServerRelativePath(siteUrl, String(options.folderPath))
		: rootFolder.ServerRelativeUrl;

	const results = await list.addValidateUpdateItemUsingPath(
		formValues,
		folderPath,
		false,
		undefined,
		options.asFolder ? { objectType: 1 } : undefined,
	);
	assertNoFieldErrors.call(this, results, itemIndex);

	const idEntry = results.find((entry) => entry.FieldName === 'Id');
	const id = Number(idEntry?.FieldValue);
	if (!Number.isFinite(id) || id <= 0) {
		throw new NodeOperationError(
			this.getNode(),
			'SharePoint did not return the ID of the new item',
			{
				itemIndex,
			},
		);
	}
	return (await list.items.getById(id).select(...DEFAULT_SELECT)()) as IDataObject;
}

async function get(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const { list, itemId } = await resolveTarget.call(this, context);
	const options = this.getNodeParameter('getOptions', itemIndex, {}) as {
		selectFields?: string;
		returnValuesAs?: string;
	};
	const item = list.items.getById(itemId);

	if (options.returnValuesAs === 'text') {
		return (await item.fieldValuesAsText()) as IDataObject;
	}
	const select = parseCommaList(options.selectFields);
	if (!select.length) return (await item.select(...DEFAULT_SELECT)()) as IDataObject;
	return (await item.select(...select).expand(...expandFromSelect(select))()) as IDataObject;
}

function buildFolderFilter(folderPath: string, recursive: boolean): string {
	const escaped = escapeODataString(folderPath);
	return recursive
		? `(FileDirRef eq '${escaped}' or startswith(FileDirRef,'${escaped}/'))`
		: `FileDirRef eq '${escaped}'`;
}

async function readCamlPages(
	list: IList,
	viewXml: string,
	folderPath: string | undefined,
	max: number,
): Promise<IDataObject[]> {
	const parameters = {
		ViewXml: viewXml,
		RenderOptions: 2,
		...(folderPath ? { FolderServerRelativeUrl: folderPath } : {}),
	};
	const rows: IDataObject[] = [];
	let page = await list.renderListDataAsStream(parameters);
	rows.push(...(page.Row as IDataObject[]));
	while (page.NextHref && rows.length < max) {
		const query = new Map(new URLSearchParams(page.NextHref.replace(/^\?/, '')).entries());
		page = await list.renderListDataAsStream(parameters, undefined, query);
		rows.push(...(page.Row as IDataObject[]));
	}
	return rows.slice(0, max);
}

async function getAll(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const list = getList.call(this, sp, itemIndex);
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const limit = returnAll
		? Number.POSITIVE_INFINITY
		: (this.getNodeParameter('limit', itemIndex) as number);
	const queryMode = this.getNodeParameter('queryMode', itemIndex) as string;
	const options = this.getNodeParameter('getAllOptions', itemIndex, {}) as {
		folderPath?: string;
		recursive?: boolean;
	};
	const folderPath = options.folderPath?.trim()
		? toServerRelativePath(siteUrl, String(options.folderPath))
		: undefined;
	const recursive = options.recursive === true;

	if (queryMode === 'caml') {
		const caml = this.getNodeParameter('camlQuery', itemIndex) as string;
		const pageSize = returnAll ? STREAM_PAGE_SIZE : Math.min(limit, MAX_ROW_LIMIT);
		return await readCamlPages(
			list,
			buildViewXml(caml, { recursive, rowLimit: pageSize }),
			folderPath,
			limit,
		);
	}

	const requested = parseCommaList(this.getNodeParameter('selectFields', itemIndex, '') as string);
	const select = requested.length ? requested : DEFAULT_SELECT;
	const conditions: string[] = [];
	const odataFilter = (this.getNodeParameter('odataFilter', itemIndex, '') as string).trim();
	if (odataFilter) conditions.push(`(${odataFilter})`);
	if (folderPath) conditions.push(buildFolderFilter(folderPath, recursive));
	const orderBy = (this.getNodeParameter('orderBy', itemIndex, '') as string).trim();

	let query = list.items;
	query = query.select(...select).expand(...expandFromSelect(select));
	if (conditions.length) query = query.filter(conditions.join(' and '));
	if (orderBy) {
		const [field, direction] = orderBy.split(/\s+/);
		query = query.orderBy(field, direction?.toLowerCase() !== 'desc');
	}

	if (!returnAll) return (await query.top(Math.min(limit, MAX_ROW_LIMIT))()) as IDataObject[];

	const results: IDataObject[] = [];
	for await (const page of query.top(MAX_ROW_LIMIT)) results.push(...(page as IDataObject[]));
	return results;
}

async function getMultiValueCurrent(
	list: IList,
	itemId: number,
	names: string[],
	types: Map<string, string>,
): Promise<Record<string, unknown>> {
	const selects = names.map((name) => (types.get(name) === 'LookupMulti' ? `${name}Id` : name));
	const current = (await list.items.getById(itemId).select(...selects)()) as Record<
		string,
		unknown
	>;
	return Object.fromEntries(
		names.map((name) => [
			name,
			types.get(name) === 'LookupMulti' ? current[`${name}Id`] : current[name],
		]),
	);
}

/** Reads the current tags of multi-value managed metadata fields, for "Append to Multi-Value Fields". */
async function getTagCurrent(
	list: IList,
	itemId: number,
	names: string[],
	types: Map<string, string>,
): Promise<Record<string, unknown> | undefined> {
	const tagFields = names.filter((name) => types.get(name) === 'TaxonomyFieldTypeMulti');
	if (!tagFields.length) return undefined;
	const current = (await list.items.getById(itemId).select(...tagFields)()) as Record<string, unknown>;
	return Object.fromEntries(
		tagFields.map((name) => {
			const value = current[name];
			const results = (value as { results?: unknown } | null)?.results;
			return [name, Array.isArray(results) ? results : value];
		}),
	);
}

/**
 * validateUpdateListItem always sets Modified and Editor. A system update sends the current values
 * back, so they do not change. SharePoint reads the date in the time zone of the site, and a user
 * can only be sent by sign-in name (an app principal cannot be resolved, so Editor may change).
 */
async function getSystemUpdateValues(
	sp: SPFI,
	item: () => IItem,
	formValues: FieldFormValue[],
): Promise<FieldFormValue[]> {
	const current = (await item().select('Modified', 'Editor/Name').expand('Editor')()) as {
		Modified?: string;
		Editor?: { Name?: string };
	};
	const given = new Set(formValues.map((entry) => entry.FieldName));
	const extra: FieldFormValue[] = [];
	if (current.Modified && !given.has('Modified')) {
		const local = await sp.web.regionalSettings.timeZone.utcToLocalTime(current.Modified);
		extra.push({
			FieldName: 'Modified',
			FieldValue: local.replace('T', ' ').replace(/\.\d+$/, ''),
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

async function update(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const { list, itemId } = await resolveTarget.call(this, context);
	const options = this.getNodeParameter('updateOptions', itemIndex, {}) as {
		systemUpdate?: boolean;
		appendToMultiValue?: boolean;
		createMissingTerms?: boolean;
	};
	let values = readFieldValues.call(this, itemIndex, 'update');
	const fieldInfos = await getFieldInfos(list);

	if (options.appendToMultiValue) {
		const types = new Map(fieldInfos.map((field) => [field.InternalName, field.TypeAsString]));
		const appendable = Object.keys(values).filter((name) =>
			APPENDABLE_TYPES.has(types.get(name) ?? ''),
		);
		if (appendable.length) {
			const current = await getMultiValueCurrent(list, itemId, appendable, types);
			values = { ...values };
			for (const name of appendable) values[name] = mergeMultiValues(current[name], values[name]);
		}
	}

	const fieldTypes = new Map(fieldInfos.map((field) => [field.InternalName, field.TypeAsString]));
	const currentTags = options.appendToMultiValue
		? await getTagCurrent(list, itemId, Object.keys(values), fieldTypes)
		: undefined;
	values = await resolveTags.call(
		this,
		context.siteUrl,
		values,
		fieldInfos,
		options.createMissingTerms === true,
		currentTags,
	);

	const formValues: FieldFormValue[] = buildFormValues(values, fieldInfos);
	if (!formValues.length) {
		throw new NodeOperationError(this.getNode(), 'No fields to update were given', { itemIndex });
	}

	// PnPjs queryables keep their select and expand state, so each request needs a new instance.
	const item = () => list.items.getById(itemId);
	if (options.systemUpdate) {
		formValues.push(...(await getSystemUpdateValues(context.sp, item, formValues)));
	}
	const results = await item().validateUpdateListItem(formValues, options.systemUpdate === true);
	assertNoFieldErrors.call(this, results, itemIndex);
	return (await item().select(...DEFAULT_SELECT)()) as IDataObject;
}

async function deleteItem(this: IExecuteFunctions, context: OperationContext) {
	const { itemIndex } = context;
	const { list, itemId } = await resolveTarget.call(this, context);
	const options = this.getNodeParameter('deleteOptions', itemIndex, {}) as { permanent?: boolean };
	const item = list.items.getById(itemId);
	if (options.permanent) await item.delete();
	else await item.recycle();
	return { success: true, id: itemId, recycled: !options.permanent } as IDataObject;
}

export const listItemResource: ResourceModule = {
	option: {
		name: 'List Item',
		value: RESOURCE,
		description: 'Create, query and update list items and their metadata',
	},
	description,
	operations: { create, get, getAll, update, delete: deleteItem },
};
