import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { IFieldDefault } from '@pnp/sp/column-defaults/index.js';

import { getList, getPath, listRLC } from '../../shared/descriptions';
import { getRequiredString, toServerRelativePath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
import {
	escapeXml,
	filterDefaultsByPath,
	normalizeFolderPath,
	mergeFolderDefaults,
	splitViewFields,
	type DefaultColumnValue,
} from './helpers';

const show = (operation: string[]) => ({ show: { resource: ['list'], operation } });

const LIST_SELECT = [
	'Id',
	'Title',
	'Description',
	'BaseTemplate',
	'BaseType',
	'ItemCount',
	'Hidden',
	'Created',
	'LastItemModifiedDate',
	'RootFolder/ServerRelativeUrl',
];

const FIELD_SELECT = [
	'Id',
	'Title',
	'InternalName',
	'StaticName',
	'TypeAsString',
	'FieldTypeKind',
	'Required',
	'Hidden',
	'ReadOnlyField',
	'Sealed',
	'Choices',
	'Group',
	'Description',
	'DefaultValue',
];

const VIEW_SELECT = [
	'Id',
	'Title',
	'ServerRelativeUrl',
	'DefaultView',
	'PersonalView',
	'Hidden',
	'Paged',
	'RowLimit',
	'Scope',
	'ViewQuery',
	'ViewType',
];

const COPIED_VIEW_SETTINGS = [
	'ViewQuery',
	'RowLimit',
	'Paged',
	'Scope',
	'Aggregations',
	'AggregationsStatus',
	'ViewData',
	'MobileView',
	'TabularView',
	'IncludeRootFolder',
	'OrderedView',
	'Threaded',
] as const;

const LIST_TEMPLATES = [
	{ name: 'Announcements', value: 104 },
	{ name: 'Contacts', value: 105 },
	{ name: 'Discussion Board', value: 108 },
	{ name: 'Document Library', value: 101 },
	{ name: 'Events (Calendar)', value: 106 },
	{ name: 'Generic List', value: 100 },
	{ name: 'Issue Tracking', value: 1100 },
	{ name: 'Links', value: 103 },
	{ name: 'Picture Library', value: 109 },
	{ name: 'Survey', value: 102 },
	{ name: 'Tasks', value: 107 },
];

const LIST_OPERATIONS = [
	'get',
	'delete',
	'getFields',
	'getViews',
	'createView',
	'copyView',
	'getDefaultColumnValues',
	'setDefaultColumnValues',
];

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['list'] } },
		options: [
			{
				name: 'Copy View',
				value: 'copyView',
				description: 'Copy a view from another list in the same site',
				action: 'Copy a view',
			},
			{
				name: 'Create',
				value: 'create',
				description: 'Create a list or library',
				action: 'Create a list',
			},
			{
				name: 'Create View',
				value: 'createView',
				description: 'Create a view in a list',
				action: 'Create a view',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a list or library, or move it to the recycle bin',
				action: 'Delete a list',
			},
			{ name: 'Get', value: 'get', description: 'Get a list or library', action: 'Get a list' },
			{
				name: 'Get Default Column Values',
				value: 'getDefaultColumnValues',
				description: 'Get the default column values of a library',
				action: 'Get default column values',
			},
			{
				name: 'Get Fields',
				value: 'getFields',
				description: 'Get the fields (columns) of a list',
				action: 'Get fields of a list',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many lists and libraries in a site',
				action: 'Get many lists',
			},
			{
				name: 'Get Views',
				value: 'getViews',
				description: 'Get the views of a list',
				action: 'Get views of a list',
			},
			{
				name: 'Set Default Column Values',
				value: 'setDefaultColumnValues',
				description: 'Set the default column values of a folder in a library',
				action: 'Set default column values',
			},
		],
		default: 'getAll',
	},
	{ ...listRLC, displayOptions: show(LIST_OPERATIONS) },

	{
		displayName: 'Title',
		name: 'title',
		type: 'string',
		default: '',
		required: true,
		description: 'The title of the new list',
		displayOptions: show(['create']),
	},
	{
		displayName: 'Template',
		name: 'template',
		type: 'options',
		options: LIST_TEMPLATES,
		default: 100,
		description: 'The kind of list to create',
		displayOptions: show(['create']),
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
				displayName: 'Description',
				name: 'description',
				type: 'string',
				default: '',
				description: 'The description of the list',
			},
			{
				displayName: 'Show on Quick Launch',
				name: 'onQuickLaunch',
				type: 'boolean',
				default: false,
				description: 'Whether to show the list in the site navigation (Quick Launch)',
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
				description: 'Whether to delete the list for good. If off, the list goes to the recycle bin.',
			},
		],
	},

	{
		displayName: 'Filters',
		name: 'fieldFilters',
		type: 'collection',
		placeholder: 'Add Filter',
		default: {},
		displayOptions: show(['getFields']),
		options: [
			{
				displayName: 'Include Hidden',
				name: 'includeHidden',
				type: 'boolean',
				default: false,
				description: 'Whether to include hidden fields',
			},
			{
				displayName: 'Include Read-Only',
				name: 'includeReadOnly',
				type: 'boolean',
				default: false,
				description: 'Whether to include read-only fields',
			},
		],
	},

	{
		displayName: 'Options',
		name: 'getViewsOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getViews']),
		options: [
			{
				displayName: 'Include Hidden',
				name: 'includeHidden',
				type: 'boolean',
				default: false,
				description: 'Whether to include hidden views',
			},
			{
				displayName: 'Include View Fields',
				name: 'includeFields',
				type: 'boolean',
				default: false,
				description: 'Whether to add the internal names of the fields of each view. This needs one more request for each view.',
			},
		],
	},

	{
		displayName: 'View Title',
		name: 'viewTitle',
		type: 'string',
		default: '',
		required: true,
		description: 'The title of the new view',
		displayOptions: show(['createView']),
	},
	{
		displayName: 'View Field Names or IDs',
		name: 'viewFields',
		type: 'multiOptions',
		typeOptions: { loadOptionsMethod: 'getListFieldOptions', loadOptionsDependsOn: ['list.value'] },
		default: [],
		description: 'The fields to show in the view, in the order of selection. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>. Empty keeps the default fields of SharePoint. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show(['createView']),
	},
	{
		displayName: 'Options',
		name: 'createViewOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createView']),
		options: [
			{
				displayName: 'Personal View',
				name: 'personalView',
				type: 'boolean',
				default: false,
				description: 'Whether the view is a personal view of the current user. If off, it is a public view.',
			},
			{
				displayName: 'Query (CAML)',
				name: 'viewQuery',
				type: 'string',
				typeOptions: { rows: 4 },
				default: '',
				placeholder:
					'<Where><Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq></Where>',
				description: 'The content of the Query element of the view (Where, OrderBy, GroupBy)',
			},
			{
				displayName: 'Row Limit',
				name: 'rowLimit',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 30,
				description: 'The number of items for each page of the view',
			},
		],
	},

	{
		displayName: 'Source List',
		name: 'sourceList',
		type: listRLC.type,
		default: listRLC.default,
		required: true,
		description: 'The list that has the view to copy. It must be in the same site.',
		modes: listRLC.modes,
		displayOptions: show(['copyView']),
	},
	{
		displayName: 'View to Copy Name or ID',
		name: 'sourceView',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getSourceViewOptions',
			loadOptionsDependsOn: ['sourceList.value'],
		},
		default: '',
		required: true,
		description: 'The view in the source list. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show(['copyView']),
	},
	{
		displayName: 'New View Title',
		name: 'newViewTitle',
		type: 'string',
		default: '',
		description: 'The title of the copy. Empty uses the title of the source view.',
		displayOptions: show(['copyView']),
	},

	{
		displayName: 'Folder Path',
		name: 'folderPath',
		type: 'string',
		default: '',
		required: true,
		placeholder: 'Shared Documents/Folder',
		description:
			'The folder that gets the defaults. A server-relative path, a site-relative path or an absolute URL. Use the library root for library-wide defaults',
		displayOptions: show(['setDefaultColumnValues']),
	},
	{
		displayName: 'Default Values',
		name: 'defaultValues',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Default Value',
		default: {},
		description:
			'The values to set. A default for the same field in the same folder is replaced. Taxonomy defaults are not supported',
		displayOptions: show(['setDefaultColumnValues']),
		options: [
			{
				displayName: 'Value',
				name: 'values',
				values: [
					{
						displayName: 'Field Internal Name',
						name: 'name',
						type: 'string',
						default: '',
						description: 'The internal name of the column',
					},
					{
						displayName: 'Value',
						name: 'value',
						type: 'string',
						default: '',
						description: 'The default value as text. For multi-choice, separate values with a semicolon.',
					},
				],
			},
		],
	},
	{
		displayName: 'Options',
		name: 'setDefaultsOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['setDefaultColumnValues']),
		options: [
			{
				displayName: 'Remove Other Defaults of This Folder',
				name: 'removeOthers',
				type: 'boolean',
				default: false,
				description: 'Whether to remove the existing defaults of this folder that are not in the list above. Other folders are not changed.',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'getDefaultsOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getDefaultColumnValues']),
		options: [
			{
				displayName: 'Folder Path',
				name: 'folderPath',
				type: 'string',
				default: '',
				placeholder: 'Shared Documents/Folder',
				description: 'Only return the defaults of this folder. Empty returns the defaults of all folders.',
			},
		],
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
		displayOptions: { show: { resource: ['list'], operation: ['getAll'], returnAll: [false] } },
	},
	{
		displayName: 'Filters',
		name: 'filters',
		type: 'collection',
		placeholder: 'Add Filter',
		default: {},
		displayOptions: show(['getAll']),
		options: [
			{
				displayName: 'Include Hidden',
				name: 'includeHidden',
				type: 'boolean',
				default: false,
				description: 'Whether to include hidden lists',
			},
			{
				displayName: 'Type',
				name: 'baseTemplate',
				type: 'options',
				default: -1,
				options: [
					{ name: 'Any', value: -1 },
					{ name: 'Document Library', value: 101 },
					{ name: 'Generic List', value: 100 },
					{ name: 'Picture Library', value: 109 },
					{ name: 'Site Pages', value: 119 },
				],
			},
		],
	},
];

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	return (await getList
		.call(this, sp, itemIndex)
		.select(...LIST_SELECT)
		.expand('RootFolder')()) as unknown as IDataObject;
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const filters = this.getNodeParameter('filters', itemIndex, {}) as {
		includeHidden?: boolean;
		baseTemplate?: number;
	};

	const conditions: string[] = [];
	if (!filters.includeHidden) conditions.push('Hidden eq false');
	if (filters.baseTemplate !== undefined && filters.baseTemplate !== -1) {
		conditions.push(`BaseTemplate eq ${filters.baseTemplate}`);
	}

	let query = sp.web.lists.select(...LIST_SELECT).expand('RootFolder');
	if (conditions.length) query = query.filter(conditions.join(' and '));
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);

	return (await query()) as unknown as IDataObject[];
}

async function create(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const title = getRequiredString.call(this, 'title', itemIndex);
	const template = this.getNodeParameter('template', itemIndex) as number;
	const options = this.getNodeParameter('createOptions', itemIndex, {}) as {
		description?: string;
		onQuickLaunch?: boolean;
	};

	const created = await sp.web.lists.add(title, options.description ?? '', template, false);
	const list = sp.web.lists.getById(created.Id);
	if (options.onQuickLaunch) await list.update({ OnQuickLaunch: true });
	return (await list.select(...LIST_SELECT).expand('RootFolder')()) as unknown as IDataObject;
}

async function deleteList(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('deleteOptions', itemIndex, {}) as { permanent?: boolean };
	const info = await getList.call(this, sp, itemIndex).select('Id', 'Title')();
	const list = getList.call(this, sp, itemIndex);
	if (options.permanent) await list.delete();
	else await list.recycle();
	return {
		success: true,
		id: info.Id,
		title: info.Title,
		recycled: !options.permanent,
	} as IDataObject;
}

async function getFields(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const filters = this.getNodeParameter('fieldFilters', itemIndex, {}) as {
		includeHidden?: boolean;
		includeReadOnly?: boolean;
	};
	const conditions: string[] = [];
	if (!filters.includeHidden) conditions.push('Hidden eq false');
	if (!filters.includeReadOnly) conditions.push('ReadOnlyField eq false');

	let query = getList.call(this, sp, itemIndex).fields.select(...FIELD_SELECT);
	if (conditions.length) query = query.filter(conditions.join(' and '));
	return (await query()) as unknown as IDataObject[];
}

async function getViews(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('getViewsOptions', itemIndex, {}) as {
		includeHidden?: boolean;
		includeFields?: boolean;
	};
	const list = getList.call(this, sp, itemIndex);
	let query = list.views.select(...VIEW_SELECT);
	if (!options.includeHidden) query = query.filter('Hidden eq false');
	const views = (await query()) as unknown as IDataObject[];

	if (options.includeFields) {
		for (const view of views) {
			const fields = (await list.views.getById(String(view.Id)).fields()) as unknown as {
				Items?: string[];
			};
			view.ViewFields = fields.Items ?? [];
		}
	}
	return views;
}

async function applyViewFields(
	list: ReturnType<typeof getList>,
	viewId: string,
	fields: string[],
): Promise<void> {
	const viewFields = list.views.getById(viewId).fields;
	await viewFields.removeAll();
	for (const field of fields) await viewFields.add(field);
}

async function createView(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const list = getList.call(this, sp, itemIndex);
	const title = getRequiredString.call(this, 'viewTitle', itemIndex);
	const fields = this.getNodeParameter('viewFields', itemIndex, []) as string[];
	const options = this.getNodeParameter('createViewOptions', itemIndex, {}) as {
		personalView?: boolean;
		viewQuery?: string;
		rowLimit?: number;
	};

	const settings: Record<string, unknown> = { Paged: true };
	if (options.viewQuery?.trim()) settings.ViewQuery = options.viewQuery.trim();
	if (options.rowLimit) settings.RowLimit = options.rowLimit;

	const created = (await list.views.add(title, options.personalView === true, settings)) as {
		Id: string;
	};
	if (fields.length) await applyViewFields(list, created.Id, fields);

	const view = (await list.views
		.getById(created.Id)
		.select(...VIEW_SELECT)()) as unknown as IDataObject;
	view.ViewFields =
		((await list.views.getById(created.Id).fields()) as { Items?: string[] }).Items ?? [];
	return view;
}

async function copyView(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const target = getList.call(this, sp, itemIndex);
	const source = getList.call(this, sp, itemIndex, 'sourceList');
	const sourceViewId = getRequiredString.call(this, 'sourceView', itemIndex);
	const newTitle = (this.getNodeParameter('newViewTitle', itemIndex, '') as string).trim();

	const sourceView = source.views.getById(sourceViewId);
	const info = (await sourceView.select(
		'Title',
		'PersonalView',
		...COPIED_VIEW_SETTINGS,
	)()) as Record<string, unknown>;
	const sourceFields = ((await sourceView.fields()) as { Items?: string[] }).Items ?? [];
	const targetFields = (await target.fields.select('InternalName')()) as Array<{
		InternalName: string;
	}>;
	const { kept, skipped } = splitViewFields(
		sourceFields,
		targetFields.map((field) => field.InternalName),
	);

	const settings: Record<string, unknown> = {};
	for (const key of COPIED_VIEW_SETTINGS) {
		if (info[key] !== null && info[key] !== undefined) settings[key] = info[key];
	}

	const created = (await target.views.add(
		newTitle || String(info.Title),
		info.PersonalView === true,
		settings,
	)) as { Id: string };
	await applyViewFields(target, created.Id, kept);

	const view = (await target.views
		.getById(created.Id)
		.select(...VIEW_SELECT)()) as unknown as IDataObject;
	view.ViewFields = kept;
	view.SkippedFields = skipped;
	return view;
}

async function getDefaultColumnValues(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const options = this.getNodeParameter('getDefaultsOptions', itemIndex, {}) as {
		folderPath?: string;
	};
	const folderPath = options.folderPath?.trim()
		? toServerRelativePath(siteUrl, String(options.folderPath))
		: undefined;
	const defaults = (await getList
		.call(this, sp, itemIndex)
		.getDefaultColumnValues()) as DefaultColumnValue[];
	return filterDefaultsByPath(defaults, folderPath) as unknown as IDataObject[];
}

async function setDefaultColumnValues(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	// PnPjs list methods keep query state, so each call needs a new instance.
	const list = () => getList.call(this, sp, itemIndex);
	const folderPath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const collection = this.getNodeParameter('defaultValues', itemIndex, {}) as {
		values?: Array<{ name: string; value: string }>;
	};
	const options = this.getNodeParameter('setDefaultsOptions', itemIndex, {}) as {
		removeOthers?: boolean;
	};
	const updates = (collection.values ?? []).filter((entry) => entry.name.trim());
	if (!updates.length && !options.removeOthers) {
		throw new NodeOperationError(this.getNode(), 'Add at least one default value', { itemIndex });
	}

	const existing = (await list().getDefaultColumnValues()) as DefaultColumnValue[];
	const merged = mergeFolderDefaults(
		existing,
		folderPath,
		updates.map((entry) => ({ name: entry.name.trim(), value: escapeXml(entry.value) })),
		options.removeOthers === true,
	);

	const fieldTypes = new Map(
		(
			(await list().fields.select('InternalName', 'TypeAsString')()) as Array<{
				InternalName: string;
				TypeAsString: string;
			}>
		).map((field) => [field.InternalName, field.TypeAsString]),
	);
	const taxonomy = merged.find(
		(entry) =>
			normalizeFolderPath(entry.path) !== normalizeFolderPath(folderPath) &&
			fieldTypes.get(entry.name)?.startsWith('Taxonomy'),
	);
	if (taxonomy) {
		throw new NodeOperationError(
			this.getNode(),
			`The library has a taxonomy default (${taxonomy.name}) for another folder. PnPjs cannot write it back, so the defaults are not changed`,
			{ itemIndex },
		);
	}

	await list().setDefaultColumnValues(merged as unknown as IFieldDefault[]);
	return filterDefaultsByPath(
		(await list().getDefaultColumnValues()) as DefaultColumnValue[],
		folderPath,
	) as unknown as IDataObject[];
}

export const listResource: ResourceModule = {
	option: {
		name: 'List',
		value: 'list',
		description: 'Get, create and manage lists and libraries',
	},
	description,
	operations: {
		get,
		getAll,
		create,
		delete: deleteList,
		getFields,
		getViews,
		createView,
		copyView,
		getDefaultColumnValues,
		setDefaultColumnValues,
	},
};
