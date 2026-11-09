import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import type { IList } from '@pnp/sp/lists/index.js';

import { getCopyJobStatus, getJobInfoParameter } from '../../shared/copyJob';
import { getPath, pathRLC } from '../../shared/descriptions';
import {
	buildCamlFilter,
	fieldFiltersCombineOption,
	fieldFiltersProperty,
	loadFilterFieldTypes,
	readCombine,
	readFieldFilters,
} from '../../shared/fieldFilters';
import {
	buildViewXml,
	findOwningList,
	folderExists,
	getListRoots,
	hasExtension,
	normalizeRow,
	parseExtensions,
	type FileFolderFilters,
	type ListRoot,
	VIEW_FIELDS,
	undefinedIfNotFound,
	type IncludeKind,
} from '../../shared/filesFolders';
import { selectionOptions, toList, type SelectionValues } from '../../shared/selection';
import { toNodeError, toServerRelativePath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
import {
	compareTrees,
	parsePatterns,
	readTree,
	type CompareInclude,
	type FolderReader,
	type TreeEntry,
	type TreeOptions,
} from './compare';
import { toDepthRange, walkFolders, type DepthRange } from './walk';

const show = (operation: string[]) => ({ show: { resource: ['fileFolder'], operation } });

const OUTPUT_PROPERTIES =
	'name, type, serverRelativeUrl, uniqueId, listItemId, size, created, modified, author, editor';

/** Get Many reads list item fields, so the options use field names and have no Expand. */
const getAllOptions = (): INodeProperties[] => {
	const [additional, , returnAll] = selectionOptions(VIEW_FIELDS);
	const recursiveOnly = { show: { '/recursive': [true] } };
	return [
		{
			...additional,
			placeholder: 'e.g. ContentTypeId, Title',
			description: `Comma-separated internal names of list item fields to add to "fields" of each entry. Each entry always has the default properties (${OUTPUT_PROPERTIES}).`,
		},
		fieldFiltersCombineOption,
		{
			displayName: 'Include List Item Fields',
			name: 'includeListItem',
			type: 'boolean',
			default: false,
			description:
				'Whether to add all list item fields of each entry in "fields". Same as Return All Properties. Kept for older workflows.',
		},
		{
			displayName: 'Max Depth',
			name: 'maxDepth',
			type: 'number',
			typeOptions: { minValue: 0 },
			default: 0,
			description:
				'The deepest folder level to return. The direct children of the folder are level 1. Folders below this level are not read. 0 means no limit.',
			displayOptions: recursiveOnly,
		},
		{
			displayName: 'Min Depth',
			name: 'minDepth',
			type: 'number',
			typeOptions: { minValue: 1 },
			default: 1,
			description:
				'The first folder level to return. The direct children of the folder are level 1. The node still reads the levels above it to find the deeper entries.',
			displayOptions: recursiveOnly,
		},
		{
			...returnAll,
			description:
				'Whether to add all list item fields of each entry in "fields", in addition to the default properties',
		},
	];
};

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['fileFolder'] } },
		options: [
			{
				name: 'Check Exists',
				value: 'checkExists',
				description: 'Check if files or folders exist',
				action: 'Check if files or folders exist',
			},
			{
				name: 'Compare Structure',
				value: 'compareStructure',
				description:
					'Compare the folders and files of a source folder, such as a template, with a target folder',
				action: 'Compare the structure of two folders',
			},
			{
				name: 'Get Copy Job Status',
				value: 'getCopyJobStatus',
				description: 'Get the progress of a copy or move job that is running in SharePoint',
				action: 'Get the status of a copy job',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get many files and folders in a folder',
				action: 'Get many files and folders',
			},
		],
		default: 'getAll',
	},
	{
		...pathRLC({ displayName: 'Folder', name: 'folderPath' }),
		displayOptions: show(['getAll']),
	},
	{
		displayName: 'Recursive',
		name: 'recursive',
		type: 'boolean',
		default: false,
		description:
			'Whether to also return the content of all subfolders. Each entry then has a "depth" property: 1 for the direct children of the folder.',
		displayOptions: show(['getAll']),
	},
	{
		displayName: 'Include',
		name: 'include',
		type: 'options',
		default: 'all',
		options: [
			{ name: 'Files and Folders', value: 'all' },
			{ name: 'Files Only', value: 'files' },
			{ name: 'Folders Only', value: 'folders' },
		],
		displayOptions: show(['getAll']),
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
		displayOptions: {
			show: { resource: ['fileFolder'], operation: ['getAll'], returnAll: [false] },
		},
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
				displayName: 'File Extensions',
				name: 'extensions',
				type: 'string',
				default: '',
				placeholder: 'pdf, docx',
				description: 'Comma-separated list of file extensions. Folders are not returned when set.',
			},
			{
				displayName: 'Modified After',
				name: 'modifiedAfter',
				type: 'dateTime',
				default: '',
			},
			{
				displayName: 'Modified Before',
				name: 'modifiedBefore',
				type: 'dateTime',
				default: '',
			},
			{
				displayName: 'Name Contains',
				name: 'nameContains',
				type: 'string',
				default: '',
			},
		],
	},
	fieldFiltersProperty({
		loadOptionsMethod: 'getFolderListFieldOptions',
		loadOptionsDependsOn: ['folderPath.value'],
		displayOptions: show(['getAll']),
		description:
			'Conditions on list item fields, for example Title or Status. Each folder is queried on its own, so the filters stay below the list view threshold when a folder has 5000 items or less.',
	}),
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: getAllOptions(),
	},
	{
		displayName: 'Job Info',
		name: 'jobInfo',
		type: 'json',
		required: true,
		default: '={{ $json }}',
		description:
			'The job info (jobId, jobQueueUri, encryptionKey) that a Copy Job copy or move returned. By default, the input item is used.',
		displayOptions: show(['getCopyJobStatus']),
	},
	{
		displayName: 'Paths',
		name: 'paths',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Path',
		default: { values: [{ path: '' }] },
		description:
			'The paths to check. Each path can be site-relative, server-relative or an absolute URL. You can also put many paths in one field, one on each line. The node has a True and a False output.',
		displayOptions: show(['checkExists']),
		options: [
			{
				displayName: 'Path',
				name: 'values',
				values: [
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						default: '',
						placeholder: 'Shared Documents/Folder/file.pdf',
					},
				],
			},
		],
	},
	{
		displayName: 'Combine',
		name: 'combine',
		type: 'options',
		default: 'all',
		description: 'When the input item goes to the True output',
		options: [
			{
				name: 'All',
				value: 'all',
				description: 'All paths must exist (AND)',
			},
			{
				name: 'Any',
				value: 'any',
				description: 'At least one path must exist (OR)',
			},
		],
		displayOptions: show(['checkExists']),
	},
	{
		...pathRLC({
			displayName: 'Source Folder',
			name: 'sourceFolder',
			description:
				'The folder with the expected structure, for example a template folder. Give a server-relative path, a site-relative path or an absolute URL.',
		}),
		displayOptions: show(['compareStructure']),
	},
	{
		...pathRLC({
			displayName: 'Target Folder',
			name: 'targetFolder',
			description:
				'The folder to check, for example a copy of a template folder. It can be in another library of the site.',
		}),
		displayOptions: show(['compareStructure']),
	},
	{
		displayName: 'Options',
		name: 'compareOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['compareStructure']),
		options: [
			{
				displayName: 'Case Sensitive',
				name: 'caseSensitive',
				type: 'boolean',
				default: false,
				description:
					'Whether "Folder" and "folder" are different names. SharePoint names are not case-sensitive.',
			},
			{
				displayName: 'Exclude Names',
				name: 'excludeNames',
				type: 'string',
				typeOptions: { rows: 2 },
				default: '',
				placeholder: 'Archive, ~*, Drafts/Old',
				description:
					'Names to leave out of the comparison, separated by commas or line breaks. "*" matches any text and "?" matches one character. A value with "/" matches the path below the folder. The content of an excluded folder is not read.',
			},
			{
				displayName: 'Include',
				name: 'include',
				type: 'options',
				default: 'folders',
				options: [
					{ name: 'Folders', value: 'folders' },
					{ name: 'Files', value: 'files' },
					{ name: 'Both', value: 'both' },
				],
				description: 'What to compare',
			},
			{
				displayName: 'Max Depth',
				name: 'maxDepth',
				type: 'number',
				typeOptions: { minValue: 0 },
				default: 1,
				description:
					'How many folder levels to compare. 1 compares only the direct children. 0 compares all levels.',
			},
		],
	},
];

async function findList(sp: SPFI, path: string): Promise<ListRoot | undefined> {
	return findOwningList(path, await getListRoots(sp));
}

/** Reads the rows of one folder level page by page, up to "max" rows. */
async function readFolderRows(
	list: IList,
	path: string,
	viewXml: string,
	max: number,
	includeListItem = false,
	extraFields: string[] = [],
): Promise<IDataObject[]> {
	const found: IDataObject[] = [];
	let query = new Map<string, string>();
	while (found.length < max) {
		const page = await list.renderListDataAsStream(
			{
				ViewXml: viewXml,
				FolderServerRelativeUrl: path,
				RenderOptions: 2,
				DatesInUtc: true,
				AddAllFields: includeListItem,
			},
			null,
			query,
		);
		for (const row of page.Row as unknown as IDataObject[]) {
			found.push(normalizeRow(row, includeListItem, extraFields));
			if (found.length >= max) break;
		}
		if (!page.NextHref) break;
		query = new Map(new URLSearchParams(page.NextHref).entries());
	}
	return found;
}

async function getAll(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const folderPath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const recursive = this.getNodeParameter('recursive', itemIndex) as boolean;
	const include = this.getNodeParameter('include', itemIndex) as IncludeKind;
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	const limit = returnAll ? Infinity : (this.getNodeParameter('limit', itemIndex) as number);
	const filters = this.getNodeParameter('filters', itemIndex, {}) as Omit<
		FileFolderFilters,
		'include'
	> & { extensions?: string };
	const options = this.getNodeParameter('options', itemIndex, {}) as Pick<
		SelectionValues,
		'additionalProperties' | 'returnAllProperties'
	> & {
		includeListItem?: boolean;
		fieldFiltersCombine?: string;
		minDepth?: number;
		maxDepth?: number;
	};
	const extensions = parseExtensions(filters.extensions);
	const includeListItem = options.includeListItem === true || options.returnAllProperties === true;
	const extraFields = toList(options.additionalProperties);

	const list = await findList(sp, folderPath);
	if (!list) {
		throw new NodeOperationError(this.getNode(), `No list or library contains "${folderPath}"`, {
			itemIndex,
		});
	}
	const spList = sp.web.lists.getById(list.id);

	let outputView: string;
	let foldersView: string;
	let fieldCondition: string | undefined;
	let depth: DepthRange;
	try {
		depth = toDepthRange(recursive, options.minDepth, options.maxDepth);
		const fieldFilters = readFieldFilters(this.getNodeParameter('fieldValueFilters', itemIndex, {}));
		if (fieldFilters.length) {
			fieldCondition = buildCamlFilter(
				fieldFilters,
				await loadFilterFieldTypes(spList, fieldFilters),
				readCombine(options.fieldFiltersCombine),
			);
		}
		outputView = buildViewXml(
			{
				include: extensions.length && include === 'all' ? 'files' : include,
				nameContains: filters.nameContains,
				modifiedAfter: filters.modifiedAfter,
				modifiedBefore: filters.modifiedBefore,
				fieldCondition,
			},
			500,
			extraFields,
		);
		foldersView = buildViewXml({ include: 'folders' });
	} catch (error) {
		throw toNodeError(this.getNode(), error, itemIndex);
	}
	const noCamlFilters =
		!filters.nameContains && !filters.modifiedAfter && !filters.modifiedBefore && !fieldCondition;

	const readFolder = async (path: string, viewXml: string, max: number) =>
		await readFolderRows(spList, path, viewXml, max, includeListItem, extraFields);

	// A recursive "Scope" view reads the whole library page by page, so each folder gets its own request.
	return await walkFolders(
		folderPath,
		{
			readOutput: async (path, max) => await readFolder(path, outputView, max),
			readFolders: async (path) => await readFolder(path, foldersView, Infinity),
			outputHasAllFolders: include === 'all' && noCamlFilters && !extensions.length,
			accept: extensions.length
				? (entry) => hasExtension(String(entry.name), extensions)
				: undefined,
		},
		{ depth, limit, addDepth: recursive },
	);
}

async function probe(sp: SPFI, path: string): Promise<IDataObject> {
	const check = async (read: () => Promise<{ Exists?: boolean; ServerRelativeUrl?: string }>) => {
		const info = await undefinedIfNotFound(read());
		return info?.Exists === false ? undefined : info;
	};
	const file = await check(() =>
		sp.web.getFileByServerRelativePath(path).select('Exists', 'ServerRelativeUrl')(),
	);
	if (file)
		return { exists: true, type: 'file', serverRelativeUrl: file.ServerRelativeUrl ?? path };
	const folder = await check(() =>
		sp.web.getFolderByServerRelativePath(path).select('Exists', 'ServerRelativeUrl')(),
	);
	if (folder) {
		return { exists: true, type: 'folder', serverRelativeUrl: folder.ServerRelativeUrl ?? path };
	}
	return { exists: false, type: null, serverRelativeUrl: path };
}

export type CombineMode = 'all' | 'any';

/** Builds the output of Check Exists. "exists" is the result for the Combine mode and selects the output. */
export function summarizeExists(results: IDataObject[], combine: CombineMode): IDataObject {
	const allExist = results.every((result) => result.exists === true);
	const anyExists = results.some((result) => result.exists === true);
	return {
		exists: combine === 'any' ? anyExists : allExist,
		allExist,
		anyExists,
		combine,
		results,
	};
}

async function checkExists(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const collection = this.getNodeParameter('paths', itemIndex, {}) as {
		values?: Array<{ path?: string }>;
	};
	const combine = this.getNodeParameter('combine', itemIndex, 'all') as CombineMode;
	const paths = (collection.values ?? [])
		.flatMap((entry) => String(entry.path ?? '').split(/\r?\n/))
		.map((path) => path.trim())
		.filter(Boolean);
	if (!paths.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one path to check', { itemIndex });
	}

	const results: IDataObject[] = [];
	for (const path of paths) {
		const { serverRelativeUrl, exists, type } = await probe(
			sp,
			toServerRelativePath(siteUrl, path),
		);
		results.push({ path, serverRelativeUrl, exists, type });
	}
	return summarizeExists(results, combine);
}

async function getCopyJobStatusOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const { sourcePath, destinationPath, ...info } = getJobInfoParameter.call(this, itemIndex);
	const status = await getCopyJobStatus(sp, siteUrl, {
		jobId: info.jobId,
		jobQueueUri: info.jobQueueUri,
		encryptionKey: info.encryptionKey,
	});
	return {
		...status,
		...(sourcePath ? { sourcePath } : {}),
		...(destinationPath ? { destinationPath } : {}),
	};
}

function createFolderReader(sp: SPFI, lists: ListRoot[]): FolderReader {
	const views = {
		all: buildViewXml({ include: 'all' }),
		folders: buildViewXml({ include: 'folders' }),
	};
	return async (path, foldersOnly) => {
		const list = findOwningList(path, lists);
		if (!list) throw new Error(`No list or library contains "${path}"`);
		const rows = await readFolderRows(
			sp.web.lists.getById(list.id),
			path,
			foldersOnly ? views.folders : views.all,
			Infinity,
		);
		return rows.map((row) => ({
			name: String(row.name),
			type: row.type === 'folder' ? 'folder' : 'file',
			serverRelativeUrl: String(row.serverRelativeUrl),
		}));
	};
}

async function compareStructure(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const source = getPath.call(this, siteUrl, itemIndex, 'sourceFolder');
	const target = getPath.call(this, siteUrl, itemIndex, 'targetFolder');
	const options = this.getNodeParameter('compareOptions', itemIndex, {}) as {
		include?: CompareInclude;
		maxDepth?: number;
		caseSensitive?: boolean;
		excludeNames?: string;
	};
	const treeOptions: TreeOptions = {
		include: options.include ?? 'folders',
		maxDepth: Math.max(0, Math.floor(options.maxDepth ?? 1)),
		caseSensitive: options.caseSensitive === true,
		exclude: parsePatterns(options.excludeNames),
	};

	if (!(await folderExists(sp, source))) {
		throw new NodeOperationError(this.getNode(), `The source folder "${source}" does not exist`, {
			itemIndex,
		});
	}
	const targetExists = await folderExists(sp, target);
	const read = createFolderReader(sp, await getListRoots(sp));
	let sourceTree: TreeEntry[];
	let targetTree: TreeEntry[] = [];
	try {
		sourceTree = await readTree(source, read, treeOptions);
		if (targetExists) targetTree = await readTree(target, read, treeOptions);
	} catch (error) {
		if (error instanceof Error && error.constructor === Error) {
			throw new NodeOperationError(this.getNode(), error.message, { itemIndex });
		}
		throw toNodeError(this.getNode(), error, itemIndex);
	}
	return {
		...compareTrees(sourceTree, targetTree, treeOptions.caseSensitive),
		source,
		target,
		targetExists,
	} as IDataObject;
}

export const fileFolderResource: ResourceModule = {
	option: {
		name: 'File/Folder',
		value: 'fileFolder',
		description:
			'Find files and folders, check if paths exist, compare folders, and check copy jobs',
	},
	description,
	operations: {
		getAll,
		checkExists,
		compareStructure,
		getCopyJobStatus: getCopyJobStatusOperation,
	},
	routing: {
		checkExists: {
			outputs: ['True', 'False'],
			route: (result) => ((result as IDataObject).exists === true ? 0 : 1),
			errorOutput: 1,
		},
	},
};
