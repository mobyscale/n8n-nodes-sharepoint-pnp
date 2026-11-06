import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { getPath, pathRLC } from '../../shared/descriptions';
import {
	buildViewXml,
	findOwningList,
	hasExtension,
	normalizeRow,
	parseExtensions,
	type FileFolderFilters,
	undefinedIfNotFound,
	type IncludeKind,
} from '../../shared/filesFolders';
import { toServerRelativePath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';

const show = (operation: string[]) => ({ show: { resource: ['fileFolder'], operation } });

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
		description: 'Whether to also return the content of all subfolders',
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
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getAll']),
		options: [
			{
				displayName: 'Include List Item Fields',
				name: 'includeListItem',
				type: 'boolean',
				default: false,
				description: 'Whether to add all list item fields of each entry in "fields"',
			},
		],
	},
	{
		displayName: 'Paths',
		name: 'paths',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Path',
		default: { values: [{ path: '' }] },
		description:
			'The paths to check. Each path can be site-relative, server-relative or an absolute URL. You can also put many paths in one field, one on each line.',
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
];

interface ListInfo {
	id: string;
	rootUrl: string;
}

async function findList(sp: SPFI, path: string): Promise<ListInfo | undefined> {
	const lists = await sp.web.lists
		.select('Id', 'RootFolder/ServerRelativeUrl')
		.expand('RootFolder')();
	return findOwningList(
		path,
		lists.map((list) => ({
			id: list.Id as string,
			rootUrl: (list as unknown as { RootFolder: { ServerRelativeUrl: string } }).RootFolder
				.ServerRelativeUrl,
		})),
	);
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
	const options = this.getNodeParameter('options', itemIndex, {}) as { includeListItem?: boolean };
	const extensions = parseExtensions(filters.extensions);
	const includeListItem = options.includeListItem ?? false;

	const list = await findList(sp, folderPath);
	if (!list) {
		throw new NodeOperationError(this.getNode(), `No list or library contains "${folderPath}"`, {
			itemIndex,
		});
	}

	const noCamlFilters = !filters.nameContains && !filters.modifiedAfter && !filters.modifiedBefore;
	let outputView: string;
	let foldersView: string;
	try {
		outputView = buildViewXml({
			include: extensions.length && include === 'all' ? 'files' : include,
			nameContains: filters.nameContains,
			modifiedAfter: filters.modifiedAfter,
			modifiedBefore: filters.modifiedBefore,
		});
		foldersView = buildViewXml({ include: 'folders' });
	} catch (error) {
		throw new NodeOperationError(this.getNode(), error as Error, { itemIndex });
	}
	const outputHasAllFolders = include === 'all' && noCamlFilters && !extensions.length;

	const spList = sp.web.lists.getById(list.id);
	const readFolder = async (path: string, viewXml: string, max: number) => {
		const found: IDataObject[] = [];
		let query = new Map<string, string>();
		while (found.length < max) {
			const page = await spList.renderListDataAsStream(
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
				found.push(normalizeRow(row, includeListItem));
				if (found.length >= max) break;
			}
			if (!page.NextHref) break;
			query = new Map(new URLSearchParams(page.NextHref).entries());
		}
		return found;
	};

	// A recursive "Scope" view reads the whole library page by page, so each folder gets its own request.
	const entries: IDataObject[] = [];
	const queue = [folderPath];
	while (queue.length && entries.length < limit) {
		const path = queue.shift() as string;
		const level = await readFolder(
			path,
			outputView,
			extensions.length ? Infinity : limit - entries.length,
		);
		for (const entry of level) {
			if (extensions.length && !hasExtension(String(entry.name), extensions)) continue;
			if (entries.length < limit) entries.push(entry);
		}
		if (recursive) {
			const folders = outputHasAllFolders ? level : await readFolder(path, foldersView, Infinity);
			queue.push(
				...folders
					.filter((entry) => entry.type === 'folder')
					.map((entry) => String(entry.serverRelativeUrl)),
			);
		}
	}
	return entries;
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

async function checkExists(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const collection = this.getNodeParameter('paths', itemIndex, {}) as {
		values?: Array<{ path?: string }>;
	};
	const paths = (collection.values ?? [])
		.flatMap((entry) => String(entry.path ?? '').split(/\r?\n/))
		.map((path) => path.trim())
		.filter(Boolean);
	if (!paths.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one path to check', { itemIndex });
	}

	const results: IDataObject[] = [];
	for (const path of paths) {
		const serverRelativePath = toServerRelativePath(siteUrl, path);
		results.push({ path, ...(await probe(sp, serverRelativePath)) });
	}
	return results.map(({ path, serverRelativeUrl, exists, type }) => ({
		path,
		serverRelativeUrl,
		exists,
		type,
	}));
}

export const fileFolderResource: ResourceModule = {
	option: {
		name: 'File/Folder',
		value: 'fileFolder',
		description: 'Find files and folders, and check if paths exist',
	},
	description,
	operations: { getAll, checkExists },
};
