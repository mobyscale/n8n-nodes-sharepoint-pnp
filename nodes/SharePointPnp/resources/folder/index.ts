import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { getPath, pathRLC } from '../../shared/descriptions';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import { copyJobProperties, methodProperty, runMoveCopyJob } from '../../shared/copyJob';
import {
	buildMoveCopyOptions,
	getAncestorPaths,
	getDestinationPath,
	getMissingSegments,
	undefinedIfNotFound,
	withListItemFields,
	type MoveCopyFields,
} from '../../shared/filesFolders';
import { joinPath, splitPath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';

const show = (operation: string[]) => ({ show: { resource: ['folder'], operation } });

const FOLDER_SELECT = [
	'Name',
	'ServerRelativeUrl',
	'UniqueId',
	'ItemCount',
	'TimeCreated',
	'TimeLastModified',
	'Exists',
	'IsWOPIEnabled',
	'ProgID',
	'WelcomePage',
];

const includeFieldsOption: INodeProperties = {
	displayName: 'Include List Item Fields',
	name: 'includeListItem',
	type: 'boolean',
	default: false,
	description:
		'Whether to add the list item fields of the folder. Same as adding ListItemAllFields to Expand. Kept for older workflows.',
};

function moveCopyOptions(kind: 'move' | 'copy'): INodeProperties {
	return {
		displayName: 'Direct Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: { show: { resource: ['folder'], operation: [kind], method: ['direct'] } },
		options: [
			{
				displayName: 'Bypass Shared Lock',
				name: 'bypassSharedLock',
				type: 'boolean',
				default: false,
				description:
					'Whether to allow the operation when a file in the folder has a co-authoring shared lock',
			},
			{
				displayName: 'If Exists',
				name: 'ifExists',
				type: 'options',
				default: 'fail',
				description:
					'What to do when a folder with the same name exists in the destination. The Direct method cannot replace a folder. Use the Copy Job method for that.',
				options: [
					{ name: 'Fail', value: 'fail' },
					{ name: 'Replace', value: 'replace' },
					{ name: 'Keep Both', value: 'keepBoth' },
				],
			},
			{
				displayName: 'New Name',
				name: 'newName',
				type: 'string',
				default: '',
				description: 'A new folder name for the destination. Leave empty to keep the name.',
			},
			kind === 'move'
				? {
						displayName: 'Retain Editor and Modified',
						name: 'retainEditorAndModified',
						type: 'boolean',
						default: false,
						description: 'Whether to keep the editor and modified date of the source folder',
					}
				: {
						displayName: 'Reset Author and Created',
						name: 'resetAuthorAndCreated',
						type: 'boolean',
						default: false,
						description:
							'Whether to set the author to the current user and the created date to now',
					},
		],
	};
}

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['folder'] } },
		options: [
			{ name: 'Copy', value: 'copy', description: 'Copy a folder', action: 'Copy a folder' },
			{
				name: 'Create',
				value: 'create',
				description: 'Create a folder and all missing parent folders',
				action: 'Create a folder',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a folder',
				action: 'Delete a folder',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get the properties of a folder',
				action: 'Get a folder',
			},
			{ name: 'Move', value: 'move', description: 'Move a folder', action: 'Move a folder' },
			{
				name: 'Rename',
				value: 'rename',
				description: 'Rename a folder',
				action: 'Rename a folder',
			},
		],
		default: 'create',
	},
	{
		...pathRLC({ displayName: 'Folder', name: 'folderPath' }),
		displayOptions: { show: { resource: ['folder'] } },
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: [...selectionOptions(FOLDER_SELECT), includeFieldsOption],
	},
	{
		...pathRLC({
			displayName: 'Destination Folder',
			name: 'destinationFolder',
			description: 'The folder to put the folder in',
		}),
		displayOptions: show(['copy', 'move']),
	},
	methodProperty('folder'),
	moveCopyOptions('copy'),
	moveCopyOptions('move'),
	...copyJobProperties('folder', 'copy'),
	...copyJobProperties('folder', 'move'),
	{
		displayName: 'New Name',
		name: 'newName',
		type: 'string',
		required: true,
		default: '',
		description: 'The new name of the folder',
		displayOptions: show(['rename']),
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['delete']),
		options: [
			{
				displayName: 'Bypass Shared Lock',
				name: 'bypassSharedLock',
				type: 'boolean',
				default: false,
				description:
					'Whether to allow the delete when a file in the folder has a co-authoring shared lock. Only used with Delete Permanently.',
			},
			{
				displayName: 'Delete Permanently',
				name: 'deletePermanently',
				type: 'boolean',
				default: false,
				description:
					'Whether to delete the folder for good. If off, the folder goes to the recycle bin.',
			},
		],
	},
];

function getFolder(sp: SPFI, path: string) {
	return sp.web.getFolderByServerRelativePath(path);
}

async function folderExists(sp: SPFI, path: string): Promise<boolean> {
	const folder = await undefinedIfNotFound(getFolder(sp, path).select('Exists')());
	return folder?.Exists ?? false;
}

async function readFolder(
	sp: SPFI,
	path: string,
	options: SelectionValues & { includeListItem?: boolean } = {},
): Promise<IDataObject> {
	return (await applySelection(
		getFolder(sp, path),
		FOLDER_SELECT,
		withListItemFields(options),
	)()) as unknown as IDataObject;
}

async function readResult(sp: SPFI, path: string, base: IDataObject): Promise<IDataObject> {
	const info = await undefinedIfNotFound(readFolder(sp, path));
	return info ? { ...base, ...info } : base;
}

async function create(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	if (await folderExists(sp, path)) return { ...(await readFolder(sp, path)), created: false };

	let existingAncestor: string | undefined;
	for (const candidate of getAncestorPaths(path, siteUrl).slice(1)) {
		if (await folderExists(sp, candidate)) {
			existingAncestor = candidate;
			break;
		}
	}
	if (!existingAncestor) {
		throw new NodeOperationError(
			this.getNode(),
			`No parent of "${path}" exists. The first folder level must be a library that exists.`,
			{ itemIndex },
		);
	}
	for (const missing of getMissingSegments(path, existingAncestor)) {
		await sp.web.folders.addUsingPath(missing, false);
	}
	return { ...(await readFolder(sp, path)), created: true };
}

async function get(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const options = this.getNodeParameter('options', itemIndex, {}) as SelectionValues & {
		includeListItem?: boolean;
	};
	return await readFolder(sp, path, options);
}

function moveCopy(kind: 'move' | 'copy') {
	return async function (this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
		const sourcePath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
		const destinationFolder = getPath.call(this, siteUrl, itemIndex, 'destinationFolder');
		const method = this.getNodeParameter('method', itemIndex, 'copyJob') as string;
		if (method === 'copyJob') {
			return await runMoveCopyJob.call(
				this,
				{ sp, siteUrl, itemIndex },
				{
					resource: 'folder',
					kind,
					sourcePath,
					destinationFolder,
				},
			);
		}
		const fields = this.getNodeParameter('options', itemIndex, {}) as MoveCopyFields;
		const destinationPath = getDestinationPath(sourcePath, destinationFolder, fields.newName);

		if (fields.ifExists === 'replace') {
			throw new NodeOperationError(
				this.getNode(),
				'The Direct method cannot replace a folder. Use "Fail" or "Keep Both", or use the Copy Job method.',
				{ itemIndex },
			);
		}
		const { options } = buildMoveCopyOptions(kind, fields);
		const folder = getFolder(sp, sourcePath);
		if (kind === 'move') await folder.moveByPath(destinationPath, options);
		else await folder.copyByPath(destinationPath, options);
		return await readResult(sp, destinationPath, { sourcePath, destinationPath, method });
	};
}

async function rename(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const sourcePath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const newName = (this.getNodeParameter('newName', itemIndex) as string).trim();
	if (!newName || newName.includes('/')) {
		throw new NodeOperationError(this.getNode(), 'The new name must not be empty or have a "/"', {
			itemIndex,
		});
	}
	const destinationPath = joinPath(splitPath(sourcePath).parent, newName);
	await getFolder(sp, sourcePath).moveByPath(destinationPath, {
		KeepBoth: false,
		RetainEditorAndModifiedOnMove: true,
		ShouldBypassSharedLocks: false,
	});
	return await readFolder(sp, destinationPath);
}

async function remove(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		deletePermanently?: boolean;
		bypassSharedLock?: boolean;
	};
	const folder = getFolder(sp, path);
	if (options.deletePermanently) {
		await folder.deleteWithParams({ BypassSharedLock: options.bypassSharedLock ?? false });
		return { path, deleted: true, permanent: true };
	}
	const recycleBinItemId = await folder.recycle();
	return { path, deleted: true, permanent: false, recycleBinItemId };
}

export const folderResource: ResourceModule = {
	option: { name: 'Folder', value: 'folder', description: 'Create, copy, move and manage folders' },
	description,
	operations: {
		create,
		get,
		copy: moveCopy('copy'),
		move: moveCopy('move'),
		rename,
		delete: remove,
	},
};
