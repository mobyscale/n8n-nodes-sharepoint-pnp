import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import { CheckinType } from '@pnp/sp/files/index.js';

import { getPath, pathRLC } from '../../shared/descriptions';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import { copyJobProperties, methodProperty, runMoveCopyJob } from '../../shared/copyJob';
import {
	buildMoveCopyOptions,
	ensureFolder,
	filePathRLC,
	getDestinationPath,
	undefinedIfNotFound,
	withListItemFields,
	type MoveCopyFields,
} from '../../shared/filesFolders';
import { joinPath, splitPath } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
import { finishUpload, hasPostUploadSteps, parseUploadFields, type UploadOptions } from './upload';

const show = (operation: string[]) => ({ show: { resource: ['file'], operation } });

const CHUNK_THRESHOLD = 10 * 1024 * 1024;

const FILE_SELECT = [
	'Name',
	'ServerRelativeUrl',
	'UniqueId',
	'Length',
	'TimeCreated',
	'TimeLastModified',
	'Title',
	'MajorVersion',
	'MinorVersion',
	'UIVersion',
	'UIVersionLabel',
	'Level',
	'CheckOutType',
	'CheckInComment',
	'ETag',
	'Exists',
	'LinkingUrl',
	'ContentTag',
	'CustomizedPageStatus',
	'IrmEnabled',
];

const VERSION_SELECT = [
	'ID',
	'VersionLabel',
	'Created',
	'CheckInComment',
	'IsCurrentVersion',
	'Size',
	'Url',
];

const ifExistsProperty: INodeProperties = {
	displayName: 'If Exists',
	name: 'ifExists',
	type: 'options',
	default: 'fail',
	description: 'What to do when a file with the same name exists in the destination',
	options: [
		{ name: 'Fail', value: 'fail' },
		{ name: 'Replace', value: 'replace' },
		{ name: 'Keep Both', value: 'keepBoth' },
	],
};

function moveCopyOptions(kind: 'move' | 'copy'): INodeProperties {
	return {
		displayName: 'Direct Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: { show: { resource: ['file'], operation: [kind], method: ['direct'] } },
		options: [
			{
				displayName: 'Bypass Shared Lock',
				name: 'bypassSharedLock',
				type: 'boolean',
				default: false,
				description: 'Whether to allow the operation when the file has a co-authoring shared lock',
			},
			ifExistsProperty,
			{
				displayName: 'New Name',
				name: 'newName',
				type: 'string',
				default: '',
				description: 'A new file name for the destination. Leave empty to keep the name.',
			},
			kind === 'move'
				? {
						displayName: 'Retain Editor and Modified',
						name: 'retainEditorAndModified',
						type: 'boolean',
						default: false,
						description: 'Whether to keep the editor and modified date of the source file',
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
		displayOptions: { show: { resource: ['file'] } },
		options: [
			{
				name: 'Approve',
				value: 'approve',
				description: 'Approve a file that waits for content approval',
				action: 'Approve a file',
			},
			{
				name: 'Check In',
				value: 'checkIn',
				description: 'Check in a file',
				action: 'Check in a file',
			},
			{
				name: 'Check Out',
				value: 'checkOut',
				description: 'Check out a file',
				action: 'Check out a file',
			},
			{ name: 'Copy', value: 'copy', description: 'Copy a file', action: 'Copy a file' },
			{ name: 'Delete', value: 'delete', description: 'Delete a file', action: 'Delete a file' },
			{
				name: 'Deny',
				value: 'deny',
				description: 'Deny a file that waits for content approval',
				action: 'Deny a file',
			},
			{
				name: 'Discard Check Out',
				value: 'discardCheckOut',
				description: 'Discard the check out of a file',
				action: 'Discard the check out of a file',
			},
			{
				name: 'Download',
				value: 'download',
				description: 'Download a file as binary data',
				action: 'Download a file',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get the properties of a file',
				action: 'Get a file',
			},
			{
				name: 'Get Versions',
				value: 'getVersions',
				description: 'Get the versions of a file',
				action: 'Get the versions of a file',
			},
			{ name: 'Move', value: 'move', description: 'Move a file', action: 'Move a file' },
			{
				name: 'Publish',
				value: 'publish',
				description: 'Publish the latest version of a file as a major version',
				action: 'Publish a file',
			},
			{ name: 'Rename', value: 'rename', description: 'Rename a file', action: 'Rename a file' },
			{
				name: 'Restore Version',
				value: 'restoreVersion',
				description: 'Restore a version of a file',
				action: 'Restore a version of a file',
			},
			{
				name: 'Unpublish',
				value: 'unpublish',
				description: 'Unpublish the current major version of a file',
				action: 'Unpublish a file',
			},
			{
				name: 'Upload',
				value: 'upload',
				description: 'Upload a file from binary data',
				action: 'Upload a file',
			},
		],
		default: 'upload',
	},
	{
		...pathRLC({
			displayName: 'Destination Folder',
			name: 'folderPath',
			description: 'The folder to upload the file to',
		}),
		displayOptions: show(['upload']),
	},
	{
		displayName: 'File Name',
		name: 'fileName',
		type: 'string',
		default: '',
		placeholder: 'report.pdf',
		description: 'The name of the new file. Leave empty to use the file name of the binary data.',
		displayOptions: show(['upload']),
	},
	{
		displayName: 'Input Binary Field',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		description: 'The name of the input binary field that has the file',
		displayOptions: show(['upload']),
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['upload']),
		options: [
			{
				displayName: 'Check In',
				name: 'checkIn',
				type: 'boolean',
				default: false,
				description:
					'Whether to check in the file after the upload. Only done when the file is checked out, for example in a library that requires check out.',
			},
			{
				displayName: 'Check-In Comment',
				name: 'checkInComment',
				type: 'string',
				default: '',
				description: 'The comment for the check in. Used with Check In.',
			},
			{
				displayName: 'Check-In Type',
				name: 'checkInType',
				type: 'options',
				default: 'major',
				description: 'The version type for the check in. Used with Check In.',
				options: [
					{ name: 'Major', value: 'major' },
					{ name: 'Minor', value: 'minor' },
					{ name: 'Overwrite', value: 'overwrite' },
				],
			},
			{
				displayName: 'Create Missing Folders',
				name: 'createMissingFolders',
				type: 'boolean',
				default: false,
				description:
					'Whether to create the destination folder and its missing parent folders before the upload. The library must exist.',
			},
			{
				displayName: 'Fields',
				name: 'fields',
				type: 'fixedCollection',
				typeOptions: { multipleValues: true },
				placeholder: 'Add Field',
				default: {},
				description:
					'List item fields to set on the file after the upload. These values replace the same fields in Fields (JSON).',
				options: [
					{
						displayName: 'Field',
						name: 'values',
						values: [
							{
								displayName: 'Field Name',
								name: 'name',
								type: 'string',
								default: '',
								placeholder: 'Title',
								description: 'The internal name of the field',
							},
							{
								displayName: 'Value',
								name: 'value',
								type: 'string',
								default: '',
							},
						],
					},
				],
			},
			{
				displayName: 'Fields (JSON)',
				name: 'fieldsJson',
				type: 'json',
				default: '{}',
				description:
					'An object with internal field names as keys, to set on the file after the upload. Values use the same format as List Item Create.',
			},
			{
				displayName: 'Overwrite',
				name: 'overwrite',
				type: 'boolean',
				default: true,
				description: 'Whether to replace a file that has the same name',
			},
			{
				displayName: 'Publish',
				name: 'publish',
				type: 'boolean',
				default: false,
				description:
					'Whether to publish the file as a major version after the upload. Only done when the file is a draft, in a library with minor versions or content approval.',
			},
			{
				displayName: 'Publish Comment',
				name: 'publishComment',
				type: 'string',
				default: '',
				description: 'The comment for the publish. Used with Publish. At most 1023 characters.',
			},
		],
	},
	{
		...filePathRLC({
			displayName: 'File',
			name: 'filePath',
			description:
				'A server-relative path (/sites/Example/Shared Documents/a.pdf), a site-relative path (Shared Documents/a.pdf), or an absolute URL',
		}),
		displayOptions: {
			show: { resource: ['file'] },
			hide: { operation: ['upload'] },
		},
	},
	{
		displayName: 'Output Binary Field',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		description: 'The name of the output binary field to put the file in',
		displayOptions: show(['download']),
	},
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: [
			...selectionOptions(FILE_SELECT),
			{
				displayName: 'Include List Item Fields',
				name: 'includeListItem',
				type: 'boolean',
				default: false,
				description:
					'Whether to add the list item fields of the file in ListItemAllFields. The default properties do not change.',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'versionOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getVersions']),
		options: selectionOptions(VERSION_SELECT),
	},
	{
		...pathRLC({
			displayName: 'Destination Folder',
			name: 'destinationFolder',
			description: 'The folder to put the file in',
		}),
		displayOptions: show(['copy', 'move']),
	},
	methodProperty('file'),
	moveCopyOptions('copy'),
	moveCopyOptions('move'),
	...copyJobProperties('file', 'copy'),
	...copyJobProperties('file', 'move'),
	{
		displayName: 'New Name',
		name: 'newName',
		type: 'string',
		required: true,
		default: '',
		description: 'The new name of the file, with the extension',
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
					'Whether to allow the delete when the file has a co-authoring shared lock. Only used with Delete Permanently.',
			},
			{
				displayName: 'Delete Permanently',
				name: 'deletePermanently',
				type: 'boolean',
				default: false,
				description:
					'Whether to delete the file for good. If off, the file goes to the recycle bin.',
			},
		],
	},
	{
		displayName: 'Comment',
		name: 'comment',
		type: 'string',
		default: '',
		description: 'The check-in comment',
		displayOptions: show(['checkIn']),
	},
	{
		displayName: 'Comment',
		name: 'comment',
		type: 'string',
		default: '',
		description:
			'The comment for the version history. At most 1023 characters. Publish and Unpublish need a library with minor versions or content approval. Approve and Deny need content approval.',
		displayOptions: show(['publish', 'unpublish', 'approve', 'deny']),
	},
	{
		displayName: 'Check-In Type',
		name: 'checkInType',
		type: 'options',
		default: 'major',
		options: [
			{ name: 'Major', value: 'major' },
			{ name: 'Minor', value: 'minor' },
			{ name: 'Overwrite', value: 'overwrite' },
		],
		displayOptions: show(['checkIn']),
	},
	{
		displayName: 'Version',
		name: 'versionLabel',
		type: 'resourceLocator',
		required: true,
		default: { mode: 'list', value: '' },
		description: 'The version to restore',
		displayOptions: show(['restoreVersion']),
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod: 'searchFileVersions', searchable: true },
			},
			{
				displayName: 'By Label',
				name: 'label',
				type: 'string',
				placeholder: '1.0',
				validation: [
					{
						type: 'regex',
						properties: { regex: '^\\d+\\.\\d+$', errorMessage: 'Not a valid version label' },
					},
				],
			},
		],
	},
];

function getFile(sp: SPFI, path: string) {
	return sp.web.getFileByServerRelativePath(path);
}

async function readFile(sp: SPFI, path: string): Promise<IDataObject> {
	return (await getFile(sp, path)()) as unknown as IDataObject;
}

/** Returns the properties of a moved or copied file. A "Keep Both" copy can have another name. */
async function readResult(sp: SPFI, path: string, base: IDataObject): Promise<IDataObject> {
	const info = await undefinedIfNotFound(readFile(sp, path));
	return info ? { ...base, ...info } : base;
}

async function upload(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const folderPath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const binaryPropertyName = this.getNodeParameter('binaryPropertyName', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as UploadOptions;
	let fields: Record<string, unknown>;
	try {
		fields = parseUploadFields(options);
	} catch (error) {
		throw new NodeOperationError(this.getNode(), error as Error, { itemIndex });
	}
	const binary = this.getInputData()[itemIndex].binary?.[binaryPropertyName];
	if (!binary) {
		throw new NodeOperationError(
			this.getNode(),
			`The item has no binary data in the field "${binaryPropertyName}"`,
			{ itemIndex },
		);
	}
	const fileName =
		(this.getNodeParameter('fileName', itemIndex, '') as string).trim() || binary.fileName;
	if (!fileName) {
		throw new NodeOperationError(this.getNode(), 'Set a file name for the upload', { itemIndex });
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, binaryPropertyName);
	if (options.createMissingFolders) await ensureFolder(sp, siteUrl, folderPath);
	const files = sp.web.getFolderByServerRelativePath(folderPath).files;
	const props = { Overwrite: options.overwrite ?? true };
	const info =
		buffer.length > CHUNK_THRESHOLD
			? await files.addChunked(fileName, new Blob([new Uint8Array(buffer)]), props)
			: await files.addUsingPath(
					fileName,
					buffer.buffer.slice(
						buffer.byteOffset,
						buffer.byteOffset + buffer.byteLength,
					) as ArrayBuffer,
					props,
				);
	if (!hasPostUploadSteps(options, fields)) return info as unknown as IDataObject;
	const path = String(info.ServerRelativeUrl ?? joinPath(folderPath, fileName));
	return await finishUpload(sp, path, options, fields);
}

async function download(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const binaryPropertyName = this.getNodeParameter('binaryPropertyName', itemIndex) as string;
	const info = await readFile(sp, path);
	const buffer = Buffer.from(await getFile(sp, path).getBuffer());
	const binary = await this.helpers.prepareBinaryData(buffer, String(info.Name));
	return [{ json: info, binary: { [binaryPropertyName]: binary } }];
}

async function get(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const options = this.getNodeParameter('options', itemIndex, {}) as SelectionValues & {
		includeListItem?: boolean;
	};
	return (await applySelection(
		getFile(sp, path),
		FILE_SELECT,
		withListItemFields(options),
	)()) as unknown as IDataObject;
}

function moveCopy(kind: 'move' | 'copy') {
	return async function (this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
		const sourcePath = getPath.call(this, siteUrl, itemIndex, 'filePath');
		const destinationFolder = getPath.call(this, siteUrl, itemIndex, 'destinationFolder');
		if (this.getNodeParameter('method', itemIndex, 'copyJob') === 'copyJob') {
			return await runMoveCopyJob.call(
				this,
				{ sp, siteUrl, itemIndex },
				{
					resource: 'file',
					kind,
					sourcePath,
					destinationFolder,
				},
			);
		}
		const fields = this.getNodeParameter('options', itemIndex, {}) as MoveCopyFields;
		const destinationPath = getDestinationPath(sourcePath, destinationFolder, fields.newName);
		const { overwrite, options } = buildMoveCopyOptions(kind, fields);
		const file = getFile(sp, sourcePath);
		if (kind === 'move') await file.moveByPath(destinationPath, overwrite, options);
		else await file.copyByPath(destinationPath, overwrite, options);
		return await readResult(sp, destinationPath, { sourcePath, destinationPath });
	};
}

async function rename(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const sourcePath = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const newName = (this.getNodeParameter('newName', itemIndex) as string).trim();
	if (!newName || newName.includes('/')) {
		throw new NodeOperationError(this.getNode(), 'The new name must not be empty or have a "/"', {
			itemIndex,
		});
	}
	const destinationPath = joinPath(splitPath(sourcePath).parent, newName);
	await getFile(sp, sourcePath).moveByPath(destinationPath, false, {
		KeepBoth: false,
		RetainEditorAndModifiedOnMove: true,
		ShouldBypassSharedLocks: false,
	});
	return await readFile(sp, destinationPath);
}

async function remove(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		deletePermanently?: boolean;
		bypassSharedLock?: boolean;
	};
	const file = getFile(sp, path);
	if (options.deletePermanently) {
		await file.deleteWithParams({ BypassSharedLock: options.bypassSharedLock ?? false });
		return { path, deleted: true, permanent: true };
	}
	const recycleBinItemId = await file.recycle();
	return { path, deleted: true, permanent: false, recycleBinItemId };
}

const CHECKIN_TYPES: Record<string, CheckinType> = {
	minor: CheckinType.Minor,
	major: CheckinType.Major,
	overwrite: CheckinType.Overwrite,
};

async function checkIn(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const comment = this.getNodeParameter('comment', itemIndex, '') as string;
	const type = this.getNodeParameter('checkInType', itemIndex) as string;
	await getFile(sp, path).checkin(comment, CHECKIN_TYPES[type]);
	return await readFile(sp, path);
}

const COMMENT_LIMIT = 1023;

/** Reads the file and, if the library has content approval, its approval status. */
async function readPublishState(sp: SPFI, path: string): Promise<IDataObject> {
	const info = await readFile(sp, path);
	const status = await getFile(sp, path)
		.listItemAllFields.select('OData__ModerationStatus', 'OData__ModerationComments')()
		.catch(() => undefined);
	if (!status) return info;
	return {
		...info,
		moderationStatus: status.OData__ModerationStatus ?? null,
		moderationComments: status.OData__ModerationComments ?? null,
	};
}

function publishOperation(action: 'publish' | 'unpublish' | 'approve' | 'deny') {
	return async function (this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
		const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
		const comment = this.getNodeParameter('comment', itemIndex, '') as string;
		if (comment.length > COMMENT_LIMIT) {
			throw new NodeOperationError(
				this.getNode(),
				`The comment has ${comment.length} characters. SharePoint accepts at most ${COMMENT_LIMIT}.`,
				{ itemIndex },
			);
		}
		await getFile(sp, path)[action](comment);
		return await readPublishState(sp, path);
	};
}

async function checkOut(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	await getFile(sp, path).checkout();
	return await readFile(sp, path);
}

async function discardCheckOut(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	await getFile(sp, path).undoCheckout();
	return await readFile(sp, path);
}

async function getVersions(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const options = this.getNodeParameter('versionOptions', itemIndex, {}) as SelectionValues;
	return (await applySelection(
		getFile(sp, path).versions,
		VERSION_SELECT,
		options,
	)()) as unknown as IDataObject[];
}

async function restoreVersion(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const path = getPath.call(this, siteUrl, itemIndex, 'filePath');
	const versionLabel = String(
		this.getNodeParameter('versionLabel', itemIndex, '', { extractValue: true }) ?? '',
	).trim();
	if (!versionLabel) {
		throw new NodeOperationError(this.getNode(), 'Select the version to restore', { itemIndex });
	}
	await getFile(sp, path).versions.restoreByLabel(versionLabel);
	return { ...(await readFile(sp, path)), restoredVersionLabel: versionLabel };
}

export const fileResource: ResourceModule = {
	option: {
		name: 'File',
		value: 'file',
		description: 'Upload, download, copy, move and manage files',
	},
	description,
	operations: {
		upload,
		download,
		get,
		copy: moveCopy('copy'),
		move: moveCopy('move'),
		rename,
		delete: remove,
		checkIn,
		checkOut,
		discardCheckOut,
		publish: publishOperation('publish'),
		unpublish: publishOperation('unpublish'),
		approve: publishOperation('approve'),
		deny: publishOperation('deny'),
		getVersions,
		restoreVersion,
	},
};
