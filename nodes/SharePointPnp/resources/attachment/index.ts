import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import '@pnp/sp/attachments/index.js';

import { getList, listRLC } from '../../shared/descriptions';
import { applySelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import { undefinedIfNotFound } from '../../shared/filesFolders';
import { itemIdRLC, locatorValue, parseItemId } from '../../shared/pickers';
import { getRequiredString } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
import { getAttachmentSupportError, toArrayBuffer } from './helpers';

const RESOURCE = 'attachment';
const show = (operations: string[]) => ({ show: { resource: [RESOURCE], operation: operations } });
const ALL = ['add', 'getAll', 'download', 'delete', 'updateContent'];

const ATTACHMENT_SELECT = ['FileName', 'ServerRelativeUrl'];

const nameProperty = (operations: string[], description: string): INodeProperties => ({
	displayName: 'Attachment Name',
	name: 'attachmentName',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description,
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: { searchListMethod: 'searchItemAttachments', searchable: true },
		},
		{ displayName: 'By Name', name: 'name', type: 'string', placeholder: 'Report.pdf' },
	],
	displayOptions: show(operations),
});

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Add an attachment to a list item',
				action: 'Add an attachment',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Move an attachment to the recycle bin or delete it for good',
				action: 'Delete an attachment',
			},
			{
				name: 'Download',
				value: 'download',
				description: 'Download an attachment as binary data',
				action: 'Download an attachment',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get the attachments of a list item',
				action: 'Get many attachments',
			},
			{
				name: 'Update Content',
				value: 'updateContent',
				description: 'Replace the content of an attachment',
				action: 'Update the content of an attachment',
			},
		],
		default: 'getAll',
	},
	{
		...listRLC,
		description: 'The list that has the item. Document libraries do not support attachments.',
		displayOptions: show(ALL),
	},
	itemIdRLC(
		{ displayOptions: show(ALL) },
		'The list item. Choose from the list or give the numeric ID.',
	),
	nameProperty(['download', 'delete', 'updateContent'], 'The file name of the attachment'),
	{
		displayName: 'File Name',
		name: 'fileName',
		type: 'string',
		default: '',
		description: 'The name of the attachment. Leave empty to use the file name of the binary data.',
		displayOptions: show(['add']),
	},
	{
		displayName: 'Input Binary Field',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'The name of the input binary field that has the content',
		displayOptions: show(['add', 'updateContent']),
	},
	{
		displayName: 'Output Binary Field',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'The name of the output binary field to put the attachment in',
		displayOptions: show(['download']),
	},
	{
		displayName: 'Options',
		name: 'addOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['add']),
		options: [
			{
				displayName: 'Overwrite',
				name: 'overwrite',
				type: 'boolean',
				default: false,
				description:
					'Whether to replace the content of an attachment that has the same name. If off, the operation fails.',
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
				description:
					'Whether to delete the attachment for good. If off, it goes to the recycle bin.',
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
		options: selectionOptions(ATTACHMENT_SELECT),
	},
];

async function getAttachments(this: IExecuteFunctions, sp: SPFI, itemIndex: number) {
	const list = getList.call(this, sp, itemIndex);
	let itemId: number;
	try {
		itemId = parseItemId(this.getNodeParameter('itemId', itemIndex));
	} catch (error) {
		throw new NodeOperationError(this.getNode(), (error as Error).message, { itemIndex });
	}
	const info = await list.select('BaseType', 'EnableAttachments')();
	const problem = getAttachmentSupportError(info as { BaseType?: number });
	if (problem) throw new NodeOperationError(this.getNode(), problem, { itemIndex });
	return { attachments: list.items.getById(itemId).attachmentFiles, itemId };
}

function readAttachmentName(this: IExecuteFunctions, itemIndex: number): string {
	const name = locatorValue(this.getNodeParameter('attachmentName', itemIndex, ''));
	if (!name) {
		throw new NodeOperationError(this.getNode(), 'The attachment name must not be empty', {
			itemIndex,
		});
	}
	return name;
}

function readBinaryName(this: IExecuteFunctions, itemIndex: number): string {
	return getRequiredString.call(this, 'binaryPropertyName', itemIndex);
}

async function readInput(this: IExecuteFunctions, itemIndex: number) {
	const binaryPropertyName = readBinaryName.call(this, itemIndex);
	const binary = this.getInputData()[itemIndex].binary?.[binaryPropertyName];
	if (!binary) {
		throw new NodeOperationError(
			this.getNode(),
			`The item has no binary data in the field "${binaryPropertyName}"`,
			{ itemIndex },
		);
	}
	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, binaryPropertyName);
	return { binary, content: toArrayBuffer(buffer) };
}

async function add(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const { attachments, itemId } = await getAttachments.call(this, sp, itemIndex);
	const { binary, content } = await readInput.call(this, itemIndex);
	const fileName =
		(this.getNodeParameter('fileName', itemIndex, '') as string).trim() || binary.fileName;
	if (!fileName) {
		throw new NodeOperationError(this.getNode(), 'Set a file name for the attachment', {
			itemIndex,
		});
	}
	const options = this.getNodeParameter('addOptions', itemIndex, {}) as { overwrite?: boolean };

	let overwritten = false;
	if (options.overwrite) {
		overwritten =
			(await undefinedIfNotFound(attachments.getByName(fileName).select('FileName')())) !==
			undefined;
	}
	if (overwritten) await attachments.getByName(fileName).setContent(content);
	else await attachments.add(fileName, content);

	const info = await attachments.getByName(fileName).select(...ATTACHMENT_SELECT)();
	return { itemId, overwritten, ...(info as unknown as IDataObject) };
}

async function getAll(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const { attachments, itemId } = await getAttachments.call(this, sp, itemIndex);
	const options = this.getNodeParameter('getAllOptions', itemIndex, {}) as SelectionValues;
	const result = await applySelection(attachments, ATTACHMENT_SELECT, options)();
	return (result as unknown as IDataObject[]).map((entry) => ({ itemId, ...entry }));
}

async function download(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const { attachments, itemId } = await getAttachments.call(this, sp, itemIndex);
	const name = readAttachmentName.call(this, itemIndex);
	const binaryPropertyName = readBinaryName.call(this, itemIndex);
	const attachment = attachments.getByName(name);
	const info = (await attachment.select(...ATTACHMENT_SELECT)()) as unknown as IDataObject;
	const buffer = Buffer.from(await attachment.getBuffer());
	const binary = await this.helpers.prepareBinaryData(buffer, String(info.FileName ?? name));
	return [{ json: { itemId, ...info }, binary: { [binaryPropertyName]: binary } }];
}

async function deleteAttachment(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const { attachments, itemId } = await getAttachments.call(this, sp, itemIndex);
	const name = readAttachmentName.call(this, itemIndex);
	const options = this.getNodeParameter('deleteOptions', itemIndex, {}) as { permanent?: boolean };
	const attachment = attachments.getByName(name);
	if (options.permanent) await attachment.delete();
	else await attachment.recycle();
	return { itemId, fileName: name, deleted: true, permanent: options.permanent ?? false };
}

async function updateContent(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const { attachments, itemId } = await getAttachments.call(this, sp, itemIndex);
	const name = readAttachmentName.call(this, itemIndex);
	const { content } = await readInput.call(this, itemIndex);
	const attachment = attachments.getByName(name);
	await attachment.setContent(content);
	const info = await attachment.select(...ATTACHMENT_SELECT)();
	return { itemId, updated: true, ...(info as unknown as IDataObject) };
}

export const attachmentResource: ResourceModule = {
	option: {
		name: 'Attachment',
		value: RESOURCE,
		description: 'Add, get and delete list item attachments',
	},
	description,
	operations: { add, getAll, download, delete: deleteAttachment, updateContent },
};
