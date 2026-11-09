import type { IExecuteFunctions } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import { getPath } from '../../shared/descriptions';
import type { OperationContext } from '../types';
import {
	copyChildrenIfMissing,
	createCopyChildrenGateway,
	type CopyChildrenOptions,
} from './copyChildren';
import {
	ensureFolders,
	foldersFromJson,
	foldersFromRows,
	toPerFolderItems,
	type FolderEntryRow,
} from './ensureFolders';
import {
	createEnsurePathGateway,
	ensurePath,
	segmentsFromJson,
	segmentsFromRows,
	type SegmentRow,
} from './ensurePath';
import {
	fromKeyValueRows,
	loadListFields,
	parseFieldsObject,
	type KeyValueRow,
	type ListFields,
} from './fieldValues';
import {
	createFolderValueReader,
	inheritMetadataFromParent,
	parseFieldNames,
	type InheritOptions,
} from './inheritMetadata';
import {
	createItemWriter,
	createPageReader,
	readContentTypeFields,
	readFolderItself,
	setMetadataRecursive,
	type MetadataOptions,
	type SetMetadataSource,
} from './setMetadata';

export async function ensurePathOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const basePath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const mode = this.getNodeParameter('segmentsInputMode', itemIndex, 'rows') as string;
	const segments =
		mode === 'json'
			? segmentsFromJson(this.getNodeParameter('segmentsJson', itemIndex, '[]'))
			: segmentsFromRows(
					(this.getNodeParameter('segments', itemIndex, {}) as { values?: SegmentRow[] }).values,
				);
	return await ensurePath(basePath, segments, createEnsurePathGateway(sp, basePath));
}

export async function ensureFoldersOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const basePath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const mode = this.getNodeParameter('foldersInputMode', itemIndex, 'rows') as string;
	const entries =
		mode === 'json'
			? foldersFromJson(this.getNodeParameter('foldersJson', itemIndex, '[]'))
			: foldersFromRows(
					(this.getNodeParameter('folders', itemIndex, {}) as { values?: FolderEntryRow[] }).values,
				);
	const result = await ensureFolders(
		basePath,
		entries,
		createEnsurePathGateway(sp, basePath),
		siteUrl,
	);
	const output = this.getNodeParameter('foldersOutput', itemIndex, 'single') as string;
	return output === 'perFolder' ? toPerFolderItems(result) : result;
}

export async function copyChildrenIfMissingOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const sourcePath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const destination = getPath.call(this, siteUrl, itemIndex, 'childrenDestinationFolder');
	const options = this.getNodeParameter(
		'copyChildrenOptions',
		itemIndex,
		{},
	) as CopyChildrenOptions;
	return await copyChildrenIfMissing(
		sourcePath,
		destination,
		options,
		createCopyChildrenGateway(sp, siteUrl, this.getNode(), itemIndex),
	);
}

export async function setMetadataRecursiveOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const folderPath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const mode = this.getNodeParameter('metadataInputMode', itemIndex, 'keyValue') as string;
	const fields =
		mode === 'json'
			? parseFieldsObject(
					this.getNodeParameter('metadataFieldsJson', itemIndex, '{}'),
					'Fields (JSON)',
				)
			: fromKeyValueRows(
					(this.getNodeParameter('metadataFields', itemIndex, {}) as { values?: KeyValueRow[] })
						.values,
				);
	const options = this.getNodeParameter('metadataOptions', itemIndex, {}) as MetadataOptions;
	const listFields = await loadListFields(sp, folderPath);
	const names = Object.keys(fields).filter((name) => listFields.types.has(name));
	return await setMetadataRecursive(
		folderPath,
		fields,
		options,
		metadataSource(sp, folderPath, listFields, names, options.systemUpdate === true),
	);
}

function metadataSource(
	sp: SPFI,
	folderPath: string,
	listFields: ListFields,
	names: string[],
	systemUpdate: boolean,
): SetMetadataSource {
	return {
		listFields,
		readPage: createPageReader(sp, listFields.listId, names),
		readSelf: async () => await readFolderItself(sp, folderPath, names, listFields.types),
		readContentTypeFields: async () => await readContentTypeFields(sp, listFields.listId),
		writer: createItemWriter(sp, listFields.listId, listFields.infos, systemUpdate),
	};
}

export async function inheritMetadataFromParentOperation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const folderPath = getPath.call(this, siteUrl, itemIndex, 'folderPath');
	const fieldNames = parseFieldNames(
		this.getNodeParameter('inheritFieldsPicker', itemIndex, []),
		this.getNodeParameter('inheritFields', itemIndex, ''),
	);
	const options = this.getNodeParameter('inheritOptions', itemIndex, {}) as InheritOptions;
	const listFields = await loadListFields(sp, folderPath);
	const names = fieldNames.filter((name) => listFields.types.has(name));
	return await inheritMetadataFromParent(folderPath, fieldNames, options, {
		...metadataSource(sp, folderPath, listFields, names, options.systemUpdate === true),
		libraryRoot: listFields.rootUrl ?? '',
		readValues: createFolderValueReader(sp, names, listFields.types),
	});
}
