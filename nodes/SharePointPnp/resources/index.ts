import { attachmentResource } from './attachment';
import { changeResource } from './change';
import { fileResource } from './file';
import { fileFolderResource } from './fileFolder';
import { folderResource } from './folder';
import { groupResource } from './group';
import { userResource } from './user';
import { listResource } from './list';
import { listItemResource } from './listItem';
import { permissionResource } from './permission';
import { recycleBinResource } from './recycleBin';
import { searchResource } from './search';
import { sharingResource } from './sharing';
import { siteResource } from './site';
import { taxonomyResource } from './taxonomy';
import type { ResourceModule } from './types';

export const resources: ResourceModule[] = [
	siteResource,
	fileResource,
	folderResource,
	fileFolderResource,
	listResource,
	listItemResource,
	attachmentResource,
	permissionResource,
	sharingResource,
	groupResource,
	userResource,
	searchResource,
	taxonomyResource,
	changeResource,
	recycleBinResource,
];
