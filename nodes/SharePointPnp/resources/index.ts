import { changeResource } from './change';
import { fileResource } from './file';
import { fileFolderResource } from './fileFolder';
import { folderResource } from './folder';
import { groupResource } from './group';
import { listResource } from './list';
import { listItemResource } from './listItem';
import { permissionResource } from './permission';
import { recycleBinResource } from './recycleBin';
import { searchResource } from './search';
import type { ResourceModule } from './types';

export const resources: ResourceModule[] = [
	fileResource,
	folderResource,
	fileFolderResource,
	listResource,
	listItemResource,
	permissionResource,
	groupResource,
	searchResource,
	changeResource,
	recycleBinResource,
];
