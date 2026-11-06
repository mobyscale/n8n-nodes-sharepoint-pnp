import { describe, expect, it } from 'vitest';
import type { INodeProperties } from 'n8n-workflow';

import { groupResource } from '../../nodes/SharePointPnp/resources/group';
import { listResource } from '../../nodes/SharePointPnp/resources/list';
import {
	hasCustomSelection,
	toAssignmentOutput,
} from '../../nodes/SharePointPnp/resources/permission/helpers';
import { permissionResource } from '../../nodes/SharePointPnp/resources/permission';
import {
	extraProperties,
	mapItem,
	RECYCLE_BIN_SELECT,
	type RecycleBinRawItem,
} from '../../nodes/SharePointPnp/resources/recycleBin/helpers';
import { buildPageUrlPart } from '../../nodes/SharePointPnp/resources/recycleBin/paging';
import { recycleBinResource } from '../../nodes/SharePointPnp/resources/recycleBin';
import { fileResource } from '../../nodes/SharePointPnp/resources/file';
import { fileFolderResource } from '../../nodes/SharePointPnp/resources/fileFolder';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';
import { userResource } from '../../nodes/SharePointPnp/resources/user';
import type { ResourceModule } from '../../nodes/SharePointPnp/resources/types';

function optionNames(
	resource: ResourceModule,
	collection: string,
	operation: string,
): string[] | undefined {
	const property = resource.description.find(
		(entry: INodeProperties) =>
			entry.name === collection && entry.displayOptions?.show?.operation?.includes(operation),
	);
	return (property?.options as INodeProperties[] | undefined)?.map((option) => option.name);
}

const CONTRACT = ['additionalProperties', 'expand', 'returnAllProperties'];

describe('selection options of the Get and Get Many operations', () => {
	const cases: Array<[string, ResourceModule, string, string]> = [
		['file get', fileResource, 'options', 'get'],
		['file getVersions', fileResource, 'versionOptions', 'getVersions'],
		['folder get', folderResource, 'options', 'get'],
		['list get', listResource, 'getOptions', 'get'],
		['list getAll', listResource, 'getAllOptions', 'getAll'],
		['list getFields', listResource, 'getFieldsOptions', 'getFields'],
		['list getViews', listResource, 'getViewsOptions', 'getViews'],
		['permission getAll', permissionResource, 'getAllOptions', 'getAll'],
		['group get', groupResource, 'getOptions', 'get'],
		['group getAll', groupResource, 'getAllOptions', 'getAll'],
		['group getMembers', groupResource, 'getMembersOptions', 'getMembers'],
		['user get', userResource, 'getOptions', 'get'],
		['user getAll', userResource, 'getAllOptions', 'getAll'],
		['recycleBin getAll', recycleBinResource, 'getAllOptions', 'getAll'],
	];

	it.each(cases)('%s has the contract options', (_label, resource, collection, operation) => {
		expect(optionNames(resource, collection, operation)).toEqual(expect.arrayContaining(CONTRACT));
	});

	it('keeps Include List Item Fields for older workflows', () => {
		expect(optionNames(fileResource, 'options', 'get')).toContain('includeListItem');
		expect(optionNames(folderResource, 'options', 'get')).toContain('includeListItem');
		expect(optionNames(fileFolderResource, 'options', 'getAll')).toContain('includeListItem');
	});

	it('offers no Expand for a CAML query', () => {
		const names = optionNames(fileFolderResource, 'options', 'getAll');
		expect(names).toEqual(expect.arrayContaining(['additionalProperties', 'returnAllProperties']));
		expect(names).not.toContain('expand');
	});
});

describe('permission Get Many output', () => {
	const assignment = {
		PrincipalId: 7,
		Member: { Id: 7, Title: 'Example Members', LoginName: 'Example Members', PrincipalType: 8 },
		RoleDefinitionBindings: [{ Id: 1073741827, Name: 'Contribute' }],
	};

	it('keeps the default shape without selection options', () => {
		expect(toAssignmentOutput(assignment, true)).toEqual({
			hasUniqueRoleAssignments: true,
			principalId: 7,
			member: {
				id: 7,
				title: 'Example Members',
				loginName: 'Example Members',
				principalType: 8,
				email: undefined,
			},
			roleDefinitions: [{ id: 1073741827, name: 'Contribute' }],
		});
	});

	it('adds the SharePoint object when the selection changes', () => {
		expect(toAssignmentOutput(assignment, false, { additionalProperties: 'Member/IsHiddenInUI' }))
			.toHaveProperty('properties', assignment);
		expect(hasCustomSelection({})).toBe(false);
		expect(hasCustomSelection({ additionalProperties: '' })).toBe(false);
		expect(hasCustomSelection({ expand: 'Member' })).toBe(true);
		expect(hasCustomSelection({ returnAllProperties: true })).toBe(true);
	});
});

describe('recycle bin Get Many output', () => {
	const item = {
		Id: '00000000-0000-0000-0000-000000000001',
		Title: 'a.txt',
		ItemType: 1,
		ItemState: 1,
		DeletedDate: '2026-01-01T00:00:00Z',
		Size: '5',
		Author: { Title: 'Example User' },
		'odata.type': 'SP.RecycleBinItem',
	} as unknown as RecycleBinRawItem;

	it('lists the properties that the filters and the paging need', () => {
		expect(RECYCLE_BIN_SELECT).toEqual(
			expect.arrayContaining(['Id', 'Title', 'ItemType', 'ItemState', 'DeletedDate', 'DirName']),
		);
	});

	it('returns extra properties without SharePoint metadata', () => {
		expect(extraProperties(item)).toEqual({ Author: { Title: 'Example User' } });
		expect(mapItem(item)).not.toHaveProperty('Author');
		expect(mapItem(item, true)).toMatchObject({ id: item.Id, Author: { Title: 'Example User' } });
	});

	it('does not change the paged request', () => {
		expect(buildPageUrlPart(1, undefined, 200)).toContain('rowLimit=200');
	});
});
