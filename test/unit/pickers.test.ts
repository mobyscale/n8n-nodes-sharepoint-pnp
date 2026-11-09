import type { ILoadOptionsFunctions, INodeProperties } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const sp = vi.hoisted(() => ({ current: {} as Record<string, unknown> }));

vi.mock('../../nodes/SharePointPnp/transport', () => ({
	createSp: () => sp.current,
	getTenantUrl: async () => 'https://contoso.sharepoint.com',
}));

import { getLookupListFieldOptions } from '../../nodes/SharePointPnp/methods/listSearchFields';
import {
	searchColumnGroups,
	searchItemAttachments,
	searchListFolders,
} from '../../nodes/SharePointPnp/methods/listSearchScoped';
import { attachmentResource } from '../../nodes/SharePointPnp/resources/attachment';
import { fileFolderResource } from '../../nodes/SharePointPnp/resources/fileFolder';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';
import { listResource } from '../../nodes/SharePointPnp/resources/list';
import { listItemResource } from '../../nodes/SharePointPnp/resources/listItem';
import {
	locatorValue,
	mergeFieldNames,
	parseItemId,
} from '../../nodes/SharePointPnp/shared/pickers';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

function loadContext(params: Record<string, unknown>): ILoadOptionsFunctions {
	return {
		getNodeParameter: (name: string, fallback?: unknown) =>
			name === 'site' ? siteUrl : name in params ? params[name] : fallback,
		getCredentials: async () => ({ subdomain: 'contoso' }),
	} as unknown as ILoadOptionsFunctions;
}

const query = (rows: unknown) => {
	const run = vi.fn().mockResolvedValue(rows);
	return Object.assign(run, {
		select: vi.fn().mockReturnValue(run),
		filter: vi.fn().mockReturnThis(),
	});
};

describe('picker value parsing', () => {
	it('reads locator values and plain values', () => {
		expect(locatorValue({ __rl: true, mode: 'list', value: ' 5 ' })).toBe('5');
		expect(locatorValue('Docs/A')).toBe('Docs/A');
		expect(locatorValue(undefined)).toBe('');
	});

	it('parses item IDs from locators, numbers and text', () => {
		expect(parseItemId({ mode: 'id', value: '12' })).toBe(12);
		expect(parseItemId(7)).toBe(7);
		expect(parseItemId('3')).toBe(3);
		expect(() => parseItemId({ mode: 'list', value: '' })).toThrow('not a valid item ID');
		expect(() => parseItemId('x')).toThrow('not a valid item ID');
	});

	it('merges picked and typed field names without duplicates', () => {
		expect(mergeFieldNames(['Title', 'Id'], 'Author/Title, Title')).toEqual([
			'Title',
			'Id',
			'Author/Title',
		]);
		expect(mergeFieldNames(undefined, '')).toEqual([]);
		expect(mergeFieldNames('Title,Id', undefined)).toEqual(['Title', 'Id']);
	});
});

describe('picker wiring', () => {
	const find = (properties: INodeProperties[], name: string) =>
		properties.filter((property) => property.name === name);

	it('uses resource locators for item IDs and attachment names', () => {
		for (const property of find(listItemResource.description, 'itemId')) {
			expect(property.type).toBe('resourceLocator');
		}
		expect(find(attachmentResource.description, 'itemId')[0].type).toBe('resourceLocator');
		expect(find(attachmentResource.description, 'attachmentName')[0].type).toBe('resourceLocator');
	});

	it('registers every search and load method that the properties use', () => {
		const methods = new SharePointPnp().methods;
		const used = new Set<string>();
		const visit = (properties: INodeProperties[]) => {
			for (const property of properties) {
				const typeOptions = property.typeOptions as Record<string, string> | undefined;
				if (typeOptions?.loadOptionsMethod)
					used.add(`loadOptions:${typeOptions.loadOptionsMethod}`);
				for (const mode of property.modes ?? []) {
					if (mode.typeOptions?.searchListMethod) {
						used.add(`listSearch:${mode.typeOptions.searchListMethod}`);
					}
				}
				for (const option of (property.options ?? []) as INodeProperties[]) {
					if ('name' in option && 'type' in option) visit([option]);
					const values = (option as { values?: INodeProperties[] }).values;
					if (values) visit(values);
				}
			}
		};
		visit(listItemResource.description);
		visit(attachmentResource.description);
		visit(listResource.description);
		visit(fileFolderResource.description);
		visit(folderResource.description);
		expect(used).toContain('loadOptions:getFolderListFieldOptions');
		for (const key of used) {
			const [kind, name] = key.split(':');
			expect(Object.keys(methods[kind as 'loadOptions' | 'listSearch'])).toContain(name);
		}
	});
});

describe('searchListFolders', () => {
	beforeEach(() => {
		const folders = query([
			{ Name: 'Forms', ServerRelativeUrl: '/sites/Example/Lists/Tasks/Forms' },
			{ Name: 'Beta', ServerRelativeUrl: '/sites/Example/Lists/Tasks/Beta' },
			{ Name: 'Alpha', ServerRelativeUrl: '/sites/Example/Lists/Tasks/Alpha' },
		]);
		const getFolder = vi.fn().mockReturnValue({ folders });
		sp.current = {
			web: {
				lists: {
					getById: () => ({
						rootFolder: {
							select: () => async () => ({ ServerRelativeUrl: '/sites/Example/Lists/Tasks' }),
						},
					}),
				},
				getFolderByServerRelativePath: getFolder,
			},
			getFolder,
		};
	});

	const ctx = () => loadContext({ list: { mode: 'id', value: 'guid' } });

	it('lists the root and the child folders for empty text', async () => {
		const result = await searchListFolders.call(ctx(), '');
		expect(result.results.map((entry) => entry.name)).toEqual(['(list root)', 'Alpha/', 'Beta/']);
		expect((sp.current.getFolder as ReturnType<typeof vi.fn>).mock.calls[0][0]).toBe(
			'/sites/Example/Lists/Tasks',
		);
	});

	it('looks in the parent folder and filters by the last segment', async () => {
		const result = await searchListFolders.call(ctx(), 'Parent/Al');
		expect((sp.current.getFolder as ReturnType<typeof vi.fn>).mock.calls[0][0]).toBe(
			'/sites/Example/Lists/Tasks/Parent',
		);
		expect(result.results).toEqual([
			{
				name: 'Parent/Alpha/',
				value: '/sites/Example/Lists/Tasks/Alpha',
				url: 'https://contoso.sharepoint.com/sites/Example/Lists/Tasks/Alpha',
			},
		]);
	});
});

describe('searchItemAttachments', () => {
	it('lists the attachment names of the selected item', async () => {
		const files = query([
			{ FileName: 'b.txt', ServerRelativeUrl: '/sites/Example/Lists/Tasks/Attachments/5/b.txt' },
			{ FileName: 'a.pdf', ServerRelativeUrl: '/sites/Example/Lists/Tasks/Attachments/5/a.pdf' },
		]);
		const getById = vi.fn().mockReturnValue({ attachmentFiles: files });
		sp.current = { web: { lists: { getById: () => ({ items: { getById } }) } } };
		const ctx = loadContext({
			list: { mode: 'id', value: 'guid' },
			itemId: { mode: 'list', value: '5' },
		});
		const result = await searchItemAttachments.call(ctx, '');
		expect(getById).toHaveBeenCalledWith(5);
		expect(result.results.map((entry) => entry.value)).toEqual(['a.pdf', 'b.txt']);
		const filtered = await searchItemAttachments.call(ctx, 'PDF');
		expect(filtered.results.map((entry) => entry.value)).toEqual(['a.pdf']);
	});
});

describe('searchColumnGroups', () => {
	it('returns the distinct groups', async () => {
		sp.current = {
			web: {
				fields: {
					select: () => async () => [
						{ Group: 'B' },
						{ Group: 'A' },
						{ Group: 'B' },
						{ Group: '' },
						{},
					],
				},
			},
		};
		const result = await searchColumnGroups.call(loadContext({}), '');
		expect(result.results.map((entry) => entry.value)).toEqual(['A', 'B']);
	});
});

describe('getLookupListFieldOptions', () => {
	it('reads the fields of the lookup list', async () => {
		const getById = vi.fn().mockReturnValue({
			fields: {
				filter: () => ({
					select: () => async () => [
						{ InternalName: 'Title', Title: 'Title' },
						{ InternalName: 'Code', Title: 'Code' },
					],
				}),
			},
		});
		sp.current = { web: { lists: { getById } } };
		const result = await getLookupListFieldOptions.call(
			loadContext({ lookupList: { mode: 'id', value: 'lookup-guid' } }),
		);
		expect(getById).toHaveBeenCalledWith('lookup-guid');
		expect(result.map((entry) => entry.value)).toEqual(['Code', 'Title']);
	});
});
