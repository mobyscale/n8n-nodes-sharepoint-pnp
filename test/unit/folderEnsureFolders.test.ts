import { describe, expect, it, vi } from 'vitest';

import {
	ensureFolders,
	folderUrls,
	foldersFromJson,
	foldersFromRows,
	planFolderTree,
	splitRelativePath,
	toPerFolderItems,
	type FolderEntrySpec,
} from '../../nodes/SharePointPnp/resources/folder/ensureFolders';
import type {
	EnsurePathGateway,
	FolderState,
} from '../../nodes/SharePointPnp/resources/folder/ensurePath';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';

const SITE = 'https://contoso.sharepoint.com/sites/Example';
const LIB = '/sites/Example/Shared Documents';

const entry = (
	key: string,
	path: string,
	extra: Partial<FolderEntrySpec> = {},
): FolderEntrySpec => ({
	key,
	segments: path.split('/'),
	fieldsWhenCreated: {},
	fieldsAlways: {},
	...extra,
});

/** A reports folder with an archive folder, a year folder with two quarter folders. */
const TREE: FolderEntrySpec[] = [
	entry('reports', 'Reports', { fieldsWhenCreated: { Title: 'Reports' } }),
	entry('archive', 'Reports/Archive'),
	entry('year', 'Reports/2024', { fieldsAlways: { Title: '2024' } }),
	entry('q2', 'Reports/2024/Q2'),
	entry('quarter', 'Reports/2024/Q1', {
		appendToMultiChoice: { field: 'Tags', values: ['Active'] },
	}),
];

function fakeGateway(existing: Record<string, Record<string, unknown>> = {}) {
	const folders = new Map(
		Object.entries(existing).map(([path, item]) => [path.toLowerCase(), { path, item }]),
	);
	let nextId = 100;
	const ids = new Map<string, number>();
	const gateway = {
		getBaseFolder: vi.fn(async (path: string) => (path === LIB ? LIB : undefined)),
		getFolder: vi.fn(async (path: string): Promise<FolderState | undefined> => {
			const found = folders.get(path.toLowerCase());
			if (!found) return undefined;
			if (!ids.has(found.path)) ids.set(found.path, nextId++);
			const id = ids.get(found.path)!;
			return {
				serverRelativeUrl: found.path,
				listItemId: id,
				uniqueId: `guid-${id}`,
				item: found.item,
			};
		}),
		createFolder: vi.fn(async (path: string) => {
			folders.set(path.toLowerCase(), { path, item: {} });
		}),
		getFieldTypes: vi.fn(
			async () =>
				new Map([
					['Title', 'Text'],
					['Tags', 'MultiChoice'],
				]),
		),
		updateItem: vi.fn(async (id: number, values: Record<string, unknown>) => {
			for (const folder of folders.values()) {
				if (ids.get(folder.path) === id) Object.assign(folder.item, values);
			}
		}),
	} satisfies EnsurePathGateway;
	return gateway;
}

describe('Ensure Folders input', () => {
	it('normalizes relative paths', () => {
		expect(splitRelativePath(' /Reports\\ Drafts/./ ', 'X')).toEqual(['Reports', 'Drafts']);
		expect(() => splitRelativePath('a/../b', 'X')).toThrow('must not have ".."');
		expect(() => splitRelativePath(' / ', 'X')).toThrow('X: the path is empty');
	});

	it('reads folder rows', () => {
		expect(
			foldersFromRows([
				{
					key: ' quarter ',
					path: 'Reports/Q1',
					fieldsWhenCreated: '{"Title":"P"}',
					fieldsAlways: '',
					appendField: 'Tags',
					appendValues: 'Active, New',
				},
			]),
		).toEqual([
			{
				key: 'quarter',
				segments: ['Reports', 'Q1'],
				fieldsWhenCreated: { Title: 'P' },
				fieldsAlways: {},
				appendToMultiChoice: { field: 'Tags', values: ['Active', 'New'] },
			},
		]);
	});

	it('reads JSON folders', () => {
		expect(
			foldersFromJson(
				JSON.stringify([
					{ key: 'reports', path: 'Reports', fieldsAlways: { Title: 'A' } },
					{
						key: 'quarter',
						path: 'Reports/Q1',
						appendToMultiChoice: { field: 'Tags', values: ['x'] },
					},
				]),
			),
		).toEqual([
			{
				key: 'reports',
				segments: ['Reports'],
				fieldsWhenCreated: {},
				fieldsAlways: { Title: 'A' },
				appendToMultiChoice: undefined,
			},
			{
				key: 'quarter',
				segments: ['Reports', 'Q1'],
				fieldsWhenCreated: {},
				fieldsAlways: {},
				appendToMultiChoice: { field: 'Tags', values: ['x'] },
			},
		]);
	});

	it('rejects bad folders', () => {
		expect(() => foldersFromRows([{ key: ' ', path: 'A' }])).toThrow('Folder 1 has no key');
		expect(() =>
			foldersFromRows([
				{ key: 'a', path: 'A' },
				{ key: 'a', path: 'B' },
			]),
		).toThrow('The key "a" is used more than once');
		expect(() =>
			foldersFromRows([
				{ key: 'a', path: 'Reports/X' },
				{ key: 'b', path: 'reports/x/' },
			]),
		).toThrow('The keys "a" and "b" have the same path');
		expect(() => foldersFromJson('{}')).toThrow('must be an array');
		expect(() => foldersFromJson('["A"]')).toThrow('must be an object with "key" and "path"');
		expect(() => foldersFromJson('{bad')).toThrow('not valid JSON');
	});
});

describe('Ensure Folders tree plan', () => {
	it('puts parents before children and keeps each shared parent once', () => {
		const nodes = planFolderTree(TREE);
		expect(nodes.map((node) => node.path)).toEqual([
			'Reports',
			'Reports/Archive',
			'Reports/2024',
			'Reports/2024/Q2',
			'Reports/2024/Q1',
		]);
		expect(nodes.map((node) => node.entry?.key)).toEqual([
			'reports',
			'archive',
			'year',
			'q2',
			'quarter',
		]);
	});

	it('adds missing parents first, also when a child comes before its parent', () => {
		const nodes = planFolderTree([
			entry('quarter', 'Reports/2024/Q1'),
			entry('drafts', 'Reports/2024/Q2'),
			entry('reports', 'reports'),
		]);
		expect(nodes.map((node) => [node.path, node.entry?.key])).toEqual([
			['Reports', 'reports'],
			['Reports/2024', undefined],
			['Reports/2024/Q1', 'quarter'],
			['Reports/2024/Q2', 'drafts'],
		]);
		expect(nodes.map((node) => node.parentId)).toEqual([
			null,
			'reports',
			'reports/2024',
			'reports/2024',
		]);
	});
});

describe('Ensure Folders run', () => {
	it('creates the tree once, parents first, and returns the folders by key', async () => {
		const gateway = fakeGateway();
		const result = await ensureFolders(LIB, TREE, gateway, SITE);
		expect(gateway.createFolder.mock.calls.map(([path]) => path)).toEqual([
			`${LIB}/Reports`,
			`${LIB}/Reports/Archive`,
			`${LIB}/Reports/2024`,
			`${LIB}/Reports/2024/Q2`,
			`${LIB}/Reports/2024/Q1`,
		]);
		// Only the first folder is looked up before it is created. The others have a new parent.
		expect(gateway.getFolder).toHaveBeenCalledTimes(6);
		expect(result).toMatchObject({ created: true, changed: true, baseFolder: LIB });
		expect(Object.keys(result.folders)).toEqual(TREE.map((spec) => spec.key));
		expect(result.folders.quarter).toEqual({
			path: 'Reports/2024/Q1',
			serverRelativeUrl: `${LIB}/Reports/2024/Q1`,
			siteRelativeUrl: 'Shared Documents/Reports/2024/Q1',
			absoluteUrl:
				'https://contoso.sharepoint.com/sites/Example/Shared%20Documents/Reports/2024/Q1',
			listItemId: expect.any(Number),
			uniqueId: expect.stringMatching(/^guid-/),
			created: true,
			fieldsUpdated: ['Tags'],
		});
		expect(result.folders.reports.fieldsUpdated).toEqual(['Title']);
		expect(result.folders.year.fieldsUpdated).toEqual(['Title']);
		expect(result.folders.archive.fieldsUpdated).toEqual([]);
	});

	it('creates a shared parent without a key only once', async () => {
		const gateway = fakeGateway();
		const result = await ensureFolders(
			LIB,
			[entry('quarter', 'Reports/2024/Q1'), entry('drafts', 'Reports/2024/Q2')],
			gateway,
			SITE,
		);
		expect(gateway.createFolder.mock.calls.map(([path]) => path)).toEqual([
			`${LIB}/Reports`,
			`${LIB}/Reports/2024`,
			`${LIB}/Reports/2024/Q1`,
			`${LIB}/Reports/2024/Q2`,
		]);
		expect(Object.keys(result.folders)).toEqual(['quarter', 'drafts']);
	});

	it('reports no change on a second run', async () => {
		const gateway = fakeGateway();
		await ensureFolders(LIB, TREE, gateway, SITE);
		gateway.createFolder.mockClear();
		gateway.updateItem.mockClear();
		const result = await ensureFolders(LIB, TREE, gateway, SITE);
		expect(gateway.createFolder).not.toHaveBeenCalled();
		expect(gateway.updateItem).not.toHaveBeenCalled();
		expect(result).toMatchObject({ created: false, changed: false });
		expect(Object.values(result.folders).every((folder) => !folder.created)).toBe(true);
	});

	it('sets created fields only on new folders and always fields when they differ', async () => {
		const gateway = fakeGateway({
			[`${LIB}/reports`]: { Title: 'Old' },
			[`${LIB}/reports/2024`]: { Title: 'Old' },
		});
		const result = await ensureFolders(LIB, TREE.slice(0, 3), gateway, SITE);
		expect(gateway.createFolder.mock.calls.map(([path]) => path)).toEqual([
			`${LIB}/reports/Archive`,
		]);
		expect(result.folders.reports).toMatchObject({
			created: false,
			fieldsUpdated: [],
			serverRelativeUrl: `${LIB}/reports`,
		});
		expect(result.folders.year).toMatchObject({ created: false, fieldsUpdated: ['Title'] });
		expect(result.created).toBe(true);
	});

	it('makes one item for each folder', async () => {
		const result = await ensureFolders(LIB, TREE.slice(0, 2), fakeGateway(), SITE);
		const items = toPerFolderItems(result);
		expect(items.map((item) => item.key)).toEqual(['reports', 'archive']);
		expect(items[1]).toMatchObject({
			key: 'archive',
			path: 'Reports/Archive',
			serverRelativeUrl: `${LIB}/Reports/Archive`,
			created: true,
			baseFolder: LIB,
		});
	});

	it('fails for no folders, unknown fields and a missing base folder', async () => {
		await expect(ensureFolders(LIB, [], fakeGateway(), SITE)).rejects.toThrow(
			'at least one folder',
		);
		await expect(
			ensureFolders(LIB, [entry('a', 'A', { fieldsAlways: { Nope: 1 } })], fakeGateway(), SITE),
		).rejects.toThrow('do not exist in the library: Nope');
		await expect(
			ensureFolders(`${LIB}/Missing`, [entry('a', 'A')], fakeGateway(), SITE),
		).rejects.toThrow('does not exist');
	});

	it('makes site-relative and absolute URLs', () => {
		expect(folderUrls(SITE, `${LIB}/A #1`)).toEqual({
			siteRelativeUrl: 'Shared Documents/A #1',
			absoluteUrl: 'https://contoso.sharepoint.com/sites/Example/Shared%20Documents/A%20%231',
		});
		expect(folderUrls('https://contoso.sharepoint.com', '/Shared Documents/A')).toEqual({
			siteRelativeUrl: 'Shared Documents/A',
			absoluteUrl: 'https://contoso.sharepoint.com/Shared%20Documents/A',
		});
	});
});

describe('Folder description for Ensure Folders', () => {
	it('lists Ensure Folders in order', () => {
		const names = (folderResource.description[0].options as Array<{ name: string }>).map(
			(option) => option.name,
		);
		expect(names).toEqual([...names].sort());
		expect(names).toContain('Ensure Folders');
		expect(Object.keys(folderResource.operations)).toContain('ensureFolders');
	});
});
