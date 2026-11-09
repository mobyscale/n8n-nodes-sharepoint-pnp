import type { IDataObject, IExecuteFunctions, ILoadOptionsFunctions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const sp = vi.hoisted(() => ({ current: {} as Record<string, unknown> }));

vi.mock('../../nodes/SharePointPnp/transport', () => ({
	createSp: () => sp.current,
	getTenantUrl: async () => 'https://contoso.sharepoint.com',
}));

import { getFolderListFieldOptions } from '../../nodes/SharePointPnp/methods/listSearchFields';
import { fileFolderResource } from '../../nodes/SharePointPnp/resources/fileFolder';
import { toDepthRange, walkFolders } from '../../nodes/SharePointPnp/resources/fileFolder/walk';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';
const ROOT = '/sites/Example/Shared Documents/Reports';

/** Reports/a.txt, Reports/2024/b.txt, Reports/2024/Q1/c.txt, Reports/2024/Q1/Final/d.txt */
const TREE: Record<string, IDataObject[]> = {
	[ROOT]: [
		{ name: '2024', type: 'folder', serverRelativeUrl: `${ROOT}/2024` },
		{ name: 'a.txt', type: 'file', serverRelativeUrl: `${ROOT}/a.txt` },
	],
	[`${ROOT}/2024`]: [
		{ name: 'Q1', type: 'folder', serverRelativeUrl: `${ROOT}/2024/Q1` },
		{ name: 'b.txt', type: 'file', serverRelativeUrl: `${ROOT}/2024/b.txt` },
	],
	[`${ROOT}/2024/Q1`]: [
		{ name: 'Final', type: 'folder', serverRelativeUrl: `${ROOT}/2024/Q1/Final` },
		{ name: 'c.txt', type: 'file', serverRelativeUrl: `${ROOT}/2024/Q1/c.txt` },
	],
	[`${ROOT}/2024/Q1/Final`]: [
		{ name: 'd.txt', type: 'file', serverRelativeUrl: `${ROOT}/2024/Q1/Final/d.txt` },
	],
};

function readers(outputHasAllFolders = true) {
	const readOutput = vi.fn(async (path: string, max: number) => (TREE[path] ?? []).slice(0, max));
	const readFolders = vi.fn(async (path: string) =>
		(TREE[path] ?? []).filter((entry) => entry.type === 'folder'),
	);
	return { readOutput, readFolders, outputHasAllFolders };
}

const names = (entries: IDataObject[]) => entries.map((entry) => entry.name);

describe('depth range', () => {
	it('reads only depth 1 without Recursive', () => {
		expect(toDepthRange(false, 3, 5)).toEqual({ min: 1, max: 1 });
	});

	it('uses no limit for Max Depth 0 and at least 1 for Min Depth', () => {
		expect(toDepthRange(true)).toEqual({ min: 1, max: Infinity });
		expect(toDepthRange(true, 0, 0)).toEqual({ min: 1, max: Infinity });
		expect(toDepthRange(true, 2.7, 3)).toEqual({ min: 2, max: 3 });
	});

	it('rejects a Min Depth above the Max Depth', () => {
		expect(() => toDepthRange(true, 3, 2)).toThrow('must not be greater than Max Depth');
	});
});

describe('folder walk', () => {
	it('walks all levels and adds the depth', async () => {
		const result = await walkFolders(ROOT, readers(), {
			depth: toDepthRange(true),
			limit: Infinity,
			addDepth: true,
		});
		expect(result.map((entry) => [entry.name, entry.depth])).toEqual([
			['2024', 1],
			['a.txt', 1],
			['Q1', 2],
			['b.txt', 2],
			['Final', 3],
			['c.txt', 3],
			['d.txt', 4],
		]);
	});

	it('does not read folders below Max Depth', async () => {
		const read = readers();
		const result = await walkFolders(ROOT, read, {
			depth: toDepthRange(true, 1, 2),
			limit: Infinity,
			addDepth: true,
		});
		expect(names(result)).toEqual(['2024', 'a.txt', 'Q1', 'b.txt']);
		expect(read.readOutput.mock.calls.map(([path]) => path)).toEqual([ROOT, `${ROOT}/2024`]);
		expect(read.readFolders).not.toHaveBeenCalled();
	});

	it('walks but does not return the levels above Min Depth', async () => {
		const read = readers();
		const result = await walkFolders(ROOT, read, {
			depth: toDepthRange(true, 3, 3),
			limit: Infinity,
			addDepth: true,
		});
		expect(names(result)).toEqual(['Final', 'c.txt']);
		expect(read.readFolders.mock.calls.map(([path]) => path)).toEqual([ROOT, `${ROOT}/2024`]);
		expect(read.readOutput.mock.calls.map(([path]) => path)).toEqual([`${ROOT}/2024/Q1`]);
	});

	it('reads the subfolders on their own when the output is filtered', async () => {
		const read = readers(false);
		read.readOutput.mockImplementation(async (path: string) =>
			(TREE[path] ?? []).filter((entry) => entry.type === 'file'),
		);
		const result = await walkFolders(ROOT, read, {
			depth: toDepthRange(true),
			limit: Infinity,
			addDepth: false,
		});
		expect(names(result)).toEqual(['a.txt', 'b.txt', 'c.txt', 'd.txt']);
		expect(result[0].depth).toBeUndefined();
		expect(read.readFolders).toHaveBeenCalledTimes(4);
	});

	it('stops at the limit and applies the accept filter', async () => {
		const limited = await walkFolders(ROOT, readers(), {
			depth: toDepthRange(true),
			limit: 3,
			addDepth: false,
		});
		expect(names(limited)).toEqual(['2024', 'a.txt', 'Q1']);
		const accepted = await walkFolders(
			ROOT,
			{ ...readers(), accept: (entry) => String(entry.name).endsWith('.txt') },
			{ depth: toDepthRange(true), limit: 2, addDepth: false },
		);
		expect(names(accepted)).toEqual(['a.txt', 'b.txt']);
	});

	it('reads only the folder itself without Recursive', async () => {
		const read = readers();
		const result = await walkFolders(ROOT, read, {
			depth: toDepthRange(false),
			limit: Infinity,
			addDepth: false,
		});
		expect(names(result)).toEqual(['2024', 'a.txt']);
		expect(read.readOutput).toHaveBeenCalledTimes(1);
	});
});

function rowsFor(path: string, viewXml: string): IDataObject[] {
	const foldersOnly = viewXml.includes('<Value Type="Integer">1</Value>');
	return (TREE[path] ?? [])
		.filter((entry) => !foldersOnly || entry.type === 'folder')
		.map((entry) => ({
			FileLeafRef: entry.name,
			FileRef: entry.serverRelativeUrl,
			FSObjType: entry.type === 'folder' ? '1' : '0',
		}));
}

describe('File/Folder Get Many with depth and field filters', () => {
	let render: ReturnType<typeof vi.fn>;
	let fields: ReturnType<typeof vi.fn>;

	beforeEach(() => {
		render = vi.fn(async (parameters: { ViewXml: string; FolderServerRelativeUrl: string }) => ({
			Row: rowsFor(parameters.FolderServerRelativeUrl, parameters.ViewXml),
		}));
		fields = vi.fn().mockResolvedValue([
			{ InternalName: 'Title', TypeAsString: 'Text' },
			{ InternalName: 'Status', TypeAsString: 'Choice' },
		]);
		const lists = Object.assign(
			{
				select: () => ({
					expand: () => async () => [
						{ Id: 'lib', RootFolder: { ServerRelativeUrl: '/sites/Example/Shared Documents' } },
					],
				}),
			},
			{
				getById: () => ({
					renderListDataAsStream: render,
					fields: { select: () => fields },
				}),
			},
		);
		sp.current = { web: { lists } };
	});

	const run = async (parameters: IDataObject) => {
		const all: IDataObject = {
			folderPath: { __rl: true, mode: 'path', value: ROOT },
			include: 'all',
			returnAll: true,
			filters: {},
			options: {},
			...parameters,
		};
		const context = {
			getNode: () => ({ name: 'SharePoint PnP' }),
			getNodeParameter: (
				name: string,
				_index: number,
				fallback?: unknown,
				options?: { extractValue?: boolean },
			) => {
				const value = name in all ? all[name] : fallback;
				return options?.extractValue && value && typeof value === 'object'
					? (value as IDataObject).value
					: value;
			},
		} as unknown as IExecuteFunctions;
		return (await fileFolderResource.operations.getAll.call(context, {
			sp: sp.current as never,
			siteUrl,
			itemIndex: 0,
		})) as IDataObject[];
	};

	it('keeps the output without depth when Recursive is off', async () => {
		const result = await run({ recursive: false, options: { minDepth: 2 } });
		expect(names(result)).toEqual(['2024', 'a.txt']);
		expect(result[0]).not.toHaveProperty('depth');
	});

	it('returns only the levels between Min Depth and Max Depth', async () => {
		const result = await run({ recursive: true, options: { minDepth: 2, maxDepth: 3 } });
		expect(result.map((entry) => [entry.name, entry.depth])).toEqual([
			['Q1', 2],
			['b.txt', 2],
			['Final', 3],
			['c.txt', 3],
		]);
		const read = render.mock.calls.map(([parameters]) => parameters.FolderServerRelativeUrl);
		expect(read).not.toContain(`${ROOT}/2024/Q1/Final`);
	});

	it('adds the field filters to the CAML of each folder', async () => {
		await run({
			recursive: true,
			fieldValueFilters: {
				conditions: [
					{ field: 'Title', operator: 'eq', value: 'Draft' },
					{ field: 'Status', operator: 'isEmpty' },
				],
			},
			options: { fieldFiltersCombine: 'or' },
		});
		const views = render.mock.calls.map(([parameters]) => parameters.ViewXml as string);
		const filtered = views.filter((xml) => xml.includes('<Or>'));
		expect(filtered).toHaveLength(4);
		expect(filtered[0]).toContain(
			'<Or><Eq><FieldRef Name="Title"/><Value Type="Text">Draft</Value></Eq><IsNull><FieldRef Name="Status"/></IsNull></Or>',
		);
		// The filtered output does not have all subfolders, so the folders are read with their own view.
		expect(views.length).toBe(8);
	});

	it('fails with the name of a filter field that does not exist', async () => {
		await expect(
			run({
				recursive: false,
				fieldValueFilters: { conditions: [{ field: 'Category', operator: 'eq', value: 'a' }] },
			}),
		).rejects.toThrow('Category');
	});
});

describe('folder field picker', () => {
	const context = (folderPath: unknown) =>
		({
			getNode: () => ({ name: 'SharePoint PnP' }),
			getCredentials: async () => ({ subdomain: 'contoso' }),
			getNodeParameter: (name: string, fallback?: unknown) =>
				name === 'site' ? siteUrl : name === 'folderPath' ? folderPath : fallback,
		}) as unknown as ILoadOptionsFunctions;

	beforeEach(() => {
		const fieldQuery = vi.fn().mockResolvedValue([
			{ InternalName: 'Status', Title: 'Status' },
			{ InternalName: 'Title', Title: 'Title' },
			{ InternalName: 'Category', Title: 'Category' },
		]);
		const getById = vi.fn().mockReturnValue({
			fields: { filter: vi.fn().mockReturnValue({ select: () => fieldQuery }) },
		});
		sp.current = {
			web: {
				lists: Object.assign(
					{
						select: () => ({
							expand: () => async () => [
								{ Id: 'lib', RootFolder: { ServerRelativeUrl: '/sites/Example/Shared Documents' } },
								{ Id: 'other', RootFolder: { ServerRelativeUrl: '/sites/Example/Lists/Other' } },
							],
						}),
					},
					{ getById },
				),
			},
			getById,
		};
	});

	it('lists the fields of the library that contains the folder, A-Z', async () => {
		const result = await getFolderListFieldOptions.call(
			context({ __rl: true, mode: 'path', value: 'Shared Documents/Reports/2024' }),
		);
		expect(result.map((entry) => entry.value)).toEqual(['Category', 'Status', 'Title']);
		expect(result[0].name).toBe('Category (Category)');
		expect((sp.current.getById as ReturnType<typeof vi.fn>).mock.calls[0][0]).toBe('lib');
	});

	it('returns nothing while no folder is set', async () => {
		expect(await getFolderListFieldOptions.call(context({ mode: 'path', value: '' }))).toEqual([]);
	});

	it('fails when no library contains the folder', async () => {
		await expect(getFolderListFieldOptions.call(context('/sites/Example/Archive'))).rejects.toThrow(
			'No list or library contains',
		);
	});
});
