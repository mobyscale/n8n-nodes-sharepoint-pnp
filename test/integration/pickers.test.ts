import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	getFolderListFieldOptions,
	getListFieldOptions,
	getLookupListFieldOptions,
} from '../../nodes/SharePointPnp/methods/listSearchFields';
import {
	searchColumnGroups,
	searchListFolders,
} from '../../nodes/SharePointPnp/methods/listSearchScoped';
import { searchListItems } from '../../nodes/SharePointPnp/methods/listSearchItems';
import { joinPath } from '../../nodes/SharePointPnp/shared/utils';
import {
	createLoadOptionsContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

describe.runIf(hasConfig)('pickers', () => {
	let testFolder: string;
	let relativeFolder: string;
	let listId: string;
	let subId: number;

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const sp = createTestSp();
		await sp.web.folders.addUsingPath(joinPath(testFolder, 'PickerSub'));

		const parent = (await sp.web
			.getFolderByServerRelativePath(testFolder)
			.listItemAllFields.select('Id', 'ParentList/Id', 'ParentList/RootFolder/ServerRelativeUrl')
			.expand('ParentList', 'ParentList/RootFolder')()) as unknown as {
			ParentList: { Id: string; RootFolder: { ServerRelativeUrl: string } };
		};
		listId = parent.ParentList.Id;
		relativeFolder = testFolder.slice(parent.ParentList.RootFolder.ServerRelativeUrl.length + 1);
		const sub = await sp.web
			.getFolderByServerRelativePath(joinPath(testFolder, 'PickerSub'))
			.listItemAllFields.select('Id')();
		subId = sub.Id as number;
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	const context = (extra: Record<string, unknown> = {}) =>
		createLoadOptionsContext({ list: locator(listId, 'id'), ...extra });

	it('finds an item of the test folder with the item search', async () => {
		const byId = await searchListItems.call(context(), String(subId));
		expect(byId.results.map((entry) => entry.value)).toEqual([String(subId)]);
		// Above the list view threshold the text filter fails and the newest items are matched instead.
		const byText = await searchListItems.call(context(), 'PickerSub');
		expect(byText.results.map((entry) => entry.value)).toContain(String(subId));
	});

	it('pages the item search fallback through older items (read-only)', async () => {
		const first = await searchListItems.call(context(), 'zz-no-such-item-zz');
		expect(first.results).toEqual([]);
		// The library has more than 300 items, so the scan has a next page.
		expect(first.paginationToken).toMatch(/^scan:\d+$/);
		const second = await searchListItems.call(
			context(),
			'zz-no-such-item-zz',
			first.paginationToken,
		);
		expect(second.paginationToken).toMatch(/^scan:\d+$/);
		expect(Number(second.paginationToken?.slice(5))).toBeLessThan(
			Number(first.paginationToken?.slice(5)),
		);
	});

	it('lists the folders of the list below a path', async () => {
		const result = await searchListFolders.call(context(), `${relativeFolder}/Picker`);
		expect(result.results.map((entry) => entry.value)).toEqual([joinPath(testFolder, 'PickerSub')]);
		const root = await searchListFolders.call(context(), '');
		expect(root.results[0].name).toBe('(list root)');
	});

	it('lists the fields of the list and of a lookup list', async () => {
		const fields = await getListFieldOptions.call(context());
		expect(fields.map((entry) => entry.value)).toContain('FileLeafRef');
		const lookup = await getLookupListFieldOptions.call(
			context({ lookupList: locator(listId, 'id') }),
		);
		expect(lookup.map((entry) => entry.value)).toContain('FileLeafRef');
	});

	it('lists the fields of the library that contains a folder', async () => {
		const fields = await getFolderListFieldOptions.call(
			createLoadOptionsContext({ folderPath: locator(testFolder) }),
		);
		const values = fields.map((entry) => entry.value);
		expect(values).toContain('Title');
		expect(values).toContain('FileLeafRef');
		const labels = fields.map((entry) => entry.name);
		expect(labels).toEqual([...labels].sort((a, b) => a.localeCompare(b)));
		const empty = await getFolderListFieldOptions.call(
			createLoadOptionsContext({ folderPath: locator('') }),
		);
		expect(empty).toEqual([]);
	});

	it('lists the column groups of the site', async () => {
		const result = await searchColumnGroups.call(context(), '');
		expect(Array.isArray(result.results)).toBe(true);
	});

	it('uses the picker values in the Get and Get Many operations', async () => {
		const [item] = await runNode({
			resource: 'listItem',
			operation: 'get',
			identifyBy: 'id',
			list: locator(listId, 'id'),
			itemId: locator(String(subId), 'list'),
			getOptions: { selectFieldsPicker: ['Id', 'FileLeafRef'], selectFields: 'FileRef' },
		});
		expect(item.json).toMatchObject({ Id: subId, FileLeafRef: 'PickerSub' });
		expect(item.json.FileRef).toBe(joinPath(testFolder, 'PickerSub'));

		const picked = await runNode({
			resource: 'listItem',
			operation: 'getAll',
			list: locator(listId, 'id'),
			queryMode: 'filter',
			odataFilter: `Id eq ${subId}`,
			returnAll: true,
			selectFieldsPicker: ['Id', 'FileLeafRef'],
			getAllOptions: {},
		});
		expect(picked.map((row) => row.json.FileLeafRef)).toEqual(['PickerSub']);

		// The folder filter of the filter mode can pass the list view threshold on a large library.
		const inFolder = await runNode({
			resource: 'listItem',
			operation: 'getAll',
			list: locator(listId, 'id'),
			queryMode: 'caml',
			camlQuery: '<Where><IsNotNull><FieldRef Name="ID"/></IsNotNull></Where>',
			returnAll: true,
			getAllOptions: { folderPath: locator(testFolder, 'list') },
		});
		expect(inFolder.map((row) => row.json.FileLeafRef)).toContain('PickerSub');
	});

	it('still accepts a plain item ID from an older workflow', async () => {
		const [item] = await runNode({
			resource: 'listItem',
			operation: 'get',
			identifyBy: 'id',
			list: locator(listId, 'id'),
			itemId: subId,
		});
		expect(item.json.Id).toBe(subId);
	});
});
