import { sleep } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { joinPath } from '../../nodes/SharePointPnp/shared/utils';
import {
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

describe.runIf(hasConfig)('list items', () => {
	let testFolder: string;
	let subFolder: string;
	let deepFolder: string;
	let listId: string;
	let subId: number;

	const itemByPath = async (path: string) =>
		await createTestSp()
			.web.getFolderByServerRelativePath(path)
			.listItemAllFields.select('Id', 'Modified', 'Title')();

	beforeAll(async () => {
		testFolder = await createTestFolder();
		subFolder = joinPath(testFolder, 'Sub');
		deepFolder = joinPath(subFolder, 'Deep');
		const sp = createTestSp();
		await sp.web.folders.addUsingPath(subFolder);
		await sp.web.folders.addUsingPath(deepFolder);

		const parent = await sp.web
			.getFolderByServerRelativePath(testFolder)
			.listItemAllFields.select('Id', 'ParentList/Id')
			.expand('ParentList')();
		listId = (parent as unknown as { ParentList: { Id: string } }).ParentList.Id;
		subId = (await itemByPath(subFolder)).Id as number;
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	const listParams = () => ({ list: locator(listId, 'id') });

	it('gets an item by path', async () => {
		const [result] = await runNode({
			resource: 'listItem',
			operation: 'get',
			identifyBy: 'path',
			itemPath: locator(subFolder),
		});
		expect(result.json.Id).toBe(subId);
		expect(result.json.FileLeafRef).toBe('Sub');
	});

	it('gets an item by ID, as raw and as text', async () => {
		const base = { resource: 'listItem', operation: 'get', identifyBy: 'id', ...listParams() };
		const [raw] = await runNode({
			...base,
			itemId: subId,
			getOptions: { selectFields: 'Id,FileLeafRef' },
		});
		expect(raw.json.FileLeafRef).toBe('Sub');

		const [text] = await runNode({
			...base,
			itemId: subId,
			getOptions: { returnValuesAs: 'text' },
		});
		expect(text.json.FileLeafRef).toBe('Sub');
	});

	it('gets many items in the test folder with an OData filter', async () => {
		const base = {
			resource: 'listItem',
			operation: 'getAll',
			...listParams(),
			queryMode: 'filter',
			returnAll: true,
			selectFields: 'Id,FileLeafRef',
		};
		const deepId = (await itemByPath(deepFolder)).Id as number;
		const direct = await runNode({
			...base,
			odataFilter: `Id eq ${subId}`,
			getAllOptions: { folderPath: testFolder },
		});
		expect(direct.map((entry) => entry.json.FileLeafRef)).toEqual(['Sub']);

		const notDirect = await runNode({
			...base,
			odataFilter: `Id eq ${deepId}`,
			getAllOptions: { folderPath: testFolder },
		});
		expect(notDirect).toHaveLength(0);

		const recursive = await runNode({
			...base,
			odataFilter: `Id eq ${deepId}`,
			getAllOptions: { folderPath: testFolder, recursive: true },
		});
		expect(recursive.map((entry) => entry.json.FileLeafRef)).toEqual(['Deep']);
	});

	it('gets many items in the test folder with a CAML query', async () => {
		const caml =
			'<Where><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">1</Value></Eq></Where>';
		const base = {
			resource: 'listItem',
			operation: 'getAll',
			...listParams(),
			queryMode: 'caml',
			camlQuery: caml,
			returnAll: true,
		};
		const direct = await runNode({ ...base, getAllOptions: { folderPath: testFolder } });
		expect(direct).toHaveLength(1);

		const recursive = await runNode({
			...base,
			getAllOptions: { folderPath: testFolder, recursive: true },
		});
		expect(recursive).toHaveLength(2);

		const limited = await runNode({
			...base,
			returnAll: false,
			limit: 1,
			getAllOptions: { folderPath: testFolder, recursive: true },
		});
		expect(limited).toHaveLength(1);
	});

	it('updates an item and changes Modified', async () => {
		const before = await itemByPath(deepFolder);
		await sleep(1500);
		const [updated] = await runNode({
			resource: 'listItem',
			operation: 'update',
			identifyBy: 'path',
			itemPath: locator(deepFolder),
			fieldsJson: JSON.stringify({ Title: 'Changed' }),
		});
		expect(updated.json.Title).toBe('Changed');
		expect((await itemByPath(deepFolder)).Modified).not.toBe(before.Modified);
	});

	it('updates an item with System Update and keeps Modified', async () => {
		const before = await itemByPath(subFolder);
		await sleep(1500);
		const [updated] = await runNode({
			resource: 'listItem',
			operation: 'update',
			identifyBy: 'path',
			itemPath: locator(subFolder),
			fieldsJson: JSON.stringify({ Title: 'System Changed' }),
			updateOptions: { systemUpdate: true },
		});
		expect(updated.json.Title).toBe('System Changed');
		const after = await itemByPath(subFolder);
		expect(after.Modified).toBe(before.Modified);
	});

	it('updates an item by ID with the JSON input', async () => {
		const [updated] = await runNode({
			resource: 'listItem',
			operation: 'update',
			identifyBy: 'id',
			...listParams(),
			itemId: subId,
			fieldsInputMode: 'json',
			fieldsJson: { Title: 'By ID' },
		});
		expect(updated.json.Title).toBe('By ID');
	});

	it('deletes an item permanently', async () => {
		const sp = createTestSp();
		const doomed = joinPath(testFolder, 'Doomed');
		await sp.web.folders.addUsingPath(doomed);
		const [result] = await runNode({
			resource: 'listItem',
			operation: 'delete',
			identifyBy: 'path',
			itemPath: locator(doomed),
			deleteOptions: { permanent: true },
		});
		expect(result.json.recycled).toBe(false);
		const folder = await sp.web.getFolderByServerRelativePath(doomed).select('Exists')();
		expect(folder.Exists).toBe(false);
	});

	it('gets fields, views and default column values of the library', async () => {
		const fields = await runNode({
			resource: 'list',
			operation: 'getFields',
			...listParams(),
			fieldFilters: {},
		});
		expect(fields.some((entry) => entry.json.InternalName === 'Title')).toBe(true);

		const hidden = await runNode({
			resource: 'list',
			operation: 'getFields',
			...listParams(),
			fieldFilters: { includeHidden: true, includeReadOnly: true },
		});
		expect(hidden.length).toBeGreaterThan(fields.length);

		const views = await runNode({
			resource: 'list',
			operation: 'getViews',
			...listParams(),
			getViewsOptions: { includeFields: true },
		});
		expect(views.length).toBeGreaterThan(0);
		expect(Array.isArray(views[0].json.ViewFields)).toBe(true);

		const defaults = await runNode({
			resource: 'list',
			operation: 'getDefaultColumnValues',
			...listParams(),
			getDefaultsOptions: { folderPath: testFolder },
		});
		expect(defaults).toHaveLength(0);
	});

	// These tests change the tenant and need manual cleanup. Run them by hand in a sandbox site.
	it.skip('creates and deletes a list', async () => {
		const [created] = await runNode({
			resource: 'list',
			operation: 'create',
			title: `__pnp-test-${Date.now()}`,
			template: 100,
			createOptions: { onQuickLaunch: false },
		});
		await runNode({
			resource: 'list',
			operation: 'delete',
			list: locator(created.json.Id as string, 'id'),
			deleteOptions: { permanent: true },
		});
	});

	it.skip('creates a view and copies it to another list', async () => {
		await runNode({
			resource: 'list',
			operation: 'createView',
			...listParams(),
			viewTitle: `__pnp-test-${Date.now()}`,
			viewFields: ['LinkFilename'],
			createViewOptions: { rowLimit: 10 },
		});
	});

	it.skip('sets default column values of a folder', async () => {
		await runNode({
			resource: 'list',
			operation: 'setDefaultColumnValues',
			...listParams(),
			defaultsFolderPath: testFolder,
			defaultValues: { values: [{ name: 'Title', value: 'Example' }] },
		});
	});

	it.skip('creates a folder item with the field mapper', async () => {
		await runNode({
			resource: 'listItem',
			operation: 'create',
			...listParams(),
			fieldsInputMode: 'json',
			fieldsJson: { FileLeafRef: 'Created' },
			createOptions: { asFolder: true, folderPath: testFolder },
		});
	});
});
