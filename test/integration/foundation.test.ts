import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { createTestFolder, createTestSp, hasConfig, removeTestFolder, runNode } from './harness';

describe.runIf(hasConfig)('foundation', () => {
	let testFolder: string | undefined;

	beforeAll(async () => {
		testFolder = await createTestFolder();
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('gets lists through the node with the certificate credential', async () => {
		const items = await runNode({
			resource: 'list',
			operation: 'getAll',
			returnAll: true,
			filters: { baseTemplate: 101 },
		});
		expect(items.length).toBeGreaterThan(0);
		expect(items[0].json).toHaveProperty('Title');
		expect(items[0].pairedItem).toEqual({ item: 0 });
	});

	it('applies the selection options to list reads', async () => {
		const lists = await runNode({
			resource: 'list',
			operation: 'getAll',
			returnAll: false,
			limit: 1,
			filters: { baseTemplate: 101 },
			getAllOptions: { additionalProperties: 'EntityTypeName' },
		});
		const first = lists[0].json;
		expect(first).toHaveProperty('Title');
		expect(first).toHaveProperty('EntityTypeName');
		expect(first).toHaveProperty('RootFolder');
		const list = { __rl: true, mode: 'id', value: String(first.Id) };

		const plain = await runNode({ resource: 'list', operation: 'get', list, getOptions: {} });
		expect(plain[0].json).toHaveProperty('ItemCount');
		expect(plain[0].json).not.toHaveProperty('EntityTypeName');
		const everything = await runNode({
			resource: 'list',
			operation: 'get',
			list,
			getOptions: { returnAllProperties: true },
		});
		expect(everything[0].json).toHaveProperty('EntityTypeName');
		expect(everything[0].json).toHaveProperty('RootFolder');

		const fields = await runNode({
			resource: 'list',
			operation: 'getFields',
			list,
			fieldFilters: {},
			getFieldsOptions: { additionalProperties: 'SchemaXml' },
		});
		expect(fields.length).toBeGreaterThan(0);
		expect(fields[0].json).toHaveProperty('InternalName');
		expect(fields[0].json).toHaveProperty('SchemaXml');

		const views = await runNode({
			resource: 'list',
			operation: 'getViews',
			list,
			getViewsOptions: { additionalProperties: 'ListViewXml' },
		});
		expect(views.length).toBeGreaterThan(0);
		expect(views[0].json).toHaveProperty('ServerRelativeUrl');
		expect(views[0].json).toHaveProperty('ListViewXml');
	});

	it('created the test folder', async () => {
		const folder = await createTestSp()
			.web.getFolderByServerRelativePath(testFolder!)
			.select('Exists', 'ServerRelativeUrl')();
		expect(folder.Exists).toBe(true);
	});
});
