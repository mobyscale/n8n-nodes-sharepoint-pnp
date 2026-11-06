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

	it('created the test folder', async () => {
		const folder = await createTestSp()
			.web.getFolderByServerRelativePath(testFolder!)
			.select('Exists', 'ServerRelativeUrl')();
		expect(folder.Exists).toBe(true);
	});
});
