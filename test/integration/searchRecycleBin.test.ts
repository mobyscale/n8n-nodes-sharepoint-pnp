import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	config,
	createTestFolder,
	createTestSp,
	hasConfig,
	removeTestFolder,
	runNode,
} from './harness';

describe.runIf(hasConfig)('search (read-only)', () => {
	it('runs a query and returns flattened rows', async () => {
		const items = await runNode({
			resource: 'search',
			operation: 'query',
			query: 'contentclass:STS_Site',
			selectProperties: 'Title,Path,SiteId',
			returnAll: false,
			limit: 3,
		});
		expect(items.length).toBeLessThanOrEqual(3);
		for (const item of items) {
			expect(item.json).toHaveProperty('Path');
			expect(item.pairedItem).toEqual({ item: 0 });
		}
	});

	it('pages with a small page size and respects the limit', async () => {
		const items = await runNode({
			resource: 'search',
			operation: 'query',
			query: 'contentclass:STS_Site OR contentclass:STS_Web',
			selectProperties: 'Title,Path',
			returnAll: false,
			limit: 5,
			options: {
				rowLimitPerPage: 2,
				sort: { values: [{ property: 'LastModifiedTime', direction: 'descending' }] },
			},
		});
		expect(items.length).toBeLessThanOrEqual(5);
		const paths = items.map((item) => item.json.Path);
		expect(new Set(paths).size).toBe(paths.length);
	});

	it('scopes the query to the site', async () => {
		const items = await runNode({
			resource: 'search',
			operation: 'query',
			query: 'IsDocument:true',
			selectProperties: 'Path',
			scopeToSite: true,
			returnAll: false,
			limit: 5,
		});
		for (const item of items) {
			expect(String(item.json.Path).toLowerCase()).toContain(config!.siteUrl.toLowerCase());
		}
	});

	it('adds the raw result on request', async () => {
		const items = await runNode({
			resource: 'search',
			operation: 'query',
			query: 'contentclass:STS_Site',
			selectProperties: 'Path',
			returnAll: false,
			limit: 1,
			options: { includeRawResult: true },
		});
		expect(items.length).toBeGreaterThan(0);
		expect(items[items.length - 1].json).toHaveProperty('rawResult');
	});
});

describe.runIf(hasConfig)('recycle bin', () => {
	let testFolder: string | undefined;
	const idsToClean: string[] = [];

	const fileName = `__pnp-test-file-${Date.now()}.txt`;
	const deletedAfter = new Date(Date.now() - 60_000).toISOString();
	const bin = (operation: string, extra: Record<string, unknown> = {}) =>
		runNode({ resource: 'recycleBin', operation, scope: 'siteCollection', ...extra });

	async function recycleFile(create: boolean): Promise<string> {
		const sp = createTestSp();
		const path = `${testFolder}/${fileName}`;
		if (create) {
			await sp.web
				.getFolderByServerRelativePath(testFolder!)
				.files.addUsingPath(fileName, 'recycle bin test');
		}
		const id = await sp.web.getFileByServerRelativePath(path).recycle();
		idsToClean.push(id);
		return id;
	}

	async function findInBin(stage: number): Promise<string[]> {
		const items = await bin('getAll', {
			returnAll: true,
			filters: { text: fileName, itemType: 1, stage, deletedAfter },
		});
		return items.map((item) => String(item.json.id));
	}

	beforeAll(async () => {
		testFolder = await createTestFolder();
	});

	afterAll(async () => {
		const sp = createTestSp();
		for (const id of idsToClean) {
			try {
				await sp.site.recycleBin.getById(id).delete();
			} catch {
				// Already deleted by the test.
			}
		}
		await removeTestFolder(testFolder);
	});

	it('finds, restores, moves and deletes a recycled file', async () => {
		const firstId = await recycleFile(true);

		const found = await bin('getAll', {
			returnAll: false,
			limit: 10,
			filters: { text: fileName, itemType: 1, stage: 1, deletedAfter },
		});
		const match = found.find((item) => item.json.id === firstId);
		expect(match).toBeDefined();
		expect(match!.json).toMatchObject({
			itemType: 'File',
			itemState: 'FirstStage',
			leafName: fileName,
		});
		expect(String(match!.json.dirName)).toContain('__pnp-test-');

		const inSite = await runNode({
			resource: 'recycleBin',
			operation: 'getAll',
			scope: 'site',
			returnAll: true,
			filters: { text: fileName, deletedAfter },
		});
		expect(inSite.map((item) => item.json.id)).toContain(firstId);

		const restored = await bin('restore', { ids: firstId });
		expect(restored[0].json).toMatchObject({ id: firstId, restored: true });
		expect(await findInBin(1)).not.toContain(firstId);
		const exists = await createTestSp()
			.web.getFolderByServerRelativePath(testFolder!)
			.files.getByUrl(fileName)
			.select('Name')();
		expect(exists.Name).toBe(fileName);

		const secondId = await recycleFile(false);
		expect(await findInBin(1)).toContain(secondId);

		const moved = await bin('moveToSecondStage', { ids: [secondId] });
		expect(moved[0].json).toMatchObject({ id: secondId, moved: true });
		expect(await findInBin(2)).toContain(secondId);
		expect(await findInBin(1)).not.toContain(secondId);

		const deleted = await bin('deletePermanently', { ids: secondId });
		expect(deleted[0].json).toMatchObject({ id: secondId, deleted: true });
		expect(await findInBin(2)).not.toContain(secondId);
		expect(await findInBin(1)).not.toContain(secondId);
	});

	it('rejects values that are not IDs', async () => {
		await expect(bin('restore', { ids: 'not-an-id' })).rejects.toThrow(/not valid/);
	});
});
