import type { IDataObject } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

/**
 * Builds two small trees in the test folder:
 *   Template: A, A/B, C and the file x.txt
 *   Copy:     A, D
 */
describe.runIf(hasConfig)('file/folder compare structure', () => {
	let testFolder: string;
	let source: string;
	let target: string;

	beforeAll(async () => {
		testFolder = await createTestFolder();
		source = `${testFolder}/Template`;
		target = `${testFolder}/Copy`;
		const sp = createTestSp();
		for (const path of [source, `${source}/A`, `${source}/A/B`, `${source}/C`]) {
			await sp.web.folders.addUsingPath(path);
		}
		await sp.web.getFolderByServerRelativePath(source).files.addUsingPath('x.txt', 'test');
		for (const path of [target, `${target}/A`, `${target}/D`]) {
			await sp.web.folders.addUsingPath(path);
		}
	}, 120_000);

	afterAll(async () => {
		await removeTestFolder(testFolder);
	}, 120_000);

	const compare = async (options: IDataObject, from = source, to = target) => {
		const [result] = await runNode({
			resource: 'fileFolder',
			operation: 'compareStructure',
			sourceFolder: locator(from),
			targetFolder: locator(to),
			compareOptions: options,
		});
		return result.json;
	};

	it('compares the first folder level by default', async () => {
		expect(await compare({})).toEqual({
			isComplete: false,
			missing: ['C'],
			extra: ['D'],
			matched: 1,
			source,
			target,
			targetExists: true,
		});
	});

	it('compares all folder levels with Max Depth 0', async () => {
		expect(await compare({ maxDepth: 0 })).toMatchObject({
			isComplete: false,
			missing: ['A/B', 'C'],
			extra: ['D'],
			matched: 1,
		});
	});

	it('compares files only', async () => {
		expect(await compare({ include: 'files', maxDepth: 0 })).toMatchObject({
			isComplete: false,
			missing: ['x.txt'],
			extra: [],
			matched: 0,
		});
	});

	it('compares folders and files', async () => {
		expect(await compare({ include: 'both', maxDepth: 0 })).toMatchObject({
			isComplete: false,
			missing: ['A/B', 'C', 'x.txt'],
			extra: ['D'],
			matched: 1,
		});
		expect(await compare({ include: 'both', maxDepth: 1 })).toMatchObject({
			missing: ['C', 'x.txt'],
			extra: ['D'],
			matched: 1,
		});
	});

	it('leaves out excluded names', async () => {
		expect(await compare({ include: 'both', maxDepth: 0, excludeNames: 'C, D\nx*' })).toMatchObject(
			{ isComplete: false, missing: ['A/B'], extra: [], matched: 1 },
		);
	});

	it('is complete when a folder is compared with itself', async () => {
		expect(await compare({ include: 'both', maxDepth: 0 }, source, source)).toMatchObject({
			isComplete: true,
			missing: [],
			extra: [],
			matched: 4,
		});
	});

	it('reports all entries as missing when the target does not exist', async () => {
		expect(await compare({ maxDepth: 0 }, source, `${testFolder}/Missing`)).toMatchObject({
			isComplete: false,
			missing: ['A', 'A/B', 'C'],
			extra: [],
			matched: 0,
			targetExists: false,
		});
	});
});
