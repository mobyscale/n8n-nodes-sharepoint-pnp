import type { IDataObject, INodeExecutionData } from 'n8n-workflow';
import { sleep } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { joinPath } from '../../nodes/SharePointPnp/shared/utils';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';
import {
	createExecuteContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
} from './harness';

type PerItem = (itemIndex: number) => unknown;

/** Runs the node with several input items. Parameters given as functions get a value for each item. */
async function runBulk(
	parameters: IDataObject,
	perItem: Record<string, PerItem>,
	itemCount: number,
): Promise<INodeExecutionData[]> {
	const items = Array.from({ length: itemCount }, (_, index) => ({ json: { index } }));
	const base = createExecuteContext(parameters, items);
	const context = {
		...base,
		getNodeParameter: (name: string, itemIndex: number, fallback?: unknown, options?: object) =>
			name in perItem
				? perItem[name](itemIndex)
				: base.getNodeParameter(name, itemIndex, fallback, options),
	} as typeof base;
	return (await new SharePointPnp().execute.call(context))[0];
}

describe.runIf(hasConfig)('list item bulk operations', () => {
	let testFolder: string;
	let listId: string;
	const suffix = Date.now();
	const names = [0, 1, 2].map((index) => `bulk-${suffix}-${index}.txt`);
	const ids: number[] = [];
	let goneId: number;

	const readItem = async (id: number) =>
		(await createTestSp()
			.web.lists.getById(listId)
			.items.getById(id)
			.select('Id', 'Title', 'Modified', 'FileLeafRef')()) as {
			Id: number;
			Title: string | null;
			Modified: string;
			FileLeafRef: string;
		};

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const sp = createTestSp();
		const folder = sp.web.getFolderByServerRelativePath(testFolder);
		for (const name of names) {
			await folder.files.addUsingPath(name, `content of ${name}`);
			const item = (await sp.web
				.getFileByServerRelativePath(joinPath(testFolder, name))
				.listItemAllFields.select('Id', 'ParentList/Id')
				.expand('ParentList')()) as unknown as { Id: number; ParentList: { Id: string } };
			ids.push(item.Id);
			listId = item.ParentList.Id;
		}

		// An ID that no item has: the ID of a file this test created and deleted. IDs are not reused.
		const gonePath = joinPath(testFolder, `bulk-${suffix}-gone.txt`);
		await folder.files.addUsingPath(`bulk-${suffix}-gone.txt`, 'gone');
		const gone = sp.web.getFileByServerRelativePath(gonePath);
		goneId = ((await gone.listItemAllFields.select('Id')()) as unknown as { Id: number }).Id;
		await gone.delete();
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	const listParams = () => ({ resource: 'listItem', list: locator(listId, 'id') });

	it('updates the Title of several files in one batch', async () => {
		const output = await runBulk(
			{ ...listParams(), operation: 'updateMany', fieldsInputMode: 'json' },
			{
				bulkItemId: (i) => String(ids[i]),
				fieldsJson: (i) => JSON.stringify({ Title: `Bulk ${suffix} ${i}` }),
			},
			3,
		);
		expect(output.map((entry) => entry.json)).toEqual(ids.map((id) => ({ id, success: true })));
		expect(output.map((entry) => entry.pairedItem)).toEqual([
			{ item: 0 },
			{ item: 1 },
			{ item: 2 },
		]);
		for (const [index, id] of ids.entries()) {
			expect((await readItem(id)).Title).toBe(`Bulk ${suffix} ${index}`);
		}
	});

	it('reports an item that does not exist and updates the others', async () => {
		const output = await runBulk(
			{
				...listParams(),
				operation: 'updateMany',
				fieldsInputMode: 'json',
				bulkUpdateOptions: { continueOnItemError: true },
			},
			{
				bulkItemId: (i) => String(i === 1 ? goneId : ids[i]),
				fieldsJson: (i) => ({ Title: `Partial ${suffix} ${i}` }),
			},
			3,
		);
		const results = output.map((entry) => entry.json);
		expect(results[0]).toEqual({ id: ids[0], success: true });
		expect(results[1]).toMatchObject({ id: goneId, success: false });
		expect(String(results[1].error)).not.toBe('');
		expect(results[2]).toEqual({ id: ids[2], success: true });
		expect((await readItem(ids[2])).Title).toBe(`Partial ${suffix} 2`);
	});

	it('updates several files with System Update and keeps Modified', async () => {
		const before = await Promise.all(ids.map(readItem));
		await sleep(1500);
		const output = await runBulk(
			{
				...listParams(),
				operation: 'updateMany',
				fieldsInputMode: 'json',
				bulkUpdateOptions: { systemUpdate: true },
			},
			{
				bulkItemId: (i) => ids[i],
				fieldsJson: (i) => ({ Title: `System ${suffix} ${i}` }),
			},
			3,
		);
		expect(output.every((entry) => entry.json.success === true)).toBe(true);
		const after = await Promise.all(ids.map(readItem));
		expect(after.map((item) => item.Title)).toEqual(ids.map((_, i) => `System ${suffix} ${i}`));
		expect(after.map((item) => item.Modified)).toEqual(before.map((item) => item.Modified));
	});

	it('upserts by file name: updates, then skips the unchanged item', async () => {
		const parameters = {
			...listParams(),
			operation: 'upsert',
			keyField: 'FileLeafRef',
			fieldsInputMode: 'json',
			fieldsJson: { Title: `Upserted ${suffix}` },
			upsertOptions: { skipUnchanged: true, ifSeveralMatch: 'fail', findInFolder: testFolder },
		};
		const perItem = { keyValue: () => names[0] };

		const [first] = await runBulk(parameters, perItem, 1);
		expect(first.json).toEqual({ id: ids[0], success: true, action: 'updated' });
		expect((await readItem(ids[0])).Title).toBe(`Upserted ${suffix}`);

		const before = await readItem(ids[0]);
		const [second] = await runBulk(parameters, perItem, 1);
		expect(second.json).toEqual({ id: ids[0], success: true, action: 'unchanged' });
		expect((await readItem(ids[0])).Modified).toBe(before.Modified);
	});

	it('upserts by Title in a batch', async () => {
		const output = await runBulk(
			{
				...listParams(),
				operation: 'upsert',
				keyField: 'Title',
				fieldsInputMode: 'json',
				upsertOptions: { ifSeveralMatch: 'fail', findInFolder: testFolder },
			},
			{
				keyValue: (i) => `System ${suffix} ${i + 1}`,
				fieldsJson: (i) => ({ Title: `By Title ${suffix} ${i + 1}` }),
			},
			2,
		);
		expect(output.map((entry) => entry.json)).toEqual([
			{ id: ids[1], success: true, action: 'updated' },
			{ id: ids[2], success: true, action: 'updated' },
		]);
		expect((await readItem(ids[2])).Title).toBe(`By Title ${suffix} 2`);
	});

	it('moves several files to the recycle bin', async () => {
		const output = await runBulk(
			{ ...listParams(), operation: 'deleteMany' },
			{ bulkItemId: (i) => ids[i] },
			3,
		);
		expect(output.map((entry) => entry.json)).toEqual(
			ids.map((id) => ({ id, success: true, recycled: true })),
		);
		const sp = createTestSp();
		for (const name of names) {
			const file = await sp.web
				.getFolderByServerRelativePath(testFolder)
				.files.filter(`Name eq '${name}'`)();
			expect(file).toHaveLength(0);
		}
	});
});
