import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import * as descriptions from '../../nodes/SharePointPnp/shared/descriptions';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';
import * as transport from '../../nodes/SharePointPnp/transport';

vi.mock('../../nodes/SharePointPnp/transport', async (importOriginal) => {
	const actual = await importOriginal<typeof transport>();
	return { ...actual, createSp: vi.fn() };
});

vi.mock('../../nodes/SharePointPnp/shared/descriptions', async (importOriginal) => {
	const actual = await importOriginal<typeof descriptions>();
	return { ...actual, getSiteUrl: vi.fn() };
});

const SITE = 'https://contoso.sharepoint.com/sites/Example';
const OTHER_SITE = 'https://contoso.sharepoint.com/sites/Other';
const LIST_A = { __rl: true, mode: 'id', value: 'aaaaaaaa-0000-0000-0000-000000000000' };
const LIST_B = { __rl: true, mode: 'title', value: 'Other List' };

const FIELDS = [
	{ InternalName: 'Title', TypeAsString: 'Text' },
	{ InternalName: 'Amount', TypeAsString: 'Number' },
	{ InternalName: 'Department', TypeAsString: 'Lookup' },
	{ InternalName: 'Notes', TypeAsString: 'Note' },
	{ InternalName: 'ID', TypeAsString: 'Counter' },
];

interface RequestLog {
	batch: number;
	kind: string;
	list?: string;
	args: unknown[];
}

interface FakeHandlers {
	fields?: (list: string) => unknown[];
	add?: (formValues: Array<{ FieldName: string; FieldValue: string }>, folder: string) => unknown[];
	update?: (id: number, formValues: Array<{ FieldName: string; FieldValue: string }>) => unknown[];
	find?: (filter: string, select: string[], top: number) => unknown[];
	read?: (id: number) => unknown;
	remove?: (id: number) => void;
}

let nextId = 100;

/**
 * A fake SPFI. Requests on a batched instance must start before `execute` runs, as with PnPjs.
 * Every request is logged with the number of its batch (0 for requests outside a batch).
 */
function createFakeSp(handlers: FakeHandlers = {}) {
	const log: RequestLog[] = [];
	let batches = 0;

	const build = (state?: { index: number; executed: boolean }) => {
		const request = (
			kind: string,
			list: string | undefined,
			args: unknown[],
			run: () => unknown,
		) => {
			if (state?.executed) throw new Error('A request started after the batch was executed');
			log.push({ batch: state?.index ?? 0, kind, list, args });
			return Promise.resolve().then(run);
		};

		const listApi = (list: string) => {
			const item = (id: number) => ({
				validateUpdateListItem: (
					formValues: Array<{ FieldName: string; FieldValue: string }>,
					system: boolean,
				) =>
					request('update', list, [id, formValues, system], () =>
						handlers.update
							? handlers.update(id, formValues)
							: formValues.map((entry) => ({ ...entry, HasException: false })),
					),
				recycle: () => request('recycle', list, [id], () => handlers.remove?.(id)),
				delete: () => request('delete', list, [id], () => handlers.remove?.(id)),
				select: (...select: string[]) => ({
					expand: () => () =>
						request('read', list, [id, select], () =>
							handlers.read
								? handlers.read(id)
								: {
										Modified: '2024-01-02T03:04:05Z',
										Editor: { Name: 'i:0#.f|membership|a@contoso.com' },
									},
						),
				}),
			});

			const findQuery = (select: string[]) => {
				let filter = '';
				let top = 0;
				const query = {
					expand: () => query,
					filter: (value: string) => {
						filter = value;
						return query;
					},
					orderBy: () => query,
					top: (value: number) => {
						top = value;
						return () =>
							request(
								'find',
								list,
								[filter, select, top],
								() => handlers.find?.(filter, select, top) ?? [],
							);
					},
				};
				return query;
			};

			return {
				fields: {
					select: () => () => request('fields', list, [], () => handlers.fields?.(list) ?? FIELDS),
				},
				rootFolder: {
					select: () => () =>
						request('rootFolder', list, [], () => ({
							ServerRelativeUrl: '/sites/Example/Lists/Example',
						})),
				},
				addValidateUpdateItemUsingPath: (
					formValues: Array<{ FieldName: string; FieldValue: string }>,
					folder: string,
				) =>
					request('add', list, [formValues, folder], () =>
						handlers.add
							? handlers.add(formValues, folder)
							: [{ FieldName: 'Id', FieldValue: String(nextId++), HasException: false }],
					),
				getItemsByCAMLQuery: (
					query: { ViewXml: string; FolderServerRelativeUrl?: string },
					...expand: string[]
				) =>
					request(
						'caml',
						list,
						[query, expand],
						() => handlers.find?.(query.ViewXml, expand, 0) ?? [],
					),
				items: { getById: item, select: (...select: string[]) => findQuery(select) },
			};
		};

		return {
			web: {
				lists: {
					getById: (id: string) => listApi(id),
					getByTitle: (title: string) => listApi(title),
				},
				regionalSettings: {
					timeZone: {
						utcToLocalTime: (utc: string) =>
							request('localTime', undefined, [utc], () =>
								utc.replace('Z', '').replace(/^2024/, '2023'),
							),
					},
				},
			},
			batched: () => {
				const batchState = { index: ++batches, executed: false };
				return [
					build(batchState),
					async () => {
						batchState.executed = true;
					},
				];
			},
		};
	};

	return { sp: build(), log, batchCount: () => batches };
}

type Parameter = unknown | ((itemIndex: number) => unknown);

function createContext(
	parameters: Record<string, Parameter>,
	itemCount: number,
	continueOnFail = false,
): IExecuteFunctions {
	return {
		getInputData: () => Array.from({ length: itemCount }, (_, index) => ({ json: { index } })),
		getNode: () => ({ name: 'SharePoint PnP', type: 'test', typeVersion: 1, parameters: {} }),
		continueOnFail: () => continueOnFail,
		getNodeParameter: (name: string, itemIndex: number, fallback?: unknown) => {
			if (!(name in parameters)) return fallback;
			const value = parameters[name];
			return typeof value === 'function' ? (value as (index: number) => unknown)(itemIndex) : value;
		},
	} as unknown as IExecuteFunctions;
}

async function run(
	parameters: Record<string, Parameter>,
	itemCount: number,
	fake = createFakeSp(),
	continueOnFail = false,
): Promise<INodeExecutionData[]> {
	vi.mocked(transport.createSp).mockReturnValue(fake.sp as never);
	const context = createContext(
		{ resource: 'listItem', list: LIST_A, ...parameters },
		itemCount,
		continueOnFail,
	);
	return (await new SharePointPnp().execute.call(context))[0];
}

const json = (output: INodeExecutionData[]) => output.map((entry) => entry.json);
const pairs = (output: INodeExecutionData[]) =>
	output.map((entry) => (entry.pairedItem as { item: number }).item);
const kinds = (log: RequestLog[], kind: string) => log.filter((entry) => entry.kind === kind);

beforeEach(() => {
	nextId = 100;
	vi.mocked(transport.createSp).mockReset();
	vi.mocked(descriptions.getSiteUrl).mockReset();
	vi.mocked(descriptions.getSiteUrl).mockResolvedValue(SITE);
});

describe('Create Many', () => {
	it('creates all input items in batches of the batch size', async () => {
		const fake = createFakeSp();
		const output = await run(
			{
				operation: 'createMany',
				fieldsInputMode: 'json',
				fieldsJson: (i: number) => JSON.stringify({ Title: `Item ${i}`, Amount: i }),
				bulkCreateOptions: { batchSize: 2 },
			},
			5,
			fake,
		);

		expect(json(output)).toEqual([100, 101, 102, 103, 104].map((id) => ({ id, success: true })));
		expect(pairs(output)).toEqual([0, 1, 2, 3, 4]);
		const adds = kinds(fake.log, 'add');
		expect(adds.map((entry) => entry.batch)).toEqual([1, 1, 2, 2, 3]);
		expect(adds[1].args).toEqual([
			[
				{ FieldName: 'Title', FieldValue: 'Item 1' },
				{ FieldName: 'Amount', FieldValue: '1' },
			],
			'/sites/Example/Lists/Example',
		]);
		expect(kinds(fake.log, 'fields')).toHaveLength(1);
		expect(kinds(fake.log, 'rootFolder')).toHaveLength(1);
	});

	it('uses one batch for each 100 items by default', async () => {
		const fake = createFakeSp();
		const output = await run(
			{
				operation: 'createMany',
				fieldsInputMode: 'json',
				fieldsJson: '{"Title":"x"}',
				bulkCreateOptions: { batchSize: 500 },
			},
			250,
			fake,
		);
		expect(output).toHaveLength(250);
		expect(fake.batchCount()).toBe(3);
		const sizes = [1, 2, 3].map(
			(batch) => fake.log.filter((entry) => entry.batch === batch).length,
		);
		expect(sizes).toEqual([100, 100, 50]);
	});

	it('creates items in the given folder', async () => {
		const fake = createFakeSp();
		await run(
			{
				operation: 'createMany',
				fieldsInputMode: 'json',
				fieldsJson: '{"Title":"x"}',
				bulkCreateOptions: { folderPath: 'Lists/Example/Sub' },
			},
			1,
			fake,
		);
		expect(kinds(fake.log, 'add')[0].args[1]).toBe('/sites/Example/Lists/Example/Sub');
		expect(kinds(fake.log, 'rootFolder')).toHaveLength(0);
	});

	it('reports rejected field values for each item with Continue on Item Error', async () => {
		const fake = createFakeSp({
			add: (formValues) =>
				formValues[0].FieldValue === 'bad'
					? [{ FieldName: 'Title', HasException: true, ErrorMessage: 'Invalid text' }]
					: [{ FieldName: 'Id', FieldValue: String(nextId++), HasException: false }],
		});
		const output = await run(
			{
				operation: 'createMany',
				fieldsInputMode: 'json',
				fieldsJson: (i: number) => JSON.stringify({ Title: i === 1 ? 'bad' : 'good' }),
				bulkCreateOptions: { continueOnItemError: true },
			},
			3,
			fake,
		);
		expect(json(output)).toEqual([
			{ id: 100, success: true },
			{ success: false, error: 'SharePoint rejected the field values. Title: Invalid text' },
			{ id: 101, success: true },
		]);
		expect(pairs(output)).toEqual([0, 1, 2]);
	});

	it('reports invalid JSON for an item without a request for it', async () => {
		const fake = createFakeSp();
		const output = await run(
			{
				operation: 'createMany',
				fieldsInputMode: 'json',
				fieldsJson: (i: number) => (i === 0 ? '{nope' : '{"Title":"x"}'),
				bulkCreateOptions: { continueOnItemError: true },
			},
			2,
			fake,
		);
		expect(json(output)[0]).toMatchObject({
			success: false,
			error: expect.stringContaining('not valid JSON'),
		});
		expect(json(output)[1]).toEqual({ id: 100, success: true });
		expect(kinds(fake.log, 'add')).toHaveLength(1);
	});
});

describe('Update Many', () => {
	const base = {
		operation: 'updateMany',
		bulkItemId: (i: number) => String(i + 1),
		fieldsInputMode: 'json',
		fieldsJson: (i: number) => ({ Title: `New ${i}` }),
	};

	it('updates all items in one batch', async () => {
		const fake = createFakeSp();
		const output = await run(base, 3, fake);
		expect(json(output)).toEqual([1, 2, 3].map((id) => ({ id, success: true })));
		const updates = kinds(fake.log, 'update');
		expect(updates.map((entry) => entry.batch)).toEqual([1, 1, 1]);
		expect(updates[2].args).toEqual([3, [{ FieldName: 'Title', FieldValue: 'New 2' }], false]);
	});

	it('stops with the error of the first failed item when no error handling is on', async () => {
		const fake = createFakeSp({
			update: (id) => {
				if (id === 2)
					throw Object.assign(new Error('Item does not exist'), {
						isHttpRequestError: true,
						status: 404,
					});
				return [];
			},
		});
		await expect(run(base, 3, fake)).rejects.toBeInstanceOf(NodeApiError);
		expect(kinds(fake.log, 'update')).toHaveLength(3);
	});

	it('outputs an error item for the failed input item when the node continues on fail', async () => {
		const fake = createFakeSp({
			update: (id) =>
				id === 2 ? [{ FieldName: 'Title', HasException: true, ErrorMessage: 'Bad' }] : [],
		});
		const output = await run(base, 3, fake, true);
		expect(json(output)).toEqual([
			{ id: 1, success: true },
			{ error: 'SharePoint rejected the field values. Title: Bad' },
			{ id: 3, success: true },
		]);
		expect(pairs(output)).toEqual([0, 1, 2]);
	});

	it('reports invalid IDs and empty fields before the batch', async () => {
		const fake = createFakeSp();
		const output = await run(
			{
				...base,
				bulkItemId: (i: number) => ['x', '2', '3'][i],
				fieldsJson: (i: number) => (i === 2 ? {} : { Title: 'x' }),
				bulkUpdateOptions: { continueOnItemError: true },
			},
			3,
			fake,
		);
		expect(json(output)).toEqual([
			{ success: false, error: '"x" is not a valid list item ID' },
			{ id: 2, success: true },
			{ id: 3, success: false, error: 'No fields to update were given' },
		]);
		expect(kinds(fake.log, 'update')).toHaveLength(1);
	});

	it('keeps Modified and Editor with System Update', async () => {
		const fake = createFakeSp();
		await run({ ...base, bulkUpdateOptions: { systemUpdate: true } }, 2, fake);
		expect(kinds(fake.log, 'read').map((entry) => entry.batch)).toEqual([1, 1]);
		expect(kinds(fake.log, 'localTime').map((entry) => entry.batch)).toEqual([2, 2]);
		const updates = kinds(fake.log, 'update');
		expect(updates.map((entry) => entry.batch)).toEqual([3, 3]);
		expect(updates[0].args).toEqual([
			1,
			[
				{ FieldName: 'Title', FieldValue: 'New 0' },
				{ FieldName: 'Modified', FieldValue: '2023-01-02 03:04:05' },
				{ FieldName: 'Editor', FieldValue: '[{"Key":"i:0#.f|membership|a@contoso.com"}]' },
			],
			true,
		]);
	});

	it('fails an item whose current values cannot be read for System Update', async () => {
		const fake = createFakeSp({
			read: (id) => {
				if (id === 1) throw new Error('Item does not exist');
				return { Modified: '2024-01-02T03:04:05Z' };
			},
		});
		const output = await run(
			{ ...base, bulkUpdateOptions: { systemUpdate: true, continueOnItemError: true } },
			2,
			fake,
		);
		expect(json(output)).toEqual([
			{ id: 1, success: false, error: 'Item does not exist' },
			{ id: 2, success: true },
		]);
		expect(kinds(fake.log, 'update').map((entry) => entry.args[0])).toEqual([2]);
	});

	it('groups items by list and fails only the items of a list that cannot be read', async () => {
		const fake = createFakeSp({
			fields: (list) => {
				if (list === 'Other List') throw new Error('List does not exist');
				return FIELDS;
			},
		});
		const output = await run(
			{ ...base, list: (i: number) => (i === 1 ? LIST_B : LIST_A) },
			3,
			fake,
			true,
		);
		expect(json(output)).toEqual([
			{ id: 1, success: true },
			{ error: 'List does not exist' },
			{ id: 3, success: true },
		]);
		expect(kinds(fake.log, 'update').map((entry) => entry.list)).toEqual([
			LIST_A.value,
			LIST_A.value,
		]);
	});

	it('calls the operation once for each site', async () => {
		vi.mocked(descriptions.getSiteUrl).mockImplementation(async (_ctx, itemIndex) =>
			itemIndex === 1 ? OTHER_SITE : SITE,
		);
		const fake = createFakeSp();
		const output = await run(base, 3, fake);
		expect(vi.mocked(transport.createSp).mock.calls.map((call) => call[1])).toEqual([
			SITE,
			OTHER_SITE,
		]);
		expect(pairs(output)).toEqual([0, 1, 2]);
		expect(json(output).map((entry) => entry.id)).toEqual([1, 2, 3]);
	});

	it('fails the item when the site cannot be resolved', async () => {
		vi.mocked(descriptions.getSiteUrl).mockImplementation(async (_ctx, itemIndex) => {
			if (itemIndex === 0) throw new NodeOperationError({ name: 'x' } as never, 'No site');
			return SITE;
		});
		const output = await run(base, 2, createFakeSp(), true);
		expect(json(output)).toEqual([{ error: 'No site' }, { id: 2, success: true }]);
	});
});

describe('Delete Many', () => {
	it('recycles items by default and deletes them permanently when set', async () => {
		const fake = createFakeSp();
		const output = await run(
			{ operation: 'deleteMany', bulkItemId: (i: number) => i + 10 },
			2,
			fake,
		);
		expect(json(output)).toEqual([
			{ id: 10, success: true, recycled: true },
			{ id: 11, success: true, recycled: true },
		]);
		expect(kinds(fake.log, 'recycle').map((entry) => entry.args[0])).toEqual([10, 11]);

		const permanent = createFakeSp();
		const deleted = await run(
			{ operation: 'deleteMany', bulkItemId: 5, bulkDeleteOptions: { permanent: true } },
			1,
			permanent,
		);
		expect(json(deleted)).toEqual([{ id: 5, success: true, recycled: false }]);
		expect(kinds(permanent.log, 'delete')).toHaveLength(1);
		expect(kinds(permanent.log, 'recycle')).toHaveLength(0);
	});

	it('reports a failed delete for its item only', async () => {
		const fake = createFakeSp({
			remove: (id) => {
				if (id === 2) throw new Error('Item does not exist');
			},
		});
		const output = await run(
			{
				operation: 'deleteMany',
				bulkItemId: (i: number) => i + 1,
				bulkDeleteOptions: { continueOnItemError: true },
			},
			3,
			fake,
		);
		expect(json(output)).toEqual([
			{ id: 1, success: true, recycled: true },
			{ id: 2, success: false, error: 'Item does not exist' },
			{ id: 3, success: true, recycled: true },
		]);
	});
});

describe('Create or Update (upsert)', () => {
	const base = {
		operation: 'upsert',
		keyField: 'Title',
		keyValue: (i: number) => `Key ${i}`,
		fieldsInputMode: 'json',
		fieldsJson: { Amount: 5 },
	};

	it('creates an item when no item has the key value, and sets the key field', async () => {
		const fake = createFakeSp();
		const output = await run({ ...base, keyValue: "O'Brien" }, 1, fake);
		expect(json(output)).toEqual([{ id: 100, success: true, action: 'created' }]);
		expect(kinds(fake.log, 'find')[0].args).toEqual(["Title eq 'O''Brien'", ['Id'], 1]);
		expect(kinds(fake.log, 'add')[0].args[0]).toEqual([
			{ FieldName: 'Amount', FieldValue: '5' },
			{ FieldName: 'Title', FieldValue: "O'Brien" },
		]);
	});

	it('updates the first match and batches the lookups and the writes', async () => {
		const fake = createFakeSp({ find: (filter) => (filter.includes('Key 1') ? [] : [{ Id: 7 }]) });
		const output = await run(base, 3, fake);
		expect(json(output)).toEqual([
			{ id: 7, success: true, action: 'updated' },
			{ id: 100, success: true, action: 'created' },
			{ id: 7, success: true, action: 'updated' },
		]);
		expect(kinds(fake.log, 'find').map((entry) => entry.batch)).toEqual([1, 1, 1]);
		expect(
			[...kinds(fake.log, 'update'), ...kinds(fake.log, 'add')].map((entry) => entry.batch),
		).toEqual([2, 2, 2]);
	});

	it('skips the write when the values are unchanged', async () => {
		const fake = createFakeSp({
			find: (filter) => [
				{ Id: filter.includes('Key 0') ? 1 : 2, Amount: filter.includes('Key 0') ? 5 : 4 },
			],
		});
		const output = await run({ ...base, upsertOptions: { skipUnchanged: true } }, 2, fake);
		expect(json(output)).toEqual([
			{ id: 1, success: true, action: 'unchanged' },
			{ id: 2, success: true, action: 'updated' },
		]);
		expect(kinds(fake.log, 'find')[0].args[1]).toEqual(['Id', 'Amount']);
		expect(kinds(fake.log, 'update').map((entry) => entry.args[0])).toEqual([2]);
	});

	it('always writes fields it cannot compare', async () => {
		const fake = createFakeSp({ find: () => [{ Id: 1 }] });
		const output = await run(
			{ ...base, fieldsJson: { Unknown: 'x' }, upsertOptions: { skipUnchanged: true } },
			1,
			fake,
		);
		expect(json(output)[0].action).toBe('updated');
	});

	it('updates all matches or fails, as set in If Several Match', async () => {
		const all = createFakeSp({ find: () => [{ Id: 3 }, { Id: 4 }] });
		const updated = await run({ ...base, upsertOptions: { ifSeveralMatch: 'all' } }, 1, all);
		expect(json(updated)).toEqual([{ id: 3, ids: [3, 4], success: true, action: 'updated' }]);
		expect(kinds(all.log, 'find')[0].args[2]).toBe(5000);
		expect(kinds(all.log, 'update').map((entry) => entry.args[0])).toEqual([3, 4]);

		const fail = createFakeSp({ find: () => [{ Id: 3 }, { Id: 4 }] });
		const failed = await run(
			{ ...base, upsertOptions: { ifSeveralMatch: 'fail', continueOnItemError: true } },
			1,
			fail,
		);
		expect(json(failed)).toEqual([
			{ success: false, error: 'Several items have the value "Key 0" in the field "Title"' },
		]);
		expect(kinds(fail.log, 'find')[0].args[2]).toBe(2);
		expect(kinds(fail.log, 'update')).toHaveLength(0);
	});

	it('finds items by number and lookup ID keys', async () => {
		const fake = createFakeSp({ find: () => [{ Id: 1 }] });
		await run({ ...base, keyField: 'Amount', keyValue: '12' }, 1, fake);
		await run({ ...base, keyField: 'Department', keyValue: 4 }, 1, fake);
		expect(kinds(fake.log, 'find').map((entry) => entry.args[0])).toEqual([
			'Amount eq 12',
			'DepartmentId eq 4',
		]);
	});

	it('rejects unknown and unsupported key fields', async () => {
		const output = await run(
			{
				...base,
				keyField: (i: number) => (i === 0 ? 'Missing' : 'Notes'),
				upsertOptions: { continueOnItemError: true },
			},
			2,
		);
		expect(json(output)[0].error).toBe('The list has no field with the internal name "Missing"');
		expect(json(output)[1].error).toContain('cannot be used as a key');
	});

	it('uses the Modified value of the match for System Update without another read', async () => {
		const fake = createFakeSp({
			find: () => [
				{
					Id: 9,
					Modified: '2024-05-06T07:08:09Z',
					Editor: { Name: 'i:0#.f|membership|b@contoso.com' },
				},
			],
		});
		await run({ ...base, upsertOptions: { systemUpdate: true } }, 1, fake);
		expect(kinds(fake.log, 'find')[0].args[1]).toEqual(['Id', 'Modified', 'Editor/Name']);
		expect(kinds(fake.log, 'read')).toHaveLength(0);
		expect(kinds(fake.log, 'localTime')[0].args).toEqual(['2024-05-06T07:08:09Z']);
		expect(kinds(fake.log, 'update')[0].args).toEqual([
			9,
			[
				{ FieldName: 'Amount', FieldValue: '5' },
				{ FieldName: 'Modified', FieldValue: '2023-05-06 07:08:09' },
				{ FieldName: 'Editor', FieldValue: '[{"Key":"i:0#.f|membership|b@contoso.com"}]' },
			],
			true,
		]);
	});

	it('finds the item with a CAML query in the folder of Find in Folder and creates new items there', async () => {
		const fake = createFakeSp({
			find: (viewXml) => (viewXml.includes('Key 0') ? [{ Id: 4, Amount: 5 }] : []),
		});
		const output = await run(
			{
				...base,
				upsertOptions: {
					findInFolder: 'Lists/Example/Sub',
					skipUnchanged: true,
					systemUpdate: true,
				},
			},
			2,
			fake,
		);
		expect(json(output)).toEqual([
			{ id: 4, success: true, action: 'unchanged' },
			{ id: 100, success: true, action: 'created' },
		]);
		const camls = kinds(fake.log, 'caml');
		expect(camls.map((entry) => entry.batch)).toEqual([1, 1]);
		expect(camls[0].args).toEqual([
			{
				ViewXml:
					'<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">Key 0</Value></Eq></Where><OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy></Query><RowLimit>1</RowLimit></View>',
				FolderServerRelativeUrl: '/sites/Example/Lists/Example/Sub',
			},
			['Editor'],
		]);
		expect(kinds(fake.log, 'find')).toHaveLength(0);
		expect(kinds(fake.log, 'add')[0].args[1]).toBe('/sites/Example/Lists/Example/Sub');
	});

	it('reports a failed write with the ID of the matched item', async () => {
		const fake = createFakeSp({
			find: () => [{ Id: 9 }],
			update: () => [{ FieldName: 'Amount', HasException: true, ErrorMessage: 'Not a number' }],
		});
		const output = await run({ ...base, upsertOptions: { continueOnItemError: true } }, 1, fake);
		expect(json(output)).toEqual([
			{
				id: 9,
				success: false,
				error: 'SharePoint rejected the field values. Amount: Not a number',
			},
		]);
	});
});
