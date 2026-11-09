import { describe, expect, it, vi } from 'vitest';
import type { SPFI } from '@pnp/sp/index.js';

import {
	copyChildrenIfMissing,
	matchesAny,
	parseNamePatterns,
	planCopy,
	wildcardToRegExp,
	type Child,
	type CopyChildrenGateway,
} from '../../nodes/SharePointPnp/resources/folder/copyChildren';
import {
	ensurePath,
	fieldsToRead,
	planSegmentValues,
	segmentsFromJson,
	segmentsFromRows,
	type EnsurePathGateway,
	type FolderState,
	type SegmentSpec,
} from '../../nodes/SharePointPnp/resources/folder/ensurePath';
import {
	describeFieldErrors,
	fromKeyValueRows,
	parseFieldsObject,
	readCurrentValue,
	selectNameFor,
	toValueList,
	valuesEqual,
} from '../../nodes/SharePointPnp/resources/folder/fieldValues';
import {
	applyPlans,
	collectTargets,
	createItemWriter,
	planItemValues,
	prepareFields,
	readAllPages,
	rowToTarget,
	setMetadataRecursive,
	toFormDate,
	type ItemPlan,
	type ItemWriter,
	type PageReader,
	type TargetItem,
} from '../../nodes/SharePointPnp/resources/folder/setMetadata';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';

const LIB = '/sites/Example/Shared Documents';

describe('field value helpers', () => {
	it('reads JSON objects and rejects other values', () => {
		expect(parseFieldsObject('', 'X')).toEqual({});
		expect(parseFieldsObject('{"Title":"A"}', 'X')).toEqual({ Title: 'A' });
		expect(parseFieldsObject({ Title: 'A' }, 'X')).toEqual({ Title: 'A' });
		expect(() => parseFieldsObject('{bad', 'X')).toThrow('"X" is not valid JSON');
		expect(() => parseFieldsObject('[1]', 'X')).toThrow('must be a JSON object');
	});

	it('reads key-value rows and skips rows without a name', () => {
		expect(
			fromKeyValueRows([
				{ name: ' Title ', value: 'A' },
				{ name: '', value: 'B' },
			]),
		).toEqual({
			Title: 'A',
		});
		expect(fromKeyValueRows(undefined)).toEqual({});
	});

	it('splits value lists', () => {
		expect(toValueList('a, b\nc')).toEqual(['a', 'b', 'c']);
		expect(toValueList('["a","b"]')).toEqual(['a', 'b']);
		expect(toValueList(['a', 2])).toEqual(['a', '2']);
		expect(toValueList(null)).toEqual([]);
	});

	it('compares values by field type', () => {
		expect(valuesEqual('Text', 'Team A', 'Team A')).toBe(true);
		expect(valuesEqual('Text', 'Team A', 'Team B')).toBe(false);
		expect(valuesEqual('Text', null, '')).toBe(true);
		expect(valuesEqual('Number', 5, '5')).toBe(true);
		expect(valuesEqual('Boolean', true, 'yes')).toBe(true);
		expect(valuesEqual('Boolean', false, 'true')).toBe(false);
		expect(valuesEqual('MultiChoice', ['b', 'a'], 'a, b')).toBe(true);
		expect(valuesEqual('MultiChoice', ['a'], ['a', 'b'])).toBe(false);
		expect(valuesEqual('Lookup', 3, { Id: 3 })).toBe(true);
		expect(valuesEqual('DateTime', '2026-01-01T00:00:00Z', '2026-01-01T00:00:00.000Z')).toBe(true);
	});

	it('reads lookup and person values by their ID names', () => {
		expect(selectNameFor('Department', 'Lookup')).toBe('DepartmentId');
		expect(selectNameFor('Title', 'Text')).toBe('Title');
		expect(readCurrentValue({ TagsId: { results: [1, 2] } }, 'Tags', 'LookupMulti')).toEqual([
			1, 2,
		]);
	});

	it('describes rejected field values', () => {
		expect(describeFieldErrors([{ FieldName: 'Title', HasException: false }])).toBeUndefined();
		expect(
			describeFieldErrors([{ FieldName: 'Status', HasException: true, ErrorMessage: 'Bad' }]),
		).toBe('Status: Bad');
	});
});

describe('Ensure Path segments', () => {
	it('reads segment rows', () => {
		expect(
			segmentsFromRows([
				{
					name: ' Team A ',
					fieldsWhenCreated: '{"Title":"A"}',
					fieldsAlways: '',
					appendField: 'Tags',
					appendValues: 'Active, New',
				},
			]),
		).toEqual([
			{
				name: 'Team A',
				fieldsWhenCreated: { Title: 'A' },
				fieldsAlways: {},
				appendToMultiChoice: { field: 'Tags', values: ['Active', 'New'] },
			},
		]);
	});

	it('reads JSON segments as names or objects', () => {
		expect(
			segmentsFromJson(
				JSON.stringify([
					'Team A',
					{
						name: 'Q1',
						fieldsAlways: { Title: 'P' },
						appendToMultiChoice: { field: 'Tags', values: ['x'] },
					},
				]),
			),
		).toEqual([
			{ name: 'Team A', fieldsWhenCreated: {}, fieldsAlways: {} },
			{
				name: 'Q1',
				fieldsWhenCreated: {},
				fieldsAlways: { Title: 'P' },
				appendToMultiChoice: { field: 'Tags', values: ['x'] },
			},
		]);
	});

	it('rejects bad segments', () => {
		expect(() => segmentsFromRows([{ name: '  ' }])).toThrow('Segment 1 has no name');
		expect(() => segmentsFromRows([{ name: 'a/b' }])).toThrow('must not have "/"');
		expect(() => segmentsFromJson('{}')).toThrow('must be an array');
		expect(() => segmentsFromJson('[1]')).toThrow('must be a name or an object');
	});

	const types = new Map([
		['Title', 'Text'],
		['Status', 'Choice'],
		['Tags', 'MultiChoice'],
	]);
	const segment: SegmentSpec = {
		name: 'Team A',
		fieldsWhenCreated: { Title: 'Team A', Status: 'New' },
		fieldsAlways: { Status: 'Active' },
		appendToMultiChoice: { field: 'Tags', values: ['Final'] },
	};

	it('reads only the always and append fields of an existing folder', () => {
		expect(fieldsToRead(segment)).toEqual(['Status', 'Tags']);
	});

	it('sets created and always values on a new folder', () => {
		expect(planSegmentValues(segment, true, {}, types)).toEqual({
			Title: 'Team A',
			Status: 'Active',
			Tags: ['Final'],
		});
	});

	it('sets only changed always values on an existing folder', () => {
		expect(planSegmentValues(segment, false, { Status: 'Active', Tags: ['Final'] }, types)).toEqual(
			{},
		);
		expect(planSegmentValues(segment, false, { Status: 'New', Tags: ['Final'] }, types)).toEqual({
			Status: 'Active',
		});
	});

	it('merges append values with the current choices', () => {
		expect(planSegmentValues(segment, false, { Status: 'Active', Tags: ['Draft'] }, types)).toEqual(
			{
				Tags: ['Draft', 'Final'],
			},
		);
	});
});

function fakeEnsureGateway(existing: Record<string, Record<string, unknown>>) {
	const folders = new Map(
		Object.entries(existing).map(([path, item]) => [path.toLowerCase(), { path, item }]),
	);
	let nextId = 100;
	const ids = new Map<string, number>();
	const gateway = {
		getBaseFolder: vi.fn(async (path: string) => (path === LIB ? LIB : undefined)),
		getFolder: vi.fn(async (path: string): Promise<FolderState | undefined> => {
			const entry = folders.get(path.toLowerCase());
			if (!entry) return undefined;
			if (!ids.has(entry.path)) ids.set(entry.path, nextId++);
			return { serverRelativeUrl: entry.path, listItemId: ids.get(entry.path)!, item: entry.item };
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
			for (const entry of folders.values()) {
				if (ids.get(entry.path) === id) Object.assign(entry.item, values);
			}
		}),
	} satisfies EnsurePathGateway;
	return gateway;
}

describe('Ensure Path walk', () => {
	const segments: SegmentSpec[] = [
		{ name: 'Team A', fieldsWhenCreated: { Title: 'Team A' }, fieldsAlways: {} },
		{ name: '2024', fieldsWhenCreated: {}, fieldsAlways: { Title: '2024' } },
		{
			name: 'Q1',
			fieldsWhenCreated: {},
			fieldsAlways: {},
			appendToMultiChoice: { field: 'Tags', values: ['Active'] },
		},
	];

	it('creates all missing folders and sets their fields', async () => {
		const gateway = fakeEnsureGateway({});
		const result = await ensurePath(LIB, segments, gateway);
		expect(gateway.createFolder.mock.calls.map(([path]) => path)).toEqual([
			`${LIB}/Team A`,
			`${LIB}/Team A/2024`,
			`${LIB}/Team A/2024/Q1`,
		]);
		// After a folder is created, its children are not looked up before they are created.
		expect(gateway.getFolder).toHaveBeenCalledTimes(4);
		expect(result).toMatchObject({
			folderPath: `${LIB}/Team A/2024/Q1`,
			created: true,
			changed: true,
		});
		expect(result.segments).toEqual([
			expect.objectContaining({ created: true, fieldsUpdated: ['Title'] }),
			expect.objectContaining({ created: true, fieldsUpdated: ['Title'] }),
			expect.objectContaining({ created: true, fieldsUpdated: ['Tags'] }),
		]);
	});

	it('reports no change on a second run', async () => {
		const gateway = fakeEnsureGateway({});
		await ensurePath(LIB, segments, gateway);
		gateway.updateItem.mockClear();
		gateway.createFolder.mockClear();
		const result = await ensurePath(LIB, segments, gateway);
		expect(gateway.createFolder).not.toHaveBeenCalled();
		expect(gateway.updateItem).not.toHaveBeenCalled();
		expect(result).toMatchObject({ created: false, changed: false });
		expect(
			(result.segments as Array<{ fieldsUpdated: string[] }>).every(
				(entry) => !entry.fieldsUpdated.length,
			),
		).toBe(true);
	});

	it('creates only the missing part and keeps the casing of existing folders', async () => {
		const gateway = fakeEnsureGateway({ [`${LIB}/team a`]: {} });
		const result = await ensurePath(LIB, segments.slice(0, 2), gateway);
		expect(gateway.createFolder.mock.calls.map(([path]) => path)).toEqual([`${LIB}/team a/2024`]);
		expect(result.segments).toEqual([
			expect.objectContaining({
				created: false,
				fieldsUpdated: [],
				serverRelativeUrl: `${LIB}/team a`,
			}),
			expect.objectContaining({ created: true, fieldsUpdated: ['Title'] }),
		]);
	});

	it('needs no field lookup when no segment has fields', async () => {
		const gateway = fakeEnsureGateway({});
		await ensurePath(LIB, [{ name: 'A', fieldsWhenCreated: {}, fieldsAlways: {} }], gateway);
		expect(gateway.getFieldTypes).not.toHaveBeenCalled();
	});

	it('fails for unknown fields, a missing base folder and a bad append field', async () => {
		await expect(
			ensurePath(
				LIB,
				[{ name: 'A', fieldsWhenCreated: { Nope: 1 }, fieldsAlways: {} }],
				fakeEnsureGateway({}),
			),
		).rejects.toThrow('do not exist in the library: Nope');
		await expect(
			ensurePath(`${LIB}/Missing`, segments.slice(0, 1), fakeEnsureGateway({})),
		).rejects.toThrow('does not exist');
		await expect(
			ensurePath(
				LIB,
				[
					{
						name: 'A',
						fieldsWhenCreated: {},
						fieldsAlways: {},
						appendToMultiChoice: { field: 'Title', values: ['x'] },
					},
				],
				fakeEnsureGateway({}),
			),
		).rejects.toThrow('is not a multi-choice field');
		await expect(ensurePath(LIB, [], fakeEnsureGateway({}))).rejects.toThrow(
			'at least one segment',
		);
	});
});

describe('name patterns', () => {
	it('matches wildcards without case', () => {
		expect(wildcardToRegExp('01 *').test('01 Admin')).toBe(true);
		expect(wildcardToRegExp('01 *').test('02 Admin')).toBe(false);
		expect(wildcardToRegExp('*.TXT').test('notes.txt')).toBe(true);
		expect(wildcardToRegExp('a.txt').test('abtxt')).toBe(false);
		expect(wildcardToRegExp('file?.pdf').test('file1.pdf')).toBe(true);
		expect(wildcardToRegExp('(x)+').test('(x)+')).toBe(true);
	});

	it('reads comma and line separated patterns', () => {
		const patterns = parseNamePatterns('Images,\n01 *\r\n ');
		expect(patterns).toHaveLength(2);
		expect(matchesAny('images', patterns)).toBe(true);
		expect(matchesAny('01 Admin', patterns)).toBe(true);
		expect(matchesAny('Other', patterns)).toBe(false);
	});
});

const child = (name: string, type: 'file' | 'folder' = 'folder'): Child => ({
	name,
	type,
	serverRelativeUrl: `${LIB}/Template/${name}`,
});

describe('copy plan', () => {
	const children = [
		child('01 Admin'),
		child('02 Images'),
		child('Archive'),
		child('readme.txt', 'file'),
	];

	it('copies only missing folders by default', () => {
		const plan = planCopy(children, new Set(['01 admin']), {});
		expect(plan.toCopy.map((entry) => entry.name)).toEqual(['02 Images', 'Archive']);
		expect(plan.skipped).toEqual([
			expect.objectContaining({ name: '01 Admin', reason: 'exists' }),
			expect.objectContaining({ name: 'readme.txt', reason: 'type' }),
		]);
	});

	it('applies include and exclude names', () => {
		const plan = planCopy(children, new Set(), {
			include: 'both',
			includeNames: '0*, *.txt',
			excludeNames: '02*',
		});
		expect(plan.toCopy.map((entry) => entry.name)).toEqual(['01 Admin', 'readme.txt']);
		expect(plan.skipped.map((entry) => [entry.name, entry.reason])).toEqual([
			['02 Images', 'excluded'],
			['Archive', 'notIncluded'],
		]);
	});

	it('handles names that exist by the conflict behavior', () => {
		const existing = new Set(['archive']);
		expect(planCopy(children, existing, { nameConflictBehavior: 'fail' }).conflicts).toEqual([
			child('Archive'),
		]);
		expect(
			planCopy(children, existing, { nameConflictBehavior: 'replace' }).toCopy.find(
				(entry) => entry.name === 'Archive',
			),
		).toMatchObject({ exists: true });
		expect(planCopy(children, existing, { include: 'files' }).toCopy).toEqual([
			{ ...child('readme.txt', 'file'), exists: false },
		]);
	});
});

function fakeCopyGateway(tree: Record<string, Child[]>, jobLogs: Record<string, unknown[]> = {}) {
	let job = 0;
	const gateway = {
		listChildren: vi.fn(async (path: string) => tree[path]),
		copyDirect: vi.fn(async () => {}),
		startCopyJob: vi.fn(async (entry: Child) => ({
			jobId: `job-${++job}-${entry.name}`,
			jobQueueUri: 'https://contoso.example/queue',
			encryptionKey: 'key',
		})),
		readProgress: vi.fn(async (info: { jobId: string }) => {
			const name = info.jobId.split('-').slice(2).join('-');
			return {
				JobState: 0,
				Logs: jobLogs[name] ?? [
					JSON.stringify({
						Event: 'JobEnd',
						MigrationDirection: 'Import',
						FilesCreated: '0',
						TotalErrors: '0',
					}),
				],
			};
		}),
		toPath: (url: string) => `/sites/Example/${url}`,
		pause: async () => {},
	} satisfies CopyChildrenGateway;
	return gateway;
}

describe('Copy Children If Missing', () => {
	const source = `${LIB}/Template`;
	const destination = `${LIB}/Q1`;

	it('skips or fails when the destination is missing', async () => {
		const gateway = fakeCopyGateway({ [source]: [child('A')] });
		await expect(
			copyChildrenIfMissing(source, destination, { onlyIfDestinationExists: true }, gateway),
		).resolves.toMatchObject({ changed: false, destinationExists: false, copied: [] });
		await expect(copyChildrenIfMissing(source, destination, {}, gateway)).rejects.toThrow(
			'destination folder',
		);
	});

	it('starts one copy job for each missing child and waits for them', async () => {
		const gateway = fakeCopyGateway({
			[source]: [child('A'), child('B')],
			[destination]: [child('b')],
		});
		const result = await copyChildrenIfMissing(source, destination, {}, gateway);
		expect(gateway.startCopyJob).toHaveBeenCalledTimes(1);
		expect(result).toMatchObject({
			changed: true,
			method: 'copyJob',
			copied: [{ name: 'A', status: 'completed', destinationPath: `${destination}/A` }],
			skipped: [{ name: 'B', reason: 'exists' }],
		});
	});

	it('copies nothing when all children exist', async () => {
		const gateway = fakeCopyGateway({ [source]: [child('A')], [destination]: [child('A')] });
		const result = await copyChildrenIfMissing(source, destination, {}, gateway);
		expect(gateway.startCopyJob).not.toHaveBeenCalled();
		expect(result).toMatchObject({ changed: false, copied: [] });
	});

	it('fails on job errors unless the option says to continue', async () => {
		const logs = {
			A: [
				JSON.stringify({ Event: 'JobError', Message: 'Access denied' }),
				JSON.stringify({ Event: 'JobEnd', MigrationDirection: 'Import', TotalErrors: '1' }),
			],
		};
		const gateway = fakeCopyGateway({ [source]: [child('A')], [destination]: [] }, logs);
		await expect(copyChildrenIfMissing(source, destination, {}, gateway)).rejects.toThrow(
			'Access denied',
		);
		const result = await copyChildrenIfMissing(
			source,
			destination,
			{ continueOnJobErrors: true },
			gateway,
		);
		expect(result.copied).toEqual([expect.objectContaining({ name: 'A', status: 'failed' })]);
	});

	it('copies directly with the given options', async () => {
		const gateway = fakeCopyGateway({
			[source]: [child('A'), child('f.txt', 'file')],
			[destination]: [],
		});
		const result = await copyChildrenIfMissing(
			source,
			destination,
			{ copyMethod: 'direct', include: 'both', resetAuthorAndCreated: true },
			gateway,
		);
		expect(gateway.copyDirect.mock.calls.map((call) => call[1])).toEqual([
			`${destination}/A`,
			`${destination}/f.txt`,
		]);
		expect(result.changed).toBe(true);
	});

	it('does not replace a folder with the Direct method', async () => {
		const gateway = fakeCopyGateway({ [source]: [child('A')], [destination]: [child('A')] });
		await expect(
			copyChildrenIfMissing(
				source,
				destination,
				{ copyMethod: 'direct', nameConflictBehavior: 'replace' },
				gateway,
			),
		).rejects.toThrow('cannot replace the folder "A"');
	});

	it('fails before copying when a name exists and the behavior is Fail', async () => {
		const gateway = fakeCopyGateway({
			[source]: [child('A'), child('B')],
			[destination]: [child('A')],
		});
		await expect(
			copyChildrenIfMissing(source, destination, { nameConflictBehavior: 'fail' }, gateway),
		).rejects.toThrow('exist in the destination: A');
		expect(gateway.startCopyJob).not.toHaveBeenCalled();
	});
});

const item = (
	id: number,
	type: 'file' | 'folder',
	path: string,
	values: Record<string, unknown> = {},
): TargetItem => ({
	id,
	type,
	serverRelativeUrl: path,
	values,
});

describe('Set Metadata Recursive enumeration', () => {
	it('reads all pages of a folder', async () => {
		const reader = vi.fn<PageReader>(async (_path, next) =>
			next === 'page2'
				? { rows: [item(2, 'file', '/b')] }
				: { rows: [item(1, 'file', '/a')], next: 'page2' },
		);
		expect((await readAllPages(reader, '/root')).map((row) => row.id)).toEqual([1, 2]);
		expect(reader).toHaveBeenCalledTimes(2);
	});

	const tree: Record<string, TargetItem[]> = {
		'/root': [item(1, 'folder', '/root/a'), item(2, 'file', '/root/x.txt')],
		'/root/a': [item(3, 'folder', '/root/a/b'), item(4, 'file', '/root/a/y.txt')],
		'/root/a/b': [item(5, 'file', '/root/a/b/z.txt')],
	};
	const reader: PageReader = async (path) => ({ rows: tree[path] ?? [] });

	it('walks the tree folder by folder', async () => {
		expect((await collectTargets('/root', reader, {})).map((row) => row.id)).toEqual([
			1, 2, 3, 4, 5,
		]);
		expect(
			(await collectTargets('/root', reader, { include: 'files' })).map((row) => row.id),
		).toEqual([2, 4, 5]);
		expect(
			(await collectTargets('/root', reader, { include: 'folders' })).map((row) => row.id),
		).toEqual([1, 3]);
	});

	it('stops at the max depth', async () => {
		expect((await collectTargets('/root', reader, { maxDepth: 1 })).map((row) => row.id)).toEqual([
			1, 2,
		]);
		expect((await collectTargets('/root', reader, { maxDepth: 2 })).map((row) => row.id)).toEqual([
			1, 2, 3, 4,
		]);
	});

	it('maps view rows', () => {
		expect(
			rowToTarget(
				{
					ID: '7',
					FSObjType: '1',
					FileRef: '/root/a',
					ContentTypeId: '0x0120',
					Modified: '1/1/2026 10:00 AM',
					'Modified.': '2026-01-01T10:00:00Z',
					Editor: [{ id: '12' }],
					Title: 'A',
				},
				['Title'],
			),
		).toEqual({
			id: 7,
			type: 'folder',
			serverRelativeUrl: '/root/a',
			contentTypeId: '0x0120',
			modifiedUtc: '2026-01-01T10:00:00Z',
			editorId: 12,
			values: { Title: 'A' },
		});
		expect(toFormDate('2026-01-01T05:00:00.000')).toBe('2026-01-01 05:00:00');
	});
});

describe('Set Metadata Recursive values', () => {
	const types = new Map([
		['Title', 'Text'],
		['Count', 'Number'],
	]);

	it('leaves out empty values when asked', () => {
		expect(prepareFields({ Title: '', Count: 0, Tags: [] }, true)).toEqual({ Count: 0 });
		expect(prepareFields({ Title: '' }, false)).toEqual({ Title: '' });
	});

	it('keeps only the fields of the content type', () => {
		const contentTypeFields = new Map([['0x0101', new Set(['Title'])]]);
		const target = { ...item(1, 'file', '/a'), contentTypeId: '0x0101' };
		expect(planItemValues(target, { Title: 'A', Count: 1 }, types, { contentTypeFields })).toEqual({
			Title: 'A',
		});
		const unknown = { ...item(1, 'file', '/a'), contentTypeId: '0x0199' };
		expect(planItemValues(unknown, { Title: 'A', Count: 1 }, types, { contentTypeFields })).toEqual(
			{
				Title: 'A',
				Count: 1,
			},
		);
	});

	it('skips equal text values but writes other types', () => {
		const target = item(1, 'file', '/a', { Title: 'A', Count: '1' });
		expect(
			planItemValues(target, { Title: 'A', Count: 1 }, types, { skipUnchanged: true }),
		).toEqual({ Count: 1 });
		expect(planItemValues(target, { Title: 'A' }, types, { skipUnchanged: false })).toEqual({
			Title: 'A',
		});
	});
});

function plans(count: number): ItemPlan[] {
	return Array.from({ length: count }, (_, index) => ({
		item: item(index + 1, 'file', `/f${index + 1}`),
		values: { Title: 'A' },
	}));
}

describe('Set Metadata Recursive batches', () => {
	const writer = (failIds: number[] = []) => {
		const chunks: number[][] = [];
		const value: ItemWriter = {
			writeChunk: async (chunk) => {
				chunks.push(chunk.map((plan) => plan.item.id));
				return chunk.map((plan, index) =>
					failIds.includes(plan.item.id)
						? { index, ok: false, error: new Error(`fail ${plan.item.id}`) }
						: { index, ok: true },
				);
			},
		};
		return { value, chunks };
	};

	it('splits the work into batches of the given size, at most 100', async () => {
		const small = writer();
		await applyPlans(plans(5), small.value, { batchSize: 2 });
		expect(small.chunks).toEqual([[1, 2], [3, 4], [5]]);
		const large = writer();
		const result = await applyPlans(plans(250), large.value, { batchSize: 500 });
		expect(large.chunks.map((chunk) => chunk.length)).toEqual([100, 100, 50]);
		expect(result).toEqual({ updated: 250, failed: 0, errors: [] });
	});

	it('collects partial failures when asked to continue', async () => {
		const result = await applyPlans(plans(5), writer([2, 5]).value, {
			batchSize: 2,
			continueOnItemError: true,
		});
		expect(result.updated).toBe(3);
		expect(result.failed).toBe(2);
		expect(result.errors).toEqual([
			{ listItemId: 2, serverRelativeUrl: '/f2', message: 'fail 2' },
			{ listItemId: 5, serverRelativeUrl: '/f5', message: 'fail 5' },
		]);
	});

	it('stops after the batch with the first failure', async () => {
		const failing = writer([2]);
		await expect(applyPlans(plans(5), failing.value, { batchSize: 2 })).rejects.toThrow(
			'The update of "/f2" failed: fail 2. 1 item(s) were updated',
		);
		expect(failing.chunks).toEqual([[1, 2]]);
	});

	it('runs end to end with counts', async () => {
		const tree: Record<string, TargetItem[]> = {
			'/root': [item(1, 'folder', '/root/a', { Title: 'Done' }), item(2, 'file', '/root/x.txt')],
			'/root/a': [item(3, 'file', '/root/a/y.txt')],
		};
		const write = writer([3]);
		const result = await setMetadataRecursive(
			'/root',
			{ Title: 'Done' },
			{ continueOnItemError: true, includeFolderItself: true },
			{
				listFields: { listId: 'list', infos: [], types: new Map([['Title', 'Text']]) },
				readPage: async (path) => ({ rows: tree[path] ?? [] }),
				readSelf: async () => item(9, 'folder', '/root'),
				readContentTypeFields: async () => new Map(),
				writer: write.value,
			},
		);
		expect(result).toMatchObject({ matched: 4, updated: 2, skipped: 1, failed: 1 });
		expect(write.chunks).toEqual([[9, 2, 3]]);
	});

	it('fails for unknown fields and empty input', async () => {
		const source = {
			listFields: { listId: 'list', infos: [], types: new Map([['Title', 'Text']]) },
			readPage: async () => ({ rows: [] }),
			readSelf: async () => undefined,
			readContentTypeFields: async () => new Map(),
			writer: writer().value,
		};
		await expect(setMetadataRecursive('/root', { Nope: 1 }, {}, source)).rejects.toThrow('Nope');
		await expect(
			setMetadataRecursive('/root', { Title: '' }, { ignoreEmptyValues: true }, source),
		).rejects.toThrow('No fields to set');
	});
});

describe('Set Metadata Recursive writer', () => {
	function fakeSp(rejectId?: number) {
		const updates: Array<{ id: number; formValues: unknown; systemUpdate: boolean }> = [];
		let batches = 0;
		let executed = 0;
		const web = {
			lists: {
				getById: () => ({
					items: {
						getById: (id: number) => ({
							validateUpdateListItem: (formValues: unknown, systemUpdate: boolean) => {
								updates.push({ id, formValues, systemUpdate });
								return Promise.resolve(
									id === rejectId
										? [{ FieldName: 'Title', HasException: true, ErrorMessage: 'Locked' }]
										: [{ FieldName: 'Title', HasException: false }],
								);
							},
						}),
					},
				}),
			},
			siteUsers: {
				getById: (id: number) => ({
					select: () => () =>
						Promise.resolve({ LoginName: `i:0#.f|membership|user${id}@contoso.com` }),
				}),
			},
			regionalSettings: {
				timeZone: { utcToLocalTime: () => Promise.resolve('2026-01-01T05:00:00.000') },
			},
		};
		const sp = {
			web,
			batched: () => {
				batches++;
				return [
					{ web },
					async () => {
						executed++;
					},
				];
			},
		} as unknown as SPFI;
		return { sp, updates, counts: () => ({ batches, executed }) };
	}

	const infos = [{ InternalName: 'Title', TypeAsString: 'Text' }];

	it('sends each chunk as one batch and reports rejected field values', async () => {
		const fake = fakeSp(2);
		const results = await createItemWriter(fake.sp, 'list', infos, false).writeChunk(plans(3));
		expect(fake.counts()).toEqual({ batches: 1, executed: 1 });
		expect(results.map((result) => result.ok)).toEqual([true, false, true]);
		expect((results[1].error as Error).message).toBe('Title: Locked');
		expect(fake.updates[0]).toEqual({
			id: 1,
			formValues: [{ FieldName: 'Title', FieldValue: 'A' }],
			systemUpdate: false,
		});
	});

	it('keeps Modified and Editor for a system update', async () => {
		const fake = fakeSp();
		const chunk = plans(1);
		chunk[0].item.modifiedUtc = '2026-01-01T10:00:00Z';
		chunk[0].item.editorId = 12;
		await createItemWriter(fake.sp, 'list', infos, true).writeChunk(chunk);
		expect(fake.counts().batches).toBe(3);
		expect(fake.updates[0]).toEqual({
			id: 1,
			formValues: [
				{ FieldName: 'Title', FieldValue: 'A' },
				{ FieldName: 'Modified', FieldValue: '2026-01-01 05:00:00' },
				{
					FieldName: 'Editor',
					FieldValue: JSON.stringify([{ Key: 'i:0#.f|membership|user12@contoso.com' }]),
				},
			],
			systemUpdate: true,
		});
	});
});

describe('Folder description', () => {
	it('lists the provisioning operations in alphabetical order', () => {
		const operation = folderResource.description[0];
		const names = (operation.options as Array<{ name: string }>).map((option) => option.name);
		expect(names).toEqual([...names].sort());
		expect(names).toEqual(
			expect.arrayContaining(['Ensure Path', 'Copy Children If Missing', 'Set Metadata Recursive']),
		);
		expect(Object.keys(folderResource.operations)).toEqual(
			expect.arrayContaining(['ensurePath', 'copyChildrenIfMissing', 'setMetadataRecursive']),
		);
	});
});
