import { describe, expect, it, vi } from 'vitest';

import {
	RecycleBinLookup,
	type RecycleBinEntry,
	type RecycleBinSource,
	buildChangeQuery,
	CHANGE_EVENTS,
	createChangeTokenFromDate,
	formatChange,
	getChangeTypeName,
	getEventFlags,
	getMatchingEvents,
	getObjectType,
	isUnderFolder,
	runChangeQuery,
	type ChangeQuery,
	type ChangeTarget,
} from '../../nodes/SharePointPnp/shared/changes';
import type { SPFI } from '@pnp/sp/index.js';

const LIST_ID = '11111111-2222-3333-4444-555555555555';

describe('createChangeTokenFromDate', () => {
	const date = new Date('2024-01-01T00:00:00.000Z');
	// .NET ticks for 2024-01-01T00:00:00Z
	const ticks = '638396640000000000';

	it.each([
		['siteCollection', 1],
		['web', 2],
		['list', 3],
	] as const)('uses the scope code of %s', (scope, code) => {
		expect(createChangeTokenFromDate(scope, LIST_ID, date)).toBe(
			`1;${code};${LIST_ID};${ticks};-1`,
		);
	});

	it('normalizes the ID and keeps millisecond precision', () => {
		const token = createChangeTokenFromDate(
			'list',
			`{${LIST_ID.toUpperCase()}}`,
			new Date('2024-01-01T00:00:00.123Z'),
		);
		expect(token).toBe(`1;3;${LIST_ID};638396640001230000;-1`);
	});

	it('rejects a date that is not valid', () => {
		expect(() => createChangeTokenFromDate('web', LIST_ID, new Date('nope'))).toThrow();
	});
});

describe('getChangeTypeName', () => {
	it.each([
		[1, 'Add'],
		[2, 'Update'],
		[3, 'DeleteObject'],
		[4, 'Rename'],
		[5, 'MoveAway'],
		[6, 'MoveInto'],
		[7, 'Restore'],
		[11, 'AssignmentAdd'],
		[12, 'AssignmentDelete'],
		[15, 'SystemUpdate'],
		[17, 'ScopeAdd'],
		[22, 'Activity'],
	])('maps %i to %s', (value, name) => {
		expect(getChangeTypeName(value)).toBe(name);
	});

	it('marks unknown values', () => {
		expect(getChangeTypeName(99)).toBe('Unknown (99)');
	});
});

describe('getObjectType', () => {
	it('uses the OData type', () => {
		expect(getObjectType({ 'odata.type': 'SP.ChangeItem' })).toBe('Item');
		expect(getObjectType({ 'odata.type': 'SP.ChangeFolder' })).toBe('Folder');
		expect(getObjectType({ __metadata: { type: 'SP.ChangeWeb' } })).toBe('Web');
	});

	it('falls back to the properties', () => {
		expect(getObjectType({ ItemId: 1, ListId: LIST_ID })).toBe('Item');
		expect(getObjectType({ ListId: LIST_ID, WebId: LIST_ID })).toBe('List');
		expect(getObjectType({ FieldId: LIST_ID, ListId: LIST_ID })).toBe('Field');
		expect(getObjectType({})).toBe('Unknown');
	});
});

describe('buildChangeQuery', () => {
	it('sets the selected flags and ignores unknown names', () => {
		const query = buildChangeQuery({
			objectTypes: ['Item', 'Folder', 'NotAFlag'],
			changeTypes: ['Add', 'DeleteObject', 'RoleAssignmentAdd'],
		});
		expect(query).toEqual({
			Item: true,
			Folder: true,
			Add: true,
			DeleteObject: true,
			RoleAssignmentAdd: true,
		});
	});

	it('adds the tokens and the query options', () => {
		const query = buildChangeQuery({
			objectTypes: ['File'],
			changeTypes: ['Update'],
			startToken: ' 1;3;abc;1;2 ',
			queryOptions: {
				endPosition: '1;3;abc;9;9',
				fetchLimit: 25,
				latestFirst: true,
				recursiveAll: true,
				requireSecurityTrim: true,
				ignoreStartTokenNotFoundError: true,
			},
		});
		expect(query).toEqual({
			File: true,
			Update: true,
			ChangeTokenStart: { StringValue: '1;3;abc;1;2' },
			ChangeTokenEnd: { StringValue: '1;3;abc;9;9' },
			FetchLimit: '25',
			LatestFirst: true,
			RecursiveAll: true,
			RequireSecurityTrim: true,
			IgnoreStartTokenNotFoundError: true,
		});
	});

	it('leaves out empty values', () => {
		const query = buildChangeQuery({
			objectTypes: ['Item'],
			changeTypes: ['Add'],
			startToken: '',
			queryOptions: { endPosition: '', fetchLimit: 0, latestFirst: false },
		});
		expect(query).toEqual({ Item: true, Add: true });
	});
});

describe('trigger events', () => {
	it('every event has flags', () => {
		for (const event of CHANGE_EVENTS) {
			expect(event.objectTypes.length).toBeGreaterThan(0);
			expect(event.changeTypeFlags.length).toBeGreaterThan(0);
		}
	});

	it('maps events to the union of their flags', () => {
		expect(getEventFlags(['fileCreated', 'fileDeleted'])).toEqual({
			objectTypes: ['Item', 'File', 'Folder'],
			changeTypes: ['Add', 'DeleteObject'],
		});
		const permissions = getEventFlags(['permissionsChanged']);
		expect(permissions.changeTypes).toEqual(['RoleAssignmentAdd', 'RoleAssignmentDelete']);
		expect(permissions.objectTypes).toContain('SecurityPolicy');
		expect(getEventFlags(['moved']).changeTypes).toEqual(['Move']);
		expect(getEventFlags(['unknown'])).toEqual({ objectTypes: [], changeTypes: [] });
	});

	it('matches a change by change type and kind', () => {
		const all = CHANGE_EVENTS.map((event) => event.value);
		expect(getMatchingEvents(1, 'file', all)).toEqual(['fileCreated']);
		expect(getMatchingEvents(1, 'folder', all)).toEqual(['folderCreated']);
		expect(getMatchingEvents(1, 'item', all)).toEqual(['itemCreated']);
		expect(getMatchingEvents(3, 'file', all)).toEqual(['fileDeleted']);
		expect(getMatchingEvents(4, 'folder', all)).toEqual(['renamed']);
		expect(getMatchingEvents(5, 'file', all)).toEqual(['moved']);
		expect(getMatchingEvents(6, 'item', all)).toEqual(['moved']);
		expect(getMatchingEvents(7, 'file', all)).toEqual(['restored']);
		expect(getMatchingEvents(11, 'other', all)).toEqual(['permissionsChanged']);
		expect(getMatchingEvents(2, 'file', ['fileCreated'])).toEqual([]);
	});
});

describe('isUnderFolder', () => {
	it('keeps only paths inside the folder', () => {
		const folder = '/sites/Example/Shared Documents/A/';
		expect(isUnderFolder('/sites/Example/Shared Documents/A/b.txt', folder)).toBe(true);
		expect(isUnderFolder('/sites/example/shared documents/a/B/c.txt', folder)).toBe(true);
		expect(isUnderFolder('/sites/Example/Shared Documents/A', folder)).toBe(false);
		expect(isUnderFolder('/sites/Example/Shared Documents/AB/c.txt', folder)).toBe(false);
		expect(isUnderFolder(undefined, folder)).toBe(false);
	});
});

describe('formatChange', () => {
	it('adds readable names and keeps the SharePoint values', () => {
		const output = formatChange({
			'odata.type': 'SP.ChangeItem',
			'odata.id': 'x',
			ChangeToken: { StringValue: '1;3;abc;1;7' },
			ChangeType: 4,
			Time: '2024-01-01T00:00:00Z',
			ItemId: 5,
			ListId: LIST_ID,
			UniqueId: LIST_ID,
		});
		expect(output).toEqual({
			changeToken: '1;3;abc;1;7',
			changeType: 'Rename',
			objectType: 'Item',
			time: '2024-01-01T00:00:00Z',
			ChangeToken: { StringValue: '1;3;abc;1;7' },
			ChangeType: 4,
			Time: '2024-01-01T00:00:00Z',
			ItemId: 5,
			ListId: LIST_ID,
			UniqueId: LIST_ID,
		});
	});
});

describe('runChangeQuery', () => {
	const change = (n: number, changeType = 1) => ({
		'odata.type': 'SP.ChangeWeb',
		ChangeToken: { StringValue: `1;2;abc;${n};${n}` },
		ChangeType: changeType,
		WebId: LIST_ID,
	});

	function fakeTarget(pages: object[][]): { target: ChangeTarget; queries: ChangeQuery[] } {
		const queries: ChangeQuery[] = [];
		const target: ChangeTarget = {
			scope: 'web',
			getChanges: async (query) => {
				queries.push(query);
				return (pages.shift() ?? []) as never;
			},
			getCurrentToken: async () => 'current',
			getId: async () => LIST_ID,
		};
		return { target, queries };
	}

	const sp = {} as SPFI;

	it('reads pages with the last change token until no more changes', async () => {
		const { target, queries } = fakeTarget([[change(1), change(2)], [change(3)], []]);
		const result = await runChangeQuery(sp, target, { Web: true, Add: true }, { returnAll: true });
		expect(result.changes.map((entry) => entry.changeToken)).toEqual([
			'1;2;abc;1;1',
			'1;2;abc;2;2',
			'1;2;abc;3;3',
		]);
		expect(result.nextChangeToken).toBe('1;2;abc;3;3');
		expect(result.hasMore).toBe(false);
		expect(queries.map((query) => query.ChangeTokenStart?.StringValue)).toEqual([
			undefined,
			'1;2;abc;2;2',
			'1;2;abc;3;3',
		]);
	});

	it('stops at the limit and returns the token of the last returned change', async () => {
		const { target, queries } = fakeTarget([[change(1), change(2), change(3)]]);
		const result = await runChangeQuery(sp, target, { Web: true }, { returnAll: false, limit: 2 });
		expect(result.changes).toHaveLength(2);
		expect(result.nextChangeToken).toBe('1;2;abc;2;2');
		expect(result.hasMore).toBe(true);
		expect(queries).toHaveLength(1);
	});

	it('sends only one request when latest first is set', async () => {
		const { target, queries } = fakeTarget([[change(2), change(1)], [change(0)]]);
		const result = await runChangeQuery(
			sp,
			target,
			{ Web: true, LatestFirst: true },
			{ returnAll: true },
		);
		expect(result.changes).toHaveLength(2);
		expect(queries).toHaveLength(1);
	});

	it('filters by event and still moves the token past skipped changes', async () => {
		const { target } = fakeTarget([[change(1, 11), change(2, 2)], []]);
		const result = await runChangeQuery(
			sp,
			target,
			{ Web: true },
			{ returnAll: true, events: ['permissionsChanged'] },
		);
		expect(result.changes).toHaveLength(1);
		expect(result.changes[0].events).toEqual(['permissionsChanged']);
		expect(result.nextChangeToken).toBe('1;2;abc;2;2');
	});

	it('keeps changes without a known path in a folder change log', async () => {
		const target: ChangeTarget = {
			scope: 'list',
			folderPath: '/sites/Example/Shared Documents/A',
			getChanges: vi
				.fn()
				.mockResolvedValueOnce([
					{ 'odata.type': 'SP.ChangeList', ChangeType: 2, ChangeToken: { StringValue: 't1' } },
				])
				.mockResolvedValueOnce([]),
			getCurrentToken: async () => 't0',
			getId: async () => LIST_ID,
		};
		const result = await runChangeQuery(sp, target, { List: true }, { returnAll: true });
		expect(result.changes).toHaveLength(1);
		expect(result.nextChangeToken).toBe('t1');
	});
});

describe('isUnderFolder without recursion', () => {
	it('keeps only direct children', () => {
		const folder = '/sites/Example/Shared Documents/A';
		expect(isUnderFolder('/sites/Example/Shared Documents/A/b.txt', folder, false)).toBe(true);
		expect(isUnderFolder('/sites/Example/Shared Documents/A/B/', folder, false)).toBe(true);
		expect(isUnderFolder('/sites/Example/Shared Documents/A/B/c.txt', folder, false)).toBe(false);
	});
});

describe('RecycleBinLookup', () => {
	const entry = (id: string, deletedDate: string): RecycleBinEntry => ({
		Id: `rb-${id}`,
		UniqueId: id,
		DeletedDate: deletedDate,
		DirName: 'sites/Example/Shared Documents/A',
		LeafName: `${id}.txt`,
		ItemType: 1,
	});

	it('uses the UniqueId query when the recycle bin allows it', async () => {
		const source: RecycleBinSource = {
			findByUniqueId: vi.fn(async (id: string) => entry(id, '2024-01-01T00:00:00Z')),
			getPage: vi.fn(),
		};
		const lookup = new RecycleBinLookup(source);
		expect((await lookup.find('AAA'))?.Id).toBe('rb-AAA');
		expect(await lookup.find('aaa')).toBeDefined();
		expect(source.findByUniqueId).toHaveBeenCalledTimes(1);
		expect(source.getPage).not.toHaveBeenCalled();
	});

	it('reads pages newest first when the UniqueId query fails, and stops after the change time', async () => {
		const pages = [
			Array.from({ length: 200 }, (_, i) => entry(`n${i}`, '2024-01-02T00:00:00Z')),
			Array.from({ length: 200 }, (_, i) => entry(`m${i}`, '2023-12-01T00:00:00Z')),
			[entry('old', '2023-01-01T00:00:00Z')],
		];
		const source: RecycleBinSource = {
			findByUniqueId: vi.fn(async () => {
				throw Object.assign(new Error('threshold'), { status: 500 });
			}),
			getPage: vi.fn(async () => pages.shift() ?? []),
		};
		const lookup = new RecycleBinLookup(source);

		expect((await lookup.find('n5', '2024-01-02T00:00:00Z'))?.LeafName).toBe('n5.txt');
		expect(source.getPage).toHaveBeenCalledTimes(1);

		expect(await lookup.find('missing', '2024-01-01T12:00:00Z')).toBeUndefined();
		expect(source.getPage).toHaveBeenCalledTimes(2);
		expect(vi.mocked(source.getPage).mock.calls[1][0]?.UniqueId).toBe('n199');
		expect(source.findByUniqueId).toHaveBeenCalledTimes(1);
	});

	it('does not hide access errors', async () => {
		const source: RecycleBinSource = {
			findByUniqueId: async () => {
				throw Object.assign(new Error('denied'), { status: 403 });
			},
			getPage: async () => [],
		};
		await expect(new RecycleBinLookup(source).find('x')).rejects.toThrow('denied');
	});
});
