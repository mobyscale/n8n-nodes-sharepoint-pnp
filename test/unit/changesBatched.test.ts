import type { IDataObject, IExecuteFunctions, IHttpRequestOptions } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import {
	ChangeResolver,
	buildItemIdQuery,
	runChangeQuery,
	type ChangeRunOptions,
	type ChangeTarget,
} from '../../nodes/SharePointPnp/shared/changes';
import { createSp } from '../../nodes/SharePointPnp/transport';

const SITE = 'https://contoso.sharepoint.com/sites/Example';
const WEB_ID = 'aaaaaaaa-0000-0000-0000-000000000001';
const LIBRARY_ID = 'bbbbbbbb-0000-0000-0000-000000000001';
const LIST_ID = 'cccccccc-0000-0000-0000-000000000001';
const LIBRARY_ROOT = '/sites/Example/Shared Documents';
const TIME = '2024-01-02T00:00:00Z';

interface FakeItem extends IDataObject {
	Id: number;
	FileSystemObjectType: number;
	FileRef: string;
	FileLeafRef: string;
}

interface FakeState {
	lists: Record<string, { baseType: number; items: Map<number, FakeItem> }>;
	files: Map<string, { Name: string; ServerRelativeUrl: string; isFolder: boolean }>;
	recycleBin: IDataObject[];
	/** When true, a recycle bin filter fails as on a large recycle bin. */
	recycleBinThrottled?: boolean;
	siteUsers: IDataObject[];
}

interface LoggedRequest {
	method: string;
	path: string;
	body?: string;
}

const notFound = {
	status: 404,
	json: { 'odata.error': { code: '-2147024894', message: { lang: 'en-US', value: 'Not found' } } },
};

function guidIn(path: string, name: string): string | undefined {
	const match = path.match(new RegExp(`${name}\\((?:guid)?'([^']+)'\\)`, 'i'));
	return match?.[1].toLowerCase();
}

function listOf(state: FakeState, path: string) {
	const id = guidIn(path, 'lists');
	return id ? state.lists[id] : undefined;
}

function idsFromViewXml(body: string): number[] {
	const xml = (JSON.parse(body) as { query: { ViewXml: string } }).query.ViewXml;
	return [...xml.matchAll(/<Value Type="Counter">(\d+)<\/Value>/g)].map((m) => Number(m[1]));
}

function route(
	state: FakeState,
	method: string,
	path: string,
	body: string | undefined,
): { status: number; json: unknown } {
	const lower = path.toLowerCase();
	if (/\/_api\/web\?/.test(lower) && lower.includes('select=id')) {
		return { status: 200, json: { Id: WEB_ID } };
	}
	if (lower.includes('/getitems') && method === 'POST') {
		const list = listOf(state, path);
		if (!list) return notFound;
		const rows = idsFromViewXml(body ?? '')
			.map((id) => list.items.get(id))
			.filter(Boolean);
		return { status: 200, json: { value: rows } };
	}
	const itemMatch = path.match(/\/items\((\d+)\)/i);
	if (itemMatch) {
		const item = listOf(state, path)?.items.get(Number(itemMatch[1]));
		return item ? { status: 200, json: item } : notFound;
	}
	if (/\/lists\(/i.test(path) && lower.includes('select=basetype')) {
		const list = listOf(state, path);
		return list ? { status: 200, json: { BaseType: list.baseType } } : notFound;
	}
	const fileId = guidIn(path, 'getFileById') ?? guidIn(path, 'getFolderById');
	if (fileId) {
		const file = state.files.get(fileId);
		return file ? { status: 200, json: file } : notFound;
	}
	if (lower.includes('/siteusers')) {
		const name =
			decodeURIComponent(path.replace(/\+/g, ' ')).match(/Title eq '((?:[^']|'')*)'/)?.[1] ?? '';
		return {
			status: 200,
			json: { value: state.siteUsers.filter((user) => user.Title === name.replace(/''/g, "'")) },
		};
	}
	if (lower.includes('/recyclebin')) {
		if (state.recycleBinThrottled) {
			return {
				status: 500,
				json: { 'odata.error': { code: '-2147024860', message: { value: 'threshold' } } },
			};
		}
		const ids = [...decodeURIComponent(path).matchAll(/guid'([^']+)'/g)].map((m) =>
			m[1].toLowerCase(),
		);
		return {
			status: 200,
			json: {
				value: state.recycleBin.filter((entry) =>
					ids.includes(String(entry.UniqueId).toLowerCase()),
				),
			},
		};
	}
	if (lower.includes('/getrecyclebinitems')) {
		return { status: 200, json: { value: state.recycleBin } };
	}
	throw new Error(`Unexpected request: ${method} ${path}`);
}

function batchResponse(parts: Array<{ status: number; json: unknown }>): string {
	const lines: string[] = [];
	for (const part of parts) {
		lines.push(
			'--batchresponse_1',
			'Content-Type: application/http',
			'Content-Transfer-Encoding: binary',
			'',
			`HTTP/1.1 ${part.status} ${part.status === 200 ? 'OK' : 'Error'}`,
			'Content-Type: application/json;odata=verbose;charset=utf-8',
			'',
			JSON.stringify(part.json),
		);
	}
	lines.push('--batchresponse_1--', '');
	return lines.join('\n');
}

/** A fake SharePoint behind the real PnPjs pipeline. It logs every HTTP request and $batch part. */
function createFakeSharePoint(state: FakeState) {
	const requests: LoggedRequest[] = [];
	const batchParts: LoggedRequest[] = [];
	const httpRequestWithAuthentication = async (_type: string, options: IHttpRequestOptions) => {
		const method = String(options.method ?? 'GET');
		const url = decodeURI(String(options.url));
		const body = options.body === undefined ? undefined : String(options.body);
		requests.push({ method, path: url, body });
		let status: number;
		let text: string;
		let contentType = 'application/json';
		if (url.includes('/_api/$batch')) {
			const parts = [...(body ?? '').matchAll(/^(GET|POST) (\S+) HTTP\/1\.1$/gm)].map((m) => {
				const part = { method: m[1], path: decodeURI(m[2]) };
				batchParts.push(part);
				return route(state, part.method, part.path, undefined);
			});
			status = 200;
			text = batchResponse(parts);
			contentType = 'multipart/mixed; boundary=batchresponse_1';
		} else {
			const result = route(state, method, url, body);
			status = result.status;
			text = JSON.stringify(result.json);
		}
		return {
			statusCode: status,
			statusMessage: status === 200 ? 'OK' : 'Error',
			headers: { 'content-type': contentType },
			body: Buffer.from(text),
		};
	};
	const ctx = {
		getNode: () => ({ name: 'SharePoint PnP' }),
		getNodeParameter: () => 'certificate',
		getInputData: () => [{ json: {} }],
		helpers: { httpRequestWithAuthentication },
	} as unknown as IExecuteFunctions;
	const sp = createSp(ctx, SITE);
	const count = (pattern: RegExp) => requests.filter((r) => pattern.test(r.path)).length;
	return { sp, requests, batchParts, count };
}

function itemChange(id: number, extra: IDataObject = {}): IDataObject {
	return {
		'odata.type': 'SP.ChangeItem',
		ChangeToken: { StringValue: `1;3;${LIBRARY_ID};${id};${id}` },
		ChangeType: 1,
		ItemId: id,
		ListId: LIBRARY_ID,
		WebId: WEB_ID,
		UniqueId: `00000000-0000-0000-0001-${String(id).padStart(12, '0')}`,
		Time: TIME,
		Editor: 'Person A',
		EditorEmailHint: 'a@contoso.com',
		...extra,
	};
}

function fileItem(id: number, path: string, isFolder = false, fields: IDataObject = {}): FakeItem {
	return {
		Id: id,
		ID: id,
		FileSystemObjectType: isFolder ? 1 : 0,
		FileRef: `${LIBRARY_ROOT}/${path}`,
		FileLeafRef: path.split('/').pop() ?? path,
		Title: null,
		...fields,
	};
}

function libraryWith(items: FakeItem[]): FakeState {
	return {
		lists: {
			[LIBRARY_ID]: { baseType: 1, items: new Map(items.map((item) => [item.Id, item])) },
			[LIST_ID]: { baseType: 0, items: new Map() },
		},
		files: new Map(),
		recycleBin: [],
		siteUsers: [],
	};
}

function fakeTarget(pages: IDataObject[][], rootPath = LIBRARY_ROOT): ChangeTarget {
	return {
		scope: 'list',
		getChanges: async () => pages.shift() ?? [],
		getCurrentToken: async () => 'current',
		getId: async () => LIBRARY_ID,
		getRootPath: async () => rootPath,
	};
}

async function run(state: FakeState, pages: IDataObject[][], options: Partial<ChangeRunOptions>) {
	const fake = createFakeSharePoint(state);
	const result = await runChangeQuery(
		fake.sp,
		fakeTarget(pages),
		{},
		{
			returnAll: true,
			...options,
		},
	);
	return { ...fake, result };
}

const range = (count: number, from = 1) => Array.from({ length: count }, (_, i) => from + i);

describe('buildItemIdQuery', () => {
	it('builds a recursive CAML In query on the ID field', () => {
		expect(buildItemIdQuery([3, 7])).toBe(
			'<View Scope="RecursiveAll"><Query><Where><In><FieldRef Name="ID"/><Values><Value Type="Counter">3</Value><Value Type="Counter">7</Value></Values></In></Where></Query><RowLimit>2</RowLimit></View>',
		);
	});
});

describe('batched item details', () => {
	it('reads the details of 250 changed items in one list with one CAML request', async () => {
		const ids = range(250);
		const state = libraryWith(ids.map((id) => fileItem(id, `Reports/file-${id}.txt`)));
		const { result, requests, count } = await run(state, [ids.map((id) => itemChange(id))], {
			includeItemDetails: true,
		});

		expect(result.changes).toHaveLength(250);
		expect(count(/\/getitems/i)).toBe(1);
		expect(count(/\/items\(\d+\)/i)).toBe(0);
		// The web ID, the list base type and the CAML query. Before, it was one request per item.
		expect(requests).toHaveLength(3);
		expect(result.changes[249].item).toEqual({
			exists: true,
			kind: 'file',
			name: 'file-250.txt',
			serverRelativeUrl: `${LIBRARY_ROOT}/Reports/file-250.txt`,
			isFolder: false,
			isFile: true,
		});
		expect(result.changes[0].editorEmail).toBe('a@contoso.com');
		expect(result.nextChangeToken).toBe(`1;3;${LIBRARY_ID};250;250`);
	});

	it('sends one CAML request for each 500 IDs', async () => {
		const ids = range(1200);
		const state = libraryWith(ids.map((id) => fileItem(id, `file-${id}.txt`)));
		const { result, requests } = await run(state, [ids.map((id) => itemChange(id))], {
			includeItemDetails: true,
		});
		const queries = requests.filter((r) => /\/getitems/i.test(r.path));
		expect(queries.map((r) => idsFromViewXml(r.body ?? '').length)).toEqual([500, 500, 200]);
		expect(result.changes.every((change) => (change.item as IDataObject).exists)).toBe(true);
	});

	it('reads each item once when it changed more than once', async () => {
		const state = libraryWith([fileItem(1, 'a.txt')]);
		const { result, requests } = await run(
			state,
			[[itemChange(1), itemChange(1, { ChangeType: 2 })], [itemChange(1, { ChangeType: 2 })]],
			{ includeItemDetails: true },
		);
		expect(result.changes).toHaveLength(3);
		const queries = requests.filter((r) => /\/getitems/i.test(r.path));
		expect(queries.map((r) => idsFromViewXml(r.body ?? ''))).toEqual([[1]]);
	});

	it('uses $batch requests of 100 when an item field needs an expand', async () => {
		const ids = range(250);
		const state = libraryWith(
			ids.map((id) => fileItem(id, `file-${id}.txt`, false, { Editor: { Title: 'Person A' } })),
		);
		const { result, batchParts, count } = await run(state, [ids.map((id) => itemChange(id))], {
			includeItemDetails: true,
			itemFields: ['Editor/Title'],
		});
		expect(count(/\/_api\/\$batch/)).toBe(3);
		expect(batchParts.filter((part) => /\/items\(\d+\)/i.test(part.path))).toHaveLength(250);
		expect(count(/\/getitems/i)).toBe(0);
		expect((result.changes[0].item as IDataObject).fields).toEqual({
			'Editor/Title': 'Person A',
		});
	});

	it('resolves items, files, folders and deleted items together', async () => {
		const state = libraryWith([fileItem(1, 'Archive', true), fileItem(2, 'Archive/a.txt')]);
		state.lists[LIST_ID].items.set(5, {
			Id: 5,
			FileSystemObjectType: 0,
			FileRef: '/sites/Example/Lists/Tasks/5_.000',
			FileLeafRef: '5_.000',
			Title: 'Task five',
		});
		const fileId = 'dddddddd-0000-0000-0000-000000000001';
		const folderId = 'dddddddd-0000-0000-0000-000000000002';
		const goneFileId = 'dddddddd-0000-0000-0000-000000000003';
		state.files.set(fileId, {
			Name: 'b.txt',
			ServerRelativeUrl: `${LIBRARY_ROOT}/b.txt`,
			isFolder: false,
		});
		state.files.set(folderId, {
			Name: 'Team A',
			ServerRelativeUrl: `${LIBRARY_ROOT}/Team A`,
			isFolder: true,
		});
		const deletedItem = itemChange(3, { ChangeType: 3 });
		state.recycleBin = [
			{
				Id: 'rb-1',
				UniqueId: deletedItem.UniqueId,
				DeletedDate: TIME,
				DirName: 'sites/Example/Shared Documents/Archive',
				LeafName: 'old.txt',
				ItemType: 1,
			},
			{
				Id: 'rb-2',
				UniqueId: goneFileId,
				DeletedDate: TIME,
				DirName: 'sites/Example/Shared Documents',
				LeafName: 'gone.txt',
				ItemType: 1,
			},
		];
		const fileChange = (uniqueId: string, type: string, changeType = 1) => ({
			'odata.type': `SP.Change${type}`,
			ChangeToken: { StringValue: `1;2;${WEB_ID};9;${uniqueId.slice(-1)}` },
			ChangeType: changeType,
			UniqueId: uniqueId,
			WebId: WEB_ID,
			Time: TIME,
		});

		const { result, count, batchParts } = await run(
			state,
			[
				[
					itemChange(1),
					itemChange(2),
					deletedItem,
					itemChange(5, { ListId: LIST_ID }),
					fileChange(fileId, 'File'),
					fileChange(folderId, 'Folder'),
					fileChange(goneFileId, 'File', 3),
				],
			],
			{ includeItemDetails: true, resolveDeletedItems: true },
		);

		const items = result.changes.map((change) => change.item as IDataObject);
		expect(items.map((item) => [item.kind, item.name, item.exists])).toEqual([
			['folder', 'Archive', true],
			['file', 'a.txt', true],
			['file', 'old.txt', false],
			['item', '5_.000', true],
			['file', 'b.txt', true],
			['folder', 'Team A', true],
			['file', 'gone.txt', false],
		]);
		expect(items[2]).toMatchObject({ deleted: true, recycleBinItemId: 'rb-1' });
		// One CAML query per list, one $batch for the files and folders, one recycle bin query.
		expect(count(/\/getitems/i)).toBe(2);
		expect(count(/\/_api\/\$batch/)).toBe(1);
		expect(batchParts).toHaveLength(3);
		expect(count(/\/recyclebin/i)).toBe(1);
	});

	it('reads recycle bin pages once when the recycle bin rejects the UniqueId query', async () => {
		const state = libraryWith([]);
		state.recycleBinThrottled = true;
		const changes = range(30).map((id) => itemChange(id, { ChangeType: 3 }));
		state.recycleBin = changes.map((change, i) => ({
			Id: `rb-${i}`,
			UniqueId: change.UniqueId,
			DeletedDate: TIME,
			DirName: 'sites/Example/Shared Documents',
			LeafName: `deleted-${i}.txt`,
			ItemType: 1,
		}));
		const { result, count } = await run(state, [changes], {
			includeItemDetails: true,
			resolveDeletedItems: true,
		});
		expect(result.changes.every((change) => (change.item as IDataObject).deleted)).toBe(true);
		expect(count(/\/recyclebin/i)).toBe(1);
		expect(count(/getrecyclebinitems/i)).toBe(1);
	});

	it('marks the items of a deleted list as not found', async () => {
		const state = libraryWith([]);
		delete state.lists[LIBRARY_ID];
		const { result } = await run(state, [[itemChange(1)]], { includeItemDetails: true });
		expect(result.changes[0].item).toEqual({ exists: false });
	});

	it('reads only the details up to the limit when no filter needs them', async () => {
		const ids = range(600);
		const state = libraryWith(ids.map((id) => fileItem(id, `file-${id}.txt`)));
		const { result, requests } = await run(state, [ids.map((id) => itemChange(id))], {
			includeItemDetails: true,
			returnAll: false,
			limit: 20,
		});
		expect(result.changes).toHaveLength(20);
		expect(result.nextChangeToken).toBe(`1;3;${LIBRARY_ID};20;20`);
		const queries = requests.filter((r) => /\/getitems/i.test(r.path));
		expect(queries.map((r) => idsFromViewXml(r.body ?? '').length)).toEqual([20]);
	});

	it('keeps getDetails working without a prefetch', async () => {
		const state = libraryWith([fileItem(4, 'c.txt')]);
		const { sp, count } = createFakeSharePoint(state);
		const resolver = new ChangeResolver(sp, { itemFields: [], resolveDeletedItems: false });
		expect((await resolver.getDetails(itemChange(4), 'Item'))?.name).toBe('c.txt');
		expect((await resolver.getDetails(itemChange(4), 'Item'))?.name).toBe('c.txt');
		expect(count(/\/getitems/i)).toBe(1);
	});
});

describe('change filters', () => {
	const mixedState = () =>
		libraryWith([
			fileItem(1, 'Reports', true, { Category: 'A', Status: null }),
			fileItem(2, 'Reports/2024', true, { Category: 'B', Status: 'Open' }),
			fileItem(3, 'Reports/2024/Q1/summary.txt', false, { Category: 'A', Status: 'Open' }),
			fileItem(4, 'top.txt', false, { Category: 'a', Status: 'Closed' }),
		]);
	const mixedChanges = () => range(4).map((id) => itemChange(id));
	const names = (changes: IDataObject[]) =>
		changes.map((change) => (change.item as IDataObject | undefined)?.name ?? change.ItemId);

	it('keeps only the selected item kinds', async () => {
		const { result } = await run(mixedState(), [mixedChanges()], {
			includeItemDetails: true,
			itemKinds: ['folder'],
		});
		expect(names(result.changes)).toEqual(['Reports', '2024']);
	});

	it('uses FileSystemObjectType of the change without reading item details', async () => {
		const changes = [
			itemChange(1, { FileSystemObjectType: 1 }),
			itemChange(2, { FileSystemObjectType: 0 }),
			{ 'odata.type': 'SP.ChangeList', ChangeType: 2, ChangeToken: { StringValue: 't9' } },
		];
		const { result, count } = await run(libraryWith([]), [changes], { itemKinds: ['file'] });
		expect(result.changes.map((change) => change.ItemId)).toEqual([2]);
		expect(count(/\/getitems/i)).toBe(0);
		expect(result.nextChangeToken).toBe('t9');
	});

	it('keeps changes up to the max depth below the scope root', async () => {
		const { result } = await run(mixedState(), [mixedChanges()], {
			includeItemDetails: true,
			maxDepth: 2,
		});
		expect(names(result.changes)).toEqual(['Reports', '2024', 'top.txt']);
	});

	it('counts the depth below the folder path filter and uses the path of the change', async () => {
		const paths = ['Reports/2024/a.txt', 'Reports/2024/Q1/b.txt', 'Reports/c.txt'];
		const state = libraryWith(paths.map((path, i) => fileItem(i + 1, path)));
		const changes = () =>
			paths.map((path, i) => itemChange(i + 1, { ServerRelativeUrl: `${LIBRARY_ROOT}/${path}` }));

		const filtered = await run(state, [changes()], {
			maxDepth: 1,
			folderPathFilter: `${LIBRARY_ROOT}/Reports/2024`,
		});
		expect(filtered.result.changes.map((change) => change.ItemId)).toEqual([1]);

		// Without other filters, the depth uses the path of the change and reads no item details.
		const depthOnly = await run(state, [changes()], { maxDepth: 2 });
		expect(depthOnly.result.changes.map((change) => change.ItemId)).toEqual([3]);
		expect(depthOnly.count(/\/getitems/i)).toBe(0);
	});

	it('excludes editors by display name, email hint or resolved login', async () => {
		const changes = [
			itemChange(1, { Editor: 'Robot', EditorEmailHint: null }),
			itemChange(2, { Editor: 'Person B', EditorEmailHint: 'B@contoso.com' }),
			itemChange(3, { Editor: 'Person C', EditorEmailHint: null }),
			itemChange(4, { Editor: 'Person D', EditorEmailHint: null }),
			itemChange(5, { Editor: 'Person C', EditorEmailHint: null }),
		];
		const state = libraryWith([]);
		state.siteUsers = [
			{ Title: 'Person C', Email: '', LoginName: 'i:0#.f|membership|c@contoso.com' },
			{ Title: 'Person D', Email: 'd@contoso.com', LoginName: 'i:0#.f|membership|d@contoso.com' },
		];
		const { result, count } = await run(state, [changes], {
			excludeEditors: 'robot, b@contoso.com, C@contoso.com',
		});
		expect(result.changes.map((change) => change.ItemId)).toEqual([4]);
		// One site user request for each distinct editor name that needs a login.
		expect(count(/\/siteusers/i)).toBe(2);
		expect(result.nextChangeToken).toBe(`1;3;${LIBRARY_ID};5;5`);
	});

	it('adds the editor email only when one site user has the display name', async () => {
		const changes = [
			itemChange(1, { Editor: 'Person D', EditorEmailHint: null }),
			itemChange(2, { Editor: 'Same Name', EditorEmailHint: null }),
		];
		const state = libraryWith([fileItem(1, 'a.txt'), fileItem(2, 'b.txt')]);
		state.siteUsers = [
			{ Title: 'Person D', Email: 'd@contoso.com', LoginName: 'd' },
			{ Title: 'Same Name', Email: 'x@contoso.com', LoginName: 'x' },
			{ Title: 'Same Name', Email: 'y@contoso.com', LoginName: 'y' },
		];
		const { result } = await run(state, [changes], { includeItemDetails: true });
		expect(result.changes[0].editorEmail).toBe('d@contoso.com');
		expect(result.changes[1].editorEmail).toBeUndefined();
	});

	it('applies field filters on the batched item details', async () => {
		const { result, requests } = await run(mixedState(), [mixedChanges()], {
			fieldFilters: {
				filter: [
					{ field: 'Category', operator: 'equals', value: 'A' },
					{ field: 'Status', operator: 'isNotEmpty', value: '' },
				],
			},
		});
		expect(result.changes.map((change) => change.ItemId)).toEqual([3, 4]);
		// The filter fields are read, but they are not in the output without Item Fields.
		const select = requests.find((r) => /\/getitems/i.test(r.path))?.path ?? '';
		expect(select).toContain('Category');
		expect(result.changes[0].item).toBeUndefined();
	});

	it('does not match field filters for deleted items', async () => {
		const { result } = await run(libraryWith([]), [[itemChange(9, { ChangeType: 3 })]], {
			fieldFilters: { filter: [{ field: 'Status', operator: 'isEmpty', value: '' }] },
		});
		expect(result.changes).toHaveLength(0);
	});
});
