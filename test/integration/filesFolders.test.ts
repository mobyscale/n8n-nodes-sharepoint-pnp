import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { undefinedIfNotFound } from '../../nodes/SharePointPnp/shared/filesFolders';
import { joinPath, toServerRelativePath } from '../../nodes/SharePointPnp/shared/utils';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';
import {
	config,
	createExecuteContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	purgeTestRecycleBinItems,
} from './harness';

function toBinaryItem(buffer: Buffer, fileName: string): INodeExecutionData {
	return {
		json: {},
		binary: { data: { data: buffer.toString('base64'), mimeType: 'text/plain', fileName } },
	};
}

async function run(
	parameters: IDataObject,
	items: INodeExecutionData[] = [{ json: {} }],
): Promise<INodeExecutionData[]> {
	const context = createExecuteContext(parameters, items);
	context.helpers = {
		...context.helpers,
		getBinaryDataBuffer: async (itemIndex: number, property: string) =>
			Buffer.from(items[itemIndex].binary![property].data, 'base64'),
		prepareBinaryData: async (buffer: Buffer, fileName: string) => ({
			data: buffer.toString('base64'),
			fileName,
			mimeType: 'application/octet-stream',
		}),
	} as unknown as IExecuteFunctions['helpers'];
	return (await new SharePointPnp().execute.call(context))[0];
}

const file = (operation: string, parameters: IDataObject = {}, items?: INodeExecutionData[]) =>
	run({ resource: 'file', operation, ...parameters }, items);
const folder = (operation: string, parameters: IDataObject = {}) =>
	run({ resource: 'folder', operation, ...parameters });
const fileFolder = (operation: string, parameters: IDataObject = {}) =>
	run({ resource: 'fileFolder', operation, ...parameters });

/** SharePoint answers 400 or 404 when a recycle bin item is already gone. */
async function ignoreMissing(request: Promise<void>): Promise<void> {
	await request.catch((error: { status?: number }) =>
		error.status === 400 || error.status === 404 ? undefined : Promise.reject(error),
	);
}

/**
 * Removes the test folder and the recycle bin entries that the tests made. Only the recorded IDs
 * are removed, first from stage one, then from stage two.
 */
async function cleanUp(testFolder: string | undefined, ids: string[]): Promise<void> {
	if (!testFolder?.includes('/__pnp-test-')) return;
	const sp = createTestSp();
	const folderId = await undefinedIfNotFound(
		sp.web.getFolderByServerRelativePath(testFolder).recycle(),
	);
	for (const id of new Set([...ids, ...(folderId ? [folderId] : [])])) {
		await ignoreMissing(sp.web.recycleBin.getById(id).delete());
		await ignoreMissing(sp.site.recycleBin.getById(id).delete());
	}
	await purgeTestRecycleBinItems(testFolder);
}

describe.runIf(hasConfig)('files and folders', () => {
	let testFolder: string;
	const recycledIds: string[] = [];
	const at = (...parts: string[]) => joinPath(testFolder, ...parts);

	beforeAll(async () => {
		testFolder = await createTestFolder();
	});

	afterAll(async () => {
		await cleanUp(testFolder, recycledIds);
	});

	it('creates nested folders and does not fail when they exist', async () => {
		const first = await folder('create', { folderPath: locator(at('a', 'b', 'c')) });
		expect(first[0].json.created).toBe(true);
		expect(first[0].json.ServerRelativeUrl).toBe(at('a', 'b', 'c'));

		const second = await folder('create', { folderPath: locator(at('a', 'b', 'c')) });
		expect(second[0].json.created).toBe(false);
	});

	it('uploads, gets and downloads a file with the same bytes', async () => {
		const content = Buffer.from('hello from the integration test\n'.repeat(50));
		const uploaded = await file(
			'upload',
			{
				folderPath: locator(at('a')),
				fileName: 'upload.txt',
				binaryPropertyName: 'data',
				options: {},
			},
			[toBinaryItem(content, 'ignored.txt')],
		);
		expect(uploaded[0].json.Name).toBe('upload.txt');

		const info = await file('get', {
			filePath: locator(at('a', 'upload.txt')),
			options: { includeListItem: true },
		});
		expect(info[0].json.Length).toBe(String(content.length));
		expect(info[0].json.ListItemAllFields).toBeDefined();

		const downloaded = await file('download', {
			filePath: locator(at('a', 'upload.txt')),
			binaryPropertyName: 'data',
		});
		expect(downloaded[0].binary!.data.fileName).toBe('upload.txt');
		expect(Buffer.from(downloaded[0].binary!.data.data, 'base64').equals(content)).toBe(true);
	});

	it('uploads a large file in chunks', async () => {
		const content = Buffer.alloc(11 * 1024 * 1024 + 123, 'abcdefghij');
		await file(
			'upload',
			{
				folderPath: locator(at('a')),
				fileName: 'large.bin',
				binaryPropertyName: 'data',
				options: {},
			},
			[toBinaryItem(content, 'large.bin')],
		);
		const downloaded = await file('download', {
			filePath: locator(at('a', 'large.bin')),
			binaryPropertyName: 'data',
		});
		expect(Buffer.from(downloaded[0].binary!.data.data, 'base64').equals(content)).toBe(true);
	});

	it('does not overwrite when the option is off', async () => {
		await expect(
			file(
				'upload',
				{
					folderPath: locator(at('a')),
					fileName: 'upload.txt',
					binaryPropertyName: 'data',
					options: { overwrite: false },
				},
				[toBinaryItem(Buffer.from('x'), 'upload.txt')],
			),
		).rejects.toThrow();
	});

	it('gets a folder', async () => {
		const result = await folder('get', {
			folderPath: locator(at('a')),
			options: { includeListItem: true },
		});
		expect(result[0].json.Name).toBe('a');
		expect(result[0].json.ListItemAllFields).toBeDefined();
	});

	it('checks which paths exist', async () => {
		const libraryRoot = toServerRelativePath(config!.siteUrl, config!.libraryPath);
		const result = await fileFolder('checkExists', {
			paths: {
				values: [
					{ path: at('a', 'upload.txt') },
					{ path: at('a', 'b') },
					{ path: at('a', 'missing.txt') },
					{ path: at('missing-folder', 'deep') },
					{ path: libraryRoot },
				],
			},
		});
		expect(result.map((item) => item.json.exists)).toEqual([true, true, false, false, true]);
		expect(result.map((item) => item.json.type)).toEqual(['file', 'folder', null, null, 'folder']);
		expect(result[0].json.serverRelativeUrl).toBe(at('a', 'upload.txt'));
	});

	it('gets many entries, recursive or not, with filters', async () => {
		await folder('create', { folderPath: locator(at('many', 'sub')) });
		for (const [name, parent] of [
			['one.txt', at('many')],
			['two.pdf', at('many')],
			['three.txt', at('many', 'sub')],
		]) {
			await file(
				'upload',
				{ folderPath: locator(parent), fileName: name, binaryPropertyName: 'data', options: {} },
				[toBinaryItem(Buffer.from(name), name)],
			);
		}
		const base = {
			folderPath: locator(at('many')),
			returnAll: true,
			include: 'all',
			filters: {},
			options: {},
		};

		const flat = await fileFolder('getAll', { ...base, recursive: false });
		expect(flat.map((item) => item.json.name).sort()).toEqual(['one.txt', 'sub', 'two.pdf']);
		const one = flat.find((item) => item.json.name === 'one.txt')!.json;
		expect(one.type).toBe('file');
		expect(one.size).toBe(7);
		expect(one.serverRelativeUrl).toBe(at('many', 'one.txt'));
		expect(typeof one.listItemId).toBe('number');
		expect(typeof one.modified).toBe('string');

		const deep = await fileFolder('getAll', { ...base, recursive: true });
		expect(deep.map((item) => item.json.name).sort()).toEqual([
			'one.txt',
			'sub',
			'three.txt',
			'two.pdf',
		]);

		const filesOnly = await fileFolder('getAll', { ...base, recursive: true, include: 'files' });
		expect(filesOnly).toHaveLength(3);
		const foldersOnly = await fileFolder('getAll', {
			...base,
			recursive: true,
			include: 'folders',
		});
		expect(foldersOnly.map((item) => item.json.name)).toEqual(['sub']);

		const pdf = await fileFolder('getAll', {
			...base,
			recursive: true,
			filters: { extensions: 'pdf' },
		});
		expect(pdf.map((item) => item.json.name)).toEqual(['two.pdf']);
		const contains = await fileFolder('getAll', {
			...base,
			recursive: true,
			filters: { nameContains: 'thr' },
		});
		expect(contains.map((item) => item.json.name)).toEqual(['three.txt']);
		const future = await fileFolder('getAll', {
			...base,
			recursive: true,
			filters: { modifiedAfter: '2999-01-01T00:00:00Z' },
		});
		expect(future).toHaveLength(0);
		const past = await fileFolder('getAll', {
			...base,
			recursive: true,
			filters: { modifiedBefore: '2999-01-01T00:00:00Z' },
		});
		expect(past).toHaveLength(4);

		const limited = await fileFolder('getAll', {
			...base,
			recursive: true,
			returnAll: false,
			limit: 2,
		});
		expect(limited).toHaveLength(2);

		const withFields = await fileFolder('getAll', {
			...base,
			options: { includeListItem: true },
		});
		expect(withFields[0].json.fields).toBeDefined();
	});

	it('copies, moves and renames a file', async () => {
		await folder('create', { folderPath: locator(at('target')) });
		const copied = await file('copy', {
			filePath: locator(at('a', 'upload.txt')),
			destinationFolder: locator(at('target')),
			options: { newName: 'copy.txt' },
		});
		expect(copied[0].json.ServerRelativeUrl).toBe(at('target', 'copy.txt'));

		await expect(
			file('copy', {
				filePath: locator(at('a', 'upload.txt')),
				destinationFolder: locator(at('target')),
				options: { newName: 'copy.txt' },
			}),
		).rejects.toThrow();
		const kept = await file('copy', {
			filePath: locator(at('a', 'upload.txt')),
			destinationFolder: locator(at('target')),
			options: { newName: 'copy.txt', ifExists: 'keepBoth' },
		});
		expect(kept[0].json.sourcePath).toBe(at('a', 'upload.txt'));

		const moved = await file('move', {
			filePath: locator(at('target', 'copy.txt')),
			destinationFolder: locator(at('a', 'b')),
			options: { ifExists: 'replace' },
		});
		expect(moved[0].json.ServerRelativeUrl).toBe(at('a', 'b', 'copy.txt'));

		const renamed = await file('rename', {
			filePath: locator(at('a', 'b', 'copy.txt')),
			newName: 'renamed.txt',
		});
		expect(renamed[0].json.Name).toBe('renamed.txt');
		const old = await fileFolder('checkExists', {
			paths: {
				values: [{ path: at('a', 'b', 'copy.txt') }, { path: at('a', 'b', 'renamed.txt') }],
			},
		});
		expect(old.map((item) => item.json.exists)).toEqual([false, true]);
	});

	it('checks out, checks in, lists and restores versions', async () => {
		const path = at('a', 'upload.txt');
		const checkedOut = await file('checkOut', { filePath: locator(path) });
		expect(checkedOut[0].json.CheckOutType).toBe(0);

		const discarded = await file('discardCheckOut', { filePath: locator(path) });
		expect(discarded[0].json.CheckOutType).toBe(2);

		await file('checkOut', { filePath: locator(path) });
		const checkedIn = await file('checkIn', {
			filePath: locator(path),
			comment: 'integration test',
			checkInType: 'major',
		});
		expect(checkedIn[0].json.CheckOutType).toBe(2);

		await file(
			'upload',
			{
				folderPath: locator(at('a')),
				fileName: 'upload.txt',
				binaryPropertyName: 'data',
				options: { overwrite: true },
			},
			[toBinaryItem(Buffer.from('second version'), 'upload.txt')],
		);
		const versions = await file('getVersions', { filePath: locator(path) });
		expect(versions.length).toBeGreaterThan(0);
		const label = String(versions[0].json.VersionLabel);

		const restored = await file('restoreVersion', {
			filePath: locator(path),
			versionLabel: label,
		});
		expect(restored[0].json.restoredVersionLabel).toBe(label);
	});

	it('copies and moves a folder directly', async () => {
		const copied = await folder('copy', {
			folderPath: locator(at('many')),
			destinationFolder: locator(at('target')),
			method: 'direct',
			options: { newName: 'many-copy' },
		});
		expect(copied[0].json.ServerRelativeUrl).toBe(at('target', 'many-copy'));
		const deep = await fileFolder('getAll', {
			folderPath: locator(at('target', 'many-copy')),
			recursive: true,
			returnAll: true,
			include: 'files',
			filters: {},
			options: {},
		});
		expect(deep).toHaveLength(3);

		const moved = await folder('move', {
			folderPath: locator(at('target', 'many-copy')),
			destinationFolder: locator(at('a')),
			method: 'direct',
			options: {},
		});
		expect(moved[0].json.ServerRelativeUrl).toBe(at('a', 'many-copy'));
	});

	it('copies and moves a folder with a copy job', async () => {
		await folder('create', { folderPath: locator(at('target', 'jobs')) });
		const copied = await folder('copy', {
			folderPath: locator(at('many')),
			destinationFolder: locator(at('target', 'jobs')),
			method: 'copyJob',
			options: { newName: 'job-copy' },
			copyJobOptions: { waitForCompletion: true },
		});
		expect(copied[0].json.completed).toBe(true);
		const copyEntries = await fileFolder('getAll', {
			folderPath: locator(at('target', 'jobs', 'job-copy')),
			recursive: true,
			returnAll: true,
			include: 'files',
			filters: {},
			options: {},
		});
		expect(copyEntries).toHaveLength(3);

		const moved = await folder('move', {
			folderPath: locator(at('target', 'jobs', 'job-copy')),
			destinationFolder: locator(at('a')),
			method: 'copyJob',
			options: { newName: 'job-moved' },
			copyJobOptions: {},
		});
		expect(moved[0].json.completed).toBe(true);
		const listFiles = (path: string) =>
			fileFolder('getAll', {
				folderPath: locator(path),
				recursive: true,
				returnAll: true,
				include: 'files',
				filters: {},
				options: {},
			});
		expect(await listFiles(at('a', 'job-moved'))).toHaveLength(3);
		expect(await listFiles(at('target', 'jobs', 'job-copy'))).toHaveLength(0);
	});

	it('renames a folder', async () => {
		const renamed = await folder('rename', {
			folderPath: locator(at('target', 'jobs')),
			newName: 'jobs-renamed',
		});
		expect(renamed[0].json.Name).toBe('jobs-renamed');
	});

	it('deletes files and folders, to the recycle bin or for good', async () => {
		const recycled = await file('delete', {
			filePath: locator(at('a', 'b', 'renamed.txt')),
			options: {},
		});
		expect(recycled[0].json.permanent).toBe(false);
		expect(typeof recycled[0].json.recycleBinItemId).toBe('string');
		recycledIds.push(String(recycled[0].json.recycleBinItemId));

		const permanent = await file('delete', {
			filePath: locator(at('a', 'upload.txt')),
			options: { deletePermanently: true },
		});
		expect(permanent[0].json.permanent).toBe(true);

		const folderRecycled = await folder('delete', {
			folderPath: locator(at('target')),
			options: {},
		});
		expect(folderRecycled[0].json.deleted).toBe(true);
		recycledIds.push(String(folderRecycled[0].json.recycleBinItemId));
		const folderPermanent = await folder('delete', {
			folderPath: locator(at('many')),
			options: { deletePermanently: true },
		});
		expect(folderPermanent[0].json.permanent).toBe(true);

		const result = await fileFolder('checkExists', {
			paths: {
				values: [
					{ path: at('a', 'b', 'renamed.txt') },
					{ path: at('a', 'upload.txt') },
					{ path: at('target') },
					{ path: at('many') },
				],
			},
		});
		expect(result.map((item) => item.json.exists)).toEqual([false, false, false, false]);
	});

	it('checks one path for each input item', async () => {
		const result = await run(
			{ resource: 'fileFolder', operation: 'checkExists', paths: { values: [{ path: at('a') }] } },
			[{ json: {} }, { json: {} }],
		);
		expect(result).toHaveLength(2);
		expect(result.map((item) => item.pairedItem)).toEqual([{ item: 0 }, { item: 1 }]);
	});
});
