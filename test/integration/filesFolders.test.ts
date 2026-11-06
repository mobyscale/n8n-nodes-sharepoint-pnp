import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { sleep } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { deleteSourceFolderIfEmpty } from '../../nodes/SharePointPnp/shared/copyJob';
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

const listFiles = (path: string) =>
	fileFolder('getAll', {
		folderPath: locator(path),
		recursive: true,
		returnAll: true,
		include: 'files',
		filters: {},
		options: {},
	});

const exists = async (...paths: string[]) =>
	(await fileFolder('checkExists', { paths: { values: paths.map((path) => ({ path })) } })).map(
		(item) => item.json.exists,
	);

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

		const defaults = await file('get', { filePath: locator(at('a', 'upload.txt')), options: {} });
		expect(defaults[0].json).toMatchObject({
			Name: 'upload.txt',
			UIVersionLabel: expect.any(String),
		});
		expect(defaults[0].json.ListItemAllFields).toBeUndefined();

		const expanded = await file('get', {
			filePath: locator(at('a', 'upload.txt')),
			options: { additionalProperties: 'Author/Title', expand: 'ListItemAllFields' },
		});
		expect(expanded[0].json.Author).toBeDefined();
		expect(expanded[0].json.ListItemAllFields).toBeDefined();

		const everything = await file('get', {
			filePath: locator(at('a', 'upload.txt')),
			options: { returnAllProperties: true },
		});
		expect(everything[0].json.Name).toBe('upload.txt');

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

		const defaults = await folder('get', { folderPath: locator(at('a')), options: {} });
		expect(defaults[0].json).toMatchObject({ Name: 'a', ServerRelativeUrl: at('a') });
		expect(defaults[0].json.ItemCount).toBeDefined();
		expect(defaults[0].json.ListItemAllFields).toBeUndefined();
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

		const plain = await fileFolder('getAll', { ...base, options: {} });
		expect(plain[0].json.fields).toBeUndefined();
		const selected = await fileFolder('getAll', {
			...base,
			options: { additionalProperties: 'ContentTypeId' },
		});
		expect(Object.keys(selected[0].json.fields as object)).toEqual(
			expect.arrayContaining(['ContentTypeId']),
		);
		expect(Object.keys(selected[0].json.fields as object)).not.toContain('Author');
		const everything = await fileFolder('getAll', {
			...base,
			options: { returnAllProperties: true },
		});
		expect(everything[0].json.fields).toBeDefined();
	});

	it('copies, moves and renames a file', async () => {
		await folder('create', { folderPath: locator(at('target')) });
		const copied = await file('copy', {
			filePath: locator(at('a', 'upload.txt')),
			destinationFolder: locator(at('target')),
			method: 'direct',
			options: { newName: 'copy.txt' },
		});
		expect(copied[0].json.ServerRelativeUrl).toBe(at('target', 'copy.txt'));

		await expect(
			file('copy', {
				filePath: locator(at('a', 'upload.txt')),
				destinationFolder: locator(at('target')),
				method: 'direct',
				options: { newName: 'copy.txt' },
			}),
		).rejects.toThrow();
		const kept = await file('copy', {
			filePath: locator(at('a', 'upload.txt')),
			destinationFolder: locator(at('target')),
			method: 'direct',
			options: { newName: 'copy.txt', ifExists: 'keepBoth' },
		});
		expect(kept[0].json.sourcePath).toBe(at('a', 'upload.txt'));

		const moved = await file('move', {
			filePath: locator(at('target', 'copy.txt')),
			destinationFolder: locator(at('a', 'b')),
			method: 'direct',
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

	describe('publishing', () => {
		interface LibrarySettings {
			EnableMinorVersions: boolean;
			EnableModeration: boolean;
		}

		async function readSettings(): Promise<LibrarySettings> {
			const root = toServerRelativePath(config!.siteUrl, config!.libraryPath).toLowerCase();
			const lists = (await createTestSp()
				.web.lists.select('EnableMinorVersions', 'EnableModeration', 'RootFolder/ServerRelativeUrl')
				.expand('RootFolder')()) as unknown as Array<
				LibrarySettings & { RootFolder: { ServerRelativeUrl: string } }
			>;
			const library = lists.find(
				(list) => list.RootFolder.ServerRelativeUrl.toLowerCase() === root,
			);
			return library ?? { EnableMinorVersions: false, EnableModeration: false };
		}

		async function createFile(name: string): Promise<string> {
			await file(
				'upload',
				{ folderPath: locator(testFolder), fileName: name, binaryPropertyName: 'data', options: {} },
				[toBinaryItem(Buffer.from(name), name)],
			);
			return at(name);
		}

		const run = (operation: string, path: string) =>
			file(operation, { filePath: locator(path), comment: 'integration test' });

		it('publishes and unpublishes a file in a library with minor versions', async (context) => {
			const settings = await readSettings();
			if (!settings.EnableMinorVersions) {
				context.skip('The test library has no minor versions. Settings are never changed.');
			}
			const path = await createFile('publish.txt');
			const published = await run('publish', path);
			expect(published[0].json.Level).toBe(1);
			const unpublished = await run('unpublish', path);
			expect(unpublished[0].json.Level).toBe(2);
		});

		it('approves and denies a file in a library with content approval', async (context) => {
			const settings = await readSettings();
			if (!settings.EnableModeration) {
				context.skip('The test library has no content approval. Settings are never changed.');
			}
			const path = await createFile('approve.txt');
			if (settings.EnableMinorVersions) await run('publish', path);
			const approved = await run('approve', path);
			expect(approved[0].json.moderationStatus).toBe(0);
			await file(
				'upload',
				{
					folderPath: locator(testFolder),
					fileName: 'approve.txt',
					binaryPropertyName: 'data',
					options: { overwrite: true },
				},
				[toBinaryItem(Buffer.from('second'), 'approve.txt')],
			);
			if (settings.EnableMinorVersions) await run('publish', path);
			const denied = await run('deny', path);
			expect(denied[0].json.moderationStatus).toBe(1);
		});

		it('returns the SharePoint error when the library does not support the operation', async (context) => {
			const settings = await readSettings();
			if (settings.EnableMinorVersions && settings.EnableModeration) {
				context.skip('The test library supports every publishing operation.');
			}
			const path = await createFile('unsupported.txt');
			if (!settings.EnableMinorVersions) {
				await expect(run('publish', path)).rejects.toThrow(/minor version/);
				await expect(run('unpublish', path)).rejects.toThrow(/minor version/);
			}
			if (!settings.EnableModeration) {
				await expect(run('approve', path)).rejects.toThrow(/content approval/);
				await expect(run('deny', path)).rejects.toThrow(/content approval/);
			}
		});
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
		expect(versions[0].json).toMatchObject({ VersionLabel: expect.any(String) });
		const withCreator = await file('getVersions', {
			filePath: locator(path),
			versionOptions: { additionalProperties: 'CreatedBy/Title' },
		});
		expect(withCreator[0].json.CreatedBy).toBeDefined();
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
			copyJobNewName: 'job-copy',
			waitForCompletion: true,
		});
		expect(copied[0].json.status).toBe('completed');
		expect(copied[0].json.completed).toBe(true);
		expect(copied[0].json.destinationPath).toBe(at('target', 'jobs', 'job-copy'));
		expect(copied[0].json.summary).toMatchObject({ files: 3, errors: 0 });
		expect(copied[0].json.jobId).toEqual(expect.any(String));
		expect(await listFiles(at('target', 'jobs', 'job-copy'))).toHaveLength(3);

		const moved = await folder('move', {
			folderPath: locator(at('target', 'jobs', 'job-copy')),
			destinationFolder: locator(at('a')),
			method: 'copyJob',
			copyJobNewName: 'job-moved',
			copyJobOptions: {},
		});
		expect(moved[0].json.status).toBe('completed');
		expect(await listFiles(at('a', 'job-moved'))).toHaveLength(3);
		expect(await exists(at('target', 'jobs', 'job-copy'))).toEqual([false]);
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

	describe('copy jobs', () => {
		const uploadText = (folderPath: string, name: string) =>
			file(
				'upload',
				{
					folderPath: locator(folderPath),
					fileName: name,
					binaryPropertyName: 'data',
					options: {},
				},
				[toBinaryItem(Buffer.from(`content of ${name}`), name)],
			);

		async function makeTree(root: string) {
			await folder('create', { folderPath: locator(joinPath(root, 'sub')) });
			await uploadText(root, 'one.txt');
			await uploadText(joinPath(root, 'sub'), 'two.txt');
		}

		async function waitForFinish(job: IDataObject): Promise<IDataObject> {
			for (let attempt = 0; attempt < 60; attempt++) {
				const [status] = await fileFolder('getCopyJobStatus', { jobInfo: job });
				if (status.json.finished) return status.json;
				await sleep(2000);
			}
			throw new Error('The copy job did not finish in time');
		}

		it('copies and moves a single file with a copy job', async () => {
			await folder('create', { folderPath: locator(at('cj-file', 'dst')) });
			await uploadText(at('cj-file'), 'single.txt');

			const copied = await file('copy', {
				filePath: locator(at('cj-file', 'single.txt')),
				destinationFolder: locator(at('cj-file', 'dst')),
				method: 'copyJob',
				copyJobNewName: 'copied.txt',
			});
			expect(copied[0].json.status).toBe('completed');
			expect(copied[0].json.destinationPath).toBe(at('cj-file', 'dst', 'copied.txt'));
			expect(copied[0].json.Name).toBe('copied.txt');
			expect(await exists(at('cj-file', 'single.txt'))).toEqual([true]);

			const moved = await file('move', {
				filePath: locator(at('cj-file', 'single.txt')),
				destinationFolder: locator(at('cj-file', 'dst')),
				method: 'copyJob',
			});
			expect(moved[0].json.status).toBe('completed');
			expect(moved[0].json.Name).toBe('single.txt');
			expect(await exists(at('cj-file', 'single.txt'), at('cj-file', 'dst', 'single.txt'))).toEqual(
				[false, true],
			);
		});

		it('moves a folder and leaves no source folder behind', async () => {
			await makeTree(at('cj-move', 'src'));
			await folder('create', { folderPath: locator(at('cj-move', 'dst')) });

			const moved = await folder('move', {
				folderPath: locator(at('cj-move', 'src')),
				destinationFolder: locator(at('cj-move', 'dst')),
				method: 'copyJob',
				copyJobOptions: { deleteEmptySourceFolder: true },
			});
			expect(moved[0].json.status).toBe('completed');
			expect(moved[0].json.summary).toMatchObject({ files: 2, errors: 0 });
			expect(await exists(at('cj-move', 'src'))).toEqual([false]);
			expect(await listFiles(at('cj-move', 'dst', 'src'))).toHaveLength(2);
		});

		it('keeps the source with the option "Move But Keep Source"', async () => {
			await makeTree(at('cj-keep', 'src'));
			await folder('create', { folderPath: locator(at('cj-keep', 'dst')) });

			const moved = await folder('move', {
				folderPath: locator(at('cj-keep', 'src')),
				destinationFolder: locator(at('cj-keep', 'dst')),
				method: 'copyJob',
				copyJobOptions: { moveButKeepSource: true },
			});
			expect(moved[0].json.status).toBe('completed');
			expect(moved[0].json.sourceFolderDeleted).toBeUndefined();
			expect(await listFiles(at('cj-keep', 'src'))).toHaveLength(2);
			expect(await listFiles(at('cj-keep', 'dst', 'src'))).toHaveLength(2);
		});

		it('deletes a source folder only when it is empty', async () => {
			await folder('create', { folderPath: locator(at('cj-empty', 'empty')) });
			await folder('create', { folderPath: locator(at('cj-empty', 'full')) });
			await uploadText(at('cj-empty', 'full'), 'keep.txt');
			const sp = createTestSp();

			expect(await deleteSourceFolderIfEmpty(sp, at('cj-empty', 'full'))).toBe(false);
			expect(await deleteSourceFolderIfEmpty(sp, at('cj-empty', 'empty'))).toBe(true);
			expect(await deleteSourceFolderIfEmpty(sp, at('cj-empty', 'empty'))).toBe(false);
			expect(await exists(at('cj-empty', 'full', 'keep.txt'), at('cj-empty', 'empty'))).toEqual([
				true,
				false,
			]);
		});

		it('resets role inheritance of the moved folder', async () => {
			await makeTree(at('cj-role', 'src'));
			await folder('create', { folderPath: locator(at('cj-role', 'dst')) });

			const moved = await folder('move', {
				folderPath: locator(at('cj-role', 'src')),
				destinationFolder: locator(at('cj-role', 'dst')),
				method: 'copyJob',
				copyJobOptions: { resetRoleInheritanceAfterMove: true },
			});
			expect(moved[0].json.status).toBe('completed');
			expect(moved[0].json.roleInheritanceReset).toBe(true);
		});

		it('starts a job without waiting and follows it with Get Copy Job Status', async () => {
			await makeTree(at('cj-async', 'src'));
			await folder('create', { folderPath: locator(at('cj-async', 'dst')) });

			const started = await folder('copy', {
				folderPath: locator(at('cj-async', 'src')),
				destinationFolder: locator(at('cj-async', 'dst')),
				method: 'copyJob',
				waitForCompletion: false,
			});
			const job = started[0].json;
			expect(job.status).toBe('running');
			expect(job.finished).toBe(false);
			expect(job.jobId).toEqual(expect.any(String));
			expect(job.jobQueueUri).toEqual(expect.any(String));
			expect(job.encryptionKey).toEqual(expect.any(String));
			expect(job.sourcePath).toBe(at('cj-async', 'src'));
			expect(job.destinationPath).toBe(at('cj-async', 'dst', 'src'));

			const status = await waitForFinish(job);
			expect(status.status).toBe('completed');
			expect(status.sourcePath).toBe(at('cj-async', 'src'));
			expect(await listFiles(at('cj-async', 'dst', 'src'))).toHaveLength(2);
		});

		it('returns a running job when the maximum wait passes', async () => {
			await makeTree(at('cj-slow', 'src'));
			await folder('create', { folderPath: locator(at('cj-slow', 'dst')) });

			const result = await folder('copy', {
				folderPath: locator(at('cj-slow', 'src')),
				destinationFolder: locator(at('cj-slow', 'dst')),
				method: 'copyJob',
				waitForCompletion: true,
				maxWaitSeconds: 1,
			});
			expect(['running', 'completed']).toContain(result[0].json.status);
			if (result[0].json.status === 'running') {
				expect((await waitForFinish(result[0].json)).status).toBe('completed');
			}
		});

		it('fails or keeps both when the name exists', async () => {
			await makeTree(at('cj-conflict', 'src'));
			await folder('create', { folderPath: locator(at('cj-conflict', 'dst')) });
			const parameters = {
				folderPath: locator(at('cj-conflict', 'src')),
				destinationFolder: locator(at('cj-conflict', 'dst')),
				method: 'copyJob',
			};
			await folder('copy', parameters);

			await expect(folder('copy', { ...parameters, nameConflictBehavior: 'fail' })).rejects.toThrow(
				/copy job/,
			);
			const reported = await folder('copy', {
				...parameters,
				nameConflictBehavior: 'fail',
				copyJobOptions: { continueOnJobErrors: true },
			});
			expect(reported[0].json.status).toBe('failed');
			expect((reported[0].json.summary as IDataObject).errors).toBeGreaterThan(0);

			const kept = await folder('copy', { ...parameters, nameConflictBehavior: 'keepBoth' });
			expect(kept[0].json.status).toBe('completed');
			expect(kept[0].json.destinationPath).not.toBe(at('cj-conflict', 'dst', 'src'));
			const entries = await fileFolder('getAll', {
				folderPath: locator(at('cj-conflict', 'dst')),
				recursive: false,
				returnAll: true,
				include: 'folders',
				filters: {},
				options: {},
			});
			expect(entries).toHaveLength(2);
		});
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
