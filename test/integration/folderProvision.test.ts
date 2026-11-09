import type { IDataObject } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { undefinedIfNotFound } from '../../nodes/SharePointPnp/shared/filesFolders';
import { joinPath } from '../../nodes/SharePointPnp/shared/utils';
import {
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

const folder = async (operation: string, parameters: IDataObject) =>
	(await runNode({ resource: 'folder', operation, ...parameters }))[0].json;
const folderItems = async (operation: string, parameters: IDataObject) =>
	(await runNode({ resource: 'folder', operation, ...parameters })).map((entry) => entry.json);

describe.runIf(hasConfig)('folder provisioning operations', () => {
	let testFolder: string;
	const at = (...parts: string[]) => joinPath(testFolder, ...parts);
	const sp = () => createTestSp();

	/** Creates the folder and its missing parents below the test folder, one level at a time. */
	const addFolder = async (path: string) => {
		let current = testFolder;
		for (const name of path.slice(testFolder.length).split('/').filter(Boolean)) {
			current = joinPath(current, name);
			if (!(await folderExists(current))) await sp().web.folders.addUsingPath(current, false);
		}
	};
	const addFile = async (folderPath: string, name: string) =>
		await sp()
			.web.getFolderByServerRelativePath(folderPath)
			.files.addUsingPath(name, `content of ${name}`);
	const folderItem = async (path: string) =>
		(await sp()
			.web.getFolderByServerRelativePath(path)
			.listItemAllFields.select('Title', 'Modified')()) as IDataObject;
	const fileItem = async (path: string) =>
		(await sp()
			.web.getFileByServerRelativePath(path)
			.listItemAllFields.select('Title', 'Modified')()) as IDataObject;
	const folderExists = async (path: string) =>
		(await undefinedIfNotFound(sp().web.getFolderByServerRelativePath(path).select('Exists')()))
			?.Exists === true;
	const fileExists = async (path: string) =>
		(await undefinedIfNotFound(sp().web.getFileByServerRelativePath(path).select('Exists')()))
			?.Exists === true;

	beforeAll(async () => {
		testFolder = await createTestFolder();
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('ensures a three-level path and reports no change on a second run', async () => {
		const parameters = {
			folderPath: locator(at()),
			segmentsInputMode: 'rows',
			segments: {
				values: [
					{ name: 'Reports', fieldsWhenCreated: '{"Title":"Reports"}', fieldsAlways: '' },
					{ name: '2024', fieldsWhenCreated: '', fieldsAlways: '{"Title":"2024"}' },
					{ name: 'Q1', fieldsWhenCreated: { Title: 'Q1' } },
				],
			},
		};
		const first = await folder('ensurePath', parameters);
		expect(first).toMatchObject({
			created: true,
			changed: true,
			folderPath: at('Reports', '2024', 'Q1'),
		});
		expect(first.segments).toEqual([
			expect.objectContaining({
				created: true,
				fieldsUpdated: ['Title'],
				listItemId: expect.any(Number),
			}),
			expect.objectContaining({ created: true, fieldsUpdated: ['Title'] }),
			expect.objectContaining({ created: true, fieldsUpdated: ['Title'] }),
		]);
		expect((await folderItem(at('Reports'))).Title).toBe('Reports');
		expect((await folderItem(at('Reports', '2024'))).Title).toBe('2024');
		expect((await folderItem(at('Reports', '2024', 'Q1'))).Title).toBe('Q1');

		const second = await folder('ensurePath', parameters);
		expect(second).toMatchObject({ created: false, changed: false });
		expect(second.segments).toEqual([
			expect.objectContaining({ created: false, fieldsUpdated: [] }),
			expect.objectContaining({ created: false, fieldsUpdated: [] }),
			expect.objectContaining({ created: false, fieldsUpdated: [] }),
		]);

		const json = await folder('ensurePath', {
			folderPath: locator(at()),
			segmentsInputMode: 'json',
			segmentsJson: JSON.stringify([
				'Reports',
				'2024',
				{
					name: 'Q1',
					fieldsWhenCreated: { Title: 'Ignored' },
					fieldsAlways: { Title: 'Q1 Renamed' },
				},
				{ name: 'Phase 1' },
			]),
		});
		expect(json).toMatchObject({ created: true, changed: true });
		expect(
			(json.segments as IDataObject[]).map((entry) => [entry.created, entry.fieldsUpdated]),
		).toEqual([
			[false, []],
			[false, []],
			[false, ['Title']],
			[true, []],
		]);
		expect((await folderItem(at('Reports', '2024', 'Q1'))).Title).toBe('Q1 Renamed');
	});

	it('copies the children of a template folder only when they are missing', async () => {
		const template = at('template');
		const dest = at('copy');
		await addFolder(joinPath(template, '01 Admin'));
		await addFolder(joinPath(template, '02 Images', 'Inner'));
		await addFile(joinPath(template, '02 Images'), 'image.txt');
		await addFolder(joinPath(template, 'Archive'));
		await addFile(template, 'readme.txt');
		await addFolder(joinPath(dest, '01 Admin'));

		const parameters = {
			folderPath: locator(template),
			childrenDestinationFolder: locator(dest),
			copyChildrenOptions: { excludeNames: 'Arch*', maxWaitSeconds: 240 },
		};
		const first = await folder('copyChildrenIfMissing', parameters);
		expect(first.changed).toBe(true);
		expect(first.copied).toEqual([
			expect.objectContaining({
				name: '02 Images',
				status: 'completed',
				destinationPath: joinPath(dest, '02 Images'),
			}),
		]);
		expect(
			(first.skipped as IDataObject[]).map((entry) => [entry.name, entry.reason]).sort(),
		).toEqual([
			['01 Admin', 'exists'],
			['Archive', 'excluded'],
			['readme.txt', 'type'],
		]);
		expect(await folderExists(joinPath(dest, '02 Images', 'Inner'))).toBe(true);
		expect(await fileExists(joinPath(dest, '02 Images', 'image.txt'))).toBe(true);
		expect(await folderExists(joinPath(dest, 'Archive'))).toBe(false);

		const second = await folder('copyChildrenIfMissing', parameters);
		expect(second).toMatchObject({ changed: false, copied: [] });

		const direct = await folder('copyChildrenIfMissing', {
			...parameters,
			copyChildrenOptions: { include: 'files', copyMethod: 'direct', resetAuthorAndCreated: true },
		});
		expect(direct.copied).toEqual([expect.objectContaining({ name: 'readme.txt', type: 'file' })]);
		expect(await fileExists(joinPath(dest, 'readme.txt'))).toBe(true);

		const missing = await folder('copyChildrenIfMissing', {
			...parameters,
			childrenDestinationFolder: locator(at('no-such-folder')),
			copyChildrenOptions: { onlyIfDestinationExists: true },
		});
		expect(missing).toMatchObject({ changed: false, destinationExists: false });
		expect(await folderExists(at('no-such-folder'))).toBe(false);
	});

	it('sets metadata on all files and folders below a folder', async () => {
		const root = at('meta');
		await addFolder(joinPath(root, 'x', 'y'));
		await addFile(root, 'top.txt');
		await addFile(joinPath(root, 'x'), 'f1.txt');
		await addFile(joinPath(root, 'x', 'y'), 'f2.txt');

		const parameters = {
			folderPath: locator(root),
			metadataInputMode: 'keyValue',
			metadataFields: { values: [{ name: 'Title', value: 'Tagged' }] },
			metadataOptions: { batchSize: 2 },
		};
		const first = await folder('setMetadataRecursive', parameters);
		expect(first).toMatchObject({ matched: 5, updated: 5, skipped: 0, failed: 0, errors: [] });
		expect((await folderItem(joinPath(root, 'x'))).Title).toBe('Tagged');
		expect((await folderItem(joinPath(root, 'x', 'y'))).Title).toBe('Tagged');
		expect((await fileItem(joinPath(root, 'top.txt'))).Title).toBe('Tagged');
		expect((await fileItem(joinPath(root, 'x', 'f1.txt'))).Title).toBe('Tagged');
		expect((await fileItem(joinPath(root, 'x', 'y', 'f2.txt'))).Title).toBe('Tagged');
		expect((await folderItem(root)).Title).not.toBe('Tagged');

		const second = await folder('setMetadataRecursive', parameters);
		expect(second).toMatchObject({ matched: 5, updated: 0, skipped: 5 });

		const before = await fileItem(joinPath(root, 'top.txt'));
		const system = await folder('setMetadataRecursive', {
			folderPath: locator(root),
			metadataInputMode: 'json',
			metadataFieldsJson: '{"Title":"System"}',
			metadataOptions: { include: 'files', maxDepth: 1, systemUpdate: true },
		});
		expect(system).toMatchObject({ matched: 1, updated: 1 });
		const after = await fileItem(joinPath(root, 'top.txt'));
		expect(after.Title).toBe('System');
		expect(after.Modified).toBe(before.Modified);
		expect((await fileItem(joinPath(root, 'x', 'f1.txt'))).Title).toBe('Tagged');

		const self = await folder('setMetadataRecursive', {
			folderPath: locator(root),
			metadataInputMode: 'keyValue',
			metadataFields: { values: [{ name: 'Title', value: 'Folder' }] },
			metadataOptions: {
				include: 'folders',
				includeFolderItself: true,
				onlyFieldsInContentType: true,
				continueOnItemError: true,
			},
		});
		expect(self).toMatchObject({ matched: 3, failed: 0 });
		expect((self.updated as number) + (self.skipped as number)).toBe(3);
		if (self.updated === 3) expect((await folderItem(root)).Title).toBe('Folder');
	});

	it('ensures a five-folder tree with shared parents and reports no change on a second run', async () => {
		const base = at('tree');
		await addFolder(base);
		const parameters = {
			folderPath: locator(base),
			foldersInputMode: 'rows',
			folders: {
				values: [
					{ key: 'reports', path: 'Reports', fieldsWhenCreated: '{"Title":"Reports"}' },
					{ key: 'archive', path: 'Reports/Archive' },
					{ key: 'year', path: 'Reports/2024' },
					{ key: 'q2', path: 'Reports/2024/Q2' },
					{
						key: 'quarter',
						path: 'Reports/2024/Q1',
						fieldsAlways: '{"Title":"Q1"}',
					},
				],
			},
		};
		const first = await folder('ensureFolders', parameters);
		expect(first).toMatchObject({ created: true, changed: true, baseFolder: base });
		const folders = first.folders as Record<string, IDataObject>;
		expect(Object.keys(folders)).toEqual(['reports', 'archive', 'year', 'q2', 'quarter']);
		expect(Object.values(folders).every((entry) => entry.created === true)).toBe(true);
		expect(folders.quarter).toMatchObject({
			path: 'Reports/2024/Q1',
			serverRelativeUrl: joinPath(base, 'Reports', '2024', 'Q1'),
			listItemId: expect.any(Number),
			uniqueId: expect.any(String),
			fieldsUpdated: ['Title'],
		});
		expect(String(folders.quarter.absoluteUrl)).toMatch(/^https:\/\/.+Q1$/);
		expect(String(folders.quarter.siteRelativeUrl)).toMatch(/\/Reports\/2024\/Q1$/);
		expect(await folderExists(joinPath(base, 'Reports', 'Archive'))).toBe(true);
		expect(await folderExists(joinPath(base, 'Reports', '2024', 'Q2'))).toBe(true);
		expect((await folderItem(joinPath(base, 'Reports'))).Title).toBe('Reports');
		expect((await folderItem(joinPath(base, 'Reports', '2024', 'Q1'))).Title).toBe('Q1');

		const second = await folderItems('ensureFolders', {
			...parameters,
			foldersOutput: 'perFolder',
		});
		expect(second.map((entry) => [entry.key, entry.created, entry.fieldsUpdated])).toEqual([
			['reports', false, []],
			['archive', false, []],
			['year', false, []],
			['q2', false, []],
			['quarter', false, []],
		]);
		expect(second[4].serverRelativeUrl).toBe(folders.quarter.serverRelativeUrl);

		const json = await folder('ensureFolders', {
			folderPath: locator(base),
			foldersInputMode: 'json',
			foldersJson: JSON.stringify([
				{ key: 'year', path: 'Reports/2024' },
				{ key: 'phase', path: 'Reports/2024/Q1/Phase 1' },
			]),
		});
		expect(json).toMatchObject({ created: true });
		expect((json.folders as Record<string, IDataObject>).year.created).toBe(false);
		expect((json.folders as Record<string, IDataObject>).phase.created).toBe(true);
	});

	it('inherits metadata from the parent folder to the folder and its content', async () => {
		const parent = at('inherit', 'Parent');
		const child = joinPath(parent, 'Child');
		await addFolder(joinPath(child, 'Sub'));
		await addFolder(joinPath(parent, 'Middle', 'Leaf'));
		await addFile(child, 'c.txt');
		await addFile(joinPath(child, 'Sub'), 's.txt');
		const parentItem = await sp().web.getFolderByServerRelativePath(parent).getItem();
		await parentItem.update({ Title: 'Inherited' });

		const parameters = { folderPath: locator(child), inheritFields: 'Title' };
		const first = await folder('inheritMetadataFromParent', parameters);
		expect(first).toMatchObject({
			folderPath: child,
			parentFolder: parent,
			sourceFolder: parent,
			values: { Title: 'Inherited' },
			missingFields: [],
			matched: 4,
			updated: 4,
			failed: 0,
		});
		expect((await folderItem(child)).Title).toBe('Inherited');
		expect((await folderItem(joinPath(child, 'Sub'))).Title).toBe('Inherited');
		expect((await fileItem(joinPath(child, 'c.txt'))).Title).toBe('Inherited');
		expect((await fileItem(joinPath(child, 'Sub', 's.txt'))).Title).toBe('Inherited');

		const second = await folder('inheritMetadataFromParent', parameters);
		expect(second).toMatchObject({ matched: 4, updated: 0, skipped: 4 });

		const leaf = joinPath(parent, 'Middle', 'Leaf');
		const direct = await folder('inheritMetadataFromParent', {
			folderPath: locator(leaf),
			inheritFields: 'Title',
			inheritOptions: { ignoreEmptyValues: true },
		});
		expect(direct).toMatchObject({ missingFields: ['Title'], matched: 0, updated: 0 });

		const searched = await folder('inheritMetadataFromParent', {
			folderPath: locator(leaf),
			inheritFields: 'Title',
			inheritOptions: { searchUp: true },
		});
		expect(searched).toMatchObject({
			sourceFolder: parent,
			valueSources: { Title: parent },
			updated: 1,
		});
		expect((await folderItem(leaf)).Title).toBe('Inherited');
		expect((await folderItem(joinPath(parent, 'Middle'))).Title).not.toBe('Inherited');
	});
});
