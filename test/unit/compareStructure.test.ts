import { describe, expect, it, vi } from 'vitest';

import { fileFolderResource } from '../../nodes/SharePointPnp/resources/fileFolder';
import {
	compareTrees,
	isExcluded,
	parsePatterns,
	readTree,
	type FolderChild,
	type FolderReader,
	type TreeOptions,
} from '../../nodes/SharePointPnp/resources/fileFolder/compare';

const ROOT = '/sites/Example/Shared Documents';

/** A fake folder reader. Paths ending with "/" are folders, the others are files. */
function fakeReader(paths: string[]) {
	const read = vi.fn<FolderReader>(async (url, foldersOnly) => {
		const children: FolderChild[] = [];
		for (const path of paths) {
			const isFolder = path.endsWith('/');
			const full = `${ROOT}/${path.replace(/\/$/, '')}`;
			const parent = full.slice(0, full.lastIndexOf('/'));
			if (parent !== url || (foldersOnly && !isFolder)) continue;
			children.push({
				name: full.slice(full.lastIndexOf('/') + 1),
				type: isFolder ? 'folder' : 'file',
				serverRelativeUrl: full,
			});
		}
		return children;
	});
	return read;
}

const options = (overrides: Partial<TreeOptions> = {}): TreeOptions => ({
	include: 'folders',
	maxDepth: 1,
	caseSensitive: false,
	exclude: [],
	...overrides,
});

const SOURCE = ['Template/', 'Template/A/', 'Template/A/B/', 'Template/C/', 'Template/x.txt'];
const TARGET = ['Copy/', 'Copy/A/', 'Copy/D/'];

async function compare(treeOptions: TreeOptions, source = SOURCE, target = TARGET) {
	return compareTrees(
		await readTree(`${ROOT}/Template`, fakeReader(source), treeOptions),
		await readTree(`${ROOT}/Copy`, fakeReader(target), treeOptions),
		treeOptions.caseSensitive,
	);
}

describe('readTree', () => {
	it('reads only the direct children with Max Depth 1', async () => {
		const read = fakeReader(SOURCE);
		const tree = await readTree(`${ROOT}/Template`, read, options());
		expect(tree).toEqual([
			{ path: 'A', type: 'folder' },
			{ path: 'C', type: 'folder' },
		]);
		expect(read).toHaveBeenCalledTimes(1);
		expect(read).toHaveBeenCalledWith(`${ROOT}/Template`, true);
	});

	it('reads all levels with Max Depth 0', async () => {
		const tree = await readTree(`${ROOT}/Template`, fakeReader(SOURCE), options({ maxDepth: 0 }));
		expect(tree.map((entry) => entry.path)).toEqual(['A', 'C', 'A/B']);
	});

	it('stops at the given depth', async () => {
		const deep = ['Template/', 'Template/A/', 'Template/A/B/', 'Template/A/B/C/'];
		const tree = await readTree(`${ROOT}/Template`, fakeReader(deep), options({ maxDepth: 2 }));
		expect(tree.map((entry) => entry.path)).toEqual(['A', 'A/B']);
	});

	it('reads files too when files are included, and still goes down the folders', async () => {
		const read = fakeReader([...SOURCE, 'Template/A/y.txt']);
		const tree = await readTree(
			`${ROOT}/Template`,
			read,
			options({ include: 'files', maxDepth: 0 }),
		);
		expect(tree).toEqual([
			{ path: 'x.txt', type: 'file' },
			{ path: 'A/y.txt', type: 'file' },
		]);
		expect(read).toHaveBeenCalledWith(`${ROOT}/Template`, false);
	});

	it('does not read an excluded folder', async () => {
		const read = fakeReader(SOURCE);
		const tree = await readTree(`${ROOT}/Template`, read, options({ maxDepth: 0, exclude: ['a'] }));
		expect(tree.map((entry) => entry.path)).toEqual(['C']);
		expect(read).not.toHaveBeenCalledWith(`${ROOT}/Template/A`, true);
	});
});

describe('compareTrees', () => {
	it('compares folders of the first level', async () => {
		expect(await compare(options())).toEqual({
			isComplete: false,
			missing: ['C'],
			extra: ['D'],
			matched: 1,
		});
	});

	it('compares folders of all levels', async () => {
		expect(await compare(options({ maxDepth: 0 }))).toEqual({
			isComplete: false,
			missing: ['A/B', 'C'],
			extra: ['D'],
			matched: 1,
		});
	});

	it('compares files only', async () => {
		expect(await compare(options({ include: 'files' }))).toEqual({
			isComplete: false,
			missing: ['x.txt'],
			extra: [],
			matched: 0,
		});
	});

	it('compares folders and files', async () => {
		expect(await compare(options({ include: 'both', maxDepth: 0 }))).toEqual({
			isComplete: false,
			missing: ['A/B', 'C', 'x.txt'],
			extra: ['D'],
			matched: 1,
		});
	});

	it('is complete when the target has all source entries, even with extra entries', async () => {
		const target = ['Copy/', 'Copy/a/', 'Copy/a/b/', 'Copy/C/', 'Copy/D/'];
		expect(await compare(options({ maxDepth: 0 }), SOURCE, target)).toEqual({
			isComplete: true,
			missing: [],
			extra: ['D'],
			matched: 3,
		});
	});

	it('treats names with other case as different when Case Sensitive is on', async () => {
		const target = ['Copy/', 'Copy/a/', 'Copy/C/'];
		expect(await compare(options({ caseSensitive: true }), SOURCE, target)).toEqual({
			isComplete: false,
			missing: ['A'],
			extra: ['a'],
			matched: 1,
		});
	});

	it('does not match a folder with a file of the same name', () => {
		expect(
			compareTrees([{ path: 'A', type: 'folder' }], [{ path: 'A', type: 'file' }], false),
		).toEqual({ isComplete: false, missing: ['A'], extra: ['A'], matched: 0 });
	});

	it('leaves excluded names out of both trees', async () => {
		expect(await compare(options({ maxDepth: 0, exclude: ['C', 'D'] }))).toEqual({
			isComplete: false,
			missing: ['A/B'],
			extra: [],
			matched: 1,
		});
	});
});

describe('exclude patterns', () => {
	it('splits on commas and line breaks', () => {
		expect(parsePatterns(' Archive, ~*\nDrafts/Old ,, ')).toEqual(['Archive', '~*', 'Drafts/Old']);
		expect(parsePatterns(undefined)).toEqual([]);
	});

	it('matches names with wildcards', () => {
		expect(isExcluded('~temp.docx', '~temp.docx', ['~*'], false)).toBe(true);
		expect(isExcluded('File1.txt', 'File1.txt', ['File?.txt'], false)).toBe(true);
		expect(isExcluded('File10.txt', 'File10.txt', ['File?.txt'], false)).toBe(false);
		expect(isExcluded('a.b', 'a.b', ['a?b'], false)).toBe(true);
		expect(isExcluded('axb', 'axb', ['a.b'], false)).toBe(false);
	});

	it('uses Case Sensitive for patterns', () => {
		expect(isExcluded('Archive', 'Archive', ['archive'], false)).toBe(true);
		expect(isExcluded('Archive', 'Archive', ['archive'], true)).toBe(false);
	});

	it('matches a pattern with "/" against the relative path', () => {
		expect(isExcluded('Old', 'Drafts/Old', ['Drafts/Old'], false)).toBe(true);
		expect(isExcluded('Old', 'Other/Old', ['Drafts/Old'], false)).toBe(false);
		expect(isExcluded('Old', 'Drafts/Old', ['*/Old'], false)).toBe(true);
	});
});

describe('Compare Structure description', () => {
	it('offers the operation with the source, target and options', () => {
		const operation = fileFolderResource.description[0];
		expect(
			(operation.options as Array<{ value: string }>).some(
				(option) => option.value === 'compareStructure',
			),
		).toBe(true);
		expect(fileFolderResource.operations.compareStructure).toBeTypeOf('function');
		const names = fileFolderResource.description
			.filter((entry) =>
				(entry.displayOptions?.show?.operation as string[] | undefined)?.includes(
					'compareStructure',
				),
			)
			.map((entry) => entry.name);
		expect(names).toEqual(['sourceFolder', 'targetFolder', 'compareOptions']);
		const compareOptions = fileFolderResource.description.find(
			(entry) => entry.name === 'compareOptions',
		);
		const defaults = Object.fromEntries(
			(compareOptions?.options as Array<{ name: string; default: unknown }>).map((option) => [
				option.name,
				option.default,
			]),
		);
		expect(defaults).toEqual({
			caseSensitive: false,
			excludeNames: '',
			include: 'folders',
			maxDepth: 1,
		});
	});
});
