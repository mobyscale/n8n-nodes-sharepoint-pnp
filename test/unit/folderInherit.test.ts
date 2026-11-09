import { describe, expect, it, vi } from 'vitest';
import type { INodeProperties } from 'n8n-workflow';

import {
	findInheritedValues,
	inheritMetadataFromParent,
	parseFieldNames,
	toWritableValue,
	type FolderValueReader,
	type InheritSource,
} from '../../nodes/SharePointPnp/resources/folder/inheritMetadata';
import type { ItemPlan, TargetItem } from '../../nodes/SharePointPnp/resources/folder/setMetadata';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';

const LIB = '/sites/Example/Shared Documents';

const ROOT = LIB;
const FOLDER = `${LIB}/Team A/2024/Q1`;

function reader(values: Record<string, Record<string, unknown>>) {
	return vi.fn<FolderValueReader>(async (path) => values[path]);
}

describe('Inherit source lookup', () => {
	it('reads the field names', () => {
		expect(parseFieldNames('Title, Department\nTitle')).toEqual(['Title', 'Department']);
		expect(() => parseFieldNames(' ')).toThrow('at least one field');
	});

	it('takes the values of the direct parent', async () => {
		const read = reader({
			[`${LIB}/Team A/2024`]: { Title: '2024', Department: '' },
			[`${LIB}/Team A`]: { Department: 'Team A' },
		});
		const result = await findInheritedValues(FOLDER, ROOT, ['Title', 'Department'], false, read);
		expect(read).toHaveBeenCalledTimes(1);
		expect(result).toEqual({
			parentFolder: `${LIB}/Team A/2024`,
			sourceFolder: `${LIB}/Team A/2024`,
			values: { Title: '2024', Department: null },
			valueSources: { Title: `${LIB}/Team A/2024`, Department: null },
			missingFields: ['Department'],
		});
	});

	it('searches up to the nearest folder with a value for each field', async () => {
		const read = reader({
			[`${LIB}/Team A/2024`]: { Title: '2024', Department: null },
			[`${LIB}/Team A`]: { Title: 'Team A', Department: 'Team A' },
		});
		const result = await findInheritedValues(FOLDER, ROOT, ['Title', 'Department'], true, read);
		expect(result.values).toEqual({ Title: '2024', Department: 'Team A' });
		expect(result.valueSources).toEqual({
			Title: `${LIB}/Team A/2024`,
			Department: `${LIB}/Team A`,
		});
		expect(result.sourceFolder).toBe(`${LIB}/Team A/2024`);
		expect(result.missingFields).toEqual([]);
	});

	it('stops when all fields have values', async () => {
		const read = reader({ [`${LIB}/Team A/2024`]: { Title: '2024' } });
		await findInheritedValues(FOLDER, ROOT, ['Title'], true, read);
		expect(read).toHaveBeenCalledTimes(1);
	});

	it('stops below the library root and reports missing values', async () => {
		const read = reader({});
		const result = await findInheritedValues(FOLDER, ROOT, ['Title'], true, read);
		expect(read.mock.calls.map(([path]) => path)).toEqual([`${LIB}/Team A/2024`, `${LIB}/Team A`]);
		expect(result).toMatchObject({
			sourceFolder: null,
			values: { Title: null },
			missingFields: ['Title'],
		});
	});

	it('fails for a library root folder and its direct children', async () => {
		await expect(findInheritedValues(LIB, ROOT, ['Title'], false, reader({}))).rejects.toThrow(
			'is a library root folder',
		);
		await expect(
			findInheritedValues(`${LIB}/Team A`, `${ROOT}/`, ['Title'], false, reader({})),
		).rejects.toThrow('is the library root folder');
	});

	it('converts REST values to writable values', () => {
		const item = {
			Owner: { Name: 'i:0#.f|membership|user@contoso.com' },
			Team: [{ Name: 'i:0#.f|membership|a@contoso.com' }, { Name: '' }],
			DepartmentId: 4,
			TagsId: { results: [1, 2] },
			Link: { Url: 'https://contoso.com', Description: 'Site' },
			Empty: null,
		};
		expect(toWritableValue(item, 'Owner', 'User')).toEqual({
			Key: 'i:0#.f|membership|user@contoso.com',
		});
		expect(toWritableValue(item, 'Team', 'UserMulti')).toEqual([
			{ Key: 'i:0#.f|membership|a@contoso.com' },
		]);
		expect(toWritableValue(item, 'Department', 'Lookup')).toBe(4);
		expect(toWritableValue(item, 'Tags', 'LookupMulti')).toEqual([1, 2]);
		expect(toWritableValue(item, 'Link', 'URL')).toBe('https://contoso.com, Site');
		expect(toWritableValue(item, 'Empty', 'User')).toBeNull();
		expect(toWritableValue(item, 'Missing', 'Text')).toBeNull();
	});
});

const target = (
	id: number,
	type: 'file' | 'folder',
	path: string,
	values: Record<string, unknown> = {},
): TargetItem => ({ id, type, serverRelativeUrl: path, values });

function inheritSource(parentValues: Record<string, unknown>) {
	const written: ItemPlan[] = [];
	const tree: Record<string, TargetItem[]> = {
		[FOLDER]: [target(2, 'folder', `${FOLDER}/Images`), target(3, 'file', `${FOLDER}/a.txt`)],
		[`${FOLDER}/Images`]: [target(4, 'file', `${FOLDER}/Images/b.txt`, { Title: '2024' })],
	};
	const source: InheritSource = {
		listFields: {
			listId: 'list',
			infos: [],
			types: new Map([
				['Title', 'Text'],
				['Category', 'TaxonomyFieldType'],
			]),
		},
		libraryRoot: ROOT,
		readValues: async (path) => (path === `${LIB}/Team A/2024` ? parentValues : undefined),
		readPage: async (path) => ({ rows: tree[path] ?? [] }),
		readSelf: async () => target(1, 'folder', FOLDER),
		readContentTypeFields: async () => new Map(),
		writer: {
			writeChunk: async (chunk) => {
				written.push(...chunk);
				return chunk.map((_, index) => ({ index, ok: true }));
			},
		},
	};
	return { source, written };
}

describe('Inherit Metadata From Parent', () => {
	it('applies the parent values to the folder and all items below it', async () => {
		const { source, written } = inheritSource({ Title: '2024' });
		const result = await inheritMetadataFromParent(FOLDER, ['Title'], {}, source);
		expect(written.map((plan) => [plan.item.id, plan.values])).toEqual([
			[1, { Title: '2024' }],
			[2, { Title: '2024' }],
			[3, { Title: '2024' }],
		]);
		expect(result).toMatchObject({
			folderPath: FOLDER,
			sourceFolder: `${LIB}/Team A/2024`,
			values: { Title: '2024' },
			matched: 4,
			updated: 3,
			skipped: 1,
			failed: 0,
		});
	});

	it('respects Include, Max Depth and Include The Folder Itself', async () => {
		const { source, written } = inheritSource({ Title: '2024' });
		await inheritMetadataFromParent(
			FOLDER,
			['Title'],
			{ include: 'files', maxDepth: 1, includeFolderItself: false, skipUnchanged: false },
			source,
		);
		expect(written.map((plan) => plan.item.id)).toEqual([3]);
	});

	it('does nothing when no value is found and empty values are ignored', async () => {
		const { source, written } = inheritSource({ Title: '' });
		const result = await inheritMetadataFromParent(
			FOLDER,
			['Title'],
			{ ignoreEmptyValues: true },
			source,
		);
		expect(written).toEqual([]);
		expect(result).toMatchObject({ matched: 0, updated: 0, missingFields: ['Title'] });
	});

	it('clears the values when the parent has none and empty values are not ignored', async () => {
		const { source, written } = inheritSource({});
		await inheritMetadataFromParent(FOLDER, ['Title'], {}, source);
		// Items without a value are already equal, so only the item with a value is cleared.
		expect(written.map((plan) => [plan.item.id, plan.values])).toEqual([[4, { Title: null }]]);
	});

	it('fails for unknown and unsupported fields', async () => {
		const { source } = inheritSource({});
		await expect(inheritMetadataFromParent(FOLDER, ['Nope'], {}, source)).rejects.toThrow(
			'do not exist in the library: Nope',
		);
		await expect(inheritMetadataFromParent(FOLDER, ['Category'], {}, source)).rejects.toThrow(
			'cannot be inherited: Category',
		);
	});
});

describe('Folder description for Inherit Metadata From Parent', () => {
	const property = (name: string) =>
		folderResource.description.find((entry) => entry.name === name) as INodeProperties;

	it('lists Inherit Metadata From Parent in order', () => {
		const names = (folderResource.description[0].options as Array<{ name: string }>).map(
			(option) => option.name,
		);
		expect(names).toEqual([...names].sort());
		expect(names).toContain('Inherit Metadata From Parent');
		expect(Object.keys(folderResource.operations)).toContain('inheritMetadataFromParent');
	});

	it('includes the folder itself by default and can search up', () => {
		const options = property('inheritOptions').options as INodeProperties[];
		const names = options.map((option) => option.displayName);
		expect(names).toEqual([...names].sort());
		expect(options.find((option) => option.name === 'includeFolderItself')?.default).toBe(true);
		expect(options.find((option) => option.name === 'searchUp')).toBeDefined();
		const metadata = property('metadataOptions').options as INodeProperties[];
		expect(metadata.find((option) => option.name === 'includeFolderItself')?.default).toBe(false);
		expect(metadata.find((option) => option.name === 'searchUp')).toBeUndefined();
	});
});

describe('Inherit Metadata From Parent field picker', () => {
	it('uses a field picker and keeps the text field of older workflows', () => {
		const picker = folderResource.description.find(
			(entry) => entry.name === 'inheritFieldsPicker',
		)!;
		expect(picker.type).toBe('multiOptions');
		expect(picker.typeOptions?.loadOptionsMethod).toBe('getFolderListFieldOptions');
		expect(picker.typeOptions?.loadOptionsDependsOn).toEqual(['folderPath.value']);
		const text = folderResource.description.find((entry) => entry.name === 'inheritFields')!;
		expect(text.type).toBe('string');
		expect(text.required).toBeFalsy();
	});

	it('merges picked and typed field names', () => {
		expect(parseFieldNames(['Title', 'Status'], 'Category, Title')).toEqual([
			'Title',
			'Status',
			'Category',
		]);
		expect(parseFieldNames([], 'Title\nStatus')).toEqual(['Title', 'Status']);
		expect(() => parseFieldNames([], '')).toThrow('at least one field');
	});
});
