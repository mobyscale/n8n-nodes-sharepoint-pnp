import { describe, expect, it, vi } from 'vitest';

import {
	EditorDirectory,
	getDepthBelow,
	getEditorEmail,
	isEditorExcluded,
	isWithinDepth,
	matchesFieldFilter,
	matchesFieldFilters,
	parseEditorList,
	parseFieldFilters,
	parseItemKinds,
	parseMaxDepth,
	stripClaims,
	type EditorSource,
	type FieldFilter,
} from '../../nodes/SharePointPnp/shared/changeFilters';

const filter = (field: string, operator: FieldFilter['operator'], value = ''): FieldFilter => ({
	field,
	operator,
	value,
});

describe('parse helpers', () => {
	it('reads the field filter rows and ignores rows without a field', () => {
		expect(
			parseFieldFilters({
				filter: [
					{ field: ' Status ', operator: 'contains', value: 'op' },
					{ field: '', operator: 'equals', value: 'x' },
					{ field: 'Category', operator: 'unknown', value: 'A' },
				],
			}),
		).toEqual([filter('Status', 'contains', 'op'), filter('Category', 'equals', 'A')]);
		expect(parseFieldFilters(undefined)).toEqual([]);
		expect(parseFieldFilters({})).toEqual([]);
	});

	it('reads item kinds, editors and max depth', () => {
		expect(parseItemKinds(['file', 'Folder', 'list', 'file'])).toEqual(['file', 'folder']);
		expect(parseItemKinds('item, folder')).toEqual(['item', 'folder']);
		expect(parseEditorList(' Person A, a@contoso.com ,,')).toEqual(['person a', 'a@contoso.com']);
		expect(parseMaxDepth(2.7)).toBe(2);
		expect(parseMaxDepth(0)).toBeUndefined();
		expect(parseMaxDepth('')).toBeUndefined();
		expect(parseMaxDepth(undefined)).toBeUndefined();
	});
});

describe('matchesFieldFilter', () => {
	const values = {
		Title: 'Quarterly Report',
		Category: 'A',
		Status: null,
		Count: 5,
		Done: false,
		Tags: ['Red', 'Blue'],
		Choices: { results: ['One', 'Two'] },
		'Editor/Title': 'Person A',
		Author: { Title: 'Person B' },
		Due: '2024-03-01T00:00:00Z',
	};

	it.each([
		[filter('Category', 'equals', 'a'), true],
		[filter('Category', 'equals', 'B'), false],
		[filter('Category', 'notEquals', 'B'), true],
		[filter('Title', 'contains', 'report'), true],
		[filter('Title', 'contains', 'summary'), false],
		[filter('Status', 'isEmpty'), true],
		[filter('Status', 'isNotEmpty'), false],
		[filter('Missing', 'isEmpty'), true],
		[filter('Status', 'equals', ''), true],
		[filter('Count', 'equals', '5.0'), true],
		[filter('Done', 'equals', 'false'), true],
		[filter('Tags', 'equals', 'blue'), true],
		[filter('Tags', 'notEquals', 'blue'), false],
		[filter('Choices', 'contains', 'tw'), true],
		[filter('Editor/Title', 'equals', 'person a'), true],
		[filter('Author/Title', 'equals', 'Person B'), true],
		[filter('Author', 'equals', 'Person B'), true],
		[filter('Due', 'equals', '2024-03-01'), true],
	])('%o -> %s', (entry, expected) => {
		expect(matchesFieldFilter(values, entry)).toBe(expected);
	});

	it('needs all filters and field values', () => {
		const filters = [filter('Category', 'equals', 'A'), filter('Status', 'isEmpty')];
		expect(matchesFieldFilters(values, filters)).toBe(true);
		expect(matchesFieldFilters(values, [...filters, filter('Count', 'equals', '6')])).toBe(false);
		expect(matchesFieldFilters(undefined, filters)).toBe(false);
		expect(matchesFieldFilters(undefined, [])).toBe(true);
	});
});

describe('depth', () => {
	const root = '/sites/Example/Shared Documents';

	it('counts the levels below the root', () => {
		expect(getDepthBelow(`${root}/a.txt`, root)).toBe(1);
		expect(getDepthBelow(`${root}/Reports/2024/Q1/b.txt`, `${root}/`)).toBe(4);
		expect(getDepthBelow(root.toUpperCase(), root)).toBe(0);
		expect(getDepthBelow('/sites/Other/a.txt', root)).toBeUndefined();
		expect(getDepthBelow('/a/b', '/')).toBe(2);
	});

	it('keeps changes without a path or outside the root', () => {
		expect(isWithinDepth(`${root}/Reports/a.txt`, root, 1)).toBe(false);
		expect(isWithinDepth(`${root}/Reports/a.txt`, root, 2)).toBe(true);
		expect(isWithinDepth(undefined, root, 1)).toBe(true);
		expect(isWithinDepth('/sites/Other/x/y/z', root, 1)).toBe(true);
		expect(isWithinDepth(`${root}/a/b`, undefined, 1)).toBe(true);
	});
});

describe('editors', () => {
	const source = (users: Record<string, object[]>): EditorSource => ({
		findByDisplayName: vi.fn(async (name: string) => users[name] ?? []),
	});

	it('strips the claims prefix', () => {
		expect(stripClaims('i:0#.f|membership|a@contoso.com')).toBe('a@contoso.com');
		expect(stripClaims('plain')).toBe('plain');
	});

	it('resolves a display name once and only when one user has it', async () => {
		const users = source({
			'Person A': [{ Title: 'Person A', Email: 'a@contoso.com', LoginName: 'i:0#.f|membership|a' }],
			Twin: [
				{ Title: 'Twin', Email: 'x@contoso.com' },
				{ Title: 'Twin', Email: 'y@contoso.com' },
			],
		});
		const directory = new EditorDirectory(users);
		expect(await directory.resolve('Person A')).toEqual({
			email: 'a@contoso.com',
			login: 'i:0#.f|membership|a',
			userPrincipalName: undefined,
		});
		expect(await directory.resolve('person a')).toBeDefined();
		expect(await directory.resolve('Twin')).toBeUndefined();
		expect(await directory.resolve('')).toBeUndefined();
		expect(users.findByDisplayName).toHaveBeenCalledTimes(2);
	});

	it('ignores lookup errors', async () => {
		const directory = new EditorDirectory({
			findByDisplayName: async () => {
				throw new Error('denied');
			},
		});
		expect(await directory.resolve('Person A')).toBeUndefined();
	});

	it('excludes by name or hint without a lookup, and by login with a lookup', async () => {
		const users = source({
			'Person C': [{ Title: 'Person C', LoginName: 'i:0#.f|membership|c@contoso.com' }],
		});
		const directory = new EditorDirectory(users);
		expect(await isEditorExcluded({ Editor: 'Robot' }, ['robot'], directory)).toBe(true);
		expect(
			await isEditorExcluded(
				{ Editor: 'X', EditorEmailHint: 'B@contoso.com' },
				['b@contoso.com'],
				directory,
			),
		).toBe(true);
		expect(await isEditorExcluded({ Editor: 'Person C' }, ['person d'], directory)).toBe(false);
		expect(users.findByDisplayName).not.toHaveBeenCalled();
		expect(await isEditorExcluded({ Editor: 'Person C' }, ['c@contoso.com'], directory)).toBe(true);
		expect(await isEditorExcluded({ Editor: 'Person C' }, [], directory)).toBe(false);
	});

	it('returns the email hint first, then the resolved email', async () => {
		const directory = new EditorDirectory(
			source({ 'Person D': [{ Title: 'Person D', Email: 'd@contoso.com' }] }),
		);
		expect(await getEditorEmail({ EditorEmailHint: 'h@contoso.com' }, directory)).toBe(
			'h@contoso.com',
		);
		expect(await getEditorEmail({ Editor: 'Person D', EditorEmailHint: null }, directory)).toBe(
			'd@contoso.com',
		);
		expect(await getEditorEmail({ Editor: 'Nobody' }, directory)).toBeUndefined();
	});
});
