import { describe, expect, it } from 'vitest';

import {
	findInvalidIds,
	isInSite,
	isPastDeletedAfter,
	mapItem,
	matchesFilters,
	parseIds,
	type RecycleBinRawItem,
} from '../../nodes/SharePointPnp/resources/recycleBin/helpers';

const ID_A = '11111111-1111-1111-1111-111111111111';
const ID_B = '22222222-2222-2222-2222-222222222222';

const item: RecycleBinRawItem = {
	Id: ID_A,
	Title: 'Budget.xlsx',
	ItemType: 1,
	ItemState: 1,
	DirName: 'sites/Example/Shared Documents/Finance',
	LeafName: 'Budget.xlsx',
	DeletedDate: '2026-03-10T12:00:00Z',
	DeletedByName: 'Jane Doe',
	DeletedByEmail: 'jane.doe@contoso.com',
	AuthorName: 'John Roe',
	AuthorEmail: 'john.roe@contoso.com',
	Size: '2048',
};

describe('recycle bin ID parsing', () => {
	it('parses comma and line separated strings', () => {
		expect(parseIds(` ${ID_A}, ${ID_B}\n${ID_A} `)).toEqual([ID_A, ID_B]);
	});

	it('parses arrays and ignores empty values', () => {
		expect(parseIds([ID_A, ' ', ID_B])).toEqual([ID_A, ID_B]);
		expect(parseIds(undefined)).toEqual([]);
	});

	it('finds values that are not GUIDs', () => {
		expect(findInvalidIds([ID_A, 'abc'])).toEqual(['abc']);
	});
});

describe('recycle bin filters', () => {
	it('matches with no filters and with "any" values', () => {
		expect(matchesFilters(item, {})).toBe(true);
		expect(matchesFilters(item, { itemType: -1, stage: -1 })).toBe(true);
	});

	it('filters by item type and stage', () => {
		expect(matchesFilters(item, { itemType: 1 })).toBe(true);
		expect(matchesFilters(item, { itemType: 5 })).toBe(false);
		expect(matchesFilters(item, { stage: 2 })).toBe(false);
	});

	it('filters by deleted date range', () => {
		expect(matchesFilters(item, { deletedAfter: '2026-03-01T00:00:00Z' })).toBe(true);
		expect(matchesFilters(item, { deletedAfter: '2026-04-01T00:00:00Z' })).toBe(false);
		expect(matchesFilters(item, { deletedBefore: '2026-03-11T00:00:00Z' })).toBe(true);
		expect(matchesFilters(item, { deletedBefore: '2026-03-01T00:00:00Z' })).toBe(false);
	});

	it('filters by deleted by email and by text, ignoring case', () => {
		expect(matchesFilters(item, { deletedBy: 'JANE.DOE' })).toBe(true);
		expect(matchesFilters(item, { deletedBy: 'someone' })).toBe(false);
		expect(matchesFilters(item, { text: 'finance' })).toBe(true);
		expect(matchesFilters(item, { text: 'budget' })).toBe(true);
		expect(matchesFilters(item, { text: 'invoice' })).toBe(false);
	});

	it('detects when paging can stop', () => {
		expect(isPastDeletedAfter(item, '2026-04-01T00:00:00Z')).toBe(true);
		expect(isPastDeletedAfter(item, '2026-03-01T00:00:00Z')).toBe(false);
		expect(isPastDeletedAfter(item, undefined)).toBe(false);
	});
});

describe('recycle bin site scope', () => {
	it('matches items below the site path only', () => {
		expect(isInSite(item, '/sites/Example')).toBe(true);
		expect(isInSite(item, '/sites/Example/')).toBe(true);
		expect(isInSite(item, '/sites/Other')).toBe(false);
		expect(isInSite({ ...item, DirName: 'sites/ExampleTwo/Docs' }, '/sites/Example')).toBe(false);
		expect(isInSite(item, '/')).toBe(true);
	});
});

describe('recycle bin output', () => {
	it('maps an item to readable fields', () => {
		expect(mapItem(item)).toMatchObject({
			id: ID_A,
			title: 'Budget.xlsx',
			itemType: 'File',
			itemState: 'FirstStage',
			dirName: 'sites/Example/Shared Documents/Finance',
			leafName: 'Budget.xlsx',
			deletedDate: '2026-03-10T12:00:00Z',
			deletedByName: 'Jane Doe',
			deletedByEmail: 'jane.doe@contoso.com',
			size: 2048,
			authorName: 'John Roe',
		});
	});
});
