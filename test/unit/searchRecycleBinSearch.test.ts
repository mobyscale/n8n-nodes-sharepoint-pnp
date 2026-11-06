import { describe, expect, it } from 'vitest';

import {
	buildQueryText,
	clampRowLimit,
	flattenCells,
	nextPageSize,
	parseList,
	toSortList,
} from '../../nodes/SharePointPnp/resources/search/helpers';

describe('search helpers', () => {
	it('returns the query text unchanged without a scope', () => {
		expect(buildQueryText('  contentclass:STS_Site ')).toBe('contentclass:STS_Site');
	});

	it('adds a site scope and removes a trailing slash from the site URL', () => {
		expect(buildQueryText('budget', 'https://contoso.sharepoint.com/sites/Example/')).toBe(
			'(budget) Path:"https://contoso.sharepoint.com/sites/Example/*"',
		);
	});

	it('uses only the scope when the query is empty', () => {
		expect(buildQueryText('', 'https://contoso.sharepoint.com')).toBe(
			'Path:"https://contoso.sharepoint.com/*"',
		);
	});

	it('parses lists from strings and arrays', () => {
		expect(parseList(' Title, Path ,,Author')).toEqual(['Title', 'Path', 'Author']);
		expect(parseList(['A', ' B '])).toEqual(['A', 'B']);
		expect(parseList(undefined)).toEqual([]);
	});

	it('maps sort directions and skips empty properties', () => {
		expect(
			toSortList([
				{ property: ' Size ', direction: 'descending' },
				{ property: 'Title', direction: 'ascending' },
				{ property: ' ', direction: 'ascending' },
			]),
		).toEqual([
			{ Property: 'Size', Direction: 1 },
			{ Property: 'Title', Direction: 0 },
		]);
	});

	it('flattens cells into an object', () => {
		expect(
			flattenCells({
				Cells: [
					{ Key: 'Title', Value: 'Report', ValueType: 'Edm.String' },
					{ Key: 'Size', Value: '12', ValueType: 'Edm.Int64' },
				],
			}),
		).toEqual({ Title: 'Report', Size: '12' });
		expect(flattenCells({ Cells: undefined as never })).toEqual({});
	});

	it('limits the page size to 500', () => {
		expect(clampRowLimit(900)).toBe(500);
		expect(clampRowLimit(25)).toBe(25);
		expect(clampRowLimit(0)).toBe(100);
		expect(clampRowLimit(undefined)).toBe(100);
	});

	it('plans pages until the total or the wanted count is reached', () => {
		expect(nextPageSize(100, 0, Infinity, undefined)).toBe(100);
		expect(nextPageSize(100, 100, 150, undefined)).toBe(100);
		expect(nextPageSize(100, 150, 150, undefined)).toBe(0);
		expect(nextPageSize(100, 100, 1000, 130)).toBe(30);
		expect(nextPageSize(100, 130, 1000, 130)).toBe(0);
	});
});
