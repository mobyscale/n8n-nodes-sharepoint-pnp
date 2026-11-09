import { describe, expect, it } from 'vitest';

import {
	assertFieldResults,
	buildKeyFilter,
	buildKeyViewXml,
	clampBatchSize,
	comparisonSelect,
	isSameValue,
	isUnchanged,
	newItemId,
	parseListItemId,
	shouldSetKeyOnCreate,
	systemUpdateValues,
	toODataName,
} from '../../nodes/SharePointPnp/resources/listItem/bulkHelpers';

describe('bulk helpers', () => {
	it('limits the batch size to 1 to 100', () => {
		expect(clampBatchSize(undefined)).toBe(100);
		expect(clampBatchSize(0)).toBe(100);
		expect(clampBatchSize(5)).toBe(5);
		expect(clampBatchSize(7.8)).toBe(7);
		expect(clampBatchSize(500)).toBe(100);
	});

	it('parses list item IDs', () => {
		expect(parseListItemId('12')).toBe(12);
		expect(parseListItemId(3)).toBe(3);
		expect(() => parseListItemId('')).toThrow('not a valid list item ID');
		expect(() => parseListItemId('1.5')).toThrow('not a valid list item ID');
		expect(() => parseListItemId('abc')).toThrow('not a valid list item ID');
		expect(() => parseListItemId(0)).toThrow('not a valid list item ID');
	});

	it('adds the OData prefix to names that start with an underscore', () => {
		expect(toODataName('_Status')).toBe('OData__Status');
		expect(toODataName('Title')).toBe('Title');
	});
});

describe('buildKeyFilter', () => {
	it('quotes and escapes text keys', () => {
		expect(buildKeyFilter('Title', 'Text', "O'Brien & Co")).toBe("Title eq 'O''Brien & Co'");
		expect(buildKeyFilter('FileLeafRef', 'File', 'a.txt')).toBe("FileLeafRef eq 'a.txt'");
		expect(buildKeyFilter('Status', 'Choice', 'Open')).toBe("Status eq 'Open'");
		expect(() => buildKeyFilter('Title', 'Text', ' ')).toThrow('must not be empty');
	});

	it('uses numbers, booleans and lookup IDs without quotes', () => {
		expect(buildKeyFilter('Amount', 'Number', '12.5')).toBe('Amount eq 12.5');
		expect(buildKeyFilter('ID', 'Counter', 7)).toBe('ID eq 7');
		expect(buildKeyFilter('Done', 'Boolean', 'yes')).toBe('Done eq 1');
		expect(buildKeyFilter('Done', 'Boolean', false)).toBe('Done eq 0');
		expect(buildKeyFilter('Department', 'Lookup', '4')).toBe('DepartmentId eq 4');
		expect(buildKeyFilter('Owner', 'User', 9)).toBe('OwnerId eq 9');
		expect(buildKeyFilter('_Code', 'Text', 'x')).toBe("OData__Code eq 'x'");
	});

	it('builds a CAML query with escaped values and lookup IDs', () => {
		expect(buildKeyViewXml('Title', 'Text', 'A & <B> "C"', 2)).toBe(
			'<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">A &amp; &lt;B&gt; &quot;C&quot;</Value></Eq></Where><OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy></Query><RowLimit>2</RowLimit></View>',
		);
		expect(buildKeyViewXml('Amount', 'Number', '7', 1)).toContain(
			'<FieldRef Name="Amount"/><Value Type="Number">7</Value>',
		);
		expect(buildKeyViewXml('Done', 'Boolean', true, 1)).toContain(
			'<Value Type="Boolean">1</Value>',
		);
		expect(buildKeyViewXml('Department', 'Lookup', 4, 1)).toContain(
			'<FieldRef Name="Department" LookupId="TRUE"/><Value Type="Lookup">4</Value>',
		);
		expect(buildKeyViewXml('Owner', 'User', 9, 1)).toContain(
			'<FieldRef Name="Owner" LookupId="TRUE"/><Value Type="Integer">9</Value>',
		);
		expect(() => buildKeyViewXml('Notes', 'Note', 'x', 1)).toThrow('cannot be used as a key');
	});

	it('rejects invalid values and unsupported types', () => {
		expect(() => buildKeyFilter('Amount', 'Number', 'abc')).toThrow('is not a number');
		expect(() => buildKeyFilter('Owner', 'User', 'a@contoso.com')).toThrow('is not a number');
		expect(() => buildKeyFilter('Department', 'Lookup', 1.5)).toThrow('numeric item or user ID');
		expect(() => buildKeyFilter('Notes', 'Note', 'x')).toThrow('cannot be used as a key');
	});

	it('sets the key on new items only for writable types', () => {
		expect(shouldSetKeyOnCreate('Text')).toBe(true);
		expect(shouldSetKeyOnCreate('Lookup')).toBe(true);
		expect(shouldSetKeyOnCreate('Counter')).toBe(false);
		expect(shouldSetKeyOnCreate('User')).toBe(false);
		expect(shouldSetKeyOnCreate('File')).toBe(false);
	});
});

describe('change detection', () => {
	it('selects lookup IDs and skips types it cannot compare', () => {
		expect(comparisonSelect('Title', 'Text')).toBe('Title');
		expect(comparisonSelect('Department', 'Lookup')).toBe('DepartmentId');
		expect(comparisonSelect('Tags', 'LookupMulti')).toBe('TagsId');
		expect(comparisonSelect('Owner', 'User')).toBeUndefined();
		expect(comparisonSelect('Link', 'URL')).toBeUndefined();
		expect(comparisonSelect('Unknown', undefined)).toBeUndefined();
	});

	it('compares values by field type', () => {
		expect(isSameValue('Text', 'a', 'a')).toBe(true);
		expect(isSameValue('Text', '', null)).toBe(true);
		expect(isSameValue('Text', 'a', 'b')).toBe(false);
		expect(isSameValue('Number', '5', 5)).toBe(true);
		expect(isSameValue('Number', '', null)).toBe(true);
		expect(isSameValue('Number', 0, null)).toBe(false);
		expect(isSameValue('Boolean', 'true', true)).toBe(true);
		expect(isSameValue('Boolean', false, true)).toBe(false);
		expect(isSameValue('DateTime', '2024-01-02T03:04:05Z', '2024-01-02T03:04:05.000Z')).toBe(true);
		expect(isSameValue('DateTime', 'not a date', '2024-01-02T03:04:05Z')).toBe(false);
		expect(isSameValue('Lookup', { Id: 3 }, 3)).toBe(true);
		expect(isSameValue('Lookup', null, null)).toBe(true);
		expect(isSameValue('MultiChoice', ['B', 'A'], ['A', 'B'])).toBe(true);
		expect(isSameValue('MultiChoice', 'A, B', { results: ['A', 'B'] })).toBe(true);
		expect(isSameValue('MultiChoice', ['A'], ['A', 'B'])).toBe(false);
		expect(isSameValue('LookupMulti', [1, { Id: 2 }], [2, 1])).toBe(true);
		expect(isSameValue('User', 'a@contoso.com', 'a@contoso.com')).toBe(false);
	});

	it('reports an item as unchanged only when every field is equal', () => {
		const types = new Map([
			['Title', 'Text'],
			['Department', 'Lookup'],
			['Owner', 'User'],
		]);
		expect(isUnchanged({ Title: 'a', Department: 2 }, types, { Title: 'a', DepartmentId: 2 })).toBe(
			true,
		);
		expect(isUnchanged({ Title: 'a', Department: 3 }, types, { Title: 'a', DepartmentId: 2 })).toBe(
			false,
		);
		expect(isUnchanged({ Owner: 'a@contoso.com' }, types, {})).toBe(false);
		expect(isUnchanged({}, types, {})).toBe(true);
	});
});

describe('system update values', () => {
	it('sends Modified and a user Editor back unless they are given', () => {
		const extra = systemUpdateValues(
			{ Editor: { Name: 'i:0#.f|membership|a@contoso.com' } },
			'2024-01-02T03:04:05.123',
			[{ FieldName: 'Title', FieldValue: 'x' }],
		);
		expect(extra).toEqual([
			{ FieldName: 'Modified', FieldValue: '2024-01-02 03:04:05' },
			{ FieldName: 'Editor', FieldValue: '[{"Key":"i:0#.f|membership|a@contoso.com"}]' },
		]);
		expect(
			systemUpdateValues({ Editor: { Name: 'i:0i.t|ms.sp.ext|app' } }, undefined, [
				{ FieldName: 'Modified', FieldValue: 'x' },
			]),
		).toEqual([]);
	});
});

describe('field results', () => {
	it('throws for rejected values and reads the new ID', () => {
		expect(() =>
			assertFieldResults([{ FieldName: 'Title', HasException: true, ErrorMessage: 'Too long' }]),
		).toThrow('Title: Too long');
		expect(newItemId([{ FieldName: 'Id', FieldValue: '42', HasException: false }])).toBe(42);
		expect(() => newItemId([{ FieldName: 'Title', FieldValue: 'x' }])).toThrow(
			'did not return the ID',
		);
	});
});
