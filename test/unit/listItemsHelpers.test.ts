import { describe, expect, it } from 'vitest';

import {
	buildFormValues,
	buildViewXml,
	escapeODataString,
	expandFromSelect,
	formatFieldValue,
	mergeMultiValues,
	parseCommaList,
} from '../../nodes/SharePointPnp/resources/listItem/helpers';
import { toMapperType } from '../../nodes/SharePointPnp/methods/resourceMapping';

describe('formatFieldValue', () => {
	it('formats booleans', () => {
		expect(formatFieldValue('Boolean', true)).toBe('1');
		expect(formatFieldValue('Boolean', false)).toBe('0');
		expect(formatFieldValue('Boolean', 'Yes')).toBe('1');
		expect(formatFieldValue('Boolean', 'no')).toBe('0');
	});

	it('formats multi-choice values', () => {
		expect(formatFieldValue('MultiChoice', ['A', 'B'])).toBe(';#A;#B;#');
		expect(formatFieldValue('MultiChoice', 'A, B')).toBe(';#A;#B;#');
		expect(formatFieldValue('MultiChoice', [])).toBe('');
	});

	it('formats lookups', () => {
		expect(formatFieldValue('Lookup', 5)).toBe('5');
		expect(formatFieldValue('Lookup', { Id: 7 })).toBe('7');
		expect(formatFieldValue('LookupMulti', [1, { Id: 2 }, '3'])).toBe('1;#;#2;#;#3');
	});

	it('formats users', () => {
		expect(formatFieldValue('User', 'i:0#.f|membership|a@contoso.com')).toBe(
			'[{"Key":"i:0#.f|membership|a@contoso.com"}]',
		);
		expect(formatFieldValue('UserMulti', ['a@contoso.com', { Email: 'b@contoso.com' }])).toBe(
			'[{"Key":"a@contoso.com"},{"Key":"b@contoso.com"}]',
		);
	});

	it('formats dates, text, numbers and empty values', () => {
		expect(formatFieldValue('DateTime', new Date('2024-01-02T03:04:05.000Z'))).toBe(
			'2024-01-02T03:04:05.000Z',
		);
		expect(formatFieldValue('Text', 12)).toBe('12');
		expect(formatFieldValue(undefined, { a: 1 })).toBe('{"a":1}');
		expect(formatFieldValue('Text', null)).toBe('');
		expect(formatFieldValue('Text', undefined)).toBeNull();
	});
});

describe('buildFormValues', () => {
	it('uses the field type to format each value', () => {
		const result = buildFormValues(
			{ Title: 'Example', Done: true, Tags: ['x', 'y'], Skipped: undefined },
			[
				{ InternalName: 'Title', TypeAsString: 'Text' },
				{ InternalName: 'Done', TypeAsString: 'Boolean' },
				{ InternalName: 'Tags', TypeAsString: 'MultiChoice' },
			],
		);
		expect(result).toEqual([
			{ FieldName: 'Title', FieldValue: 'Example' },
			{ FieldName: 'Done', FieldValue: '1' },
			{ FieldName: 'Tags', FieldValue: ';#x;#y;#' },
		]);
	});
});

describe('mergeMultiValues', () => {
	it('keeps current values and appends new ones without duplicates', () => {
		expect(mergeMultiValues(['A', 'B'], ['B', 'C'])).toEqual(['A', 'B', 'C']);
	});

	it('merges lookup ids and objects', () => {
		expect(mergeMultiValues([1, 2], [{ Id: 2 }, 3])).toEqual([1, 2, 3]);
	});

	it('handles empty current values and text input', () => {
		expect(mergeMultiValues(null, 'A, B')).toEqual(['A', 'B']);
		expect(mergeMultiValues(undefined, [])).toEqual([]);
	});
});

describe('buildViewXml', () => {
	it('wraps a where clause in a query', () => {
		expect(buildViewXml('<Where><Eq/></Where>')).toBe(
			'<View><Query><Where><Eq/></Where></Query></View>',
		);
	});

	it('adds scope and row limit', () => {
		expect(buildViewXml('<Where/>', { recursive: true, rowLimit: 100 })).toBe(
			'<View Scope="RecursiveAll"><Query><Where/></Query><RowLimit Paged="TRUE">100</RowLimit></View>',
		);
	});

	it('does not wrap a Query element again', () => {
		expect(buildViewXml('<Query><Where/></Query>')).toBe('<View><Query><Where/></Query></View>');
	});

	it('changes a full view only when needed', () => {
		const view = '<View><Query/><RowLimit>5</RowLimit></View>';
		expect(buildViewXml(view, { rowLimit: 100 })).toBe(view);
		expect(buildViewXml('<View><Query/></View>', { recursive: true, rowLimit: 10 })).toBe(
			'<View Scope="RecursiveAll"><Query/><RowLimit Paged="TRUE">10</RowLimit></View>',
		);
		const scoped = '<View Scope="FilesOnly"><Query/></View>';
		expect(buildViewXml(scoped, { recursive: true })).toBe(scoped);
	});
});

describe('small helpers', () => {
	it('expands navigation properties from a select list', () => {
		expect(expandFromSelect(['Id', 'Author/Title', 'Author/EMail', 'Editor/Title'])).toEqual([
			'Author',
			'Editor',
		]);
	});

	it('escapes OData strings', () => {
		expect(escapeODataString("O'Brien")).toBe("O''Brien");
	});

	it('parses comma lists', () => {
		expect(parseCommaList(' Id, Title ,,')).toEqual(['Id', 'Title']);
		expect(parseCommaList(undefined)).toEqual([]);
	});

	it('maps SharePoint field types to mapper types', () => {
		expect(toMapperType('Currency')).toBe('number');
		expect(toMapperType('Choice')).toBe('options');
		expect(toMapperType('MultiChoice')).toBe('array');
		expect(toMapperType('Note')).toBe('string');
	});
});
