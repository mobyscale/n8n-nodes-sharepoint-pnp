import type { INodeProperties } from 'n8n-workflow';
import type { IList } from '@pnp/sp/lists/index.js';
import { describe, expect, it, vi } from 'vitest';

import {
	addCamlCondition,
	buildCamlCondition,
	buildCamlFilter,
	buildODataCondition,
	buildODataFilter,
	fieldFiltersCombineOption,
	fieldFiltersProperty,
	loadFilterFieldTypes,
	readCombine,
	readFieldFilters,
	type FieldFilter,
	type FieldFilterOperator,
} from '../../nodes/SharePointPnp/shared/fieldFilters';
import { buildViewXml } from '../../nodes/SharePointPnp/shared/filesFolders';

const filter = (field: string, operator: FieldFilterOperator, value = ''): FieldFilter => ({
	field,
	operator,
	value,
});

describe('field filter description', () => {
	const property = fieldFiltersProperty({
		loadOptionsMethod: 'getListFieldOptions',
		loadOptionsDependsOn: ['list.value'],
		displayOptions: { show: { resource: ['listItem'] } },
		description: 'Conditions',
	});
	const values = (property.options as Array<{ values: INodeProperties[] }>)[0].values;

	it('has a field picker, all operators and a value that Is Empty hides', () => {
		expect(property.type).toBe('fixedCollection');
		expect(property.typeOptions?.multipleValues).toBe(true);
		const field = values.find((entry) => entry.name === 'field')!;
		expect(field.typeOptions?.loadOptionsMethod).toBe('getListFieldOptions');
		expect(field.typeOptions?.loadOptionsDependsOn).toEqual(['list.value']);
		const operator = values.find((entry) => entry.name === 'operator')!;
		expect((operator.options as Array<{ name: string }>).map((entry) => entry.name)).toEqual([
			'Equals',
			'Not Equals',
			'Contains',
			'Begins With',
			'Greater Than',
			'Less Than',
			'Is Empty',
			'Is Not Empty',
		]);
		const value = values.find((entry) => entry.name === 'value')!;
		expect(value.displayOptions?.hide?.operator).toEqual(['isEmpty', 'isNotEmpty']);
	});

	it('combines with AND by default', () => {
		expect(fieldFiltersCombineOption.default).toBe('and');
		expect(readCombine(undefined)).toBe('and');
		expect(readCombine('or')).toBe('or');
		expect(readCombine('x')).toBe('and');
	});
});

describe('reading field filters', () => {
	it('reads the rows of the parameter', () => {
		expect(readFieldFilters({})).toEqual([]);
		expect(readFieldFilters(undefined)).toEqual([]);
		expect(
			readFieldFilters({
				conditions: [
					{ field: ' Title ', operator: 'eq', value: ' Draft ' },
					{ field: 'Status', operator: 'isEmpty' },
					{ field: 'Count', operator: 'gt', value: 3 },
				],
			}),
		).toEqual([
			filter('Title', 'eq', 'Draft'),
			filter('Status', 'isEmpty'),
			filter('Count', 'gt', '3'),
		]);
	});

	it('rejects rows without a field, a value or a valid operator', () => {
		expect(() =>
			readFieldFilters({ conditions: [{ field: '', operator: 'eq', value: 'a' }] }),
		).toThrow('Choose a field');
		expect(() => readFieldFilters({ conditions: [{ field: 'Title', operator: 'eq' }] })).toThrow(
			'Give a value',
		);
		expect(() =>
			readFieldFilters({ conditions: [{ field: 'Title', operator: 'like', value: 'a' }] }),
		).toThrow('not a valid field filter operator');
	});
});

describe('CAML field filter builder', () => {
	it('builds every operator for a text field', () => {
		const ref = '<FieldRef Name="Title"/>';
		const value = '<Value Type="Text">Draft</Value>';
		expect(buildCamlCondition(filter('Title', 'eq', 'Draft'), 'Text')).toBe(
			`<Eq>${ref}${value}</Eq>`,
		);
		expect(buildCamlCondition(filter('Title', 'ne', 'Draft'), 'Text')).toBe(
			`<Neq>${ref}${value}</Neq>`,
		);
		expect(buildCamlCondition(filter('Title', 'contains', 'Draft'), 'Text')).toBe(
			`<Contains>${ref}${value}</Contains>`,
		);
		expect(buildCamlCondition(filter('Title', 'beginsWith', 'Draft'), 'Text')).toBe(
			`<BeginsWith>${ref}${value}</BeginsWith>`,
		);
		expect(buildCamlCondition(filter('Title', 'gt', 'Draft'), 'Text')).toBe(
			`<Gt>${ref}${value}</Gt>`,
		);
		expect(buildCamlCondition(filter('Title', 'lt', 'Draft'), 'Text')).toBe(
			`<Lt>${ref}${value}</Lt>`,
		);
		expect(buildCamlCondition(filter('Title', 'isEmpty'), 'Text')).toBe(`<IsNull>${ref}</IsNull>`);
		expect(buildCamlCondition(filter('Title', 'isNotEmpty'), 'Text')).toBe(
			`<IsNotNull>${ref}</IsNotNull>`,
		);
	});

	it('uses the value type of each field type', () => {
		const value = (type: string | undefined, input: string) =>
			buildCamlCondition(filter('F', 'eq', input), type).replace(/^<Eq>|<\/Eq>$/g, '');
		expect(value('Note', 'a')).toBe('<FieldRef Name="F"/><Value Type="Note">a</Value>');
		expect(value('Choice', 'Final')).toBe('<FieldRef Name="F"/><Value Type="Choice">Final</Value>');
		expect(value('MultiChoice', 'Final')).toBe(
			'<FieldRef Name="F"/><Value Type="MultiChoice">Final</Value>',
		);
		expect(value('Number', '1.5')).toBe('<FieldRef Name="F"/><Value Type="Number">1.5</Value>');
		expect(value('Currency', '10')).toBe('<FieldRef Name="F"/><Value Type="Currency">10</Value>');
		expect(value('Integer', '7')).toBe('<FieldRef Name="F"/><Value Type="Integer">7</Value>');
		expect(value('Counter', '7')).toBe('<FieldRef Name="F"/><Value Type="Counter">7</Value>');
		expect(value('Boolean', 'true')).toBe('<FieldRef Name="F"/><Value Type="Boolean">1</Value>');
		expect(value('Boolean', 'No')).toBe('<FieldRef Name="F"/><Value Type="Boolean">0</Value>');
		expect(value('Lookup', '4')).toBe(
			'<FieldRef Name="F" LookupId="TRUE"/><Value Type="Lookup">4</Value>',
		);
		expect(value('LookupMulti', '4')).toBe(
			'<FieldRef Name="F" LookupId="TRUE"/><Value Type="Lookup">4</Value>',
		);
		expect(value('User', '12')).toBe(
			'<FieldRef Name="F" LookupId="TRUE"/><Value Type="Integer">12</Value>',
		);
		expect(value('UserMulti', '12')).toBe(
			'<FieldRef Name="F" LookupId="TRUE"/><Value Type="Integer">12</Value>',
		);
		expect(value('DateTime', '2024-01-31')).toBe(
			'<FieldRef Name="F"/><Value Type="DateTime">2024-01-31</Value>',
		);
		expect(value('DateTime', '2024-01-31T10:00:00Z')).toBe(
			'<FieldRef Name="F"/><Value Type="DateTime" IncludeTimeValue="TRUE">2024-01-31T10:00:00.000Z</Value>',
		);
		expect(value(undefined, 'a')).toBe('<FieldRef Name="F"/><Value Type="Text">a</Value>');
		expect(value('Computed', 'a')).toBe('<FieldRef Name="F"/><Value Type="Text">a</Value>');
	});

	it('escapes XML in names and values', () => {
		expect(buildCamlCondition(filter('A"B', 'eq', `<a & 'b'>`), 'Text')).toBe(
			'<Eq><FieldRef Name="A&quot;B"/><Value Type="Text">&lt;a &amp; &apos;b&apos;&gt;</Value></Eq>',
		);
	});

	it('rejects values that do not match the field type', () => {
		expect(() => buildCamlCondition(filter('F', 'eq', 'x'), 'Number')).toThrow('not a number');
		expect(() => buildCamlCondition(filter('F', 'eq', 'maybe'), 'Boolean')).toThrow(
			'not a yes/no value',
		);
		expect(() => buildCamlCondition(filter('F', 'eq', 'Team A'), 'Lookup')).toThrow('not an ID');
		expect(() => buildCamlCondition(filter('F', 'eq', 'someone'), 'User')).toThrow('not an ID');
		expect(() => buildCamlCondition(filter('F', 'eq', 'soon'), 'DateTime')).toThrow(
			'not a valid date',
		);
		expect(() => buildCamlCondition(filter('F', 'contains', '1'), 'Number')).toThrow(
			'works only with text fields',
		);
		expect(() => buildCamlCondition(filter('F', 'beginsWith', '1'), 'Lookup')).toThrow(
			'works only with text fields',
		);
	});

	it('does not check values for Is Empty and Is Not Empty', () => {
		expect(buildCamlCondition(filter('F', 'isEmpty'), 'Lookup')).toBe(
			'<IsNull><FieldRef Name="F"/></IsNull>',
		);
		expect(buildCamlCondition(filter('F', 'isNotEmpty'), 'Number')).toBe(
			'<IsNotNull><FieldRef Name="F"/></IsNotNull>',
		);
	});

	it('combines many filters with And or Or', () => {
		const types = new Map([['Count', 'Number']]);
		const filters = [
			filter('Title', 'eq', 'a'),
			filter('Status', 'isEmpty'),
			filter('Count', 'gt', '2'),
		];
		expect(buildCamlFilter([], types)).toBeUndefined();
		expect(buildCamlFilter([filters[0]], types)).toBe(
			'<Eq><FieldRef Name="Title"/><Value Type="Text">a</Value></Eq>',
		);
		const and = buildCamlFilter(filters, types)!;
		expect(and.startsWith('<And><And><Eq>')).toBe(true);
		expect(and).toContain('<Value Type="Number">2</Value>');
		const or = buildCamlFilter(filters, types, 'or')!;
		expect(or.startsWith('<Or><Or><Eq>')).toBe(true);
		expect(or).not.toContain('<And>');
	});

	it('adds the condition to the folder view of File/Folder Get Many', () => {
		const condition = buildCamlFilter([filter('Title', 'eq', 'a')], new Map())!;
		const xml = buildViewXml({ include: 'files', fieldCondition: condition });
		expect(xml).toContain(
			'<Where><And><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq><Eq><FieldRef Name="Title"/>',
		);
		expect(buildViewXml({ include: 'all', fieldCondition: condition })).toContain(
			`<Where>${condition}</Where>`,
		);
	});
});

describe('OData field filter builder', () => {
	it('builds every operator for a text field', () => {
		expect(buildODataCondition(filter('Title', 'eq', 'Draft'), 'Text')).toBe("Title eq 'Draft'");
		expect(buildODataCondition(filter('Title', 'ne', 'Draft'), 'Text')).toBe("Title ne 'Draft'");
		expect(buildODataCondition(filter('Title', 'contains', 'Dra'), 'Text')).toBe(
			"substringof('Dra',Title)",
		);
		expect(buildODataCondition(filter('Title', 'beginsWith', 'Dra'), 'Text')).toBe(
			"startswith(Title,'Dra')",
		);
		expect(buildODataCondition(filter('Title', 'gt', 'B'), 'Text')).toBe("Title gt 'B'");
		expect(buildODataCondition(filter('Title', 'lt', 'B'), 'Text')).toBe("Title lt 'B'");
		expect(buildODataCondition(filter('Title', 'isEmpty'), 'Text')).toBe('Title eq null');
		expect(buildODataCondition(filter('Title', 'isNotEmpty'), 'Text')).toBe('Title ne null');
	});

	it('uses the literal of each field type', () => {
		const odata = (type: string | undefined, input: string) =>
			buildODataCondition(filter('F', 'eq', input), type);
		expect(odata('Note', 'a')).toBe("F eq 'a'");
		expect(odata('Choice', 'Final')).toBe("F eq 'Final'");
		expect(odata('MultiChoice', 'Final')).toBe("F eq 'Final'");
		expect(odata('Number', '1.50')).toBe('F eq 1.5');
		expect(odata('Currency', '10')).toBe('F eq 10');
		expect(odata('Integer', '7')).toBe('F eq 7');
		expect(odata('Counter', '7')).toBe('F eq 7');
		expect(odata('Boolean', 'yes')).toBe('F eq 1');
		expect(odata('Boolean', '0')).toBe('F eq 0');
		expect(odata('Lookup', '4')).toBe('FId eq 4');
		expect(odata('User', '12')).toBe('FId eq 12');
		expect(odata('DateTime', '2024-01-31')).toBe("F eq datetime'2024-01-31T00:00:00Z'");
		expect(odata('DateTime', '2024-01-31T10:00:00Z')).toBe(
			"F eq datetime'2024-01-31T10:00:00.000Z'",
		);
		expect(odata(undefined, 'a')).toBe("F eq 'a'");
		expect(buildODataCondition(filter('F', 'isEmpty'), 'Lookup')).toBe('FId eq null');
		expect(buildODataCondition(filter('F', 'isNotEmpty'), 'User')).toBe('FId ne null');
	});

	it('escapes quotes and rejects field names that are not internal names', () => {
		expect(buildODataCondition(filter('Title', 'eq', "it's"), 'Text')).toBe("Title eq 'it''s'");
		expect(() => buildODataCondition(filter('Title eq 1 or Id', 'eq', 'a'), 'Text')).toThrow(
			'not a valid internal field name',
		);
		expect(() => buildODataCondition(filter('F', 'contains', 'a'), 'DateTime')).toThrow(
			'works only with text fields',
		);
	});

	it('combines many filters with and or or', () => {
		const filters = [filter('Title', 'eq', 'a'), filter('Status', 'isNotEmpty')];
		expect(buildODataFilter([], new Map())).toBeUndefined();
		expect(buildODataFilter([filters[0]], new Map())).toBe("Title eq 'a'");
		expect(buildODataFilter(filters, new Map())).toBe("(Title eq 'a') and (Status ne null)");
		expect(buildODataFilter(filters, new Map(), 'or')).toBe("(Title eq 'a') or (Status ne null)");
	});
});

describe('adding a condition to a CAML query', () => {
	const condition = '<IsNull><FieldRef Name="Status"/></IsNull>';

	it('adds a Where element when there is none', () => {
		expect(addCamlCondition('', condition)).toBe(`<Where>${condition}</Where>`);
		expect(addCamlCondition('<OrderBy><FieldRef Name="ID"/></OrderBy>', condition)).toBe(
			`<Where>${condition}</Where><OrderBy><FieldRef Name="ID"/></OrderBy>`,
		);
		expect(addCamlCondition('<Query><OrderBy/></Query>', condition)).toBe(
			`<Query><Where>${condition}</Where><OrderBy/></Query>`,
		);
		expect(addCamlCondition('<View Scope="All"><RowLimit>5</RowLimit></View>', condition)).toBe(
			`<View Scope="All"><Query><Where>${condition}</Where></Query><RowLimit>5</RowLimit></View>`,
		);
	});

	it('joins an existing Where element with And', () => {
		const where = '<Eq><FieldRef Name="Title"/><Value Type="Text">a</Value></Eq>';
		expect(addCamlCondition(`<Where>${where}</Where>`, condition)).toBe(
			`<Where><And>${where}${condition}</And></Where>`,
		);
		expect(
			addCamlCondition(
				`<View><Query><Where> ${where} </Where><OrderBy/></Query></View>`,
				condition,
			),
		).toBe(`<View><Query><Where><And>${where}${condition}</And></Where><OrderBy/></Query></View>`);
	});
});

describe('loading the field types of the filters', () => {
	const list = (fields: Array<{ InternalName: string; TypeAsString: string }>) =>
		({
			fields: { select: vi.fn().mockReturnValue(vi.fn().mockResolvedValue(fields)) },
		}) as unknown as IList;

	it('maps internal names to types', async () => {
		const types = await loadFilterFieldTypes(
			list([
				{ InternalName: 'Title', TypeAsString: 'Text' },
				{ InternalName: 'Count', TypeAsString: 'Number' },
			]),
			[filter('Count', 'gt', '1')],
		);
		expect(types.get('Count')).toBe('Number');
	});

	it('names the fields that do not exist', async () => {
		await expect(
			loadFilterFieldTypes(list([{ InternalName: 'Title', TypeAsString: 'Text' }]), [
				filter('Category', 'eq', 'a'),
				filter('Status', 'isEmpty'),
			]),
		).rejects.toThrow('do not exist in the list: Category, Status');
	});
});
