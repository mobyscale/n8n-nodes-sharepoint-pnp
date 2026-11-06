import { describe, expect, it } from 'vitest';

import {
	buildContentTypeOrder,
	buildFieldPlan,
	buildFieldUpdate,
	buildListUpdate,
	buildLookupSchemaXml,
	parseChoices,
	restTypeForFieldKind,
	toBooleanDefault,
} from '../../nodes/SharePointPnp/resources/list/setupHelpers';

const common = { displayName: 'Project Code' };

describe('buildFieldPlan', () => {
	it('maps a text field and keeps the common settings', () => {
		expect(
			buildFieldPlan(
				'text',
				{
					...common,
					required: true,
					indexed: true,
					defaultValue: 'X',
					description: 'About',
					group: 'Example Group',
				},
				{ maxLength: 50 },
			),
		).toEqual({
			method: 'addText',
			title: 'Project Code',
			renameTo: undefined,
			properties: {
				Required: true,
				Indexed: true,
				Description: 'About',
				Group: 'Example Group',
				DefaultValue: 'X',
				MaxLength: 50,
			},
		});
	});

	it('creates with the internal name and renames to the display name afterward', () => {
		const plan = buildFieldPlan('text', { ...common, internalName: 'ProjectCode' });
		expect(plan.title).toBe('ProjectCode');
		expect(plan.renameTo).toBe('Project Code');
	});

	it('does not rename when both names are equal', () => {
		const plan = buildFieldPlan('text', { displayName: 'Code', internalName: 'Code' });
		expect(plan.renameTo).toBeUndefined();
	});

	it('rejects an invalid internal name', () => {
		expect(() => buildFieldPlan('text', { ...common, internalName: 'Project Code' })).toThrow();
		expect(() => buildFieldPlan('text', { ...common, internalName: '1Code' })).toThrow();
	});

	it('rejects an empty display name', () => {
		expect(() => buildFieldPlan('text', { displayName: ' ' })).toThrow();
	});

	it('maps note settings', () => {
		const plan = buildFieldPlan('note', common, { numberOfLines: 10, richText: true });
		expect(plan.method).toBe('addMultilineText');
		expect(plan.properties).toEqual({ NumberOfLines: 10, RichText: true, AppendOnly: false });
	});

	it('maps number and currency ranges', () => {
		expect(buildFieldPlan('number', common, { minimumValue: 1, maximumValue: 9 })).toMatchObject({
			method: 'addNumber',
			properties: { MinimumValue: 1, MaximumValue: 9 },
		});
		expect(buildFieldPlan('currency', common, { currencyLocaleId: 2057 })).toMatchObject({
			method: 'addCurrency',
			properties: { CurrencyLocaleId: 2057 },
		});
	});

	it('maps a Yes/No default to 1 or 0', () => {
		expect(buildFieldPlan('boolean', { ...common, defaultValue: 'true' }).properties).toEqual({
			DefaultValue: '1',
		});
		expect(buildFieldPlan('boolean', { ...common, defaultValue: 'no' }).properties).toEqual({
			DefaultValue: '0',
		});
		expect(toBooleanDefault(' Yes ')).toBe('1');
	});

	it('maps date and time formats', () => {
		expect(buildFieldPlan('dateTime', common, { dateTimeFormat: 'DateTime' })).toMatchObject({
			method: 'addDateTime',
			properties: { DisplayFormat: 1 },
		});
		expect(buildFieldPlan('dateTime', common).properties).toMatchObject({ DisplayFormat: 0 });
	});

	it('maps choice and multi choice with unique non-empty choices', () => {
		const choice = buildFieldPlan('choice', common, {
			choices: 'A\r\n B \n\nA\nC',
			choiceFormat: 'RadioButtons',
			fillInChoice: true,
		});
		expect(choice.method).toBe('addChoice');
		expect(choice.properties).toEqual({
			Choices: ['A', 'B', 'C'],
			FillInChoice: true,
			EditFormat: 1,
		});

		const multi = buildFieldPlan('multiChoice', common, { choices: 'A\nB' });
		expect(multi.method).toBe('addMultiChoice');
		expect(multi.properties).toEqual({ Choices: ['A', 'B'], FillInChoice: false });
	});

	it('rejects a choice field without choices', () => {
		expect(() => buildFieldPlan('choice', common, { choices: ' \n ' })).toThrow();
	});

	it('maps user and multi user fields', () => {
		expect(buildFieldPlan('user', common, { userSelectionMode: 'PeopleOnly' })).toMatchObject({
			method: 'addUser',
			properties: { SelectionMode: 0 },
		});
		const multi = buildFieldPlan('multiUser', common);
		expect(multi.properties).toEqual({ SelectionMode: 1, AllowMultipleValues: true });
	});

	it('maps url fields', () => {
		expect(buildFieldPlan('url', common, { urlFormat: 'Image' })).toMatchObject({
			method: 'addUrl',
			properties: { DisplayFormat: 1 },
		});
	});

	it('maps calculated fields and ignores the default value', () => {
		const plan = buildFieldPlan(
			'calculated',
			{ ...common, defaultValue: 'x' },
			{ formula: '=[Title]', calculatedOutputType: 'Number' },
		);
		expect(plan.method).toBe('addCalculated');
		expect(plan.properties).toEqual({ Formula: '=[Title]', OutputType: 9 });
		expect(() => buildFieldPlan('calculated', common, {})).toThrow();
	});

	it('builds a lookup field as schema XML', () => {
		const plan = buildFieldPlan(
			'lookup',
			{ ...common, internalName: 'ProjectRef', required: true },
			{ lookupListId: '{11111111-1111-1111-1111-111111111111}', lookupField: 'Title' },
		);
		expect(plan.method).toBe('createFieldAsXml');
		expect(plan.schemaXml).toBe(
			'<Field Type="Lookup" DisplayName="Project Code" Name="ProjectRef" StaticName="ProjectRef" List="{11111111-1111-1111-1111-111111111111}" ShowField="Title" Required="TRUE" />',
		);
	});

	it('uses LookupMulti for multiple values and escapes XML', () => {
		const xml = buildLookupSchemaXml(
			{ displayName: 'A & B', group: 'G"1' },
			{ lookupListId: '11111111-1111-1111-1111-111111111111', lookupAllowMultiple: true },
		);
		expect(xml).toContain('Type="LookupMulti"');
		expect(xml).toContain('Mult="TRUE"');
		expect(xml).toContain('DisplayName="A &amp; B"');
		expect(xml).toContain('Group="G&quot;1"');
		expect(xml).toContain('Name="AB"');
	});

	it('rejects a lookup field without a list', () => {
		expect(() => buildFieldPlan('lookup', common, {})).toThrow();
	});

	it('passes schema XML through', () => {
		const plan = buildFieldPlan('schemaXml', { displayName: '' }, { schemaXml: ' <Field /> ' });
		expect(plan).toEqual({ method: 'createFieldAsXml', title: '', schemaXml: '<Field />' });
		expect(() => buildFieldPlan('schemaXml', { displayName: '' }, { schemaXml: ' ' })).toThrow();
	});
});

describe('parseChoices', () => {
	it('handles empty input', () => {
		expect(parseChoices(undefined)).toEqual([]);
	});
});

describe('buildFieldUpdate', () => {
	it('maps only the set options', () => {
		expect(
			buildFieldUpdate({
				displayName: 'New',
				required: false,
				indexed: true,
				enforceUniqueValues: true,
				validationFormula: '=[A]>1',
			}),
		).toEqual({
			Title: 'New',
			Required: false,
			Indexed: true,
			EnforceUniqueValues: true,
			ValidationFormula: '=[A]>1',
		});
		expect(buildFieldUpdate({})).toEqual({});
	});

	it('wraps choices and merges other properties', () => {
		expect(buildFieldUpdate({ choices: 'A\nB', otherProperties: '{"MaxLength": 10}' })).toEqual({
			Choices: { results: ['A', 'B'] },
			MaxLength: 10,
		});
	});

	it('rejects other properties that are not an object', () => {
		expect(() => buildFieldUpdate({ otherProperties: '[1]' })).toThrow();
		expect(() => buildFieldUpdate({ otherProperties: 'nope' })).toThrow();
	});

	it('maps a field kind to its REST type', () => {
		expect(restTypeForFieldKind(6)).toBe('SP.FieldChoice');
		expect(restTypeForFieldKind(20)).toBe('SP.FieldUser');
		expect(restTypeForFieldKind(8)).toBe('SP.Field');
	});
});

describe('buildListUpdate', () => {
	it('maps all list settings to SharePoint properties', () => {
		expect(
			buildListUpdate({
				title: 'T',
				description: 'D',
				enableVersioning: true,
				majorVersionLimit: 50,
				enableMinorVersions: true,
				minorVersionLimit: 5,
				enableModeration: true,
				forceCheckout: true,
				enableAttachments: false,
				onQuickLaunch: true,
				hidden: false,
				noCrawl: true,
				enableFolderCreation: true,
				contentTypesEnabled: true,
				draftVersionVisibility: 2,
			}),
		).toEqual({
			Title: 'T',
			Description: 'D',
			EnableVersioning: true,
			MajorVersionLimit: 50,
			EnableMinorVersions: true,
			MajorWithMinorVersionsLimit: 5,
			EnableModeration: true,
			ForceCheckout: true,
			EnableAttachments: false,
			OnQuickLaunch: true,
			Hidden: false,
			NoCrawl: true,
			EnableFolderCreation: true,
			ContentTypesEnabled: true,
			DraftVersionVisibility: 2,
		});
	});

	it('sends nothing for options that are not set', () => {
		expect(buildListUpdate({})).toEqual({});
	});
});

describe('buildContentTypeOrder', () => {
	it('puts the default first and keeps the order of the others', () => {
		expect(buildContentTypeOrder(['0x01', '0x0101', '0x0102'], '0x0102')).toEqual([
			'0x0102',
			'0x01',
			'0x0101',
		]);
	});

	it('adds the default when it is missing and does not duplicate it by case', () => {
		expect(buildContentTypeOrder(['0xAB', '0xCD'], '0xab')).toEqual(['0xab', '0xCD']);
		expect(buildContentTypeOrder([], '0x01')).toEqual(['0x01']);
	});
});
