import { escapeXml } from './helpers';

export type FieldType =
	| 'text'
	| 'note'
	| 'number'
	| 'currency'
	| 'boolean'
	| 'dateTime'
	| 'choice'
	| 'multiChoice'
	| 'user'
	| 'multiUser'
	| 'lookup'
	| 'url'
	| 'calculated'
	| 'schemaXml';

export interface CommonFieldSettings {
	displayName: string;
	internalName?: string;
	required?: boolean;
	indexed?: boolean;
	defaultValue?: string;
	description?: string;
	group?: string;
}

export interface FieldTypeSettings {
	maxLength?: number;
	numberOfLines?: number;
	richText?: boolean;
	appendOnly?: boolean;
	minimumValue?: number;
	maximumValue?: number;
	currencyLocaleId?: number;
	dateTimeFormat?: 'DateOnly' | 'DateTime';
	choices?: string;
	choiceFormat?: 'Dropdown' | 'RadioButtons';
	fillInChoice?: boolean;
	userSelectionMode?: 'PeopleOnly' | 'PeopleAndGroups';
	lookupListId?: string;
	lookupField?: string;
	lookupAllowMultiple?: boolean;
	urlFormat?: 'Hyperlink' | 'Image';
	formula?: string;
	calculatedOutputType?: 'Text' | 'Number' | 'Currency' | 'DateTime' | 'Boolean';
	schemaXml?: string;
}

export type FieldAddMethod =
	| 'addText'
	| 'addMultilineText'
	| 'addNumber'
	| 'addCurrency'
	| 'addBoolean'
	| 'addDateTime'
	| 'addChoice'
	| 'addMultiChoice'
	| 'addUser'
	| 'addUrl'
	| 'addCalculated'
	| 'createFieldAsXml';

export interface FieldPlan {
	method: FieldAddMethod;
	/** The title used to create the field. SharePoint derives the internal name from it. */
	title: string;
	/** Set when the display name must be written after the creation, so that the internal name stays as given. */
	renameTo?: string;
	properties?: Record<string, unknown>;
	schemaXml?: string;
}

const INTERNAL_NAME_PATTERN = /^[A-Za-z_][A-Za-z0-9_]*$/;
const DATE_FORMAT = { DateOnly: 0, DateTime: 1 } as const;
const USER_SELECTION = { PeopleOnly: 0, PeopleAndGroups: 1 } as const;
const URL_FORMAT = { Hyperlink: 0, Image: 1 } as const;
const CHOICE_FORMAT = { Dropdown: 0, RadioButtons: 1 } as const;
const OUTPUT_TYPE = { Text: 2, DateTime: 4, Boolean: 8, Number: 9, Currency: 10 } as const;

export function parseChoices(value: string | undefined): string[] {
	return [
		...new Set(
			String(value ?? '')
				.split(/\r?\n/)
				.map((choice) => choice.trim())
				.filter(Boolean),
		),
	];
}

/** SharePoint stores a Yes/No default as 1 or 0. */
export function toBooleanDefault(value: string): string {
	return /^(1|true|yes|on)$/i.test(value.trim()) ? '1' : '0';
}

export function buildLookupSchemaXml(
	common: CommonFieldSettings,
	settings: FieldTypeSettings,
): string {
	const name = common.internalName?.trim() || common.displayName.replace(/[^A-Za-z0-9_]/g, '');
	const listId = settings.lookupListId?.trim().replace(/^\{|\}$/g, '') ?? '';
	if (!listId) throw new Error('A lookup field needs a source list');
	const multiple = settings.lookupAllowMultiple === true;
	const attributes: Array<[string, string]> = [
		['Type', multiple ? 'LookupMulti' : 'Lookup'],
		['DisplayName', common.displayName],
		['Name', name],
		['StaticName', name],
		['List', `{${listId}}`],
		['ShowField', settings.lookupField?.trim() || 'Title'],
		['Required', common.required ? 'TRUE' : 'FALSE'],
	];
	if (multiple) attributes.push(['Mult', 'TRUE']);
	if (common.indexed) attributes.push(['Indexed', 'TRUE']);
	if (common.group) attributes.push(['Group', common.group]);
	if (common.description) attributes.push(['Description', common.description]);
	return `<Field ${attributes.map(([key, value]) => `${key}="${escapeXml(value)}"`).join(' ')} />`;
}

/** Maps the node parameters of "Create Field" to the PnPjs add method and its properties. */
export function buildFieldPlan(
	type: FieldType,
	common: CommonFieldSettings,
	settings: FieldTypeSettings = {},
): FieldPlan {
	if (type === 'schemaXml') {
		const xml = settings.schemaXml?.trim();
		if (!xml) throw new Error('The schema XML must not be empty');
		return { method: 'createFieldAsXml', title: common.displayName, schemaXml: xml };
	}

	const displayName = common.displayName.trim();
	if (!displayName) throw new Error('The display name must not be empty');
	const internalName = common.internalName?.trim();
	if (internalName && !INTERNAL_NAME_PATTERN.test(internalName)) {
		throw new Error(
			'The internal name can only have letters, digits and underscores, and must not start with a digit',
		);
	}

	if (type === 'lookup') {
		return {
			method: 'createFieldAsXml',
			title: displayName,
			schemaXml: buildLookupSchemaXml({ ...common, displayName }, settings),
		};
	}

	const properties: Record<string, unknown> = {};
	if (common.required) properties.Required = true;
	if (common.indexed) properties.Indexed = true;
	if (common.description) properties.Description = common.description;
	if (common.group) properties.Group = common.group;
	if (common.defaultValue !== undefined && common.defaultValue !== '' && type !== 'calculated') {
		properties.DefaultValue =
			type === 'boolean' ? toBooleanDefault(common.defaultValue) : common.defaultValue;
	}

	const plan = (method: FieldAddMethod, extra: Record<string, unknown> = {}): FieldPlan => ({
		method,
		title: internalName ?? displayName,
		renameTo: internalName && internalName !== displayName ? displayName : undefined,
		properties: { ...properties, ...extra },
	});

	switch (type) {
		case 'text':
			return plan('addText', { MaxLength: settings.maxLength ?? 255 });
		case 'note':
			return plan('addMultilineText', {
				NumberOfLines: settings.numberOfLines ?? 6,
				RichText: settings.richText === true,
				AppendOnly: settings.appendOnly === true,
			});
		case 'number':
			return plan('addNumber', rangeProperties(settings));
		case 'currency':
			return plan('addCurrency', {
				...rangeProperties(settings),
				CurrencyLocaleId: settings.currencyLocaleId ?? 1033,
			});
		case 'boolean':
			return plan('addBoolean');
		case 'dateTime':
			return plan('addDateTime', { DisplayFormat: DATE_FORMAT[settings.dateTimeFormat ?? 'DateOnly'] });
		case 'choice':
		case 'multiChoice': {
			const choices = parseChoices(settings.choices);
			if (!choices.length) throw new Error('Add at least one choice');
			const extra: Record<string, unknown> = {
				Choices: choices,
				FillInChoice: settings.fillInChoice === true,
			};
			if (type === 'choice') {
				extra.EditFormat = CHOICE_FORMAT[settings.choiceFormat ?? 'Dropdown'];
			}
			return plan(type === 'choice' ? 'addChoice' : 'addMultiChoice', extra);
		}
		case 'user':
		case 'multiUser':
			return plan('addUser', {
				SelectionMode: USER_SELECTION[settings.userSelectionMode ?? 'PeopleAndGroups'],
				...(type === 'multiUser' ? { AllowMultipleValues: true } : {}),
			});
		case 'url':
			return plan('addUrl', { DisplayFormat: URL_FORMAT[settings.urlFormat ?? 'Hyperlink'] });
		case 'calculated': {
			const formula = settings.formula?.trim();
			if (!formula) throw new Error('A calculated field needs a formula');
			return plan('addCalculated', {
				Formula: formula,
				OutputType: OUTPUT_TYPE[settings.calculatedOutputType ?? 'Text'],
			});
		}
	}
}

function rangeProperties(settings: FieldTypeSettings): Record<string, unknown> {
	const result: Record<string, unknown> = {};
	if (settings.minimumValue !== undefined) result.MinimumValue = settings.minimumValue;
	if (settings.maximumValue !== undefined) result.MaximumValue = settings.maximumValue;
	return result;
}

export interface FieldUpdateOptions {
	displayName?: string;
	description?: string;
	required?: boolean;
	indexed?: boolean;
	defaultValue?: string;
	group?: string;
	hidden?: boolean;
	enforceUniqueValues?: boolean;
	validationFormula?: string;
	validationMessage?: string;
	choices?: string;
	otherProperties?: string;
}

/** Maps the "Update Field" options to the properties of the SharePoint field. Only set options are sent. */
export function buildFieldUpdate(options: FieldUpdateOptions): Record<string, unknown> {
	const properties: Record<string, unknown> = {};
	const map: Array<[keyof FieldUpdateOptions, string]> = [
		['displayName', 'Title'],
		['description', 'Description'],
		['required', 'Required'],
		['indexed', 'Indexed'],
		['defaultValue', 'DefaultValue'],
		['group', 'Group'],
		['hidden', 'Hidden'],
		['enforceUniqueValues', 'EnforceUniqueValues'],
		['validationFormula', 'ValidationFormula'],
		['validationMessage', 'ValidationMessage'],
	];
	for (const [option, property] of map) {
		if (options[option] !== undefined) properties[property] = options[option];
	}
	if (options.choices !== undefined && options.choices.trim()) {
		properties.Choices = { results: parseChoices(options.choices) };
	}
	if (options.otherProperties?.trim()) {
		const other = JSON.parse(options.otherProperties) as unknown;
		if (typeof other !== 'object' || other === null || Array.isArray(other)) {
			throw new Error('"Other Properties (JSON)" must be a JSON object');
		}
		Object.assign(properties, other);
	}
	return properties;
}

/** The REST type name that SharePoint needs to update a field of this kind. */
export function restTypeForFieldKind(kind: number): string {
	const types: Record<number, string> = {
		2: 'SP.FieldText',
		3: 'SP.FieldMultiLineText',
		4: 'SP.FieldDateTime',
		6: 'SP.FieldChoice',
		7: 'SP.FieldLookup',
		9: 'SP.FieldNumber',
		10: 'SP.FieldCurrency',
		11: 'SP.FieldUrl',
		15: 'SP.FieldMultiChoice',
		17: 'SP.FieldCalculated',
		20: 'SP.FieldUser',
	};
	return types[kind] ?? 'SP.Field';
}

export interface ListUpdateOptions {
	title?: string;
	description?: string;
	enableVersioning?: boolean;
	majorVersionLimit?: number;
	enableMinorVersions?: boolean;
	minorVersionLimit?: number;
	enableModeration?: boolean;
	forceCheckout?: boolean;
	enableAttachments?: boolean;
	onQuickLaunch?: boolean;
	hidden?: boolean;
	noCrawl?: boolean;
	enableFolderCreation?: boolean;
	contentTypesEnabled?: boolean;
	draftVersionVisibility?: number;
}

/** Maps the "Update" list options to list properties. Only set options are sent. */
export function buildListUpdate(options: ListUpdateOptions): Record<string, unknown> {
	const map: Array<[keyof ListUpdateOptions, string]> = [
		['title', 'Title'],
		['description', 'Description'],
		['enableVersioning', 'EnableVersioning'],
		['majorVersionLimit', 'MajorVersionLimit'],
		['enableMinorVersions', 'EnableMinorVersions'],
		['minorVersionLimit', 'MajorWithMinorVersionsLimit'],
		['enableModeration', 'EnableModeration'],
		['forceCheckout', 'ForceCheckout'],
		['enableAttachments', 'EnableAttachments'],
		['onQuickLaunch', 'OnQuickLaunch'],
		['hidden', 'Hidden'],
		['noCrawl', 'NoCrawl'],
		['enableFolderCreation', 'EnableFolderCreation'],
		['contentTypesEnabled', 'ContentTypesEnabled'],
		['draftVersionVisibility', 'DraftVersionVisibility'],
	];
	const properties: Record<string, unknown> = {};
	for (const [option, property] of map) {
		if (options[option] !== undefined) properties[property] = options[option];
	}
	return properties;
}

/** Puts the chosen content type first. The order of the others stays as it is. */
export function buildContentTypeOrder(current: string[], defaultId: string): string[] {
	return [defaultId, ...current.filter((id) => id.toLowerCase() !== defaultId.toLowerCase())];
}
