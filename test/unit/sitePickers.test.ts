import type { IDataObject, IExecuteFunctions, ILoadOptionsFunctions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import { getChangeItemFieldOptions } from '../../nodes/SharePointPnp/methods/changeOptions';
import { searchPropertyBagKeys } from '../../nodes/SharePointPnp/methods/listSearchPropertyBag';
import { searchTerms } from '../../nodes/SharePointPnp/methods/listSearchTerms';
import { getRecycleBinItemOptions } from '../../nodes/SharePointPnp/methods/recycleBinOptions';
import {
	getInstalledLanguageOptions,
	getPropertyBagKeyOptions,
	getSiteDesignOptions,
} from '../../nodes/SharePointPnp/methods/siteOptions';
import { getTermStoreLanguageOptions } from '../../nodes/SharePointPnp/methods/taxonomyOptions';
import { recycleBinResource } from '../../nodes/SharePointPnp/resources/recycleBin';
import { siteResource } from '../../nodes/SharePointPnp/resources/site';
import { taxonomyResource } from '../../nodes/SharePointPnp/resources/taxonomy';
import * as taxonomyService from '../../nodes/SharePointPnp/resources/taxonomy/service';
import * as paging from '../../nodes/SharePointPnp/resources/recycleBin/paging';
import { resolveFolderFilter } from '../../nodes/SharePointPnp/shared/changes';
import { getTriggerSettings } from '../../nodes/SharePointPnp/shared/changeTrigger';
import { locatorValue } from '../../nodes/SharePointPnp/shared/pickers';
import * as transport from '../../nodes/SharePointPnp/transport';

vi.mock('../../nodes/SharePointPnp/transport', () => ({
	createSp: vi.fn(),
	getTenantUrl: vi.fn(async () => 'https://contoso.sharepoint.com'),
}));
vi.mock('../../nodes/SharePointPnp/resources/taxonomy/service', async (importOriginal) => ({
	...(await importOriginal<typeof taxonomyService>()),
	getTerms: vi.fn(),
	getTerm: vi.fn(),
	getStoreInfo: vi.fn(),
}));
vi.mock('../../nodes/SharePointPnp/resources/recycleBin/paging', async (importOriginal) => ({
	...(await importOriginal<typeof paging>()),
	fetchRecycleBinPage: vi.fn(),
}));

const SET_ID = '11111111-2222-3333-4444-555555555555';
const TERM_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';

function loadContext(parameters: IDataObject): ILoadOptionsFunctions {
	const all: IDataObject = {
		site: { __rl: true, mode: 'url', value: '/sites/Example' },
		...parameters,
	};
	return {
		getNodeParameter: (name: string, ...rest: unknown[]) => {
			const fallback = typeof rest[0] === 'object' && rest[0] !== null ? undefined : rest[0];
			const value = name in all ? all[name] : (rest[0] ?? fallback);
			if (rest[rest.length - 1] && (rest[rest.length - 1] as IDataObject).extractValue) {
				return (value as IDataObject)?.value ?? value;
			}
			return value;
		},
	} as unknown as ILoadOptionsFunctions;
}

function mockSp(sp: unknown) {
	vi.mocked(transport.createSp).mockReturnValue(sp as ReturnType<typeof transport.createSp>);
}

beforeEach(() => {
	vi.clearAllMocks();
});

describe('locatorValue', () => {
	it('reads locators and plain strings', () => {
		expect(locatorValue({ __rl: true, mode: 'list', value: ' abc ' })).toBe('abc');
		expect(locatorValue(' abc ')).toBe('abc');
		expect(locatorValue(undefined)).toBe('');
		expect(locatorValue({ mode: 'list' })).toBe('');
	});
});

describe('site pickers', () => {
	it('lists built-in and custom site designs', async () => {
		mockSp({
			siteDesigns: {
				getSiteDesigns: async () => [
					{ Id: 'b', Title: 'Zeta', Description: '' },
					{ Id: 'a', Title: 'Alpha', Description: 'First' },
				],
			},
		});
		const options = await getSiteDesignOptions.call(loadContext({}));
		expect(options.map((option) => option.value)).toContain('f6cc5403-0d63-442e-96c0-285923709ffc');
		expect(options.map((option) => option.name)).toEqual([
			'Alpha',
			'Blank (Communication Site)',
			'Showcase (Communication Site)',
			'Topic (Communication Site)',
			'Zeta',
		]);
	});

	it('lists property bag keys without OData members and filters them', async () => {
		mockSp({
			web: {
				allProperties: async () => ({
					'odata.metadata': 'x',
					vti_title: 'a',
					My_x005f_Key: 'b',
				}),
			},
		});
		const keys = await getPropertyBagKeyOptions.call(loadContext({}));
		expect(keys.map((key) => key.value)).toEqual(['My_Key', 'vti_title']);
		const found = await searchPropertyBagKeys.call(loadContext({}), 'VTI');
		expect(found.results).toEqual([{ name: 'vti_title', value: 'vti_title' }]);
	});

	it('lists installed languages with the LCID as value', async () => {
		mockSp({
			web: {
				regionalSettings: {
					getInstalledLanguages: async () => [
						{ DisplayName: 'German', Lcid: 1031, LanguageTag: 'de-DE' },
						{ DisplayName: 'English', Lcid: 1033, LanguageTag: 'en-US' },
					],
				},
			},
		});
		const options = await getInstalledLanguageOptions.call(loadContext({}));
		expect(options).toEqual([
			{ name: 'English (1033)', value: 1033 },
			{ name: 'German (1031)', value: 1031 },
		]);
	});

	it('keeps the old shapes of the changed Site parameters', () => {
		const find = (name: string, operation: string) =>
			siteResource.description.find(
				(property) =>
					property.name === name && property.displayOptions?.show?.operation?.includes(operation),
			);
		expect(find('statusSiteUrl', 'getSiteCreationStatus')?.type).toBe('resourceLocator');
		expect(find('propertyKeys', 'getPropertyBag')?.type).toBe('multiOptions');
		const rows = find('propertyValues', 'setPropertyBag')?.options?.[0] as {
			values: Array<{ name: string; type: string }>;
		};
		expect(rows.values.find((value) => value.name === 'key')?.type).toBe('resourceLocator');
	});

	it('reads property bag keys from a string, a list and a locator row', async () => {
		const run = async (propertyKeys: unknown) => {
			mockSp({ web: { allProperties: async () => ({ A: '1', B: '2', C: '3' }) } });
			const context = {
				getNodeParameter: (name: string) => (name === 'propertyKeys' ? propertyKeys : undefined),
			} as unknown as IExecuteFunctions;
			const sp = transport.createSp(context, '');
			return (await siteResource.operations.getPropertyBag.call(context, {
				sp,
				siteUrl: 'https://contoso.sharepoint.com/sites/Example',
				itemIndex: 0,
			})) as { properties: IDataObject };
		};
		expect(Object.keys((await run('a, b')).properties)).toEqual(['A', 'B']);
		expect(Object.keys((await run(['c'])).properties)).toEqual(['C']);
		expect(Object.keys((await run([])).properties)).toEqual(['A', 'B', 'C']);
	});
});

describe('taxonomy pickers', () => {
	const terms = [
		{ id: 'id-1', name: 'Red', path: 'Colors;Red', labels: [], isDeprecated: false },
		{ id: 'id-2', name: 'Blue', path: 'Colors;Blue', labels: [], isDeprecated: false },
		{ id: 'id-3', name: 'Old', path: 'Old', labels: [], isDeprecated: true },
	] as unknown as taxonomyService.Term[];

	it('searches the terms of the term set and hides deprecated terms', async () => {
		vi.mocked(taxonomyService.getTerms).mockResolvedValue(terms);
		const result = await searchTerms.call(
			loadContext({ operation: 'createTerm', termSet: { value: SET_ID } }),
			're',
		);
		expect(result.results).toEqual([{ name: 'Colors > Red', value: 'id-1' }]);
		expect(vi.mocked(taxonomyService.getTerms).mock.calls[0][2]).toBe(SET_ID);
	});

	it('uses the helper term set for the operations on one term', async () => {
		vi.mocked(taxonomyService.getTerms).mockResolvedValue(terms);
		const none = await searchTerms.call(
			loadContext({ operation: 'getTerm', termSet: { value: SET_ID } }),
		);
		expect(none.results).toEqual([]);
		const result = await searchTerms.call(
			loadContext({ operation: 'getTerm', termSetHelper: { value: SET_ID } }),
		);
		expect(result.results.map((entry) => entry.value)).toEqual(['id-2', 'id-1']);
	});

	it('lists the languages of the term store', async () => {
		vi.mocked(taxonomyService.getStoreInfo).mockResolvedValue({
			id: 's',
			name: 'Store',
			defaultLanguage: 1033,
			languages: [1033, 1031],
		});
		mockSp({
			web: {
				regionalSettings: {
					getInstalledLanguages: async () => [{ DisplayName: 'English', Lcid: 1033 }],
				},
			},
		});
		const options = await getTermStoreLanguageOptions.call(loadContext({}));
		expect(options).toEqual([
			{ name: 'Default Language of the Term Store', value: 0 },
			{ name: 'English (1033)', value: 1033 },
			{ name: 'Language (1031)', value: 1031 },
		]);
	});

	it('reads the term as a plain string, a locator and an expression result', async () => {
		vi.mocked(taxonomyService.getTerm).mockResolvedValue({} as taxonomyService.Term);
		for (const termId of [TERM_ID, { __rl: true, mode: 'list', value: TERM_ID }]) {
			const context = {
				getNodeParameter: () => termId,
				getNode: () => ({ name: 'n' }),
			} as unknown as IExecuteFunctions;
			await taxonomyResource.operations.getTerm.call(context, {
				sp: {} as never,
				siteUrl: 'https://contoso.sharepoint.com/sites/Example',
				itemIndex: 0,
			});
		}
		const calls = vi.mocked(taxonomyService.getTerm).mock.calls;
		expect(calls.map((call) => call[2])).toEqual([TERM_ID, TERM_ID]);
	});

	it('reads the parent term of Get Terms from a locator and from a string', async () => {
		vi.mocked(taxonomyService.getTerms).mockResolvedValue([]);
		for (const parentTermId of [TERM_ID, { __rl: true, mode: 'list', value: TERM_ID }]) {
			const context = {
				getNodeParameter: (name: string) =>
					name === 'getTermsOptions' ? { parentTermId } : name === 'termSet' ? SET_ID : false,
				getNode: () => ({ name: 'n' }),
			} as unknown as IExecuteFunctions;
			await taxonomyResource.operations.getTerms.call(context, {
				sp: {} as never,
				siteUrl: 'https://contoso.sharepoint.com/sites/Example',
				itemIndex: 0,
			});
		}
		const calls = vi.mocked(taxonomyService.getTerms).mock.calls;
		expect(calls.map((call) => call[3]?.parentTermId)).toEqual([TERM_ID, TERM_ID]);
	});
});

describe('recycle bin picker', () => {
	const item = (id: string, dir: string, state: number) => ({
		Id: id,
		Title: `File ${id}`,
		ItemType: 1,
		ItemState: state,
		DirName: dir,
		DeletedDate: '2025-01-01T00:00:00Z',
		DeletedByName: 'Pat',
	});

	beforeEach(() => {
		mockSp({});
		vi.mocked(paging.fetchRecycleBinPage).mockImplementation(async (_sp, stage) => ({
			items:
				stage === 1
					? [item('1', 'sites/Example/Docs', 1), item('2', 'sites/Other/Docs', 1)]
					: [item('3', 'sites/Example/Docs', 2)],
		}));
	});

	it('lists the items of the site collection', async () => {
		const options = await getRecycleBinItemOptions.call(loadContext({ scope: 'siteCollection' }));
		expect(options.map((option) => option.value)).toEqual(['1', '2', '3']);
		expect(options[0].name).toBe('File 1 (sites/Example/Docs, deleted by Pat)');
	});

	it('limits the items to the selected site', async () => {
		const options = await getRecycleBinItemOptions.call(loadContext({ scope: 'site' }));
		expect(options.map((option) => option.value)).toEqual(['1', '3']);
	});

	it('uses a multi-select for the IDs', () => {
		const ids = recycleBinResource.description.filter((property) => property.name === 'ids');
		expect(ids.length).toBe(3);
		expect(ids.every((property) => property.type === 'multiOptions')).toBe(true);
	});
});

describe('change pickers', () => {
	it('returns no field options without a list', async () => {
		expect(
			await getChangeItemFieldOptions.call(loadContext({ list: { mode: 'list', value: '' } })),
		).toEqual([]);
	});

	it('lists the fields of the selected list', async () => {
		const select = vi.fn(async () => [
			{ InternalName: 'Title', Title: 'Title' },
			{ InternalName: 'Amount', Title: 'Amount' },
		]);
		mockSp({
			web: {
				lists: { getById: () => ({ fields: { filter: () => ({ select: () => select }) } }) },
			},
		});
		const options = await getChangeItemFieldOptions.call(
			loadContext({ list: { mode: 'id', value: SET_ID } }),
		);
		expect(options.map((option) => option.value)).toEqual(['Amount', 'Title']);
	});

	it('resolves the folder filter from a string and from a locator', () => {
		const site = 'https://contoso.sharepoint.com/sites/Example';
		expect(resolveFolderFilter(site, 'Shared Documents/Folder')).toBe(
			'/sites/Example/Shared Documents/Folder',
		);
		expect(
			resolveFolderFilter(site, { mode: 'list', value: '/sites/Example/Shared Documents/Folder' }),
		).toBe('/sites/Example/Shared Documents/Folder');
		expect(resolveFolderFilter(site, { mode: 'path', value: '' })).toBeUndefined();
		expect(resolveFolderFilter(site, undefined)).toBeUndefined();
	});

	it('reads the item fields of the trigger as a list and as a string', () => {
		const settings = (itemFields: unknown, folderPathFilter: unknown) =>
			getTriggerSettings(
				{
					getNode: () => ({}) as never,
					getNodeParameter: (name: string, fallback?: unknown) =>
						({
							events: ['fileCreated'],
							options: { itemFields },
							folderPathFilter,
						})[name] ?? fallback,
				},
				'https://contoso.sharepoint.com/sites/Example',
			);
		expect(settings(['Title'], { mode: 'path', value: 'Docs/A' }).runOptions.itemFields).toEqual([
			'Title',
		]);
		const old = settings('Title,Editor/Title', 'Docs/A');
		expect(old.runOptions.itemFields).toBe('Title,Editor/Title');
		expect(old.folderPathFilter).toBe('/sites/Example/Docs/A');
	});
});
