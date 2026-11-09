import { describe, expect, it } from 'vitest';

import { toServerRelativePath } from '../../nodes/SharePointPnp/shared/utils';
import { config, hasConfig, locator, runNode, itWrite } from './harness';

/**
 * Only read operations run in this file. Every operation that adds, changes or deletes site, list,
 * field or content type settings is an "it.skip" test. Enable them only on a test site that you can
 * change and reset, and replace the names below with names that do not exist yet.
 */
describe.runIf(hasConfig)('site and list setup (read only)', () => {
	let libraryId: string | undefined;

	const getLibraryId = async (): Promise<string> => {
		if (libraryId) return libraryId;
		const libraryRoot = toServerRelativePath(config!.siteUrl, config!.libraryPath).toLowerCase();
		const lists = await runNode({
			resource: 'list',
			operation: 'getAll',
			returnAll: true,
			filters: { baseTemplate: 101 },
		});
		const match = lists.find(
			(item) =>
				String(
					(item.json.RootFolder as { ServerRelativeUrl?: string } | undefined)?.ServerRelativeUrl ??
						'',
				).toLowerCase() === libraryRoot,
		);
		if (!match) throw new Error('The test library was not found');
		libraryId = String(match.json.Id);
		return libraryId;
	};

	it('gets the site with the site collection information', async () => {
		const [result] = await runNode({ resource: 'site', operation: 'get' });
		expect(result.json).toHaveProperty('Title');
		expect(result.json).toHaveProperty('Url');
		const siteCollection = result.json.SiteCollection as Record<string, unknown>;
		expect(siteCollection).toHaveProperty('Id');
		expect(siteCollection).toHaveProperty('IsHubSite');
	});

	it('gets the site without the site collection when it is switched off', async () => {
		const [result] = await runNode({
			resource: 'site',
			operation: 'get',
			getOptions: { includeSiteCollection: false },
		});
		expect(result.json).toHaveProperty('Title');
		expect(result.json).not.toHaveProperty('SiteCollection');
	});

	it('gets the subsites', async () => {
		const items = await runNode({
			resource: 'site',
			operation: 'getSubsites',
			returnAll: true,
		});
		for (const item of items) expect(item.json).toHaveProperty('Url');
	});

	it('gets all property bag values', async () => {
		const [result] = await runNode({ resource: 'site', operation: 'getPropertyBag' });
		const properties = result.json.properties as Record<string, unknown>;
		expect(typeof properties).toBe('object');
		expect(Object.keys(properties).length).toBeGreaterThan(0);
		expect(properties).not.toHaveProperty('odata.metadata');
	});

	it('gets given property bag keys and reports missing keys', async () => {
		const [all] = await runNode({ resource: 'site', operation: 'getPropertyBag' });
		const existing = Object.keys(all.json.properties as Record<string, unknown>)[0];
		const [result] = await runNode({
			resource: 'site',
			operation: 'getPropertyBag',
			propertyKeys: `${existing}, __pnp_missing_key`,
		});
		expect(Object.keys(result.json.properties as object)).toEqual([existing]);
		expect(result.json.missingKeys).toEqual(['__pnp_missing_key']);
	});

	it('gets the hub sites', async () => {
		const items = await runNode({ resource: 'site', operation: 'getHubSites', returnAll: true });
		for (const item of items) {
			expect(item.json).toHaveProperty('SiteId');
			expect(item.json).toHaveProperty('Title');
		}
	});

	it('gets the creation status of the existing test site', async () => {
		const [result] = await runNode({
			resource: 'site',
			operation: 'getSiteCreationStatus',
			statusSiteUrl: config!.siteUrl,
		});
		expect(result.json.SiteStatusName).toBe('Ready');
	});

	it('gets the content types of the site', async () => {
		const items = await runNode({
			resource: 'list',
			operation: 'getContentTypes',
			target: 'site',
		});
		expect(items.length).toBeGreaterThan(0);
		expect(items[0].json).toHaveProperty('StringId');
		expect(items[0].json).toHaveProperty('Name');
	});

	it('gets the content types of the test library', async () => {
		const items = await runNode({
			resource: 'list',
			operation: 'getContentTypes',
			target: 'list',
			list: locator(await getLibraryId(), 'id'),
		});
		expect(items.length).toBeGreaterThan(0);
		expect(items.map((item) => item.json.Name)).toContain('Document');
	});

	it('gets the fields of the test library', async () => {
		const items = await runNode({
			resource: 'list',
			operation: 'getFields',
			list: locator(await getLibraryId(), 'id'),
		});
		expect(items.map((item) => item.json.InternalName)).toContain('Title');
	});
});

describe.runIf(hasConfig)('site and list setup (write, skipped on the production site)', () => {
	const listParams = () => ({ list: locator('00000000-0000-0000-0000-000000000000', 'id') });

	itWrite('sets and reads a property bag value (changes the site)', async () => {
		await runNode({
			resource: 'site',
			operation: 'setPropertyBag',
			propertyValues: { values: [{ key: '__pnp_test_key', value: 'one' }] },
		});
		const [result] = await runNode({
			resource: 'site',
			operation: 'getPropertyBag',
			propertyKeys: '__pnp_test_key',
		});
		expect(result.json.properties).toEqual({ __pnp_test_key: 'one' });
		await runNode({
			resource: 'site',
			operation: 'setPropertyBag',
			propertyValues: { values: [{ key: '__pnp_test_key', value: '' }] },
		});
	});

	itWrite('creates a Communication site (creates a site)', async () => {
		const [result] = await runNode({
			resource: 'site',
			operation: 'createSite',
			siteType: 'communication',
			siteTitle: 'PnP Test Communication Site',
			newSiteUrl: `pnp-test-comm-${Date.now()}`,
			createSiteOptions: { owner: 'user@contoso.com', waitUntilReady: true },
		});
		expect(result.json.SiteStatusName).toBe('Ready');
	});

	itWrite('creates a Team site without a group (creates a site)', async () => {
		const [result] = await runNode({
			resource: 'site',
			operation: 'createSite',
			siteType: 'team',
			siteTitle: 'PnP Test Team Site',
			newSiteUrl: `pnp-test-team-${Date.now()}`,
			createSiteOptions: { owner: 'user@contoso.com', waitUntilReady: true },
		});
		expect(result.json.SiteStatusName).toBe('Ready');
	});

	itWrite('creates a subsite (creates a subsite)', async () => {
		const [result] = await runNode({
			resource: 'site',
			operation: 'createSubsite',
			subsiteTitle: 'PnP Test Subsite',
			subsiteUrl: `pnp-test-sub-${Date.now()}`,
			subsiteTemplate: 'STS#0',
		});
		expect(result.json).toHaveProperty('Url');
	});

	itWrite('joins and leaves a hub site (changes the hub association of the site)', async () => {
		const hubs = await runNode({ resource: 'site', operation: 'getHubSites', returnAll: true });
		const hubSiteId = String(hubs[0].json.SiteId);
		const [joined] = await runNode({ resource: 'site', operation: 'joinHubSite', hubSiteId });
		expect(joined.json.HubSiteId).toBe(hubSiteId);
		const [left] = await runNode({ resource: 'site', operation: 'leaveHubSite' });
		expect(left.json.HubSiteId).toBe('00000000-0000-0000-0000-000000000000');
	});

	itWrite('updates the settings of a list (changes a list)', async () => {
		const [result] = await runNode({
			resource: 'list',
			operation: 'update',
			...listParams(),
			newTitle: 'PnP Test List Renamed',
			updateOptions: {
				enableVersioning: true,
				majorVersionLimit: 20,
				enableModeration: true,
				enableAttachments: false,
				onQuickLaunch: false,
			},
		});
		expect(result.json.EnableVersioning).toBe(true);
		expect(result.json.MajorVersionLimit).toBe(20);
	});

	itWrite('creates, updates and deletes a field in a list (changes the list schema)', async () => {
		const [created] = await runNode({
			resource: 'list',
			operation: 'createField',
			target: 'list',
			...listParams(),
			fieldType: 'choice',
			displayName: 'PnP Test Choice',
			choices: 'A\nB',
			commonSettings: { internalName: 'PnPTestChoice', required: false, addToDefaultView: true },
		});
		expect(created.json.InternalName).toBe('PnPTestChoice');

		const [updated] = await runNode({
			resource: 'list',
			operation: 'updateField',
			target: 'list',
			...listParams(),
			field: 'PnPTestChoice',
			updateFieldOptions: { displayName: 'PnP Test Choice 2', choices: 'A\nB\nC' },
		});
		expect(updated.json.Title).toBe('PnP Test Choice 2');

		const [deleted] = await runNode({
			resource: 'list',
			operation: 'deleteField',
			target: 'list',
			...listParams(),
			field: 'PnPTestChoice',
		});
		expect(deleted.json.success).toBe(true);
	});

	itWrite(
		'creates and deletes a lookup field and a field from schema XML (changes the list schema)',
		async () => {
			await runNode({
				resource: 'list',
				operation: 'createField',
				target: 'list',
				...listParams(),
				fieldType: 'lookup',
				displayName: 'PnP Test Lookup',
				lookupList: locator('00000000-0000-0000-0000-000000000000', 'id'),
				lookupField: 'Title',
				lookupAllowMultiple: false,
				commonSettings: { internalName: 'PnPTestLookup' },
			});
			await runNode({
				resource: 'list',
				operation: 'createField',
				target: 'list',
				...listParams(),
				fieldType: 'schemaXml',
				schemaXml:
					'<Field Type="Text" DisplayName="PnP Test Xml" Name="PnPTestXml" StaticName="PnPTestXml" />',
				commonSettings: { addToDefaultView: false },
			});
			for (const field of ['PnPTestLookup', 'PnPTestXml']) {
				await runNode({
					resource: 'list',
					operation: 'deleteField',
					target: 'list',
					...listParams(),
					field,
				});
			}
		},
	);

	itWrite('creates and deletes a field in the site (changes the site schema)', async () => {
		const [created] = await runNode({
			resource: 'list',
			operation: 'createField',
			target: 'site',
			fieldType: 'text',
			displayName: 'PnP Test Site Column',
			commonSettings: { internalName: 'PnPTestSiteColumn', group: 'PnP Test' },
		});
		expect(created.json.InternalName).toBe('PnPTestSiteColumn');
		await runNode({
			resource: 'list',
			operation: 'deleteField',
			target: 'site',
			field: 'PnPTestSiteColumn',
		});
	});

	itWrite('adds, defaults and removes a content type of a list (changes the list)', async () => {
		const siteTypes = await runNode({
			resource: 'list',
			operation: 'getContentTypes',
			target: 'site',
		});
		const siteContentType = String(
			siteTypes.find((item) => item.json.Name === 'Item')?.json.StringId,
		);
		const listTypes = await runNode({
			resource: 'list',
			operation: 'addContentType',
			...listParams(),
			siteContentType,
		});
		const added = listTypes.find((item) => String(item.json.StringId).startsWith(siteContentType));
		expect(added).toBeDefined();
		const listContentType = String(added!.json.StringId);

		const [defaulted] = await runNode({
			resource: 'list',
			operation: 'setDefaultContentType',
			...listParams(),
			listContentType,
		});
		expect((defaulted.json.order as string[])[0]).toBe(listContentType);

		const [removed] = await runNode({
			resource: 'list',
			operation: 'removeContentType',
			...listParams(),
			listContentType,
		});
		expect(removed.json.success).toBe(true);
	});
});
