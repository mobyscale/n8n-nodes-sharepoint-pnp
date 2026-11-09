import { describe, expect, it } from 'vitest';

import { toServerRelativePath } from '../../nodes/SharePointPnp/shared/utils';
import { config, hasConfig, itWrite, locator, runNode } from './harness';

const ensure = (parameters: Record<string, unknown>) =>
	runNode({ resource: 'list', operation: 'ensure', ...parameters });

/**
 * Only the read path of List > Create If Missing runs here: the test library exists, and "Apply
 * Settings to Existing List" is off, so nothing changes. Tests that create a list or change list
 * settings use "itWrite" and run only with SP_RUN_WRITE_TESTS=true. Run them only on a test site that you can change and reset.
 */
describe.runIf(hasConfig)('list create if missing (read only)', () => {
	const libraryRoot = () => toServerRelativePath(config!.siteUrl, config!.libraryPath);

	it('returns the existing test library by URL and changes nothing', async () => {
		const [result] = await ensure({ title: '__pnp-unused-title', listUrl: config!.libraryPath });
		expect(result.json.created).toBe(false);
		expect(result.json.changes).toEqual([]);
		expect(String(result.json.serverRelativeUrl).toLowerCase()).toBe(libraryRoot().toLowerCase());
		expect(result.json.id).toMatch(/^[0-9a-f-]{36}$/i);
	});

	it('returns the existing test library by title and changes nothing', async () => {
		const [byUrl] = await ensure({ title: '__pnp-unused-title', listUrl: config!.libraryPath });
		const [result] = await ensure({ title: String(byUrl.json.title) });
		expect(result.json).toEqual({ ...byUrl.json });
	});
});

describe.runIf(hasConfig)('list create if missing (write)', () => {
	const title = '__pnp-test-ensure-library';

	itWrite('creates a library with settings, then finds it (creates a list)', async () => {
		const [created] = await ensure({
			title,
			template: 101,
			listDescription: 'Created by an integration test',
			ensureOptions: {
				onQuickLaunch: false,
				enableVersioning: true,
				majorVersionLimit: 50,
				enableContentTypes: true,
			},
		});
		expect(created.json.created).toBe(true);
		expect(created.json.changes).toContain(`Created the list "${title}" with template 101`);

		const [again] = await ensure({ title, ensureOptions: { applyToExisting: true } });
		expect(again.json).toMatchObject({ id: created.json.id, created: false, changes: [] });

		await runNode({
			resource: 'list',
			operation: 'delete',
			list: locator(String(created.json.id), 'id'),
			deleteOptions: { permanent: true },
		});
	});

	itWrite('applies settings to an existing list (changes a list)', async () => {
		const [result] = await ensure({
			title,
			ensureOptions: { applyToExisting: true, enableVersioning: true, majorVersionLimit: 100 },
		});
		expect(result.json.created).toBe(false);
	});
});
