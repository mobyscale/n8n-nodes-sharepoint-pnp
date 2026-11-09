import type { IDataObject, ILoadOptionsFunctions } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { getChangeItemFieldOptions } from '../../nodes/SharePointPnp/methods/changeOptions';
import { searchPropertyBagKeys } from '../../nodes/SharePointPnp/methods/listSearchPropertyBag';
import { searchTerms } from '../../nodes/SharePointPnp/methods/listSearchTerms';
import { getHubSiteOptions } from '../../nodes/SharePointPnp/methods/listSearchSite';
import { getRecycleBinItemOptions } from '../../nodes/SharePointPnp/methods/recycleBinOptions';
import {
	getInstalledLanguageOptions,
	getPropertyBagKeyOptions,
	getSiteDesignOptions,
} from '../../nodes/SharePointPnp/methods/siteOptions';
import { getTermStoreLanguageOptions } from '../../nodes/SharePointPnp/methods/taxonomyOptions';
import {
	config,
	createExecuteContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

/** A load options context. It answers "getNodeParameter(name, fallback, { extractValue })". */
function loadContext(parameters: IDataObject = {}): ILoadOptionsFunctions {
	const base = createExecuteContext(parameters) as unknown as Record<string, unknown>;
	const all: IDataObject = {
		site: { __rl: true, mode: 'url', value: config?.siteUrl },
		...parameters,
	};
	return {
		...base,
		getNodeParameter: (name: string, fallback?: unknown, options?: { extractValue?: boolean }) => {
			const value = name in all ? all[name] : fallback;
			if (options?.extractValue && value && typeof value === 'object' && '__rl' in value) {
				return (value as IDataObject).value;
			}
			return value;
		},
	} as unknown as ILoadOptionsFunctions;
}

describe.runIf(hasConfig)('pickers (read-only)', () => {
	it('lists hub sites', async () => {
		const options = await getHubSiteOptions.call(loadContext());
		for (const option of options) expect(String(option.value)).toMatch(/^[0-9a-f-]{36}$/);
	});

	it('lists site designs with the built-in communication designs', async () => {
		const options = await getSiteDesignOptions.call(loadContext());
		expect(options.map((option) => option.value)).toContain('f6cc5403-0d63-442e-96c0-285923709ffc');
	});

	it('lists the property bag keys and filters them', async () => {
		const options = await getPropertyBagKeyOptions.call(loadContext());
		expect(options.length).toBeGreaterThan(0);
		const first = String(options[0].value);
		const found = await searchPropertyBagKeys.call(loadContext(), first.toUpperCase());
		expect(found.results.map((entry) => entry.value)).toContain(first);
	});

	it('lists the installed languages', async () => {
		const options = await getInstalledLanguageOptions.call(loadContext());
		expect(options.map((option) => option.value)).toContain(1033);
	});

	it('lists the languages of the term store', async () => {
		const options = await getTermStoreLanguageOptions.call(loadContext());
		expect(options[0].value).toBe(0);
		expect(options.length).toBeGreaterThan(1);
	});

	it('searches the terms of the first term set that has terms', async () => {
		const sets = (
			await runNode({ resource: 'taxonomy', operation: 'getTermSets', returnAll: true })
		).map((item) => item.json);
		for (const set of sets.slice(0, 10)) {
			const termSet = locator(String(set.id), 'id');
			const result = await searchTerms.call(loadContext({ operation: 'createTerm', termSet }));
			if (!result.results.length) continue;
			const helper = await searchTerms.call(
				loadContext({ operation: 'getTerm', termSetHelper: termSet }),
			);
			expect(helper.results.length).toBe(result.results.length);
			const first = result.results[0];
			const filtered = await searchTerms.call(
				loadContext({ operation: 'createTerm', termSet }),
				first.name.split(' > ').pop(),
			);
			expect(filtered.results.map((entry) => entry.value)).toContain(first.value);
			return;
		}
	});

	it('returns no change field options without a list', async () => {
		expect(await getChangeItemFieldOptions.call(loadContext())).toEqual([]);
	});
});

describe.runIf(hasConfig)('recycle bin picker', () => {
	let testFolder: string | undefined;
	let recycledId: string | undefined;

	beforeAll(async () => {
		testFolder = await createTestFolder();
	});

	afterAll(async () => {
		if (recycledId) {
			try {
				await createTestSp().site.recycleBin.getById(recycledId).delete();
			} catch {
				// Already deleted by the test.
			}
		}
		await removeTestFolder(testFolder);
	});

	it('lists a file that the test created and recycled, and deletes only that item', async () => {
		const sp = createTestSp();
		const fileName = `__pnp-test-picker-${Date.now()}.txt`;
		await sp.web.getFolderByServerRelativePath(testFolder!).files.addUsingPath(fileName, 'picker');
		recycledId = await sp.web.getFileByServerRelativePath(`${testFolder}/${fileName}`).recycle();

		const options = await getRecycleBinItemOptions.call(loadContext({ scope: 'siteCollection' }));
		const match = options.find((option) => option.value === recycledId);
		expect(match).toBeDefined();
		expect(match!.name).toContain(fileName);

		const siteScope = await getRecycleBinItemOptions.call(loadContext({ scope: 'site' }));
		expect(siteScope.map((option) => option.value)).toContain(recycledId);

		await runNode({
			resource: 'recycleBin',
			operation: 'deletePermanently',
			scope: 'siteCollection',
			ids: [recycledId],
		});
		const after = await getRecycleBinItemOptions.call(loadContext({ scope: 'siteCollection' }));
		expect(after.map((option) => option.value)).not.toContain(recycledId);
		recycledId = undefined;
	});
});
