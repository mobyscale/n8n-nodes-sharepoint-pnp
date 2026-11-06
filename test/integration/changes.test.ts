import type { IDataObject, IPollFunctions } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	buildChangeQuery,
	CHANGE_LOG_RETENTION_MS,
	createChangeTokenFromDate,
	getChangeTarget,
} from '../../nodes/SharePointPnp/shared/changes';
import { SharePointPnpTrigger } from '../../nodes/SharePointPnpTrigger/SharePointPnpTrigger.node';
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

interface ChangeOutput extends IDataObject {
	changeToken: string;
	changeType: string;
	objectType: string;
	UniqueId: string;
	item?: {
		exists: boolean;
		name?: string;
		serverRelativeUrl?: string;
		isFolder?: boolean;
		isFile?: boolean;
	};
}

/** A poll context with the same credentials and request helper as the harness. */
function createPollContext(
	parameters: IDataObject,
	staticData: IDataObject,
	mode: 'trigger' | 'manual' = 'trigger',
): IPollFunctions {
	const execute = createExecuteContext({});
	const all: IDataObject = {
		authentication: 'certificate',
		site: locator(config!.siteUrl, 'url'),
		...parameters,
	};
	return {
		getNode: execute.getNode,
		getMode: () => mode,
		getWorkflowStaticData: () => staticData,
		getCredentials: execute.getCredentials,
		getNodeParameter: (name: string, fallback?: unknown, options?: { extractValue?: boolean }) => {
			const value = name in all ? all[name] : fallback;
			if (options?.extractValue && value && typeof value === 'object' && '__rl' in value) {
				return (value as IDataObject).value;
			}
			return value;
		},
		helpers: {
			httpRequestWithAuthentication: execute.helpers.httpRequestWithAuthentication,
			returnJsonArray: (data: IDataObject[]) => data.map((json) => ({ json })),
		},
	} as unknown as IPollFunctions;
}

// The tests read the change log of the test folder only (List scope with a folder path filter), so
// that item details are read only for items that the tests created. They do not read the recycle
// bin, because it lists items that other users deleted.
describe.runIf(hasConfig)('change resource and trigger', () => {
	let testFolder: string;
	let listId: string;
	const uniqueIds = new Map<string, string>();
	const recycled: string[] = [];

	const listParameters = () => ({ scope: 'list', list: locator(listId, 'id') });
	const ownChangesOnly = () => ({
		resolveDeletedItems: false,
		folderPathFilter: testFolder,
	});

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const libraryRoot = testFolder.slice(0, testFolder.lastIndexOf('/'));
		const list = await createTestSp().web.getList(libraryRoot).select('Id')();
		listId = list.Id;
	});

	afterAll(async () => {
		const sp = createTestSp();
		for (const id of recycled) {
			await sp.web.recycleBin.getById(id).delete();
		}
		await removeTestFolder(testFolder);
	});

	it('gets the current token of the library', async () => {
		const [item] = await runNode({
			resource: 'change',
			operation: 'getCurrentToken',
			...listParameters(),
		});
		expect(item.json.scope).toBe('list');
		expect(String(item.json.changeToken)).toMatch(new RegExp(`^1;3;${listId};\\d+;\\d+$`, 'i'));
	});

	it('returns the add, rename and delete changes in the test folder', async () => {
		const [tokenItem] = await runNode({
			resource: 'change',
			operation: 'getCurrentToken',
			...listParameters(),
		});
		const startToken = String(tokenItem.json.changeToken);

		const sp = createTestSp();
		const folder = sp.web.getFolderByServerRelativePath(testFolder);
		await folder.files.addUsingPath('alpha.txt', 'alpha', { Overwrite: true });
		await folder.folders.addUsingPath('beta');
		await folder.files.addUsingPath('gamma.txt', 'gamma', { Overwrite: true });
		for (const name of ['alpha.txt', 'gamma.txt']) {
			const file = await sp.web
				.getFileByServerRelativePath(`${testFolder}/${name}`)
				.select('UniqueId')();
			uniqueIds.set(name, String(file.UniqueId).toLowerCase());
		}
		const beta = await sp.web
			.getFolderByServerRelativePath(`${testFolder}/beta`)
			.select('UniqueId')();
		uniqueIds.set('beta', String(beta.UniqueId).toLowerCase());

		const gamma = await sp.web.getFileByServerRelativePath(`${testFolder}/gamma.txt`).getItem();
		await gamma.update({ FileLeafRef: 'gamma-renamed.txt' });
		recycled.push(await sp.web.getFileByServerRelativePath(`${testFolder}/alpha.txt`).recycle());

		const items = await runNode({
			resource: 'change',
			operation: 'getAll',
			...listParameters(),
			startFrom: 'token',
			startPosition: startToken,
			objectTypes: ['Item', 'File', 'Folder'],
			changeTypes: ['Add', 'Update', 'DeleteObject', 'Rename'],
			returnAll: true,
			queryOptions: {},
			options: { includeItemDetails: true, ...ownChangesOnly() },
		});
		const changes = items.map((item) => item.json as ChangeOutput);
		const ownIds = [...uniqueIds.values()];

		expect(changes.length).toBeGreaterThan(0);
		for (const change of changes) {
			expect(ownIds).toContain(String(change.UniqueId).toLowerCase());
			expect(change.changeToken).toMatch(/^1;3;/);
			expect(change.objectType).toBe('Item');
			expect(change.ListId).toBeTruthy();
			expect(change.Time).toBeTruthy();
			if (change.item?.exists) {
				expect(change.item.serverRelativeUrl?.startsWith(`${testFolder}/`)).toBe(true);
			}
		}

		const find = (changeType: string, key: string) =>
			changes.find(
				(change) =>
					change.changeType === changeType &&
					String(change.UniqueId).toLowerCase() === uniqueIds.get(key),
			);

		expect(find('Add', 'beta')?.item).toMatchObject({ exists: true, name: 'beta', isFolder: true });
		expect(find('Add', 'gamma.txt')?.item).toMatchObject({
			exists: true,
			name: 'gamma-renamed.txt',
			isFile: true,
		});
		expect(find('Rename', 'gamma.txt')?.item?.name).toBe('gamma-renamed.txt');
		expect(find('Add', 'alpha.txt')?.item).toEqual({ exists: false });
		expect(find('DeleteObject', 'alpha.txt')?.item).toEqual({ exists: false });

		const [single] = await runNode({
			resource: 'change',
			operation: 'getAll',
			...listParameters(),
			startFrom: 'token',
			startPosition: startToken,
			objectTypes: ['Item'],
			changeTypes: ['Add', 'DeleteObject', 'Rename'],
			returnAll: false,
			limit: 2,
			options: { ...ownChangesOnly(), output: 'single' },
		});
		expect(single.json.count).toBe(2);
		expect(String(single.json.nextChangeToken)).toMatch(/^1;3;/);

		const rest = await runNode({
			resource: 'change',
			operation: 'getAll',
			...listParameters(),
			startFrom: 'token',
			startPosition: single.json.nextChangeToken,
			objectTypes: ['Item'],
			changeTypes: ['Add', 'DeleteObject', 'Rename'],
			returnAll: true,
			options: ownChangesOnly(),
		});
		const firstTwo = (single.json.changes as ChangeOutput[]).map((change) => change.changeToken);
		expect(rest.length).toBeGreaterThan(0);
		for (const item of rest) {
			expect(firstTwo).not.toContain((item.json as ChangeOutput).changeToken);
		}
	});

	it('finds the changes from a start date', async () => {
		const items = await runNode({
			resource: 'change',
			operation: 'getAll',
			...listParameters(),
			startFrom: 'date',
			startDate: new Date(Date.now() - 15 * 60 * 1000).toISOString(),
			objectTypes: ['Item'],
			changeTypes: ['Add'],
			returnAll: true,
			options: { includeItemDetails: true, ...ownChangesOnly() },
		});
		const names = items.map((item) => (item.json as ChangeOutput).item?.name);
		expect(names).toContain('beta');
	});

	it('polls: the first poll stores the token, the next poll returns the new file', async () => {
		const staticData: IDataObject = {};
		const parameters = {
			...listParameters(),
			events: ['fileCreated', 'folderCreated'],
			folderPathFilter: testFolder,
			options: { includeItemDetails: true, resolveDeletedItems: false },
		};
		const trigger = new SharePointPnpTrigger();

		expect(await trigger.poll.call(createPollContext(parameters, staticData))).toBeNull();
		const firstToken = staticData.changeToken;
		expect(String(firstToken)).toMatch(/^1;3;/);

		await createTestSp()
			.web.getFolderByServerRelativePath(testFolder)
			.files.addUsingPath('delta.txt', 'delta', { Overwrite: true });

		const result = await trigger.poll.call(createPollContext(parameters, staticData));
		const changes = (result?.[0] ?? []).map((item) => item.json as ChangeOutput);
		expect(changes.map((change) => change.item?.name)).toEqual(['delta.txt']);
		expect(changes[0].events).toEqual(['fileCreated']);
		expect(staticData.changeToken).not.toBe(firstToken);

		expect(await trigger.poll.call(createPollContext(parameters, staticData))).toBeNull();

		const tokenBeforeManual = staticData.changeToken;
		const manual = await trigger.poll.call(createPollContext(parameters, staticData, 'manual'));
		const manualNames = (manual?.[0] ?? []).map((item) => (item.json as ChangeOutput).item?.name);
		expect(manualNames).toContain('delta.txt');
		expect(manualNames).toContain('beta');
		expect(staticData.changeToken).toBe(tokenBeforeManual);
	});

	// A token from before the change log takes long to read in a folder change log (about a minute).
	it('handles a change token that is older than the change log', { timeout: 600_000 }, async () => {
		const day = 24 * 60 * 60 * 1000;
		const oldToken = createChangeTokenFromDate('list', listId, new Date(Date.now() - 75 * day));
		const sp = createTestSp();
		await sp.web
			.getFolderByServerRelativePath(testFolder)
			.files.addUsingPath('epsilon.txt', 'epsilon', { Overwrite: true });

		// The real behavior of SharePoint Online: no "start token not found" error. It returns the
		// changes that are still in the change log, so the node also checks the age of the token.
		const target = getChangeTarget(sp, 'list', { mode: 'id', value: listId }, testFolder);
		const raw = await target.getChanges(
			buildChangeQuery({ objectTypes: ['Item'], changeTypes: ['Add'], startToken: oldToken }),
		);
		expect(raw.length).toBeGreaterThan(0);
		for (const change of raw) {
			expect(Date.parse(String(change.Time))).toBeGreaterThan(
				Date.now() - CHANGE_LOG_RETENTION_MS - day,
			);
		}

		const parameters = {
			...listParameters(),
			events: ['fileCreated'],
			folderPathFilter: testFolder,
			options: { includeItemDetails: true, resolveDeletedItems: false },
		};
		const trigger = new SharePointPnpTrigger();
		const expiredState = async (options: IDataObject) => {
			const staticData: IDataObject = {};
			const withOptions = { ...parameters, options: { ...parameters.options, ...options } };
			await trigger.poll.call(createPollContext(withOptions, staticData));
			staticData.changeToken = oldToken;
			return { staticData, context: () => createPollContext(withOptions, staticData) };
		};

		const now = await expiredState({ onExpiredToken: 'now' });
		expect(await trigger.poll.call(now.context())).toBeNull();
		expect(String(now.staticData.changeToken)).toMatch(/^1;3;/);
		expect(now.staticData.changeToken).not.toBe(oldToken);

		const fail = await expiredState({ onExpiredToken: 'fail' });
		await expect(trigger.poll.call(fail.context())).rejects.toThrow('expired');
		expect(fail.staticData.changeToken).toBe(oldToken);

		const oldest = await expiredState({});
		const result = await trigger.poll.call(oldest.context());
		const emitted = (result?.[0] ?? []).map((item) => item.json as ChangeOutput);
		expect(emitted.map((change) => change.item?.name)).toContain('epsilon.txt');
		for (const change of emitted) {
			expect(change.tokenExpired).toBe(true);
			if (change.item?.exists) {
				expect(change.item.serverRelativeUrl?.startsWith(`${testFolder}/`)).toBe(true);
			}
		}
		expect(oldest.staticData.changeToken).not.toBe(oldToken);

		const getMany = (onExpiredToken: string) =>
			runNode({
				resource: 'change',
				operation: 'getAll',
				...listParameters(),
				startFrom: 'date',
				startDate: new Date(Date.now() - 75 * day).toISOString(),
				objectTypes: ['Item'],
				changeTypes: ['Add'],
				returnAll: true,
				options: { ...ownChangesOnly(), onExpiredToken, output: 'single' },
			});
		await expect(getMany('fail')).rejects.toThrow('expired');
		const [single] = await getMany('now');
		expect(single.json).toMatchObject({ changes: [], count: 0, tokenExpired: true });
		expect(String(single.json.nextChangeToken)).toMatch(/^1;3;/);
	});
});
