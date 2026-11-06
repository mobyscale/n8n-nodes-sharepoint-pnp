import { randomUUID } from 'node:crypto';
import type { IDataObject, IHookFunctions } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	deleteListSubscription,
	getListSubscription,
	getListSubscriptions,
	isSameNotificationUrl,
} from '../../nodes/SharePointPnp/shared/subscriptions';
import {
	SharePointPnpWebhookTrigger,
	type WebhookState,
} from '../../nodes/SharePointPnpWebhookTrigger/SharePointPnpWebhookTrigger.node';
import {
	config,
	createExecuteContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
} from './harness';

const DAY = 24 * 60 * 60 * 1000;
const hooks = new SharePointPnpWebhookTrigger().webhookMethods.default;
const notificationUrl = config?.webhookNotificationUrl ?? '';

// Webhook subscriptions are list scoped, so the tests use the test library. They only create,
// renew and delete subscriptions with a notification URL that is unique to the test run, and they
// delete every subscription they created by its ID. They never change other subscriptions.
describe.runIf(hasConfig)('webhook trigger subscriptions', () => {
	let testFolder: string;
	let listId: string;
	const created = new Set<string>();

	const hookContext = (
		state: IDataObject,
		url: string,
		mode: 'trigger' | 'manual' = 'trigger',
	): IHookFunctions => {
		const execute = createExecuteContext({});
		const parameters: IDataObject = {
			authentication: 'certificate',
			site: locator(config!.siteUrl, 'url'),
			list: locator(listId, 'id'),
			events: ['fileCreated'],
			folderPathFilter: testFolder,
			options: {},
		};
		return {
			getNode: execute.getNode,
			getMode: () => mode,
			getWorkflowStaticData: () => state,
			getCredentials: execute.getCredentials,
			getNodeParameter: (
				name: string,
				fallback?: unknown,
				options?: { extractValue?: boolean },
			) => {
				const value = name in parameters ? parameters[name] : fallback;
				if (options?.extractValue && value && typeof value === 'object' && '__rl' in value) {
					return (value as IDataObject).value;
				}
				return value;
			},
			getNodeWebhookUrl: () => url,
			helpers: { httpRequestWithAuthentication: execute.helpers.httpRequestWithAuthentication },
		} as unknown as IHookFunctions;
	};

	const testList = () => createTestSp().web.lists.getById(listId);

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const libraryRoot = testFolder.slice(0, testFolder.lastIndexOf('/'));
		listId = (await createTestSp().web.getList(libraryRoot).select('Id')()).Id;
	});

	afterAll(async () => {
		for (const id of created) {
			const subscription = await getListSubscription(testList(), id);
			if (subscription) await deleteListSubscription(testList(), id);
		}
		await removeTestFolder(testFolder);
	});

	it('checkExists returns false for a subscription ID that does not exist', async () => {
		const url = `https://contoso.invalid/webhook/${randomUUID()}`;
		const state: WebhookState = {
			subscriptionId: randomUUID(),
			notificationUrl: url,
			listKey: `${config!.siteUrl}|id|${listId}`.toLowerCase(),
		};
		expect(await hooks.checkExists.call(hookContext(state, url))).toBe(false);
	});

	it('create fails when the notification URL does not answer the validation', async () => {
		const url = `https://contoso.invalid/webhook/${randomUUID()}`;
		const state: IDataObject = {};

		await expect(hooks.create.call(hookContext(state, url))).rejects.toMatchObject({
			message: 'SharePoint could not validate the webhook URL of n8n',
			httpCode: '400',
		});

		const leftOver = (await getListSubscriptions(testList())).filter((subscription) =>
			isSameNotificationUrl(subscription.notificationUrl, url),
		);
		leftOver.forEach((subscription) => created.add(subscription.id));
		expect(leftOver).toEqual([]);
		expect(state.subscriptionId).toBeUndefined();
	});

	// SharePoint sends a validation request to the notification URL when it creates a subscription.
	// Set SP_WEBHOOK_NOTIFICATION_URL to a public HTTPS endpoint that answers it to run this test.
	it.skipIf(!notificationUrl)(
		'creates, finds, renews and deletes a subscription (needs SP_WEBHOOK_NOTIFICATION_URL)',
		async () => {
			const url = new URL(notificationUrl);
			url.searchParams.set('pnpTest', randomUUID());
			const uniqueUrl = url.href;
			const state: WebhookState = {};

			expect(await hooks.checkExists.call(hookContext(state, uniqueUrl, 'manual'))).toBe(false);
			expect(await hooks.create.call(hookContext(state, uniqueUrl, 'manual'))).toBe(true);
			expect(state.subscriptionId).toBeTruthy();
			created.add(String(state.subscriptionId));
			expect(state.clientState).toMatch(/^[0-9a-f-]{36}$/);
			expect(String(state.changeToken)).toMatch(/^1;3;/);

			// A test subscription expires after one day, so checkExists renews it to about 179 days.
			expect(await hooks.checkExists.call(hookContext(state, uniqueUrl))).toBe(true);
			const renewed = await getListSubscription(testList(), String(state.subscriptionId));
			expect(Date.parse(String(renewed?.expirationDateTime)) - Date.now()).toBeGreaterThan(
				170 * DAY,
			);

			expect(await hooks.delete.call(hookContext(state, uniqueUrl))).toBe(true);
			expect(state.subscriptionId).toBeUndefined();
			const remaining = (await getListSubscriptions(testList())).filter((subscription) =>
				isSameNotificationUrl(subscription.notificationUrl, uniqueUrl),
			);
			expect(remaining).toEqual([]);
		},
	);
});
