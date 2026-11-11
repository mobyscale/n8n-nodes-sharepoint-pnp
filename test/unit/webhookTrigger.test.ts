import type { IDataObject, IHookFunctions, INode, IWebhookFunctions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import * as changes from '../../nodes/SharePointPnp/shared/changes';
import * as descriptions from '../../nodes/SharePointPnp/shared/descriptions';
import {
	compareChangeTokens,
	getSubscriptionExpiration,
	newestChangeToken,
	shouldRenewSubscription,
	SUBSCRIPTION_LIFETIME_MS,
	type ListSubscription,
} from '../../nodes/SharePointPnp/shared/subscriptions';
import {
	getValidationToken,
	isValidNotification,
	SharePointPnpWebhookTrigger,
	type WebhookState,
} from '../../nodes/SharePointPnpWebhookTrigger/SharePointPnpWebhookTrigger.node';

vi.mock('../../nodes/SharePointPnp/shared/changes', async (importOriginal) => {
	const actual = await importOriginal<typeof changes>();
	return { ...actual, getChangeTarget: vi.fn(), runChangeQueryWithExpiry: vi.fn() };
});

vi.mock('../../nodes/SharePointPnp/shared/descriptions', async (importOriginal) => {
	const actual = await importOriginal<typeof descriptions>();
	return { ...actual, getListFromLocator: vi.fn() };
});

const DAY = 24 * 60 * 60 * 1000;
const LIST_ID = '11111111-2222-3333-4444-555555555555';
const URL_A = 'https://n8n.example.com/webhook/node-a/webhook';
const URL_OTHER = 'https://other.example.com/hook';

const node: INode = {
	id: 'node-1',
	name: 'SharePoint PnP Webhook Trigger',
	type: '@mobyscale/n8n-nodes-sharepoint-pnp.sharePointPnpWebhookTrigger',
	typeVersion: 1,
	position: [0, 0],
	parameters: {},
};

const target = {
	scope: 'list' as const,
	getChanges: vi.fn(),
	getCurrentToken: vi.fn(async () => 'current-token'),
	getId: vi.fn(async () => LIST_ID),
};

const runWithExpiry = vi.mocked(changes.runChangeQueryWithExpiry);
const getChangeTarget = vi.mocked(changes.getChangeTarget);
const getListFromLocator = vi.mocked(descriptions.getListFromLocator);

interface FakeList {
	subscriptions: ListSubscription[];
	added: Array<{ url: string; expiration: string; clientState: string }>;
	updates: Array<{ id: string; expiration: string }>;
	deleted: string[];
	failUpdate: boolean;
	list: { subscriptions: { add: (...args: string[]) => Promise<unknown> } };
}

function useFakeList(subscriptions: ListSubscription[] = []): FakeList {
	const fake = {
		subscriptions,
		added: [],
		updates: [],
		deleted: [],
		failUpdate: false,
	} as unknown as FakeList;
	const notFound = () => Object.assign(new Error('Not found'), { status: 404 });
	const find = (id: string) => fake.subscriptions.find((entry) => entry.id === id);
	const collection = Object.assign(async () => [...fake.subscriptions], {
		add: async (url: string, expiration: string, clientState: string) => {
			fake.added.push({ url, expiration, clientState });
			const created = { id: 'new-id', notificationUrl: url, expirationDateTime: expiration };
			fake.subscriptions.push(created);
			return created;
		},
		getById: (id: string) =>
			Object.assign(
				async () => {
					const entry = find(id);
					if (!entry) throw notFound();
					return entry;
				},
				{
					update: async (expiration: string) => {
						if (fake.failUpdate) throw new Error('Update failed');
						fake.updates.push({ id, expiration });
					},
					delete: async () => {
						if (!find(id)) throw notFound();
						fake.deleted.push(id);
						fake.subscriptions = fake.subscriptions.filter((entry) => entry.id !== id);
					},
				},
			),
	});
	const list = {
		subscriptions: collection,
		select: () => async () => ({ CurrentChangeToken: { StringValue: 'list-token' } }),
	};
	getListFromLocator.mockReturnValue(list as never);
	fake.list = list;
	return fake;
}

const baseParameters: IDataObject = {
	authentication: 'certificate',
	site: { __rl: true, mode: 'url', value: '/sites/Example' },
	list: { __rl: true, mode: 'id', value: LIST_ID },
	events: ['fileCreated'],
	folderPathFilter: '',
	options: {},
};

function getParameter(all: IDataObject) {
	return (name: string, fallback?: unknown, options?: { extractValue?: boolean }) => {
		const value = name in all ? all[name] : fallback;
		if (options?.extractValue && value && typeof value === 'object' && '__rl' in value) {
			return (value as IDataObject).value;
		}
		return value;
	};
}

function hookContext(
	staticData: IDataObject,
	options: { mode?: 'trigger' | 'manual'; url?: string; parameters?: IDataObject } = {},
): IHookFunctions {
	return {
		getNode: () => node,
		getMode: () => options.mode ?? 'trigger',
		getWorkflowStaticData: () => staticData,
		getCredentials: async () => ({ subdomain: 'contoso' }),
		getNodeParameter: getParameter({ ...baseParameters, ...options.parameters }),
		getNodeWebhookUrl: () => options.url ?? URL_A,
		helpers: {},
	} as unknown as IHookFunctions;
}

function createResponse() {
	const calls: string[] = [];
	const res = {
		status: vi.fn((code: number) => {
			calls.push(`status:${code}`);
			return res;
		}),
		type: vi.fn(() => res),
		send: vi.fn(() => res),
		end: vi.fn(() => {
			calls.push('end');
			return res;
		}),
	};
	return { res, calls };
}

function webhookContext(
	staticData: IDataObject,
	request: { query?: IDataObject; body?: IDataObject; parameters?: IDataObject },
) {
	const { res, calls } = createResponse();
	const warn = vi.fn();
	const ctx = {
		getNode: () => node,
		getMode: () => 'trigger',
		getWorkflow: () => ({ id: 'workflow-1', active: true }),
		getWorkflowStaticData: () => staticData,
		getCredentials: async () => ({ subdomain: 'contoso' }),
		getNodeParameter: getParameter({ ...baseParameters, ...request.parameters }),
		getQueryData: () => request.query ?? {},
		getBodyData: () => request.body ?? {},
		getResponseObject: () => res,
		logger: { warn },
		helpers: { returnJsonArray: (data: IDataObject[]) => data.map((json) => ({ json })) },
	} as unknown as IWebhookFunctions;
	return { ctx, res, calls, warn };
}

const trigger = new SharePointPnpWebhookTrigger();
const hooks = trigger.webhookMethods.default;

function activeState(overrides: Partial<WebhookState> = {}): WebhookState {
	return {
		subscriptionId: 'sub-1',
		clientState: 'secret-state',
		notificationUrl: URL_A,
		expirationDateTime: new Date(Date.now() + 100 * DAY).toISOString(),
		listKey: `https://contoso.sharepoint.com/sites/example|id|${LIST_ID}`,
		changeToken: `1;3;${LIST_ID};638396640000000000;10`,
		...overrides,
	};
}

const notification = (overrides: IDataObject = {}) => ({
	value: [
		{
			subscriptionId: 'sub-1',
			clientState: 'secret-state',
			expirationDateTime: new Date(Date.now() + 100 * DAY).toISOString(),
			resource: LIST_ID,
			...overrides,
		},
	],
});

beforeEach(() => {
	vi.clearAllMocks();
	getChangeTarget.mockReturnValue(target);
});

describe('subscription helpers', () => {
	it('sets the expiration close to the 180 day maximum, and one day for a test', () => {
		const now = Date.parse('2026-01-01T00:00:00Z');
		expect(Date.parse(getSubscriptionExpiration(now)) - now).toBe(SUBSCRIPTION_LIFETIME_MS);
		expect(SUBSCRIPTION_LIFETIME_MS).toBeLessThanOrEqual(180 * DAY);
		expect(SUBSCRIPTION_LIFETIME_MS).toBeGreaterThan(178 * DAY);
		expect(Date.parse(getSubscriptionExpiration(now, true)) - now).toBe(DAY);
	});

	it('renews when less than 30 days remain or the date is unknown', () => {
		const now = Date.parse('2026-01-01T00:00:00Z');
		const inDays = (days: number) => new Date(now + days * DAY).toISOString();
		expect(shouldRenewSubscription(inDays(29), now)).toBe(true);
		expect(shouldRenewSubscription(inDays(-1), now)).toBe(true);
		expect(shouldRenewSubscription(inDays(31), now)).toBe(false);
		expect(shouldRenewSubscription(inDays(179), now)).toBe(false);
		expect(shouldRenewSubscription(undefined, now)).toBe(true);
		expect(shouldRenewSubscription('not a date', now)).toBe(true);
	});

	it('finds the newest change token by its change number', () => {
		const a = `1;3;${LIST_ID};638396640000000000;10`;
		const b = `1;3;${LIST_ID};638396640000000001;11`;
		expect(compareChangeTokens(a, b)).toBe(-1);
		expect(compareChangeTokens(b, a)).toBe(1);
		expect(compareChangeTokens(a, a)).toBe(0);
		expect(newestChangeToken(a, undefined, b)).toBe(b);
		expect(newestChangeToken(undefined, undefined)).toBeUndefined();
	});
});

describe('webhook lifecycle', () => {
	it('checkExists returns false without a stored subscription', async () => {
		useFakeList();
		expect(await hooks.checkExists.call(hookContext({}))).toBe(false);
	});

	it('checkExists finds the subscription by ID and notification URL', async () => {
		const state = activeState();
		const fake = useFakeList([
			{ id: 'sub-1', notificationUrl: URL_A, expirationDateTime: state.expirationDateTime! },
		]);
		expect(await hooks.checkExists.call(hookContext(state))).toBe(true);
		expect(fake.updates).toEqual([]);
	});

	it('checkExists renews a subscription that expires in less than 30 days', async () => {
		const soon = new Date(Date.now() + 10 * DAY).toISOString();
		const state = activeState({ expirationDateTime: soon });
		const fake = useFakeList([{ id: 'sub-1', notificationUrl: URL_A, expirationDateTime: soon }]);

		expect(await hooks.checkExists.call(hookContext(state))).toBe(true);

		expect(fake.updates).toHaveLength(1);
		const renewed = Date.parse(fake.updates[0].expiration) - Date.now();
		expect(renewed).toBeGreaterThan(178 * DAY);
		expect(state.expirationDateTime).toBe(fake.updates[0].expiration);
	});

	it('checkExists returns false when the subscription is missing or has another URL', async () => {
		useFakeList([]);
		expect(await hooks.checkExists.call(hookContext(activeState()))).toBe(false);

		useFakeList([{ id: 'sub-1', notificationUrl: URL_OTHER, expirationDateTime: '' }]);
		expect(await hooks.checkExists.call(hookContext(activeState()))).toBe(false);
	});

	it('checkExists returns false when the stored subscription is for another URL (test mode)', async () => {
		useFakeList([{ id: 'sub-1', notificationUrl: URL_A, expirationDateTime: '' }]);
		const ctx = hookContext(activeState(), { url: 'https://n8n.example.com/webhook-test/a' });
		expect(await hooks.checkExists.call(ctx)).toBe(false);
	});

	it('create saves the token, a random client state and the subscription ID', async () => {
		const fake = useFakeList([
			{ id: 'stale', notificationUrl: URL_A, expirationDateTime: '' },
			{ id: 'foreign', notificationUrl: URL_OTHER, expirationDateTime: '' },
		]);
		const state: IDataObject = {};

		expect(await hooks.create.call(hookContext(state))).toBe(true);

		expect(fake.deleted).toEqual(['stale']);
		expect(fake.subscriptions.map((entry) => entry.id)).toEqual(['foreign', 'new-id']);
		expect(fake.added).toHaveLength(1);
		expect(fake.added[0].url).toBe(URL_A);
		expect(fake.added[0].clientState).toMatch(/^[0-9a-f-]{36}$/);
		expect(Date.parse(fake.added[0].expiration) - Date.now()).toBeGreaterThan(178 * DAY);
		expect(state).toMatchObject({
			subscriptionId: 'new-id',
			clientState: fake.added[0].clientState,
			notificationUrl: URL_A,
			changeToken: 'list-token',
		});
	});

	it('create explains a failed validation of the notification URL', async () => {
		const fake = useFakeList();
		fake.list.subscriptions.add = async () => {
			throw Object.assign(
				new Error(
					'Error making HttpClient request in queryable [400] Bad Request ::> {"odata.error":{"message":{"value":"Failed to validate the notification URL."}}}',
				),
				{ isHttpRequestError: true, status: 400 },
			);
		};
		const state: IDataObject = {};

		await expect(hooks.create.call(hookContext(state))).rejects.toMatchObject({
			message: 'SharePoint could not validate the webhook URL of n8n',
		});
		expect(state.subscriptionId).toBeUndefined();
	});

	it('create uses a short expiration in test mode', async () => {
		const fake = useFakeList();
		await hooks.create.call(hookContext({}, { mode: 'manual' }));
		expect(Date.parse(fake.added[0].expiration) - Date.now()).toBeLessThanOrEqual(DAY);
	});

	it('delete removes only the subscriptions of this webhook URL and clears the state', async () => {
		const fake = useFakeList([
			{ id: 'sub-1', notificationUrl: URL_A, expirationDateTime: '' },
			{ id: 'foreign', notificationUrl: URL_OTHER, expirationDateTime: '' },
		]);
		const state = activeState();

		expect(await hooks.delete.call(hookContext(state))).toBe(true);

		expect(fake.deleted).toEqual(['sub-1']);
		expect(state.subscriptionId).toBeUndefined();
		expect(state.clientState).toBeUndefined();
	});

	it('delete in test mode keeps the state of the production subscription', async () => {
		const fake = useFakeList([{ id: 'sub-1', notificationUrl: URL_A, expirationDateTime: '' }]);
		const state = activeState();
		const testUrl = 'https://n8n.example.com/webhook-test/a';

		await hooks.delete.call(hookContext(state, { mode: 'manual', url: testUrl }));

		expect(fake.deleted).toEqual([]);
		expect(state.subscriptionId).toBe('sub-1');
	});
});

describe('webhook requests', () => {
	it('answers the validation request with the token as plain text', async () => {
		const { ctx, res } = webhookContext(activeState(), { query: { validationtoken: 'abc 123' } });

		const result = await trigger.webhook.call(ctx);

		expect(result).toEqual({ noWebhookResponse: true });
		expect(res.status).toHaveBeenCalledWith(200);
		expect(res.type).toHaveBeenCalledWith('text/plain');
		expect(res.send).toHaveBeenCalledWith('abc 123');
		expect(runWithExpiry).not.toHaveBeenCalled();
	});

	it('reads the validation token in any supported case', () => {
		expect(getValidationToken({ validationtoken: 'a' })).toBe('a');
		expect(getValidationToken({ validationToken: 'b' })).toBe('b');
		expect(getValidationToken({})).toBeUndefined();
	});

	it('rejects a notification with a wrong client state', async () => {
		const { ctx, res } = webhookContext(activeState(), {
			body: notification({ clientState: 'wrong' }),
		});

		const result = await trigger.webhook.call(ctx);

		expect(result).toEqual({ noWebhookResponse: true });
		expect(res.status).toHaveBeenCalledWith(403);
		expect(runWithExpiry).not.toHaveBeenCalled();
	});

	it('checks the client state and the subscription ID', () => {
		const state = activeState();
		expect(
			isValidNotification({ clientState: 'secret-state', subscriptionId: 'sub-1' }, state),
		).toBe(true);
		expect(isValidNotification({ clientState: 'secret-state', subscriptionId: 'x' }, state)).toBe(
			false,
		);
		expect(isValidNotification({ subscriptionId: 'sub-1' }, state)).toBe(false);
		expect(isValidNotification({ clientState: 'secret-state' }, {})).toBe(false);
	});

	it('responds first, then emits the matching changes and saves the next token', async () => {
		useFakeList();
		const state = activeState({ subscriptionId: 'sub-respond' });
		const { ctx, calls } = webhookContext(state, {
			body: notification({ subscriptionId: 'sub-respond' }),
			parameters: { options: { onExpiredToken: 'now', includeItemDetails: true } },
		});
		runWithExpiry.mockImplementationOnce(async () => {
			calls.push('query');
			return {
				changes: [{ changeType: 'Add' }],
				nextChangeToken: `1;3;${LIST_ID};638396640000000005;15`,
				hasMore: false,
				tokenExpired: false,
			};
		});

		const result = await trigger.webhook.call(ctx);

		expect(calls).toEqual(['status:200', 'end', 'query']);
		expect(result).toEqual({
			noWebhookResponse: true,
			workflowData: [[{ json: { changeType: 'Add' } }]],
		});
		expect(state.changeToken).toBe(`1;3;${LIST_ID};638396640000000005;15`);
		const [, , query, options] = runWithExpiry.mock.calls[0];
		expect(query.ChangeTokenStart).toEqual({ StringValue: `1;3;${LIST_ID};638396640000000000;10` });
		expect(query.File).toBe(true);
		expect(options).toMatchObject({
			events: ['fileCreated'],
			returnAll: true,
			onExpiredToken: 'now',
			includeItemDetails: true,
		});
	});

	it('starts no execution when no change matches', async () => {
		useFakeList();
		const state = activeState({ subscriptionId: 'sub-empty' });
		const { ctx } = webhookContext(state, { body: notification({ subscriptionId: 'sub-empty' }) });
		runWithExpiry.mockResolvedValueOnce({
			changes: [],
			nextChangeToken: `1;3;${LIST_ID};638396640000000003;13`,
			hasMore: false,
			tokenExpired: false,
		});

		expect(await trigger.webhook.call(ctx)).toEqual({ noWebhookResponse: true });
		expect(state.changeToken).toBe(`1;3;${LIST_ID};638396640000000003;13`);
	});

	it('uses the newest token when a second notification has old static data', async () => {
		useFakeList();
		const first = activeState({ subscriptionId: 'sub-race' });
		const second = activeState({ subscriptionId: 'sub-race' });
		const newer = `1;3;${LIST_ID};638396640000000009;19`;
		runWithExpiry
			.mockResolvedValueOnce({
				changes: [],
				nextChangeToken: newer,
				hasMore: false,
				tokenExpired: false,
			})
			.mockResolvedValueOnce({
				changes: [],
				nextChangeToken: undefined,
				hasMore: false,
				tokenExpired: false,
			});

		await Promise.all([
			trigger.webhook.call(
				webhookContext(first, { body: notification({ subscriptionId: 'sub-race' }) }).ctx,
			),
			trigger.webhook.call(
				webhookContext(second, { body: notification({ subscriptionId: 'sub-race' }) }).ctx,
			),
		]);

		const [, , secondQuery] = runWithExpiry.mock.calls[1];
		expect(secondQuery.ChangeTokenStart).toEqual({ StringValue: newer });
		expect(second.changeToken).toBe(newer);
	});

	it('renews the subscription on a notification when it expires soon', async () => {
		const fake = useFakeList([{ id: 'sub-renew', notificationUrl: URL_A, expirationDateTime: '' }]);
		const state = activeState({
			subscriptionId: 'sub-renew',
			expirationDateTime: new Date(Date.now() + 5 * DAY).toISOString(),
		});
		runWithExpiry.mockResolvedValueOnce({ changes: [], hasMore: false, tokenExpired: false });

		await trigger.webhook.call(
			webhookContext(state, { body: notification({ subscriptionId: 'sub-renew' }) }).ctx,
		);

		expect(fake.updates.map((update) => update.id)).toEqual(['sub-renew']);
		expect(Date.parse(String(state.expirationDateTime)) - Date.now()).toBeGreaterThan(178 * DAY);
	});

	it('does not renew when enough time remains, and logs a failed renewal', async () => {
		const fake = useFakeList();
		runWithExpiry.mockResolvedValue({ changes: [], hasMore: false, tokenExpired: false });

		await trigger.webhook.call(
			webhookContext(activeState({ subscriptionId: 'sub-far' }), {
				body: notification({ subscriptionId: 'sub-far' }),
			}).ctx,
		);
		expect(fake.updates).toEqual([]);

		fake.failUpdate = true;
		const soon = activeState({
			subscriptionId: 'sub-fail',
			expirationDateTime: new Date(Date.now() + DAY).toISOString(),
		});
		const { ctx, warn } = webhookContext(soon, {
			body: notification({ subscriptionId: 'sub-fail' }),
		});
		await trigger.webhook.call(ctx);
		expect(warn).toHaveBeenCalledTimes(1);
		expect(runWithExpiry).toHaveBeenCalledTimes(2);
	});
});
