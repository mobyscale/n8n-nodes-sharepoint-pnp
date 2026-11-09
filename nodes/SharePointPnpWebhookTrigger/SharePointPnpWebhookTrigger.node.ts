import { randomUUID, timingSafeEqual } from 'node:crypto';
import type {
	IDataObject,
	IHookFunctions,
	INodeType,
	INodeTypeDescription,
	IWebhookFunctions,
	IWebhookResponseData,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import { getChangeItemFieldOptions } from '../SharePointPnp/methods/changeOptions';
import * as listSearch from '../SharePointPnp/methods/listSearch';
import {
	eventsProperty,
	folderPathFilterProperty,
	getTriggerEvents,
	getTriggerSettings,
	webhookTriggerOptionsProperty,
} from '../SharePointPnp/shared/changeTrigger';
import type { LocatorValue } from '../SharePointPnp/shared/changes';
import {
	buildChangeQuery,
	getChangeTarget,
	runChangeQueryWithExpiry,
} from '../SharePointPnp/shared/changes';
import {
	getListFromLocator,
	getSiteUrl,
	listRLC,
	siteRLC,
} from '../SharePointPnp/shared/descriptions';
import type { SharePointNotification } from '../SharePointPnp/shared/subscriptions';
import {
	addListSubscription,
	deleteListSubscriptionsByUrl,
	getListSubscription,
	getSubscriptionExpiration,
	isSameNotificationUrl,
	newestChangeToken,
	renewListSubscription,
	shouldRenewSubscription,
} from '../SharePointPnp/shared/subscriptions';
import { toNodeError } from '../SharePointPnp/shared/utils';
import { createSp } from '../SharePointPnp/transport';

export interface WebhookState extends IDataObject {
	subscriptionId?: string;
	clientState?: string;
	notificationUrl?: string;
	expirationDateTime?: string;
	listKey?: string;
	changeToken?: string;
}

type SubscriptionContext = IHookFunctions | IWebhookFunctions;

export class SharePointPnpWebhookTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'SharePoint PnP Webhook Trigger',
		name: 'sharePointPnpWebhookTrigger',
		icon: { light: 'file:sharepoint-pnp.svg', dark: 'file:sharepoint-pnp.svg' },
		group: ['trigger'],
		version: 1,
		subtitle: '={{$parameter["events"].join(", ")}}',
		description:
			'Starts the workflow when SharePoint sends a webhook notification for a list or library',
		defaults: { name: 'SharePoint PnP Webhook Trigger' },
		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'sharePointPnpOAuth2Api',
				required: true,
				displayOptions: { show: { authentication: ['oAuth2'] } },
			},
			{
				name: 'sharePointPnpCertificateApi',
				required: true,
				displayOptions: { show: { authentication: ['certificate'] } },
			},
		],
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'webhook',
			},
		],
		properties: [
			{
				displayName:
					'SharePoint must reach this n8n instance from the internet over HTTPS. Set WEBHOOK_URL to the public URL of n8n. "Listen for test event" uses the test URL. A subscription lasts at most 180 days. The node renews it on activation and on notifications when less than 30 days remain. If the list has no changes for more than 150 days, save or activate the workflow again to create a new subscription. If you cannot open n8n to the internet, use the SharePoint PnP Trigger (polling).',
				name: 'notice',
				type: 'notice',
				default: '',
			},
			{
				displayName: 'Authentication',
				name: 'authentication',
				type: 'options',
				options: [
					{ name: 'OAuth2 (Delegated)', value: 'oAuth2' },
					{ name: 'Certificate (App-Only)', value: 'certificate' },
				],
				default: 'oAuth2',
			},
			siteRLC,
			{ ...listRLC, required: true },
			eventsProperty,
			{
				...folderPathFilterProperty,
				description:
					'Only trigger for files, folders and items inside this folder and its subfolders. Leave empty for all. SharePoint reads the change log of the folder.',
			},
			webhookTriggerOptionsProperty,
		],
	};

	methods = { listSearch, loadOptions: { getChangeItemFieldOptions } };

	webhookMethods = {
		default: {
			async checkExists(this: IHookFunctions): Promise<boolean> {
				return await runHook(this, checkSubscription);
			},
			async create(this: IHookFunctions): Promise<boolean> {
				return await runHook(this, createSubscription);
			},
			async delete(this: IHookFunctions): Promise<boolean> {
				return await runHook(this, deleteSubscription);
			},
		},
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		return await handleWebhook.call(this);
	}
}

async function runHook(
	ctx: IHookFunctions,
	method: (this: IHookFunctions) => Promise<boolean>,
): Promise<boolean> {
	try {
		return await method.call(ctx);
	} catch (error) {
		throw toNodeError(ctx.getNode(), error, 0);
	}
}

function getListLocator(ctx: SubscriptionContext): LocatorValue {
	const list = ctx.getNodeParameter('list') as LocatorValue;
	if (!String(list?.value ?? '').trim()) {
		throw new NodeOperationError(ctx.getNode(), 'Select a list');
	}
	return list;
}

async function getListContext(ctx: SubscriptionContext) {
	const siteUrl = await getSiteUrl(ctx);
	const sp = createSp(ctx, siteUrl);
	const locator = getListLocator(ctx);
	const listKey = [siteUrl, locator.mode, locator.value].join('|').toLowerCase();
	return { siteUrl, sp, locator, list: getListFromLocator(sp, locator), listKey };
}

async function getHookContext(ctx: IHookFunctions) {
	const notificationUrl = ctx.getNodeWebhookUrl('default');
	if (!notificationUrl) {
		throw new NodeOperationError(ctx.getNode(), 'n8n has no webhook URL for this node');
	}
	return { ...(await getListContext(ctx)), notificationUrl };
}

async function getCurrentListToken(list: ReturnType<typeof getListFromLocator>): Promise<string> {
	const data = await list.select('CurrentChangeToken')();
	return (data as { CurrentChangeToken: { StringValue: string } }).CurrentChangeToken.StringValue;
}

/** Finds the subscription by its ID and notification URL, and renews it when it expires soon. */
export async function checkSubscription(this: IHookFunctions): Promise<boolean> {
	const state = this.getWorkflowStaticData('node') as WebhookState;
	if (!state.subscriptionId) return false;
	const { list, listKey, notificationUrl } = await getHookContext(this);
	if (state.listKey !== listKey || !isSameNotificationUrl(state.notificationUrl, notificationUrl)) {
		return false;
	}
	const subscription = await getListSubscription(list, state.subscriptionId);
	if (!subscription || !isSameNotificationUrl(subscription.notificationUrl, notificationUrl)) {
		return false;
	}
	const now = Date.now();
	state.expirationDateTime = subscription.expirationDateTime;
	if (shouldRenewSubscription(subscription.expirationDateTime, now)) {
		const expiration = getSubscriptionExpiration(now, this.getMode() === 'manual');
		await renewListSubscription(list, subscription.id, expiration);
		state.expirationDateTime = expiration;
	}
	state.changeToken ??= await getCurrentListToken(list);
	return true;
}

/**
 * Creates the subscription with a new client state. The current change token is saved first, so
 * that the first notification reads all changes after the activation.
 */
export async function createSubscription(this: IHookFunctions): Promise<boolean> {
	const state = this.getWorkflowStaticData('node') as WebhookState;
	const { list, listKey, notificationUrl } = await getHookContext(this);
	await deleteListSubscriptionsByUrl(list, notificationUrl);

	const clientState = randomUUID();
	const changeToken = await getCurrentListToken(list);
	const expiration = getSubscriptionExpiration(Date.now(), this.getMode() === 'manual');
	const subscription = await addListSubscription(
		list,
		notificationUrl,
		expiration,
		clientState,
	).catch((error: unknown) => {
		if (!isValidationFailure(error)) throw error;
		throw new NodeApiError(this.getNode(), error as JsonObject, {
			httpCode: '400',
			message: 'SharePoint could not validate the webhook URL of n8n',
			description: `SharePoint sends a request to ${notificationUrl} and needs the validation token back within 5 seconds. Make sure that n8n is reachable from the internet over HTTPS, and that WEBHOOK_URL is the public URL of n8n.`,
		});
	});

	state.subscriptionId = subscription.id;
	state.clientState = clientState;
	state.notificationUrl = notificationUrl;
	state.expirationDateTime = subscription.expirationDateTime ?? expiration;
	state.listKey = listKey;
	state.changeToken = changeToken;
	return true;
}

/** Deletes the subscriptions of the list that send to the webhook URL of this node. */
export async function deleteSubscription(this: IHookFunctions): Promise<boolean> {
	const state = this.getWorkflowStaticData('node') as WebhookState;
	const { list, notificationUrl } = await getHookContext(this);
	await deleteListSubscriptionsByUrl(list, notificationUrl);
	if (isSameNotificationUrl(state.notificationUrl, notificationUrl)) {
		for (const key of [
			'subscriptionId',
			'clientState',
			'notificationUrl',
			'expirationDateTime',
			'listKey',
			'changeToken',
		]) {
			delete state[key];
		}
	}
	return true;
}

function isValidationFailure(error: unknown): boolean {
	return /Failed to validate the notification URL/i.test(
		String((error as { message?: unknown })?.message ?? ''),
	);
}

export function getValidationToken(query: IDataObject): string | undefined {
	const value = query.validationtoken ?? query.validationToken;
	return typeof value === 'string' && value ? value : undefined;
}

export function getNotifications(body: IDataObject): SharePointNotification[] {
	const value = body?.value;
	return Array.isArray(value) ? (value as SharePointNotification[]) : [];
}

function safeEqual(a: string, b: string): boolean {
	const left = Buffer.from(a);
	const right = Buffer.from(b);
	return left.length === right.length && timingSafeEqual(left, right);
}

/** A notification is valid when it has the client state and the subscription ID of this node. */
export function isValidNotification(
	notification: SharePointNotification,
	state: WebhookState,
): boolean {
	if (!state.clientState || typeof notification.clientState !== 'string') return false;
	if (!safeEqual(notification.clientState, state.clientState)) return false;
	return !state.subscriptionId || notification.subscriptionId === state.subscriptionId;
}

// n8n reads the static data of the workflow when a request starts. When two notifications arrive
// close together, the second one can see an old token. In one n8n process, the notifications of a
// subscription run one after the other, and the newest token is kept here.
const queues = new Map<string, Promise<unknown>>();
const latestTokens = new Map<string, string>();

async function runExclusive<T>(key: string, run: () => Promise<T>): Promise<T> {
	const previous = queues.get(key) ?? Promise.resolve();
	const current = previous.then(run, run);
	const settled = current.then(
		() => undefined,
		() => undefined,
	);
	queues.set(key, settled);
	try {
		return await current;
	} finally {
		if (queues.get(key) === settled) queues.delete(key);
	}
}

/**
 * Answers the validation request and the notifications of SharePoint. SharePoint needs a response
 * within 5 seconds, so the node responds before it reads the changes.
 */
export async function handleWebhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
	const res = this.getResponseObject();
	const validationToken = getValidationToken(this.getQueryData() as IDataObject);
	if (validationToken) {
		res.status(200).type('text/plain').send(validationToken);
		return { noWebhookResponse: true };
	}

	const state = this.getWorkflowStaticData('node') as WebhookState;
	const notifications = getNotifications(this.getBodyData());
	const valid = notifications.filter((notification) => isValidNotification(notification, state));
	if (!valid.length) {
		res.status(403).end();
		return { noWebhookResponse: true };
	}
	res.status(200).end();

	const key = [this.getWorkflow().id ?? '', this.getNode().id, state.subscriptionId ?? ''].join(
		':',
	);
	return await runExclusive(key, async () => {
		try {
			return await processNotification.call(this, state, valid, key);
		} catch (error) {
			throw toNodeError(this.getNode(), error, 0);
		}
	});
}

async function processNotification(
	this: IWebhookFunctions,
	state: WebhookState,
	notifications: SharePointNotification[],
	key: string,
): Promise<IWebhookResponseData> {
	getTriggerEvents(this);
	const { siteUrl, sp, locator, list } = await getListContext(this);
	const { flags, folderPathFilter, queryOptions, runOptions } = getTriggerSettings(this, siteUrl);
	const target = getChangeTarget(sp, 'list', locator, folderPathFilter);

	await renewOnNotification.call(this, list, state, notifications);

	const startToken = newestChangeToken(state.changeToken, latestTokens.get(key));
	if (!startToken) {
		state.changeToken = await target.getCurrentToken();
		latestTokens.set(key, state.changeToken);
		return { noWebhookResponse: true };
	}

	const result = await runChangeQueryWithExpiry(
		sp,
		target,
		buildChangeQuery({ ...flags, startToken, queryOptions }),
		runOptions,
	);
	state.changeToken = result.nextChangeToken ?? startToken;
	latestTokens.set(key, state.changeToken);

	if (!result.changes.length) return { noWebhookResponse: true };
	return { noWebhookResponse: true, workflowData: [this.helpers.returnJsonArray(result.changes)] };
}

/** A failed renewal does not stop the change processing. The next notification tries again. */
async function renewOnNotification(
	this: IWebhookFunctions,
	list: ReturnType<typeof getListFromLocator>,
	state: WebhookState,
	notifications: SharePointNotification[],
): Promise<void> {
	const now = Date.now();
	const expiration = state.expirationDateTime ?? notifications[0]?.expirationDateTime;
	if (!state.subscriptionId || !shouldRenewSubscription(expiration, now)) return;
	const next = getSubscriptionExpiration(now, this.getMode() === 'manual');
	try {
		await renewListSubscription(list, state.subscriptionId, next);
		state.expirationDateTime = next;
	} catch (error) {
		this.logger.warn('Could not renew the SharePoint webhook subscription', {
			error: (error as Error)?.message,
		});
	}
}
