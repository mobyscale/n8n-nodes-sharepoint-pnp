import type { IList } from '@pnp/sp/lists/index.js';

import { isNotFoundError } from './utils';

const DAY_MS = 24 * 60 * 60 * 1000;

/** SharePoint allows at most 180 days. One day less leaves room for clock differences. */
export const SUBSCRIPTION_LIFETIME_MS = 179 * DAY_MS;
/** A test subscription ("Listen for test event") only lives for a short time. */
export const TEST_SUBSCRIPTION_LIFETIME_MS = DAY_MS;
/** Renew when less than this time remains. */
export const SUBSCRIPTION_RENEW_BEFORE_MS = 30 * DAY_MS;

export interface ListSubscription {
	id: string;
	notificationUrl: string;
	expirationDateTime: string;
	resource?: string;
	clientState?: string | null;
}

/** One entry of the "value" array that SharePoint posts to the notification URL. */
export interface SharePointNotification {
	subscriptionId?: string;
	clientState?: string;
	expirationDateTime?: string;
	resource?: string;
	tenantId?: string;
	siteUrl?: string;
	webId?: string;
}

export function getSubscriptionExpiration(now: number, test = false): string {
	return new Date(
		now + (test ? TEST_SUBSCRIPTION_LIFETIME_MS : SUBSCRIPTION_LIFETIME_MS),
	).toISOString();
}

/** Returns true when the subscription expires in less than 30 days, or when the date is not known. */
export function shouldRenewSubscription(
	expirationDateTime: string | undefined,
	now: number,
): boolean {
	const expiration = expirationDateTime ? Date.parse(expirationDateTime) : NaN;
	if (!Number.isFinite(expiration)) return true;
	return expiration - now < SUBSCRIPTION_RENEW_BEFORE_MS;
}

/** SharePoint can return the URL with a different case in the host. */
export function isSameNotificationUrl(a: string | undefined, b: string | undefined): boolean {
	if (!a || !b) return false;
	try {
		return new URL(a).href === new URL(b).href;
	} catch {
		return a === b;
	}
}

function undefinedIfNotFound(error: unknown): undefined {
	if (isNotFoundError(error)) return undefined;
	throw error;
}

export async function getListSubscription(
	list: IList,
	id: string,
): Promise<ListSubscription | undefined> {
	return (await list.subscriptions.getById(id)().catch(undefinedIfNotFound)) as
		| ListSubscription
		| undefined;
}

export async function getListSubscriptions(list: IList): Promise<ListSubscription[]> {
	return (await list.subscriptions()) as ListSubscription[];
}

/** SharePoint sends a validation request to the notification URL before it creates the subscription. */
export async function addListSubscription(
	list: IList,
	notificationUrl: string,
	expirationDateTime: string,
	clientState: string,
): Promise<ListSubscription> {
	return (await list.subscriptions.add(
		notificationUrl,
		expirationDateTime,
		clientState,
	)) as ListSubscription;
}

export async function renewListSubscription(
	list: IList,
	id: string,
	expirationDateTime: string,
): Promise<void> {
	await list.subscriptions.getById(id).update(expirationDateTime);
}

/** Deletes a subscription. Returns false when it does not exist. */
export async function deleteListSubscription(list: IList, id: string): Promise<boolean> {
	return await list.subscriptions
		.getById(id)
		.delete()
		.then(() => true)
		.catch((error: unknown) => {
			if (isNotFoundError(error)) return false;
			throw error;
		});
}

/**
 * Deletes the subscriptions of a list that send to one notification URL. The n8n webhook URL of a
 * node is unique, so these subscriptions belong to that node only. Returns the deleted IDs.
 */
export async function deleteListSubscriptionsByUrl(
	list: IList,
	notificationUrl: string,
): Promise<string[]> {
	const deleted: string[] = [];
	for (const subscription of await getListSubscriptions(list)) {
		if (!isSameNotificationUrl(subscription.notificationUrl, notificationUrl)) continue;
		if (await deleteListSubscription(list, subscription.id)) deleted.push(subscription.id);
	}
	return deleted;
}

/** Compares two change tokens of the same object by their change number. */
export function compareChangeTokens(a: string, b: string): number {
	const [, , , ticksA, numberA] = a.split(';');
	const [, , , ticksB, numberB] = b.split(';');
	const toBig = (value: string | undefined) =>
		value && /^-?\d+$/.test(value) ? BigInt(value) : BigInt(-1);
	const byNumber = toBig(numberA) - toBig(numberB);
	const difference = byNumber !== BigInt(0) ? byNumber : toBig(ticksA) - toBig(ticksB);
	return difference > BigInt(0) ? 1 : difference < BigInt(0) ? -1 : 0;
}

export function newestChangeToken(...tokens: Array<string | undefined>): string | undefined {
	let newest: string | undefined;
	for (const token of tokens) {
		if (!token) continue;
		if (!newest || compareChangeTokens(token, newest) > 0) newest = token;
	}
	return newest;
}
