import type { SPFI } from '@pnp/sp/index.js';
import { describe, expect, it } from 'vitest';

import {
	CHANGE_LOG_RETENTION_MS,
	ChangeTokenExpiredError,
	createChangeTokenFromDate,
	getChangeTokenTime,
	isChangeTokenOutsideRetention,
	isExpiredChangeTokenError,
	runChangeQueryWithExpiry,
	type ChangeQuery,
	type ChangeTarget,
} from '../../nodes/SharePointPnp/shared/changes';

const LIST_ID = '11111111-2222-3333-4444-555555555555';
const NOW = Date.parse('2026-06-01T00:00:00Z');
const DAY = 24 * 60 * 60 * 1000;
const sp = {} as SPFI;

const tokenAt = (time: number) => createChangeTokenFromDate('list', LIST_ID, new Date(time));
const recentToken = tokenAt(NOW - 10 * DAY);
const oldToken = tokenAt(NOW - 75 * DAY);

const change = (n: number) => ({
	'odata.type': 'SP.ChangeItem',
	ChangeToken: { StringValue: `1;3;${LIST_ID};${n};${n}` },
	ChangeType: 1,
});

const expiredError = Object.assign(
	new Error(
		'Error making HttpClient request ::> {"odata.error":{"code":"-2146232832, Microsoft.SharePoint.SPException","message":{"value":"The changeToken refers to a time before the start of the current change log."}}}',
	),
	{ isHttpRequestError: true, status: 500 },
);

/** A target that returns pages in order. A page can be an error to throw. */
function fakeTarget(pages: Array<object[] | Error>): {
	target: ChangeTarget;
	queries: ChangeQuery[];
	tokenReads: () => number;
} {
	const queries: ChangeQuery[] = [];
	let reads = 0;
	const target: ChangeTarget = {
		scope: 'list',
		getChanges: async (query) => {
			queries.push(query);
			const page = pages.shift() ?? [];
			if (page instanceof Error) throw page;
			return page as never;
		},
		getCurrentToken: async () => {
			reads++;
			return 'current';
		},
		getId: async () => LIST_ID,
	};
	return { target, queries, tokenReads: () => reads };
}

const query = (startToken: string): ChangeQuery => ({
	Item: true,
	Add: true,
	ChangeTokenStart: { StringValue: startToken },
});

describe('change token age', () => {
	it('reads the time of a change token', () => {
		expect(getChangeTokenTime(tokenAt(NOW))).toBe(NOW);
		expect(getChangeTokenTime(`1;3;${LIST_ID};638396640000000000;42`)).toBe(
			Date.parse('2024-01-01T00:00:00Z'),
		);
		expect(getChangeTokenTime('not a token')).toBeUndefined();
		expect(getChangeTokenTime(undefined)).toBeUndefined();
	});

	it('finds tokens that are older than the change log', () => {
		expect(isChangeTokenOutsideRetention(oldToken, NOW)).toBe(true);
		expect(isChangeTokenOutsideRetention(recentToken, NOW)).toBe(false);
		expect(isChangeTokenOutsideRetention(tokenAt(NOW - CHANGE_LOG_RETENTION_MS + DAY), NOW)).toBe(
			false,
		);
		expect(isChangeTokenOutsideRetention('garbage', NOW)).toBe(false);
	});

	it('finds the expired token error and ignores other errors', () => {
		expect(isExpiredChangeTokenError(expiredError)).toBe(true);
		expect(isExpiredChangeTokenError(new Error('The start token was not found'))).toBe(true);
		expect(isExpiredChangeTokenError(new Error('Input string was not in a correct format.'))).toBe(
			false,
		);
		expect(isExpiredChangeTokenError(undefined)).toBe(false);
	});
});

describe('runChangeQueryWithExpiry', () => {
	it('does not change a query with a valid token', async () => {
		const { target, queries, tokenReads } = fakeTarget([[change(1)], []]);
		const result = await runChangeQueryWithExpiry(sp, target, query(recentToken), {
			returnAll: true,
			now: NOW,
		});
		expect(result.tokenExpired).toBe(false);
		expect(result.changes[0].tokenExpired).toBeUndefined();
		expect(result.nextChangeToken).toBe(`1;3;${LIST_ID};1;1`);
		expect(queries[0].ChangeTokenStart?.StringValue).toBe(recentToken);
		expect(queries[0].IgnoreStartTokenNotFoundError).toBe(true);
		expect(tokenReads()).toBe(0);
	});

	it('runs a query without a start token as is', async () => {
		const { target, queries } = fakeTarget([[change(1)], []]);
		const result = await runChangeQueryWithExpiry(
			sp,
			target,
			{ Item: true },
			{ returnAll: true, onExpiredToken: 'fail', now: NOW },
		);
		expect(result.tokenExpired).toBe(false);
		expect(queries[0].ChangeTokenStart).toBeUndefined();
	});

	describe('oldest (default)', () => {
		it('keeps the old token and marks the changes when the token is too old', async () => {
			const { target, queries } = fakeTarget([[change(1), change(2)], []]);
			const result = await runChangeQueryWithExpiry(sp, target, query(oldToken), {
				returnAll: true,
				now: NOW,
			});
			expect(result.tokenExpired).toBe(true);
			expect(result.changes.map((entry) => entry.tokenExpired)).toEqual([true, true]);
			expect(result.nextChangeToken).toBe(`1;3;${LIST_ID};2;2`);
			expect(queries[0].ChangeTokenStart?.StringValue).toBe(oldToken);
			expect(queries[0].IgnoreStartTokenNotFoundError).toBe(true);
		});

		it('returns the current token when an expired token finds no changes', async () => {
			const { target } = fakeTarget([[]]);
			const result = await runChangeQueryWithExpiry(sp, target, query(oldToken), {
				returnAll: true,
				now: NOW,
			});
			expect(result.changes).toEqual([]);
			expect(result.nextChangeToken).toBe('current');
			expect(result.tokenExpired).toBe(true);
		});

		it('runs the query again without a start token after the expired token error', async () => {
			const { target, queries } = fakeTarget([expiredError, [change(5)], []]);
			const result = await runChangeQueryWithExpiry(sp, target, query(recentToken), {
				returnAll: true,
				now: NOW,
			});
			expect(result.tokenExpired).toBe(true);
			expect(result.changes).toHaveLength(1);
			expect(result.changes[0].tokenExpired).toBe(true);
			expect(queries[1].ChangeTokenStart).toBeUndefined();
			expect(queries[1].Item).toBe(true);
		});

		it('rethrows other errors', async () => {
			const { target } = fakeTarget([new Error('Denied')]);
			await expect(
				runChangeQueryWithExpiry(sp, target, query(recentToken), { returnAll: true, now: NOW }),
			).rejects.toThrow('Denied');
		});
	});

	describe('now', () => {
		it('returns no changes and the current token without a query', async () => {
			const { target, queries } = fakeTarget([[change(1)]]);
			const result = await runChangeQueryWithExpiry(sp, target, query(oldToken), {
				returnAll: true,
				onExpiredToken: 'now',
				now: NOW,
			});
			expect(result).toEqual({
				changes: [],
				nextChangeToken: 'current',
				hasMore: false,
				tokenExpired: true,
			});
			expect(queries).toHaveLength(0);
		});

		it('does not ask SharePoint to ignore the error, and handles the error', async () => {
			const { target, queries } = fakeTarget([expiredError]);
			const result = await runChangeQueryWithExpiry(sp, target, query(recentToken), {
				returnAll: true,
				onExpiredToken: 'now',
				now: NOW,
			});
			expect(queries[0].IgnoreStartTokenNotFoundError).toBeUndefined();
			expect(result.nextChangeToken).toBe('current');
			expect(result.changes).toEqual([]);
			expect(queries).toHaveLength(1);
		});
	});

	describe('fail', () => {
		it('throws for a token that is too old', async () => {
			const { target, queries } = fakeTarget([[change(1)]]);
			await expect(
				runChangeQueryWithExpiry(sp, target, query(oldToken), {
					returnAll: true,
					onExpiredToken: 'fail',
					now: NOW,
				}),
			).rejects.toBeInstanceOf(ChangeTokenExpiredError);
			expect(queries).toHaveLength(0);
		});

		it('throws after the expired token error', async () => {
			const { target } = fakeTarget([expiredError]);
			await expect(
				runChangeQueryWithExpiry(sp, target, query(recentToken), {
					returnAll: true,
					onExpiredToken: 'fail',
					now: NOW,
				}),
			).rejects.toBeInstanceOf(ChangeTokenExpiredError);
		});
	});
});
