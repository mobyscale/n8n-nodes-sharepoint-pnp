import type { IExecuteFunctions, IN8nHttpFullResponse } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

vi.mock('n8n-workflow', async (importOriginal) => ({
	...(await importOriginal<typeof import('n8n-workflow')>()),
	sleep: vi.fn(async () => {}),
}));

import {
	getRetryDelayMs,
	normalizeSubdomain,
	sendThroughN8n,
	toFetchResponse,
} from '../../nodes/SharePointPnp/transport';

function createContext(responses: Array<IN8nHttpFullResponse | Error>) {
	const httpRequestWithAuthentication = vi.fn(async () => {
		const next = responses.shift();
		if (next instanceof Error) throw next;
		return next;
	});
	const ctx = {
		getInputData: () => [],
		getNode: () => ({
			name: 'SharePoint PnP',
			type: 'sharePointPnp',
			typeVersion: 1,
			parameters: {},
		}),
		helpers: { httpRequestWithAuthentication },
	} as unknown as IExecuteFunctions;
	return { ctx, httpRequestWithAuthentication };
}

const ok = (body: string, statusCode = 200): IN8nHttpFullResponse => ({
	body: Buffer.from(body),
	headers: { 'content-type': 'application/json' },
	statusCode,
});

describe('transport', () => {
	it('normalizes subdomains', () => {
		expect(normalizeSubdomain('contoso')).toBe('contoso');
		expect(normalizeSubdomain('https://contoso.sharepoint.com/sites/x')).toBe('contoso');
		expect(normalizeSubdomain(' contoso.sharepoint.com ')).toBe('contoso');
	});

	it('uses Retry-After when present', () => {
		expect(getRetryDelayMs('3', 0)).toBe(3000);
		expect(getRetryDelayMs(undefined, 2)).toBe(4000);
		expect(getRetryDelayMs('100000', 0)).toBe(120_000);
	});

	it('converts n8n responses to fetch responses', async () => {
		const response = toFetchResponse(ok('{"a":1}'));
		expect(response.status).toBe(200);
		expect(await response.json()).toEqual({ a: 1 });
		expect(toFetchResponse({ body: undefined, headers: {}, statusCode: 204 }).body).toBeNull();
	});

	it('sends the request through n8n with authentication', async () => {
		const { ctx, httpRequestWithAuthentication } = createContext([ok('{"value":[]}')]);
		const response = await sendThroughN8n(ctx, 'cred', 'https://contoso.sharepoint.com/_api/web', {
			method: 'post',
			headers: { Accept: 'application/json' },
			body: '{"x":1}',
		});

		expect(response.status).toBe(200);
		expect(httpRequestWithAuthentication).toHaveBeenCalledWith('cred', {
			method: 'POST',
			url: 'https://contoso.sharepoint.com/_api/web',
			headers: { accept: 'application/json' },
			body: '{"x":1}',
			encoding: 'arraybuffer',
			returnFullResponse: true,
			ignoreHttpStatusErrors: true,
		});
	});

	it('retries throttled requests', async () => {
		const throttled: IN8nHttpFullResponse = {
			body: '',
			headers: { 'retry-after': '1' },
			statusCode: 429,
		};
		const { ctx, httpRequestWithAuthentication } = createContext([throttled, throttled, ok('{}')]);
		const response = await sendThroughN8n(ctx, 'cred', 'https://x/_api/web', {});
		expect(response.status).toBe(200);
		expect(httpRequestWithAuthentication).toHaveBeenCalledTimes(3);
	});

	it('sends a 401 again without ignoring status errors so n8n can renew the token', async () => {
		const { ctx, httpRequestWithAuthentication } = createContext([
			{ body: '', headers: {}, statusCode: 401 },
			ok('{}'),
		]);
		await sendThroughN8n(ctx, 'cred', 'https://x/_api/web', {});
		expect(httpRequestWithAuthentication).toHaveBeenLastCalledWith(
			'cred',
			expect.objectContaining({ ignoreHttpStatusErrors: false }),
		);
	});

	it('returns SharePoint errors as responses for PnPjs to parse', async () => {
		const { ctx } = createContext([ok('{"odata.error":{}}', 404)]);
		const response = await sendThroughN8n(ctx, 'cred', 'https://x/_api/web', {});
		expect(response.status).toBe(404);
	});
});
