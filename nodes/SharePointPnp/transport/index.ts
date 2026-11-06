import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	IN8nHttpFullResponse,
	IPollFunctions,
	IWebhookFunctions,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, sleep } from 'n8n-workflow';
import type { TimelinePipe } from '@pnp/core/index.js';
import { DefaultParse, type Queryable } from '@pnp/queryable/index.js';
import { DefaultHeaders, DefaultInit, spfi, type SPFI } from '@pnp/sp/index.js';
import '@pnp/sp/webs/index.js';
import '@pnp/sp/sites/index.js';
import '@pnp/sp/lists/index.js';
import '@pnp/sp/items/index.js';
import '@pnp/sp/folders/index.js';
import '@pnp/sp/files/index.js';
import '@pnp/sp/fields/index.js';
import '@pnp/sp/views/index.js';
import '@pnp/sp/content-types/index.js';
import '@pnp/sp/column-defaults/index.js';
import '@pnp/sp/security/index.js';
import '@pnp/sp/site-users/index.js';
import '@pnp/sp/site-groups/index.js';
import '@pnp/sp/search/index.js';
import '@pnp/sp/recycle-bin/index.js';
import '@pnp/sp/subscriptions/index.js';
import '@pnp/sp/batching.js';

export type SharePointContext =
	| IExecuteFunctions
	| ILoadOptionsFunctions
	| IPollFunctions
	| IHookFunctions
	| IWebhookFunctions;

export const DELEGATED_CREDENTIAL_TYPE = 'sharePointPnpOAuth2Api';
export const APP_ONLY_CREDENTIAL_TYPE = 'sharePointPnpCertificateApi';

const RETRY_STATUS_CODES = new Set([429, 503, 504]);
const RETRY_ERROR_CODES = new Set([
	'ECONNRESET',
	'ETIMEDOUT',
	'ECONNABORTED',
	'EPIPE',
	'EAI_AGAIN',
]);
const MAX_RETRIES = 5;
const MAX_RETRY_DELAY_MS = 120_000;
const NULL_BODY_STATUS_CODES = new Set([101, 204, 205, 304]);

function isExecuteContext(ctx: SharePointContext): ctx is IExecuteFunctions {
	return 'getInputData' in ctx;
}

export function getCredentialType(ctx: SharePointContext): string {
	const authentication = isExecuteContext(ctx)
		? ctx.getNodeParameter('authentication', 0, 'oAuth2')
		: ctx.getNodeParameter('authentication', 'oAuth2');
	return authentication === 'certificate' ? APP_ONLY_CREDENTIAL_TYPE : DELEGATED_CREDENTIAL_TYPE;
}

export function normalizeSubdomain(subdomain: string): string {
	return subdomain
		.trim()
		.replace(/^https?:\/\//i, '')
		.replace(/\/.*$/, '')
		.replace(/\.sharepoint\.com$/i, '');
}

export async function getTenantUrl(ctx: SharePointContext): Promise<string> {
	const credentials = await ctx.getCredentials<{ subdomain: string }>(getCredentialType(ctx));
	return `https://${normalizeSubdomain(credentials.subdomain)}.sharepoint.com`;
}

async function toRequestBody(body: RequestInit['body']): Promise<Buffer | string | undefined> {
	if (body === undefined || body === null) return undefined;
	if (typeof body === 'string') return body;
	if (Buffer.isBuffer(body)) return body;
	if (body instanceof ArrayBuffer) return Buffer.from(body);
	if (ArrayBuffer.isView(body)) return Buffer.from(body.buffer, body.byteOffset, body.byteLength);
	if (body instanceof URLSearchParams) return body.toString();
	return Buffer.from(await new Response(body).arrayBuffer());
}

function toHeaderObject(headers: RequestInit['headers']): IDataObject {
	const result: IDataObject = {};
	new Headers(headers).forEach((value, key) => {
		result[key] = value;
	});
	return result;
}

function toBuffer(body: unknown): Buffer {
	if (body === undefined || body === null) return Buffer.alloc(0);
	if (Buffer.isBuffer(body)) return body;
	if (body instanceof ArrayBuffer) return Buffer.from(body);
	if (typeof body === 'string') return Buffer.from(body);
	return Buffer.from(JSON.stringify(body));
}

export function toFetchResponse(response: IN8nHttpFullResponse): Response {
	const headers = new Headers();
	for (const [key, value] of Object.entries(response.headers ?? {})) {
		if (value === undefined || value === null) continue;
		if (Array.isArray(value)) value.forEach((entry) => headers.append(key, String(entry)));
		else headers.set(key, String(value));
	}
	const body = NULL_BODY_STATUS_CODES.has(response.statusCode) ? null : toBuffer(response.body);
	return new Response(body, {
		status: response.statusCode,
		statusText: response.statusMessage ?? '',
		headers,
	});
}

export function getRetryDelayMs(retryAfter: unknown, attempt: number): number {
	const seconds = Number(Array.isArray(retryAfter) ? retryAfter[0] : retryAfter);
	const delay = Number.isFinite(seconds) && seconds > 0 ? seconds * 1000 : 1000 * 2 ** attempt;
	return Math.min(delay, MAX_RETRY_DELAY_MS);
}

function getErrorCode(error: unknown): string | undefined {
	const candidate = error as { code?: string; cause?: { code?: string } };
	return candidate?.code ?? candidate?.cause?.code;
}

/**
 * Sends a PnPjs request through n8n so that n8n owns authentication (OAuth2 refresh, certificate
 * token renewal). Status errors are returned as responses so that PnPjs can parse SharePoint's
 * error body. A 401 is sent again without that flag, because n8n only renews tokens on a thrown 401.
 */
export async function sendThroughN8n(
	ctx: SharePointContext,
	credentialType: string,
	url: string,
	init: RequestInit,
): Promise<Response> {
	const options: IHttpRequestOptions = {
		method: (init.method ?? 'GET').toUpperCase() as IHttpRequestMethods,
		url,
		headers: toHeaderObject(init.headers),
		body: await toRequestBody(init.body),
		encoding: 'arraybuffer',
		returnFullResponse: true,
		ignoreHttpStatusErrors: true,
	};

	for (let attempt = 0; ; attempt++) {
		let response: IN8nHttpFullResponse;
		try {
			response = (await ctx.helpers.httpRequestWithAuthentication.call(
				ctx,
				credentialType,
				options,
			)) as IN8nHttpFullResponse;

			if (response.statusCode === 401) {
				response = (await ctx.helpers.httpRequestWithAuthentication.call(ctx, credentialType, {
					...options,
					ignoreHttpStatusErrors: false,
				})) as IN8nHttpFullResponse;
			}
		} catch (error) {
			const code = getErrorCode(error);
			if (code && RETRY_ERROR_CODES.has(code) && attempt < MAX_RETRIES) {
				await sleep(getRetryDelayMs(undefined, attempt));
				continue;
			}
			throw new NodeApiError(ctx.getNode(), error as JsonObject);
		}

		if (RETRY_STATUS_CODES.has(response.statusCode) && attempt < MAX_RETRIES) {
			await sleep(getRetryDelayMs(response.headers?.['retry-after'], attempt));
			continue;
		}

		return toFetchResponse(response);
	}
}

export function N8nSend(ctx: SharePointContext, credentialType: string): TimelinePipe<Queryable> {
	return (instance: Queryable) => {
		instance.on.send.replace(async (url: URL, init: RequestInit) => {
			return await sendThroughN8n(ctx, credentialType, url.toString(), init);
		});
		return instance;
	};
}

export function createSp(ctx: SharePointContext, siteUrl: string): SPFI {
	return spfi(siteUrl).using(
		DefaultHeaders(),
		DefaultInit(),
		N8nSend(ctx, getCredentialType(ctx)),
		DefaultParse(),
	);
}
