import type {
	IDataObject,
	IExecuteFunctions,
	INode,
	INodeExecutionData,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';

export function trimTrailingSlash(value: string): string {
	return value.length > 1 ? value.replace(/\/+$/, '') : value;
}

function collapseSlashes(value: string): string {
	return value.replace(/\/{2,}/g, '/');
}

export function joinPath(...parts: string[]): string {
	const joined = parts
		.filter((part) => part !== undefined && part !== null && part !== '')
		.join('/');
	return trimTrailingSlash(collapseSlashes(joined));
}

/** Returns the server-relative URL of a site, for example "/sites/Example" or "/" for the root site. */
export function getSiteServerRelativeUrl(siteUrl: string): string {
	const pathname = decodeURIComponent(new URL(siteUrl).pathname);
	return trimTrailingSlash(pathname) || '/';
}

/**
 * Accepts an absolute URL, a server-relative path ("/sites/Example/Shared Documents/A") or a
 * site-relative path ("Shared Documents/A") and returns the server-relative path.
 */
export function toServerRelativePath(siteUrl: string, input: string): string {
	const value = input.trim();
	if (/^https?:\/\//i.test(value)) {
		return trimTrailingSlash(collapseSlashes(decodeURIComponent(new URL(value).pathname)));
	}
	if (value.startsWith('/')) {
		return trimTrailingSlash(collapseSlashes(value));
	}
	return joinPath(getSiteServerRelativeUrl(siteUrl), value);
}

export function splitPath(path: string): { parent: string; name: string } {
	const normalized = trimTrailingSlash(path);
	const index = normalized.lastIndexOf('/');
	if (index < 0) return { parent: '', name: normalized };
	return { parent: normalized.slice(0, index) || '/', name: normalized.slice(index + 1) };
}

/** Normalizes a site URL or server-relative site path into an absolute site URL without a trailing slash. */
export function resolveSiteUrl(tenantUrl: string, value: string): string {
	const input = value.trim();
	if (!input) return tenantUrl;
	if (/^https?:\/\//i.test(input)) return trimTrailingSlash(input);
	return trimTrailingSlash(`${tenantUrl}${input.startsWith('/') ? '' : '/'}${input}`);
}

export function toExecutionData(
	data: IDataObject | IDataObject[] | INodeExecutionData[],
	itemIndex: number,
): INodeExecutionData[] {
	const entries = Array.isArray(data) ? data : [data];
	return entries.map((entry) => {
		const isExecutionData =
			typeof entry === 'object' &&
			entry !== null &&
			'json' in entry &&
			typeof entry.json === 'object';
		const executionData = (isExecutionData ? entry : { json: entry }) as INodeExecutionData;
		return { ...executionData, pairedItem: { item: itemIndex } };
	});
}

interface PnPHttpError {
	isHttpRequestError?: boolean;
	status?: number;
	message: string;
}

/** Extracts SharePoint's own message from a PnPjs HttpRequestError ("... ::> {odata.error...}"). */
export function getSharePointErrorMessage(error: PnPHttpError): string {
	const separator = error.message.indexOf('::>');
	if (separator < 0) return error.message;
	const raw = error.message.slice(separator + 3).trim();
	try {
		const parsed = JSON.parse(raw) as {
			'odata.error'?: { message?: { value?: string } };
			error?: { message?: { value?: string } | string };
		};
		const odataMessage = parsed['odata.error']?.message?.value;
		const errorMessage = parsed.error?.message;
		return (
			odataMessage ?? (typeof errorMessage === 'string' ? errorMessage : errorMessage?.value) ?? raw
		);
	} catch {
		return raw || error.message;
	}
}

export function isNotFoundError(error: unknown): boolean {
	return (error as PnPHttpError)?.status === 404;
}

export function toNodeError(
	node: INode,
	error: unknown,
	itemIndex: number,
): NodeApiError | NodeOperationError {
	if (error instanceof NodeApiError || error instanceof NodeOperationError) return error;
	const httpError = error as PnPHttpError;
	if (httpError?.isHttpRequestError) {
		return new NodeApiError(node, error as JsonObject, {
			itemIndex,
			httpCode: String(httpError.status ?? ''),
			message: getSharePointErrorMessage(httpError),
		});
	}
	return new NodeOperationError(node, error as Error, { itemIndex });
}

export function getRequiredString(
	this: IExecuteFunctions,
	name: string,
	itemIndex: number,
): string {
	const value = String(this.getNodeParameter(name, itemIndex, '') ?? '').trim();
	if (!value) {
		throw new NodeOperationError(this.getNode(), `The parameter "${name}" must not be empty`, {
			itemIndex,
		});
	}
	return value;
}
