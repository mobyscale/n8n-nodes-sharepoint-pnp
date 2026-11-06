import type {
	ICredentialDataDecryptedObject,
	IDataObject,
	IExecuteFunctions,
	IHttpRequestHelper,
	IHttpRequestOptions,
	IN8nHttpFullResponse,
	INode,
	INodeExecutionData,
} from 'n8n-workflow';
import { inject } from 'vitest';

import { SharePointPnpCertificateApi } from '../../credentials/SharePointPnpCertificateApi.credentials';
import { joinPath, toServerRelativePath } from '../../nodes/SharePointPnp/shared/utils';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';
import { createSp } from '../../nodes/SharePointPnp/transport';

export interface IntegrationConfig {
	tenantId: string;
	clientId: string;
	subdomain: string;
	privateKey: string;
	passphrase: string;
	certificate: string;
	thumbprint: string;
	siteUrl: string;
	libraryPath: string;
}

declare module 'vitest' {
	export interface ProvidedContext {
		integration: IntegrationConfig | null;
	}
}

export const config = inject('integration');
export const hasConfig = config !== null;

const node: INode = {
	id: 'integration',
	name: 'SharePoint PnP',
	type: 'n8n-nodes-sharepoint-pnp.sharePointPnp',
	typeVersion: 1,
	position: [0, 0],
	parameters: {},
};

function toHeaders(headers: IDataObject | undefined): Record<string, string> {
	return Object.fromEntries(
		Object.entries(headers ?? {}).map(([key, value]) => [key, String(value)]),
	);
}

async function rawRequest(options: IHttpRequestOptions): Promise<IN8nHttpFullResponse> {
	const response = await fetch(options.url, {
		method: options.method ?? 'GET',
		headers: toHeaders(options.headers),
		body: options.body as BodyInit | undefined,
	});
	const body = Buffer.from(await response.arrayBuffer());
	const headers: IDataObject = {};
	response.headers.forEach((value, key) => {
		headers[key] = value;
	});
	if (!options.ignoreHttpStatusErrors && !response.ok) {
		throw new Error(`Request failed with status ${response.status}: ${body.toString('utf8')}`);
	}
	return { body, headers, statusCode: response.status, statusMessage: response.statusText };
}

let accessToken: string | undefined;

function getCredentials(): ICredentialDataDecryptedObject {
	if (!config) throw new Error('Integration configuration is missing');
	return { ...config } as unknown as ICredentialDataDecryptedObject;
}

async function getAccessToken(): Promise<string> {
	if (accessToken) return accessToken;
	const helper = {
		helpers: {
			httpRequest: async (options: IHttpRequestOptions) => {
				const response = await rawRequest(options);
				return JSON.parse(response.body.toString('utf8'));
			},
		},
	} as unknown as IHttpRequestHelper;
	const result = await new SharePointPnpCertificateApi().preAuthentication.call(
		helper,
		getCredentials(),
	);
	accessToken = result.accessToken as string;
	return accessToken;
}

async function httpRequestWithAuthentication(
	_credentialType: string,
	options: IHttpRequestOptions,
): Promise<IN8nHttpFullResponse | unknown> {
	const token = await getAccessToken();
	const response = await rawRequest({
		...options,
		headers: { ...options.headers, Authorization: `Bearer ${token}` },
	});
	return options.returnFullResponse ? response : JSON.parse(response.body.toString('utf8'));
}

function getParameter(
	parameters: IDataObject,
	name: string,
	fallback: unknown,
	extractValue: boolean,
) {
	const value = name in parameters ? parameters[name] : fallback;
	if (extractValue && value && typeof value === 'object' && '__rl' in (value as IDataObject)) {
		return (value as IDataObject).value;
	}
	return value;
}

/** A minimal n8n execution context. Parameters use the same shape as the node UI produces. */
export function createExecuteContext(
	parameters: IDataObject,
	items: INodeExecutionData[] = [{ json: {} }],
): IExecuteFunctions {
	const allParameters: IDataObject = {
		authentication: 'certificate',
		site: { __rl: true, mode: 'url', value: config?.siteUrl },
		...parameters,
	};
	return {
		getInputData: () => items,
		getNode: () => node,
		continueOnFail: () => false,
		getCredentials: async () => getCredentials(),
		getNodeParameter: (
			name: string,
			_itemIndex: number,
			fallback?: unknown,
			options?: { extractValue?: boolean },
		) => getParameter(allParameters, name, fallback, options?.extractValue ?? false),
		helpers: { httpRequestWithAuthentication },
	} as unknown as IExecuteFunctions;
}

export async function runNode(
	parameters: IDataObject,
	items?: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
	const result = await new SharePointPnp().execute.call(createExecuteContext(parameters, items));
	return result[0];
}

export const locator = (value: string, mode = 'path') => ({ __rl: true, mode, value });

export function createTestSp() {
	return createSp(createExecuteContext({}), config!.siteUrl);
}

/** Creates the only folder the integration tests may change. Returns its server-relative path. */
export async function createTestFolder(): Promise<string> {
	const libraryRoot = toServerRelativePath(config!.siteUrl, config!.libraryPath);
	const path = joinPath(libraryRoot, `__pnp-test-${Date.now()}`);
	await createTestSp().web.folders.addUsingPath(path);
	return path;
}

export async function removeTestFolder(path: string | undefined): Promise<void> {
	if (!path || !path.includes('/__pnp-test-')) return;
	await createTestSp().web.getFolderByServerRelativePath(path).delete();
}
