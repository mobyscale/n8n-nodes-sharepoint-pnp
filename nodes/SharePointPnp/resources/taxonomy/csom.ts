import { NodeOperationError } from 'n8n-workflow';

import { getCredentialType, sendThroughN8n, type SharePointContext } from '../../transport';

/**
 * Taxonomy (managed metadata) through the SharePoint CSOM endpoint (client.svc/ProcessQuery).
 * The REST term store API (_api/v2.1/termStore) does not accept the app-only SharePoint token,
 * but CSOM accepts it. This file builds the XML requests and reads the JSON responses.
 */

const NAMESPACE = 'http://schemas.microsoft.com/sharepoint/clientquery/2009';
const TAXONOMY_SESSION_TYPE_ID = '{981cbc68-9edc-4f8d-872f-71146fcbb84f}';
const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

export type CsomParam =
	| { type: 'String'; value: string }
	| { type: 'Number'; value: number }
	| { type: 'Boolean'; value: boolean }
	| { type: 'Guid'; value: string };

export const csomString = (value: string): CsomParam => ({ type: 'String', value });
export const csomNumber = (value: number): CsomParam => ({ type: 'Number', value });
export const csomBoolean = (value: boolean): CsomParam => ({ type: 'Boolean', value });
export const csomGuid = (value: string): CsomParam => ({ type: 'Guid', value });

export function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

export function isGuid(value: string): boolean {
	return /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i.test(value.trim());
}

export function normalizeGuid(value: string): string {
	return value
		.trim()
		.replace(/^\{|\}$/g, '')
		.toLowerCase();
}

export function isEmptyGuid(value: string | undefined | null): boolean {
	return !value || normalizeGuid(value) === EMPTY_GUID;
}

function paramXml(param: CsomParam): string {
	const text =
		param.type === 'Guid'
			? `{${normalizeGuid(param.value)}}`
			: escapeXml(typeof param.value === 'boolean' ? String(param.value) : String(param.value));
	return `<Parameter Type="${param.type}">${text}</Parameter>`;
}

function paramsXml(params: CsomParam[]): string {
	return params.length ? `<Parameters>${params.map(paramXml).join('')}</Parameters>` : '';
}

export interface QueryOptions {
	/** The object is a collection: select the properties of its items. */
	collection?: boolean;
	/** Names of child collections to load with each object, for example "Labels". */
	nested?: string[];
}

interface PathNode {
	xml: string;
	parent?: number;
}

/** Builds one ProcessQuery request. Object paths and actions share one ID counter, as CSOM requires. */
export class CsomRequest {
	private counter = 0;

	private readonly paths = new Map<number, PathNode>();

	private readonly actions: string[] = [];

	private readonly instantiated = new Set<number>();

	private nextId(): number {
		this.counter += 1;
		return this.counter;
	}

	private addPath(xml: (id: number) => string, parent?: number): number {
		const id = this.nextId();
		this.paths.set(id, { xml: xml(id), parent });
		return id;
	}

	staticMethod(name: string, typeId: string): number {
		return this.addPath((id) => `<StaticMethod Id="${id}" Name="${name}" TypeId="${typeId}" />`);
	}

	property(parent: number, name: string): number {
		return this.addPath(
			(id) => `<Property Id="${id}" ParentId="${parent}" Name="${name}" />`,
			parent,
		);
	}

	method(parent: number, name: string, params: CsomParam[] = []): number {
		return this.addPath(
			(id) =>
				`<Method Id="${id}" ParentId="${parent}" Name="${name}">${paramsXml(params)}</Method>`,
			parent,
		);
	}

	private instantiate(pathId: number): void {
		const chain: number[] = [];
		for (let current: number | undefined = pathId; current !== undefined; ) {
			chain.unshift(current);
			current = this.paths.get(current)?.parent;
		}
		for (const id of chain) {
			if (this.instantiated.has(id)) continue;
			this.instantiated.add(id);
			this.actions.push(`<ObjectPath Id="${this.nextId()}" ObjectPathId="${id}" />`);
		}
	}

	/** Runs an object-returning method (such as CreateTerm) without reading it back. */
	run(pathId: number): void {
		this.instantiate(pathId);
	}

	/** Calls a method that returns nothing (such as DeleteObject or CommitAll) on an object. */
	call(pathId: number, name: string, params: CsomParam[] = []): void {
		this.instantiate(pathId);
		this.actions.push(
			`<Method Name="${name}" Id="${this.nextId()}" ObjectPathId="${pathId}">${paramsXml(params)}</Method>`,
		);
	}

	setProperty(pathId: number, name: string, param: CsomParam): void {
		this.instantiate(pathId);
		this.actions.push(
			`<SetProperty Id="${this.nextId()}" ObjectPathId="${pathId}" Name="${name}">${paramXml(param)}</SetProperty>`,
		);
	}

	/** Reads all scalar properties of an object, or of the items of a collection. Returns the key of the result. */
	query(pathId: number, options: QueryOptions = {}): number {
		this.instantiate(pathId);
		const nested = (options.nested ?? [])
			.map(
				(name) =>
					`<Property Name="${name}" SelectAll="true"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Property>`,
			)
			.join('');
		const id = this.nextId();
		const body = options.collection
			? `<Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties>${nested}</Properties></ChildItemQuery>`
			: `<Query SelectAllProperties="true"><Properties>${nested}</Properties></Query>`;
		this.actions.push(`<Query Id="${id}" ObjectPathId="${pathId}">${body}</Query>`);
		return id;
	}

	toXml(): string {
		const paths = [...this.paths.values()].map((node) => node.xml).join('');
		return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="n8n-nodes-sharepoint-pnp" xmlns="${NAMESPACE}"><Actions>${this.actions.join('')}</Actions><ObjectPaths>${paths}</ObjectPaths></Request>`;
	}

	/** The term store object: a specific store when an ID is given, else the default site collection store. */
	termStore(storeId?: string): { session: number; store: number } {
		const session = this.staticMethod('GetTaxonomySession', TAXONOMY_SESSION_TYPE_ID);
		const store =
			storeId && !isEmptyGuid(storeId)
				? this.method(this.property(session, 'TermStores'), 'GetById', [csomGuid(storeId)])
				: this.method(session, 'GetDefaultSiteCollectionTermStore');
		return { session, store };
	}
}

export type CsomRecord = Record<string, unknown>;

export class CsomResult {
	constructor(private readonly values: Map<number, unknown>) {}

	get(key: number): unknown {
		return this.values.get(key);
	}

	record(key: number): CsomRecord {
		const value = this.values.get(key);
		return value !== null && typeof value === 'object' ? (value as CsomRecord) : {};
	}

	items(key: number): CsomRecord[] {
		const value = this.record(key)._Child_Items_;
		return Array.isArray(value) ? (value as CsomRecord[]) : [];
	}
}

/** Reads a ProcessQuery response: a header object, then pairs of an action ID and its result. */
export function parseCsomResponse(body: unknown): CsomResult {
	if (!Array.isArray(body) || body.length === 0) {
		throw new Error('SharePoint returned an unexpected response for the taxonomy request');
	}
	const header = body[0] as { ErrorInfo?: { ErrorMessage?: string } | null };
	if (header?.ErrorInfo) {
		throw new Error(header.ErrorInfo.ErrorMessage ?? 'The taxonomy request failed');
	}
	const values = new Map<number, unknown>();
	for (let index = 1; index + 1 < body.length; index += 2) {
		values.set(Number(body[index]), body[index + 1]);
	}
	return new CsomResult(values);
}

export async function sendCsom(
	ctx: SharePointContext,
	siteUrl: string,
	request: CsomRequest,
): Promise<CsomResult> {
	const response = await sendThroughN8n(
		ctx,
		getCredentialType(ctx),
		`${siteUrl.replace(/\/+$/, '')}/_vti_bin/client.svc/ProcessQuery`,
		{
			method: 'POST',
			headers: { accept: 'application/json', 'content-type': 'text/xml' },
			body: request.toXml(),
		},
	);
	const text = await response.text();
	if (!response.ok) {
		throw new NodeOperationError(
			ctx.getNode(),
			`The taxonomy request failed with status ${response.status}: ${text.slice(0, 300)}`,
		);
	}
	let parsed: unknown;
	try {
		parsed = JSON.parse(text);
	} catch {
		throw new NodeOperationError(
			ctx.getNode(),
			'SharePoint returned a response that is not JSON for the taxonomy request',
		);
	}
	return parseCsomResponse(parsed);
}

export function parseCsomGuid(value: unknown): string {
	const text = String(value ?? '');
	const match = /^\/Guid\((.*)\)\/$/.exec(text);
	return normalizeGuid(match ? match[1] : text);
}

export function parseCsomDate(value: unknown): string | null {
	const match = /^\/Date\((-?\d+)\)\/$/.exec(String(value ?? ''));
	return match ? new Date(Number(match[1])).toISOString() : null;
}
