import { trimTrailingSlash } from '../../shared/utils';

export const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

export type NewSiteType = 'communication' | 'team';

export interface NewSiteSettings {
	siteType: NewSiteType;
	title: string;
	url: string;
	owner?: string;
	language?: number;
	description?: string;
	classification?: string;
	siteDesignId?: string;
	hubSiteId?: string;
	shareByEmailEnabled?: boolean;
}

export interface SiteCreationRequest {
	Title: string;
	Url: string;
	Lcid: number;
	ShareByEmailEnabled: boolean;
	Description: string;
	Classification: string;
	WebTemplate: 'SITEPAGEPUBLISHING#0' | 'STS#3';
	WebTemplateExtensionId: string;
	SiteDesignId: string;
	HubSiteId: string;
	Owner?: string;
}

export const SITE_STATUS_NAMES: Record<number, string> = {
	0: 'NotFound',
	1: 'Provisioning',
	2: 'Ready',
	3: 'Error',
};

/**
 * Accepts an absolute URL, a server-relative path ("/sites/Example") or only the site name ("Example",
 * placed under "/sites/"). Returns an absolute URL without a trailing slash.
 */
export function resolveNewSiteUrl(tenantUrl: string, input: string): string {
	const value = input.trim();
	if (!value) throw new Error('The site URL must not be empty');
	if (/^https?:\/\//i.test(value)) return trimTrailingSlash(value);
	if (value.startsWith('/')) return trimTrailingSlash(`${tenantUrl}${value}`);
	return trimTrailingSlash(`${tenantUrl}/sites/${value}`);
}

/** Builds the body of SPSiteManager/Create. A team site here is a modern site without a Microsoft 365 group. */
export function buildSiteCreationRequest(
	tenantUrl: string,
	settings: NewSiteSettings,
): SiteCreationRequest {
	const request: SiteCreationRequest = {
		Title: settings.title.trim(),
		Url: resolveNewSiteUrl(tenantUrl, settings.url),
		Lcid: Number(settings.language ?? 1033),
		ShareByEmailEnabled: settings.shareByEmailEnabled === true,
		Description: settings.description ?? '',
		Classification: settings.classification ?? '',
		WebTemplate: settings.siteType === 'team' ? 'STS#3' : 'SITEPAGEPUBLISHING#0',
		WebTemplateExtensionId: EMPTY_GUID,
		SiteDesignId: settings.siteDesignId?.trim() || EMPTY_GUID,
		HubSiteId: settings.hubSiteId?.trim() || EMPTY_GUID,
	};
	if (settings.owner?.trim()) request.Owner = settings.owner.trim();
	return request;
}

/** Removes the OData annotations and decodes the "_x005F_" escape that SharePoint uses in property names. */
export function cleanPropertyBag(raw: Record<string, unknown>): Record<string, unknown> {
	const result: Record<string, unknown> = {};
	for (const [key, value] of Object.entries(raw)) {
		if (key.startsWith('odata.') || key === '__metadata') continue;
		result[key.replace(/_x005f_/gi, '_')] = value;
	}
	return result;
}

export function pickPropertyBagKeys(
	bag: Record<string, unknown>,
	keys: string[],
): { properties: Record<string, unknown>; missingKeys: string[] } {
	if (!keys.length) return { properties: bag, missingKeys: [] };
	const lowered = new Map(Object.keys(bag).map((key) => [key.toLowerCase(), key]));
	const properties: Record<string, unknown> = {};
	const missingKeys: string[] = [];
	for (const key of keys) {
		const found = lowered.get(key.toLowerCase());
		if (found === undefined) missingKeys.push(key);
		else properties[found] = bag[found];
	}
	return { properties, missingKeys };
}

/** The request body to merge property bag values. */
export function buildPropertyBagBody(
	entries: Array<{ key: string; value: string }>,
): Record<string, unknown> {
	const result: Record<string, unknown> = { __metadata: { type: 'SP.PropertyValues' } };
	for (const entry of entries) {
		const key = entry.key.trim();
		if (!key) throw new Error('A property bag key must not be empty');
		result[key] = entry.value;
	}
	return result;
}

export const CUSTOM_SCRIPTS_HINT =
	'SharePoint refused to write the property bag. This can happen when "custom scripts" are not allowed on the site (DenyAddAndCustomizePages is on). A SharePoint admin can allow custom scripts for the site, or use an account or app that has full control. SharePoint message: ';

export function isAccessDenied(error: unknown): boolean {
	const status = (error as { status?: number })?.status;
	return status === 401 || status === 403;
}
