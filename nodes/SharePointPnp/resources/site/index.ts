import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError, sleep } from 'n8n-workflow';
import { body, headers } from '@pnp/queryable/index.js';
import { spGet, SPQueryable, spPost, type SPFI } from '@pnp/sp/index.js';
import '@pnp/sp/hubsites/index.js';
import type { ISite } from '@pnp/sp/sites/index.js';

import { userRLC } from '../../shared/securables';
import {
	applySelection,
	selectionOptions,
	toList,
	type SelectionValues,
} from '../../shared/selection';
import { getRequiredString, getSharePointErrorMessage, toNodeError } from '../../shared/utils';
import { getTenantUrl } from '../../transport';
import type { OperationContext, ResourceModule } from '../types';
import {
	buildPropertyBagBody,
	buildSiteCreationRequest,
	cleanPropertyBag,
	CUSTOM_SCRIPTS_HINT,
	EMPTY_GUID,
	isAccessDenied,
	pickPropertyBagKeys,
	resolveNewSiteUrl,
	SITE_STATUS_NAMES,
	type NewSiteType,
} from './helpers';
import { locatorValue } from '../../shared/pickers';

type CreateSiteProps = Parameters<ISite['createCommunicationSiteFromProps']>[0];

const show = (operation: string[]) => ({ show: { resource: ['site'], operation } });

const WEB_SELECT = [
	'Id',
	'Title',
	'Description',
	'Url',
	'ServerRelativeUrl',
	'Created',
	'LastItemModifiedDate',
	'WebTemplate',
	'Configuration',
	'Language',
	'SiteLogoUrl',
];

const SITE_COLLECTION_SELECT = [
	'Id',
	'Url',
	'ServerRelativeUrl',
	'GroupId',
	'HubSiteId',
	'IsHubSite',
	'Classification',
	'ReadOnly',
	'GeoLocation',
	'CompatibilityLevel',
];

const SUBSITE_SELECT = [
	'Id',
	'Title',
	'Description',
	'Url',
	'ServerRelativeUrl',
	'Created',
	'WebTemplate',
	'Language',
];

const HUB_SITE_SELECT = [
	'ID',
	'Title',
	'SiteId',
	'SiteUrl',
	'Description',
	'LogoUrl',
	'RequiresJoinApproval',
	'ParentHubSiteId',
];

const SUBSITE_TEMPLATES = [
	{ name: 'Blank Site', value: 'STS#1' },
	{ name: 'Blog', value: 'BLOG#0' },
	{ name: 'Communication Site', value: 'SITEPAGEPUBLISHING#0' },
	{ name: 'Project Site', value: 'PROJECTSITE#0' },
	{ name: 'Team Site (Classic)', value: 'STS#0' },
	{ name: 'Team Site (No Microsoft 365 Group)', value: 'STS#3' },
	{ name: 'Wiki', value: 'WIKI#0' },
];

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: ['site'] } },
		options: [
			{
				name: 'Create Site',
				value: 'createSite',
				description:
					'Create a modern Communication site or Team site without a Microsoft 365 group',
				action: 'Create a site',
			},
			{
				name: 'Create Subsite',
				value: 'createSubsite',
				description: 'Create a classic subsite in the selected site',
				action: 'Create a subsite',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get the selected site, with its site collection information',
				action: 'Get a site',
			},
			{
				name: 'Get Hub Sites',
				value: 'getHubSites',
				description: 'Get the hub sites of the tenant',
				action: 'Get hub sites',
			},
			{
				name: 'Get Property Bag',
				value: 'getPropertyBag',
				description: 'Get all property bag values of the site, or only the given keys',
				action: 'Get property bag values',
			},
			{
				name: 'Get Site Creation Status',
				value: 'getSiteCreationStatus',
				description: 'Get the creation status of a site',
				action: 'Get site creation status',
			},
			{
				name: 'Get Subsites',
				value: 'getSubsites',
				description: 'Get the subsites of the selected site',
				action: 'Get subsites',
			},
			{
				name: 'Join Hub Site',
				value: 'joinHubSite',
				description: 'Associate the selected site with a hub site',
				action: 'Join a hub site',
			},
			{
				name: 'Leave Hub Site',
				value: 'leaveHubSite',
				description: 'Remove the association of the selected site with its hub site',
				action: 'Leave a hub site',
			},
			{
				name: 'Set Property Bag Values',
				value: 'setPropertyBag',
				description: 'Set property bag values of the site',
				action: 'Set property bag values',
			},
		],
		default: 'get',
	},

	{
		displayName: 'Options',
		name: 'getOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['get']),
		options: [
			...selectionOptions(WEB_SELECT),
			{
				displayName: 'Include Site Collection',
				name: 'includeSiteCollection',
				type: 'boolean',
				default: true,
				description: `Whether to add the site collection information as "SiteCollection" (${SITE_COLLECTION_SELECT.join(', ')})`,
			},
			{
				displayName: 'Site Collection Additional Properties',
				name: 'siteAdditionalProperties',
				type: 'string',
				default: '',
				placeholder: 'e.g. ShareByEmailEnabled',
				description:
					'Comma-separated site collection properties to load in addition to the defaults',
			},
		],
	},

	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: false,
		description: 'Whether to return all results or only up to a given limit',
		displayOptions: show(['getSubsites', 'getHubSites']),
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: { minValue: 1 },
		default: 50,
		description: 'Max number of results to return',
		displayOptions: {
			show: { resource: ['site'], operation: ['getSubsites', 'getHubSites'], returnAll: [false] },
		},
	},
	{
		displayName: 'Options',
		name: 'getSubsitesOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getSubsites']),
		options: selectionOptions(SUBSITE_SELECT),
	},
	{
		displayName: 'Options',
		name: 'getHubSitesOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getHubSites']),
		options: selectionOptions(HUB_SITE_SELECT),
	},

	{
		displayName: 'Property Key Names or IDs',
		name: 'propertyKeys',
		type: 'multiOptions',
		typeOptions: { loadOptionsMethod: 'getPropertyBagKeyOptions' },
		default: [],
		description:
			'The keys to return. Empty returns all values. Keys are not case-sensitive. An expression can return a comma-separated list. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show(['getPropertyBag']),
	},
	{
		displayName: 'Values',
		name: 'propertyValues',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Value',
		default: {},
		required: true,
		description:
			'The values to set. An existing key gets the new value. This can need "custom scripts" to be allowed on the site. If they are not, SharePoint returns an access denied error.',
		displayOptions: show(['setPropertyBag']),
		options: [
			{
				displayName: 'Value',
				name: 'values',
				values: [
					{
						displayName: 'Key',
						name: 'key',
						type: 'resourceLocator',
						default: { mode: 'list', value: '' },
						description: 'The property bag key. A key that does not exist yet is created.',
						modes: [
							{
								displayName: 'From List',
								name: 'list',
								type: 'list',
								typeOptions: { searchListMethod: 'searchPropertyBagKeys', searchable: true },
							},
							{ displayName: 'By Name', name: 'name', type: 'string', placeholder: 'MyKey' },
						],
					},
					{
						displayName: 'Value',
						name: 'value',
						type: 'string',
						default: '',
						description: 'The value as text',
					},
				],
			},
		],
	},

	{
		displayName: 'Site Type',
		name: 'siteType',
		type: 'options',
		options: [
			{ name: 'Communication Site', value: 'communication' },
			{ name: 'Team Site (No Microsoft 365 Group)', value: 'team' },
		],
		default: 'communication',
		description:
			'The kind of modern site to create. A Team site here has no Microsoft 365 group. Use the Microsoft 365 group tools for a group-connected site.',
		displayOptions: show(['createSite']),
	},
	{
		displayName: 'Title',
		name: 'siteTitle',
		type: 'string',
		default: '',
		required: true,
		description: 'The title of the new site',
		displayOptions: show(['createSite']),
	},
	{
		displayName: 'URL',
		name: 'newSiteUrl',
		type: 'string',
		default: '',
		required: true,
		placeholder: 'https://contoso.sharepoint.com/sites/Example',
		description:
			'An absolute URL, a server-relative path (/sites/Example) or only the site name. The site picker above only sets the request address.',
		displayOptions: show(['createSite']),
	},
	{
		displayName: 'Options',
		name: 'createSiteOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createSite']),
		options: [
			{
				displayName: 'Allow Sharing by Email',
				name: 'shareByEmailEnabled',
				type: 'boolean',
				default: false,
				description: 'Whether members can share files by email',
			},
			{
				displayName: 'Classification',
				name: 'classification',
				type: 'string',
				default: '',
				description: 'The site classification label',
			},
			{
				displayName: 'Description',
				name: 'description',
				type: 'string',
				default: '',
				description: 'The description of the site',
			},
			{
				displayName: 'Hub Site Name or ID',
				name: 'hubSiteId',
				type: 'options',
				typeOptions: { loadOptionsMethod: 'getHubSiteOptions' },
				default: '',
				description:
					'The hub site to join when the site is ready. Empty joins no hub site. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Language Name or ID',
				name: 'language',
				type: 'options',
				typeOptions: { loadOptionsMethod: 'getInstalledLanguageOptions' },
				default: 1033,
				description:
					'The language of the site. 1033 is English (United States). Choose from the list, or specify an LCID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				...userRLC,
				displayName: 'Owner',
				name: 'owner',
				required: false,
				description:
					'The site owner. SharePoint needs it when the node uses app-only (certificate) authentication. Pick a user, or type an email address (user principal name).',
			},
			{
				displayName: 'Site Design Name or ID',
				name: 'siteDesignId',
				type: 'options',
				typeOptions: { loadOptionsMethod: 'getSiteDesignOptions' },
				default: '',
				description:
					'The site design to apply. Empty uses the default. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Wait Timeout (Seconds)',
				name: 'timeoutSeconds',
				type: 'number',
				typeOptions: { minValue: 5 },
				default: 120,
				description: 'The longest time to wait for the site. The site can still finish later.',
			},
			{
				displayName: 'Wait Until Ready',
				name: 'waitUntilReady',
				type: 'boolean',
				default: true,
				description: 'Whether to wait until SharePoint has finished to create the site',
			},
		],
	},

	{
		displayName: 'Title',
		name: 'subsiteTitle',
		type: 'string',
		default: '',
		required: true,
		description: 'The title of the new subsite',
		displayOptions: show(['createSubsite']),
	},
	{
		displayName: 'URL Name',
		name: 'subsiteUrl',
		type: 'string',
		default: '',
		required: true,
		placeholder: 'team-a',
		description: 'The last part of the subsite URL, below the selected site',
		displayOptions: show(['createSubsite']),
	},
	{
		displayName: 'Template',
		name: 'subsiteTemplate',
		type: 'options',
		options: SUBSITE_TEMPLATES,
		default: 'STS#0',
		description: 'The web template of the subsite',
		displayOptions: show(['createSubsite']),
	},
	{
		displayName: 'Options',
		name: 'createSubsiteOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createSubsite']),
		options: [
			{
				displayName: 'Description',
				name: 'description',
				type: 'string',
				default: '',
				description: 'The description of the subsite',
			},
			{
				displayName: 'Language Name or ID',
				name: 'language',
				type: 'options',
				typeOptions: { loadOptionsMethod: 'getInstalledLanguageOptions' },
				default: 1033,
				description:
					'The language of the subsite. Choose from the list, or specify an LCID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Use Same Permissions as Parent Site',
				name: 'inheritPermissions',
				type: 'boolean',
				default: true,
				description: 'Whether the subsite uses the permissions of the parent site',
			},
		],
	},

	{
		displayName: 'Site to Check',
		name: 'statusSiteUrl',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The site to check. A new site can be missing from the list. Use a URL then.',
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: { searchListMethod: 'searchSites', searchable: true },
			},
			{
				displayName: 'By URL',
				name: 'url',
				type: 'string',
				placeholder: 'https://contoso.sharepoint.com/sites/Example',
				hint: 'An absolute URL, a server-relative path or the site name',
			},
		],
		displayOptions: show(['getSiteCreationStatus']),
	},

	{
		displayName: 'Hub Site Name or ID',
		name: 'hubSiteId',
		type: 'options',
		typeOptions: { loadOptionsMethod: 'getHubSiteOptions' },
		default: '',
		required: true,
		description:
			'The hub site to join. This needs site collection admin rights. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show(['joinHubSite']),
	},
];

function selectionFrom(ctx: IExecuteFunctions, name: string, itemIndex: number): SelectionValues {
	return ctx.getNodeParameter(name, itemIndex, {}) as SelectionValues;
}

/** Calls a SPSiteManager method on the address of the selected site. */
function siteManagerQuery(sp: SPFI, baseUrl: string, path: string) {
	return SPQueryable([sp.web, baseUrl], path);
}

interface SiteCreationResponse {
	SiteId: string;
	SiteStatus: number;
	SiteUrl: string;
}

function withStatusName(response: SiteCreationResponse): IDataObject {
	return {
		...response,
		SiteStatusName: SITE_STATUS_NAMES[response.SiteStatus] ?? 'Unknown',
	} as unknown as IDataObject;
}

async function fetchSiteStatus(
	sp: SPFI,
	baseUrl: string,
	siteUrl: string,
): Promise<SiteCreationResponse> {
	const query = siteManagerQuery(sp, baseUrl, '/_api/SPSiteManager/status');
	query.query.set('url', `'${siteUrl.replace(/'/g, "''")}'`);
	return (await spGet(query)) as SiteCreationResponse;
}

async function get(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('getOptions', itemIndex, {}) as SelectionValues & {
		includeSiteCollection?: boolean;
		siteAdditionalProperties?: string;
	};
	const web = (await applySelection(sp.web, WEB_SELECT, options)()) as unknown as IDataObject;
	if (options.includeSiteCollection === false) return web;

	const siteCollection = (await applySelection(sp.site, SITE_COLLECTION_SELECT, {
		additionalProperties: options.siteAdditionalProperties,
		returnAllProperties: options.returnAllProperties,
	})()) as unknown as IDataObject;
	return { ...web, SiteCollection: siteCollection };
}

async function getSubsites(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	let query = applySelection(
		sp.web.webs,
		SUBSITE_SELECT,
		selectionFrom(this, 'getSubsitesOptions', itemIndex),
	);
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);
	return (await query()) as unknown as IDataObject[];
}

async function getHubSites(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const returnAll = this.getNodeParameter('returnAll', itemIndex) as boolean;
	let query = applySelection(
		sp.hubSites,
		HUB_SITE_SELECT,
		selectionFrom(this, 'getHubSitesOptions', itemIndex),
	);
	if (!returnAll) query = query.top(this.getNodeParameter('limit', itemIndex) as number);
	return (await query()) as unknown as IDataObject[];
}

async function getPropertyBag(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const keys = toList(this.getNodeParameter('propertyKeys', itemIndex, '') as string | string[]);
	const raw = (await sp.web.allProperties()) as Record<string, unknown>;
	const { properties, missingKeys } = pickPropertyBagKeys(cleanPropertyBag(raw), keys);
	return keys.length ? { properties, missingKeys } : { properties };
}

async function setPropertyBag(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const collection = this.getNodeParameter('propertyValues', itemIndex, {}) as {
		values?: Array<{ key: unknown; value: string }>;
	};
	const entries = (collection.values ?? [])
		.map((entry) => ({ key: locatorValue(entry.key), value: entry.value }))
		.filter((entry) => entry.key);
	if (!entries.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one property bag value', {
			itemIndex,
		});
	}

	try {
		await spPost(
			SPQueryable(sp.web, 'allproperties'),
			body(buildPropertyBagBody(entries), headers({ 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' })),
		);
	} catch (error) {
		if (isAccessDenied(error)) {
			throw new NodeOperationError(
				this.getNode(),
				CUSTOM_SCRIPTS_HINT + getSharePointErrorMessage(error as Error),
				{ itemIndex },
			);
		}
		throw toNodeError(this.getNode(), error, itemIndex);
	}
	return { success: true, keys: entries.map((entry) => entry.key.trim()) } as IDataObject;
}

async function createSite(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('createSiteOptions', itemIndex, {}) as {
		classification?: string;
		description?: string;
		hubSiteId?: string;
		language?: number;
		owner?: unknown;
		shareByEmailEnabled?: boolean;
		siteDesignId?: string;
		waitUntilReady?: boolean;
		timeoutSeconds?: number;
	};
	const tenantUrl = await getTenantUrl(this);
	const request = buildSiteCreationRequest(tenantUrl, {
		siteType: this.getNodeParameter('siteType', itemIndex) as NewSiteType,
		title: getRequiredString.call(this, 'siteTitle', itemIndex),
		url: getRequiredString.call(this, 'newSiteUrl', itemIndex),
		...options,
		owner: locatorValue(options.owner),
	});

	let response = (await sp.site.createCommunicationSiteFromProps(
		request as CreateSiteProps,
	)) as SiteCreationResponse;

	if (options.waitUntilReady !== false) {
		const deadline = Date.now() + (options.timeoutSeconds ?? 120) * 1000;
		while (response.SiteStatus === 1 && Date.now() < deadline) {
			await sleep(5000);
			response = await fetchSiteStatus(sp, siteUrl, request.Url);
		}
	}
	if (response.SiteStatus === 3) {
		throw new NodeOperationError(
			this.getNode(),
			`SharePoint could not create the site ${request.Url}`,
			{
				itemIndex,
			},
		);
	}
	return { ...withStatusName(response), RequestedUrl: request.Url };
}

async function createSubsite(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const options = this.getNodeParameter('createSubsiteOptions', itemIndex, {}) as {
		description?: string;
		language?: number;
		inheritPermissions?: boolean;
	};
	const created = await sp.web.webs.add(
		getRequiredString.call(this, 'subsiteTitle', itemIndex),
		getRequiredString.call(this, 'subsiteUrl', itemIndex).replace(/^\/+|\/+$/g, ''),
		options.description ?? '',
		this.getNodeParameter('subsiteTemplate', itemIndex) as string,
		Number(options.language ?? 1033),
		options.inheritPermissions !== false,
	);
	return created as unknown as IDataObject;
}

function getStatusSite(this: IExecuteFunctions, itemIndex: number): string {
	const value = locatorValue(this.getNodeParameter('statusSiteUrl', itemIndex, ''));
	if (!value) {
		throw new NodeOperationError(
			this.getNode(),
			'The parameter "statusSiteUrl" must not be empty',
			{
				itemIndex,
			},
		);
	}
	return value;
}

async function getSiteCreationStatus(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const target = resolveNewSiteUrl(await getTenantUrl(this), getStatusSite.call(this, itemIndex));
	return withStatusName(await fetchSiteStatus(sp, siteUrl, target));
}

async function joinHubSite(this: IExecuteFunctions, { sp, itemIndex }: OperationContext) {
	const hubSiteId = getRequiredString.call(this, 'hubSiteId', itemIndex);
	await sp.site.joinHubSite(hubSiteId);
	return (await sp.site.select('Id', 'HubSiteId', 'IsHubSite')()) as unknown as IDataObject;
}

async function leaveHubSite(this: IExecuteFunctions, { sp }: OperationContext) {
	await sp.site.joinHubSite(EMPTY_GUID);
	return (await sp.site.select('Id', 'HubSiteId', 'IsHubSite')()) as unknown as IDataObject;
}

export const siteResource: ResourceModule = {
	option: {
		name: 'Site',
		value: 'site',
		description: 'Get site information, property bag values and create sites',
	},
	description,
	operations: {
		get,
		getSubsites,
		getPropertyBag,
		setPropertyBag,
		createSite,
		createSubsite,
		getSiteCreationStatus,
		joinHubSite,
		leaveHubSite,
		getHubSites,
	},
};
