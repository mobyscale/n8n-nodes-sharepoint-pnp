import { describe, expect, it } from 'vitest';

import {
	buildPropertyBagBody,
	buildSiteCreationRequest,
	cleanPropertyBag,
	EMPTY_GUID,
	isAccessDenied,
	pickPropertyBagKeys,
	resolveNewSiteUrl,
} from '../../nodes/SharePointPnp/resources/site/helpers';

const tenant = 'https://contoso.sharepoint.com';

describe('resolveNewSiteUrl', () => {
	it('keeps an absolute URL and removes the trailing slash', () => {
		expect(resolveNewSiteUrl(tenant, 'https://contoso.sharepoint.com/sites/Example/')).toBe(
			'https://contoso.sharepoint.com/sites/Example',
		);
	});

	it('adds the tenant to a server-relative path', () => {
		expect(resolveNewSiteUrl(tenant, '/teams/Example')).toBe(
			'https://contoso.sharepoint.com/teams/Example',
		);
	});

	it('places a plain name under /sites/', () => {
		expect(resolveNewSiteUrl(tenant, ' Example ')).toBe(
			'https://contoso.sharepoint.com/sites/Example',
		);
	});

	it('rejects an empty value', () => {
		expect(() => resolveNewSiteUrl(tenant, ' ')).toThrow();
	});
});

describe('buildSiteCreationRequest', () => {
	it('builds a Communication site request with defaults', () => {
		expect(
			buildSiteCreationRequest(tenant, { siteType: 'communication', title: ' Example ', url: 'Example' }),
		).toEqual({
			Title: 'Example',
			Url: 'https://contoso.sharepoint.com/sites/Example',
			Lcid: 1033,
			ShareByEmailEnabled: false,
			Description: '',
			Classification: '',
			WebTemplate: 'SITEPAGEPUBLISHING#0',
			WebTemplateExtensionId: EMPTY_GUID,
			SiteDesignId: EMPTY_GUID,
			HubSiteId: EMPTY_GUID,
		});
	});

	it('uses the STS#3 template for a Team site without a group, and passes all settings', () => {
		const request = buildSiteCreationRequest(tenant, {
			siteType: 'team',
			title: 'Example',
			url: '/sites/Example',
			owner: ' user@contoso.com ',
			language: 1036,
			description: 'About',
			classification: 'Internal',
			siteDesignId: '6142d2a0-63a5-4ba0-aede-d9fefca2c767',
			hubSiteId: '11111111-1111-1111-1111-111111111111',
			shareByEmailEnabled: true,
		});
		expect(request).toMatchObject({
			WebTemplate: 'STS#3',
			Owner: 'user@contoso.com',
			Lcid: 1036,
			Description: 'About',
			Classification: 'Internal',
			SiteDesignId: '6142d2a0-63a5-4ba0-aede-d9fefca2c767',
			HubSiteId: '11111111-1111-1111-1111-111111111111',
			ShareByEmailEnabled: true,
		});
	});

	it('leaves out the owner when it is empty', () => {
		const request = buildSiteCreationRequest(tenant, {
			siteType: 'communication',
			title: 'Example',
			url: 'Example',
			owner: ' ',
		});
		expect(request).not.toHaveProperty('Owner');
	});
});

describe('property bag helpers', () => {
	it('removes OData annotations and decodes the underscore escape', () => {
		expect(
			cleanPropertyBag({
				'odata.metadata': 'x',
				__metadata: {},
				vti_x005F_title: 'A',
				MyKey: 'B',
			}),
		).toEqual({ vti_title: 'A', MyKey: 'B' });
	});

	it('returns all values when no key is given', () => {
		expect(pickPropertyBagKeys({ A: '1' }, [])).toEqual({
			properties: { A: '1' },
			missingKeys: [],
		});
	});

	it('picks keys without case sensitivity and reports missing keys', () => {
		expect(pickPropertyBagKeys({ MyKey: '1', Other: '2' }, ['mykey', 'Nope'])).toEqual({
			properties: { MyKey: '1' },
			missingKeys: ['Nope'],
		});
	});

	it('builds the merge body with the property values type', () => {
		expect(buildPropertyBagBody([{ key: ' A ', value: '1' }])).toEqual({
			__metadata: { type: 'SP.PropertyValues' },
			A: '1',
		});
	});

	it('rejects an empty key', () => {
		expect(() => buildPropertyBagBody([{ key: ' ', value: '1' }])).toThrow();
	});

	it('detects access denied errors', () => {
		expect(isAccessDenied({ status: 403 })).toBe(true);
		expect(isAccessDenied({ status: 404 })).toBe(false);
		expect(isAccessDenied(undefined)).toBe(false);
	});
});
