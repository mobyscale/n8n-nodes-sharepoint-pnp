import { describe, expect, it } from 'vitest';

import {
	getSharePointErrorMessage,
	getSiteServerRelativeUrl,
	joinPath,
	resolveSiteUrl,
	splitPath,
	toServerRelativePath,
} from '../../nodes/SharePointPnp/shared/utils';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

describe('path utilities', () => {
	it('gets the server-relative site URL', () => {
		expect(getSiteServerRelativeUrl(siteUrl)).toBe('/sites/Example');
		expect(getSiteServerRelativeUrl('https://contoso.sharepoint.com')).toBe('/');
		expect(getSiteServerRelativeUrl('https://contoso.sharepoint.com/sites/A%20B/')).toBe(
			'/sites/A B',
		);
	});

	it('resolves site-relative, server-relative and absolute paths', () => {
		expect(toServerRelativePath(siteUrl, 'Shared Documents/A')).toBe(
			'/sites/Example/Shared Documents/A',
		);
		expect(toServerRelativePath(siteUrl, '/sites/Other/Lib/B/')).toBe('/sites/Other/Lib/B');
		expect(
			toServerRelativePath(
				siteUrl,
				'https://contoso.sharepoint.com/sites/Example/Shared%20Documents/C',
			),
		).toBe('/sites/Example/Shared Documents/C');
		expect(toServerRelativePath('https://contoso.sharepoint.com', 'Docs/A')).toBe('/Docs/A');
	});

	it('joins and splits paths', () => {
		expect(joinPath('/sites/Example/', '/Lib', 'A/')).toBe('/sites/Example/Lib/A');
		expect(splitPath('/sites/Example/Lib/A')).toEqual({ parent: '/sites/Example/Lib', name: 'A' });
		expect(splitPath('/A')).toEqual({ parent: '/', name: 'A' });
		expect(splitPath('A')).toEqual({ parent: '', name: 'A' });
	});

	it('resolves site URLs from a tenant URL', () => {
		const tenant = 'https://contoso.sharepoint.com';
		expect(resolveSiteUrl(tenant, '/sites/Example/')).toBe(siteUrl);
		expect(resolveSiteUrl(tenant, `${siteUrl}/`)).toBe(siteUrl);
		expect(resolveSiteUrl(tenant, '')).toBe(tenant);
	});
});

describe('getSharePointErrorMessage', () => {
	it('extracts the odata error message', () => {
		const message =
			'Error making HttpClient request in queryable [404] Not Found ::> {"odata.error":{"code":"-2147024894, System.IO.FileNotFoundException","message":{"lang":"en-US","value":"File Not Found."}}}';
		expect(getSharePointErrorMessage({ message })).toBe('File Not Found.');
	});

	it('returns the raw message without a body', () => {
		expect(getSharePointErrorMessage({ message: 'Network failure' })).toBe('Network failure');
	});
});
