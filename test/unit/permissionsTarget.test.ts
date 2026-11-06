import type { IExecuteFunctions } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import { readTarget } from '../../nodes/SharePointPnp/shared/securables';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

function context(parameters: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: () => ({ name: 'SharePoint PnP' }),
		getNodeParameter: (
			name: string,
			_index: number,
			fallback?: unknown,
			options?: { extractValue?: boolean },
		) => {
			const value = name in parameters ? parameters[name] : fallback;
			if (options?.extractValue && typeof value === 'object' && value && 'value' in value) {
				return (value as { value: unknown }).value;
			}
			return value;
		},
	} as unknown as IExecuteFunctions;
}

describe('readTarget', () => {
	it('uses the site by default', () => {
		expect(readTarget.call(context({}), siteUrl, 0)).toEqual({ kind: 'site' });
	});

	it('reads a list', () => {
		const list = { mode: 'id', value: 'abc' };
		expect(readTarget.call(context({ target: 'list', list }), siteUrl, 0)).toEqual({
			kind: 'list',
			list,
		});
	});

	it('reads a list item', () => {
		const list = { mode: 'title', value: 'Documents' };
		expect(readTarget.call(context({ target: 'listItem', list, itemId: 4 }), siteUrl, 0)).toEqual({
			kind: 'listItem',
			list,
			itemId: 4,
		});
	});

	it('converts a site-relative folder path to a server-relative path', () => {
		const folderPath = { mode: 'path', value: 'Shared Documents/Folder' };
		expect(readTarget.call(context({ target: 'folder', folderPath }), siteUrl, 0)).toEqual({
			kind: 'folder',
			path: '/sites/Example/Shared Documents/Folder',
		});
	});

	it('converts an absolute file URL', () => {
		const filePath = {
			mode: 'path',
			value: 'https://contoso.sharepoint.com/sites/Example/Shared%20Documents/a.docx',
		};
		expect(readTarget.call(context({ target: 'file', filePath }), siteUrl, 0)).toEqual({
			kind: 'file',
			path: '/sites/Example/Shared Documents/a.docx',
		});
	});

	it('rejects an empty path', () => {
		const folderPath = { mode: 'path', value: ' ' };
		expect(() => readTarget.call(context({ target: 'folder', folderPath }), siteUrl, 0)).toThrow();
	});

	it('rejects an unknown target', () => {
		expect(() => readTarget.call(context({ target: 'web' }), siteUrl, 0)).toThrow();
	});
});
