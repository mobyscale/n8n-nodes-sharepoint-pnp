import type { ILoadOptionsFunctions, INodeProperties } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import {
	parseSearchPath,
	searchFileVersions,
	toFileResults,
	toVersionResults,
} from '../../nodes/SharePointPnp/methods/listSearchFiles';
import { fileResource } from '../../nodes/SharePointPnp/resources/file';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';

const tenantUrl = 'https://contoso.sharepoint.com';
const root = '/sites/Example/Shared Documents';

describe('file search path', () => {
	it('lists the content of a folder when the text ends with "/"', () => {
		expect(parseSearchPath('Shared Documents/Reports/')).toEqual({
			parent: 'Shared Documents/Reports',
			name: '',
		});
	});

	it('filters the parent folder by the last segment', () => {
		expect(parseSearchPath('Shared Documents/Reports/re')).toEqual({
			parent: 'Shared Documents/Reports',
			name: 're',
		});
	});
});

describe('file search results', () => {
	const folders = [
		{ Name: 'Reports', ServerRelativeUrl: `${root}/Reports` },
		{ Name: 'Forms', ServerRelativeUrl: `${root}/Forms` },
		{ Name: 'Archive', ServerRelativeUrl: `${root}/Archive` },
	];
	const files = [
		{ Name: 'report.pdf', ServerRelativeUrl: `${root}/report.pdf` },
		{ Name: 'notes.txt', ServerRelativeUrl: `${root}/notes.txt` },
	];

	it('lists folders first with a trailing "/", then files, without the Forms folder', () => {
		expect(toFileResults('Shared Documents', '', folders, files, tenantUrl)).toEqual([
			{
				name: 'Shared Documents/Archive/',
				value: `${root}/Archive`,
				url: `${tenantUrl}${root}/Archive`,
			},
			{
				name: 'Shared Documents/Reports/',
				value: `${root}/Reports`,
				url: `${tenantUrl}${root}/Reports`,
			},
			{
				name: 'Shared Documents/notes.txt',
				value: `${root}/notes.txt`,
				url: `${tenantUrl}${root}/notes.txt`,
			},
			{
				name: 'Shared Documents/report.pdf',
				value: `${root}/report.pdf`,
				url: `${tenantUrl}${root}/report.pdf`,
			},
		]);
	});

	it('filters by the start of the name, without case', () => {
		expect(
			toFileResults('Shared Documents', 'REP', folders, files, tenantUrl).map(
				(result) => result.name,
			),
		).toEqual(['Shared Documents/Reports/', 'Shared Documents/report.pdf']);
	});
});

describe('version search results', () => {
	const versions = [
		{ VersionLabel: '1.0', Created: '2024-01-02T10:00:00Z', CheckInComment: 'First' },
		{ VersionLabel: '10.0', Created: '2024-03-01T10:00:00Z', CheckInComment: '' },
		{ VersionLabel: '2.1', Created: '2024-02-01T10:00:00Z' },
	];

	it('lists the newest version first with the date and the comment', () => {
		expect(toVersionResults(versions)).toEqual([
			{ name: '10.0 (2024-03-01)', value: '10.0' },
			{ name: '2.1 (2024-02-01)', value: '2.1' },
			{ name: '1.0 (2024-01-02, First)', value: '1.0' },
		]);
	});

	it('filters by the start of the label', () => {
		expect(toVersionResults(versions, '1').map((result) => result.value)).toEqual(['10.0', '1.0']);
	});

	it('returns nothing when no file is selected', async () => {
		const context = {
			getNodeParameter: () => '',
		} as unknown as ILoadOptionsFunctions;
		expect(await searchFileVersions.call(context)).toEqual({ results: [] });
	});
});

describe('file pickers in the description', () => {
	const find = (name: string) =>
		fileResource.description.find((entry: INodeProperties) => entry.name === name);

	it('the File locator has a From List mode with the file search', () => {
		const modes = find('filePath')?.modes ?? [];
		expect(modes.map((mode) => mode.name)).toEqual(['list', 'path']);
		expect(modes[0].typeOptions?.searchListMethod).toBe('searchFiles');
		expect(find('filePath')?.default).toEqual({ mode: 'path', value: '' });
	});

	it('the Restore Version label is a locator with the version search', () => {
		const property = find('versionLabel');
		expect(property?.type).toBe('resourceLocator');
		expect(property?.modes?.map((mode) => mode.name)).toEqual(['list', 'label']);
		expect(property?.modes?.[0].typeOptions?.searchListMethod).toBe('searchFileVersions');
	});

	it('the node registers the new search methods', () => {
		const { listSearch } = new SharePointPnp().methods;
		expect(listSearch.searchFiles).toBeTypeOf('function');
		expect(listSearch.searchFileVersions).toBeTypeOf('function');
		expect(listSearch.searchFolders).toBeTypeOf('function');
	});
});
