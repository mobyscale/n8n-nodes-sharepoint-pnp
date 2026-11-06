import { describe, expect, it } from 'vitest';

import {
	buildMoveCopyOptions,
	buildViewXml,
	escapeXml,
	findOwningList,
	getAncestorPaths,
	getDestinationPath,
	getMissingSegments,
	hasExtension,
	normalizeRow,
	parseExtensions,
} from '../../nodes/SharePointPnp/shared/filesFolders';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

describe('move and copy options', () => {
	it('maps the defaults for a copy', () => {
		expect(buildMoveCopyOptions('copy', {})).toEqual({
			overwrite: false,
			options: {
				KeepBoth: false,
				ShouldBypassSharedLocks: false,
				ResetAuthorAndCreatedOnCopy: false,
			},
		});
	});

	it('maps replace and move options', () => {
		expect(
			buildMoveCopyOptions('move', {
				ifExists: 'replace',
				bypassSharedLock: true,
				retainEditorAndModified: true,
			}),
		).toEqual({
			overwrite: true,
			options: {
				KeepBoth: false,
				ShouldBypassSharedLocks: true,
				RetainEditorAndModifiedOnMove: true,
			},
		});
	});

	it('maps keep both', () => {
		const result = buildMoveCopyOptions('copy', {
			ifExists: 'keepBoth',
			resetAuthorAndCreated: true,
		});
		expect(result.overwrite).toBe(false);
		expect(result.options.KeepBoth).toBe(true);
		expect(result.options.ResetAuthorAndCreatedOnCopy).toBe(true);
	});

	it('builds the destination path', () => {
		expect(getDestinationPath('/sites/Example/Lib/a/b.txt', '/sites/Example/Lib/c/')).toBe(
			'/sites/Example/Lib/c/b.txt',
		);
		expect(
			getDestinationPath('/sites/Example/Lib/a/b.txt', '/sites/Example/Lib/c', ' x.txt '),
		).toBe('/sites/Example/Lib/c/x.txt');
	});
});

describe('folder path helpers', () => {
	it('lists the ancestors below the site root, deepest first', () => {
		expect(getAncestorPaths('/sites/Example/Lib/a/b', siteUrl)).toEqual([
			'/sites/Example/Lib/a/b',
			'/sites/Example/Lib/a',
			'/sites/Example/Lib',
		]);
	});

	it('works for the root site', () => {
		expect(getAncestorPaths('/Lib/a', 'https://contoso.sharepoint.com')).toEqual([
			'/Lib/a',
			'/Lib',
		]);
	});

	it('lists the missing folders, shallowest first', () => {
		expect(getMissingSegments('/sites/Example/Lib/a/b/c', '/sites/Example/Lib/a')).toEqual([
			'/sites/Example/Lib/a/b',
			'/sites/Example/Lib/a/b/c',
		]);
	});

	it('finds the owning list by the longest root', () => {
		const lists = [
			{ id: '1', rootUrl: '/sites/Example/Lib' },
			{ id: '2', rootUrl: '/sites/Example/Lib2' },
			{ id: '3', rootUrl: '/sites/Example/Lib/Nested' },
		];
		expect(findOwningList('/sites/Example/Lib/a', lists)?.id).toBe('1');
		expect(findOwningList('/sites/example/lib2', lists)?.id).toBe('2');
		expect(findOwningList('/sites/Example/Lib/Nested/x', lists)?.id).toBe('3');
		expect(findOwningList('/sites/Example/Other', lists)).toBeUndefined();
	});
});

describe('CAML view builder', () => {
	it('builds a view without conditions', () => {
		const xml = buildViewXml({ include: 'all' });
		expect(xml).toContain('<View Scope="All">');
		expect(xml).not.toContain('<Where>');
		expect(xml).toContain('<RowLimit Paged="TRUE">500</RowLimit>');
	});

	it('filters files and folders by FSObjType', () => {
		expect(buildViewXml({ include: 'files' })).toContain(
			'<Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq>',
		);
		expect(buildViewXml({ include: 'folders' })).toContain('<Value Type="Integer">1</Value>');
	});

	it('combines many conditions with nested And', () => {
		const xml = buildViewXml({
			include: 'files',
			nameContains: 'a & b',
			modifiedAfter: '2026-01-01T00:00:00Z',
			modifiedBefore: '2026-02-01T00:00:00Z',
		});
		expect(xml).toContain('<Where><And><And><And>');
		expect(xml).toContain('a &amp; b');
		expect(xml).toContain('<Gt><FieldRef Name="Modified"/>');
		expect(xml).toContain('2026-01-01T00:00:00.000Z');
		expect(xml).toContain('<Lt><FieldRef Name="Modified"/>');
	});

	it('rejects an invalid date', () => {
		expect(() => buildViewXml({ include: 'all', modifiedAfter: 'nope' })).toThrow(
			'not a valid date',
		);
	});

	it('escapes XML characters', () => {
		expect(escapeXml(`<a href="x">'&'</a>`)).toBe(
			'&lt;a href=&quot;x&quot;&gt;&apos;&amp;&apos;&lt;/a&gt;',
		);
	});
});

describe('extension filter', () => {
	it('parses a list of extensions', () => {
		expect(parseExtensions('PDF, .docx; *.xlsx')).toEqual(['pdf', 'docx', 'xlsx']);
		expect(parseExtensions(undefined)).toEqual([]);
	});

	it('matches names', () => {
		expect(hasExtension('a.PDF', ['pdf'])).toBe(true);
		expect(hasExtension('a.txt', ['pdf'])).toBe(false);
		expect(hasExtension('noext', ['pdf'])).toBe(false);
		expect(hasExtension('anything', [])).toBe(true);
	});
});

describe('row normalization', () => {
	const row = {
		FileLeafRef: 'a.txt',
		FileRef: '/sites/Example/Lib/a.txt',
		FSObjType: '0',
		UniqueId: '{11111111-1111-1111-1111-111111111111}',
		ID: '7',
		File_x0020_Size: '12',
		Created: '2026-01-01T00:00:00Z',
		Modified: '2026-01-02T00:00:00Z',
		Author: [{ id: '1', title: 'Pat', email: 'pat@contoso.com' }],
		Editor: [{ id: '2', title: 'Sam', email: 'sam@contoso.com' }],
	};

	it('maps a file', () => {
		expect(normalizeRow(row, false)).toEqual({
			name: 'a.txt',
			type: 'file',
			serverRelativeUrl: '/sites/Example/Lib/a.txt',
			uniqueId: '11111111-1111-1111-1111-111111111111',
			listItemId: 7,
			size: 12,
			created: '2026-01-01T00:00:00Z',
			modified: '2026-01-02T00:00:00Z',
			author: { id: '1', name: 'Pat', email: 'pat@contoso.com' },
			editor: { id: '2', name: 'Sam', email: 'sam@contoso.com' },
		});
	});

	it('maps a folder with fields', () => {
		const entry = normalizeRow({ ...row, FSObjType: '1', Author: undefined }, true);
		expect(entry.type).toBe('folder');
		expect(entry.size).toBeNull();
		expect(entry.author).toBeNull();
		expect(entry.fields).toBeDefined();
	});
});
