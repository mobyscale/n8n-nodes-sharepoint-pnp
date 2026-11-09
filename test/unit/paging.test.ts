import type { ILoadOptionsFunctions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
	sp: {} as Record<string, unknown>,
	params: {} as Record<string, unknown>,
}));

vi.mock('../../nodes/SharePointPnp/transport', () => ({
	createSp: () => state.sp,
	getTenantUrl: async () => 'https://contoso.sharepoint.com',
}));
vi.mock('../../nodes/SharePointPnp/shared/descriptions', async (importOriginal) => ({
	...(await importOriginal<typeof import('../../nodes/SharePointPnp/shared/descriptions')>()),
	getSiteUrl: async () => 'https://contoso.sharepoint.com/sites/Example',
}));
vi.mock('../../nodes/SharePointPnp/resources/taxonomy/service', async (importOriginal) => ({
	...(await importOriginal<
		typeof import('../../nodes/SharePointPnp/resources/taxonomy/service')
	>()),
	getTerms: vi.fn(),
}));

import { searchListItems } from '../../nodes/SharePointPnp/methods/listSearchItems';
import { searchFiles, searchFileVersions } from '../../nodes/SharePointPnp/methods/listSearchFiles';
import {
	searchGroupMembers,
	searchGroups,
	searchPrincipals,
	searchUsers,
} from '../../nodes/SharePointPnp/methods/listSearchPrincipals';
import { searchListFolders } from '../../nodes/SharePointPnp/methods/listSearchScoped';
import { searchTerms } from '../../nodes/SharePointPnp/methods/listSearchTerms';
import { searchFolders } from '../../nodes/SharePointPnp/methods/listSearch';
import * as taxonomyService from '../../nodes/SharePointPnp/resources/taxonomy/service';
import {
	MAX_OPTIONS,
	PICKER_PAGE_SIZE,
	pageByOffset,
	parseOffset,
	sortAndCap,
} from '../../nodes/SharePointPnp/shared/paging';

const SET_ID = '11111111-1111-1111-1111-111111111111';

function loadContext(): ILoadOptionsFunctions {
	return {
		getNodeParameter: (name: string, fallback?: unknown) =>
			name in state.params ? state.params[name] : (fallback ?? ''),
	} as unknown as ILoadOptionsFunctions;
}

function queryOf<T>(result: T) {
	const query: Record<string, unknown> & (() => Promise<T>) = Object.assign(
		() => Promise.resolve(result),
		{} as Record<string, unknown>,
	);
	for (const name of ['filter', 'select', 'top', 'orderBy', 'expand']) query[name] = () => query;
	return query;
}

const names = (entries: Array<{ name: string }>) => entries.map((entry) => entry.name);

beforeEach(() => {
	state.params = {};
	state.sp = {};
});

describe('paging helpers', () => {
	it('reads offsets and ignores invalid tokens', () => {
		expect(parseOffset(undefined)).toBe(0);
		expect(parseOffset('50')).toBe(50);
		expect(parseOffset('abc')).toBe(0);
		expect(parseOffset('-5')).toBe(0);
	});

	it('pages with an offset token and ends without a token', () => {
		const entries = Array.from({ length: 120 }, (_, index) => ({
			name: `n${index}`,
			value: index,
		}));
		const first = pageByOffset(entries);
		expect(first.results).toHaveLength(PICKER_PAGE_SIZE);
		expect(first.paginationToken).toBe('50');
		const second = pageByOffset(entries, first.paginationToken);
		expect(second.results[0].value).toBe(50);
		expect(second.paginationToken).toBe('100');
		const last = pageByOffset(entries, second.paginationToken);
		expect(last.results).toHaveLength(20);
		expect(last.paginationToken).toBeUndefined();
	});

	it('does not return a token when the list fits one page', () => {
		const page = pageByOffset([{ name: 'a', value: 1 }]);
		expect(page.paginationToken).toBeUndefined();
	});

	it('sorts A-Z and keeps at most the cap', () => {
		const entries = Array.from({ length: MAX_OPTIONS + 20 }, (_, index) => ({
			name: `item ${String(MAX_OPTIONS + 20 - index).padStart(4, '0')}`,
		}));
		const result = sortAndCap(entries);
		expect(result).toHaveLength(MAX_OPTIONS);
		expect(result[0].name).toBe('item 0001');
		expect(result[MAX_OPTIONS - 1].name).toBe(`item ${String(MAX_OPTIONS).padStart(4, '0')}`);
		expect(sortAndCap([{ name: 'b' }, { name: 'a' }])).toEqual([{ name: 'a' }, { name: 'b' }]);
	});
});

describe('item search fallback paging', () => {
	function setup(pages: Array<Array<{ Id: number; Title: string }>>) {
		const conditions: string[] = [];
		let call = 0;
		const items = Object.assign(() => Promise.resolve(pages[call++] ?? []), {
			select: () => items,
			orderBy: () => items,
			top: () => items,
			filter: (value: string) => {
				if (value.includes('substringof')) {
					return () => Promise.reject(new Error('The attempted operation is prohibited'));
				}
				conditions.push(value);
				return items;
			},
		});
		state.params.list = { mode: 'id', value: '00000000-0000-0000-0000-000000000000' };
		state.sp = { web: { lists: { getById: () => ({ items }) } } };
		return conditions;
	}

	const fullPage = (from: number, title = 'Other') =>
		Array.from({ length: 300 }, (_, index) => ({ Id: from - index, Title: title }));

	it('returns a scan token when the page has 300 rows', async () => {
		const first = fullPage(1000);
		first[5].Title = 'Zeta report';
		first[6].Title = 'Alpha report';
		setup([first]);
		const { results, paginationToken } = await searchListItems.call(loadContext(), 'report');
		expect(names(results)).toEqual(['Alpha report (ID 994)', 'Zeta report (ID 995)']);
		expect(paginationToken).toBe('scan:701');
	});

	it('continues below the last scanned ID and ends on a short page', async () => {
		const conditions = setup([[{ Id: 700, Title: 'Old report' }]]);
		const { results, paginationToken } = await searchListItems.call(
			loadContext(),
			'report',
			'scan:701',
		);
		expect(names(results)).toEqual(['Old report (ID 700)']);
		expect(paginationToken).toBeUndefined();
		expect(conditions).toEqual(['Id lt 701']);
	});

	it('stops when the scan reaches ID 1', async () => {
		setup([fullPage(300)]);
		const { paginationToken } = await searchListItems.call(loadContext(), 'report', 'scan:301');
		expect(paginationToken).toBeUndefined();
	});
});

describe('principal paging', () => {
	const users = Array.from({ length: 120 }, (_, index) => ({
		Id: index + 10,
		Title: `User ${String(index).padStart(3, '0')}`,
		LoginName: `i:0#.f|membership|user${index}@contoso.com`,
		Email: `user${index}@contoso.com`,
	}));
	const clientPeoplePickerSearchUser = vi.fn();

	beforeEach(() => {
		clientPeoplePickerSearchUser.mockReset().mockResolvedValue([
			{
				DisplayText: 'Dir Person',
				Key: 'i:0#.f|membership|dir@contoso.com',
				IsResolved: true,
				EntityData: { Email: 'dir@contoso.com' },
			},
		]);
		state.sp = {
			web: {
				siteUsers: queryOf(users),
				siteGroups: Object.assign(
					queryOf(Array.from({ length: 60 }, (_, i) => ({ Id: i + 1, Title: `Group ${i}` }))),
					{ getById: () => ({ users: queryOf(users.map((u) => ({ ...u, PrincipalType: 1 }))) }) },
				),
			},
			profiles: { clientPeoplePickerSearchUser },
		};
	});

	it('pages site users and adds the directory only on the first page', async () => {
		const first = await searchUsers.call(loadContext(), 'user');
		expect(first.paginationToken).toBe('50');
		expect(first.results).toHaveLength(51);
		expect(first.results[50].value).toBe('dir@contoso.com');
		const second = await searchUsers.call(loadContext(), 'user', '50');
		expect(second.results).toHaveLength(50);
		expect(second.paginationToken).toBe('100');
		const last = await searchUsers.call(loadContext(), 'user', '100');
		expect(last.results).toHaveLength(20);
		expect(last.paginationToken).toBeUndefined();
		expect(clientPeoplePickerSearchUser).toHaveBeenCalledTimes(1);
	});

	it('pages groups and principals', async () => {
		const groups = await searchGroups.call(loadContext());
		expect(groups.results).toHaveLength(50);
		expect(groups.paginationToken).toBe('50');
		const rest = await searchGroups.call(loadContext(), undefined, '50');
		expect(rest.results).toHaveLength(10);
		expect(rest.paginationToken).toBeUndefined();

		const principals = await searchPrincipals.call(loadContext());
		expect(principals.results[0].name.startsWith('Group: ')).toBe(true);
		expect(principals.paginationToken).toBe('50');
		const later = await searchPrincipals.call(loadContext(), undefined, '100');
		expect(later.results[0].name.startsWith('User: ')).toBe(true);
	});

	it('pages group members', async () => {
		state.params.group = { mode: 'list', value: '5' };
		const page = await searchGroupMembers.call(loadContext(), undefined, '100');
		expect(page.results).toHaveLength(20);
		expect(page.paginationToken).toBeUndefined();
	});
});

describe('folder and file paging', () => {
	const entries = Array.from({ length: 70 }, (_, index) => ({
		Name: `Folder ${String(index).padStart(2, '0')}`,
		ServerRelativeUrl: `/sites/Example/Shared Documents/Folder ${String(index).padStart(2, '0')}`,
	}));

	it('pages the subfolders of a folder', async () => {
		state.sp = {
			web: {
				getFolderByServerRelativePath: () => ({ folders: queryOf(entries) }),
			},
		};
		const first = await searchFolders.call(loadContext(), 'Shared Documents/Fo');
		expect(first.results).toHaveLength(50);
		expect(first.paginationToken).toBe('50');
		const second = await searchFolders.call(loadContext(), 'Shared Documents/Fo', '50');
		expect(second.results).toHaveLength(20);
		expect(second.paginationToken).toBeUndefined();
	});

	it('pages folders and files together, folders first', async () => {
		state.sp = {
			web: {
				getFolderByServerRelativePath: () => ({
					folders: queryOf(entries.slice(0, 30)),
					files: queryOf(
						Array.from({ length: 30 }, (_, index) => ({
							Name: `file${String(index).padStart(2, '0')}.txt`,
							ServerRelativeUrl: `/sites/Example/Shared Documents/file${index}.txt`,
						})),
					),
				}),
			},
		};
		const first = await searchFiles.call(loadContext(), 'Shared Documents/');
		expect(first.paginationToken).toBe('50');
		const second = await searchFiles.call(loadContext(), 'Shared Documents/', '50');
		expect(second.results).toHaveLength(10);
		expect(second.paginationToken).toBeUndefined();
	});

	it('pages the folders of a list and keeps the root on the first page only', async () => {
		state.params.list = { mode: 'id', value: '00000000-0000-0000-0000-000000000000' };
		state.sp = {
			web: {
				lists: {
					getById: () => ({
						rootFolder: queryOf({ ServerRelativeUrl: '/sites/Example/Lists/Example' }),
					}),
				},
				getFolderByServerRelativePath: () => ({ folders: queryOf(entries) }),
			},
		};
		const first = await searchListFolders.call(loadContext());
		expect(first.results[0].name).toBe('(list root)');
		expect(first.paginationToken).toBe('50');
		const second = await searchListFolders.call(loadContext(), undefined, '50');
		expect(second.results).toHaveLength(20);
		expect(second.results[0].name).not.toBe('(list root)');
	});

	it('pages the versions of a file, newest first', async () => {
		state.params.filePath = 'Shared Documents/a.docx';
		const versions = Array.from({ length: 60 }, (_, index) => ({ VersionLabel: `${index + 1}.0` }));
		state.sp = {
			web: { getFileByServerRelativePath: () => ({ versions: queryOf(versions) }) },
		};
		const first = await searchFileVersions.call(loadContext());
		expect(first.results[0].value).toBe('60.0');
		expect(first.paginationToken).toBe('50');
		const second = await searchFileVersions.call(loadContext(), undefined, '50');
		expect(second.results).toHaveLength(10);
		expect(second.paginationToken).toBeUndefined();
	});
});

describe('term paging', () => {
	it('pages all terms A-Z without a cap', async () => {
		const terms = Array.from({ length: 620 }, (_, index) => ({
			id: `id-${index}`,
			name: `Term ${String(index).padStart(3, '0')}`,
			path: `Term ${String(index).padStart(3, '0')}`,
			labels: [],
			isDeprecated: false,
		})) as unknown as taxonomyService.Term[];
		vi.mocked(taxonomyService.getTerms).mockResolvedValue(terms);
		state.params = { operation: 'createTerm', termSet: { value: SET_ID } };
		let token: string | undefined;
		let total = 0;
		let pages = 0;
		do {
			const page = await searchTerms.call(loadContext(), undefined, token);
			total += page.results.length;
			token = page.paginationToken;
			pages++;
		} while (token);
		expect(total).toBe(620);
		expect(pages).toBe(13);
	});
});
