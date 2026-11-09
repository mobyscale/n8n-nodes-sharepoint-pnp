import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';

import '@pnp/sp/profiles/index.js';
import type { SPFI } from '@pnp/sp/index.js';

import { getSiteUrl } from '../shared/descriptions';
import { isSystemAccount } from '../shared/securables';
import { pageByOffset, parseOffset, sortAndCap, sortByName } from '../shared/paging';
import { createSp } from '../transport';
import { locatorValue } from '../shared/pickers';

function matches(value: string, filter?: string): boolean {
	return !filter || value.toLowerCase().includes(filter.toLowerCase());
}

async function getGroupEntries(
	sp: SPFI,
	filter?: string,
	prefix = '',
): Promise<INodeListSearchItems[]> {
	const groups = await sp.web.siteGroups.select('Id', 'Title')();
	return sortByName(
		groups
			.filter((group) => matches(group.Title, filter))
			.map((group) => ({ name: `${prefix}${group.Title}`, value: String(group.Id) })),
	);
}

/** Site groups, sorted A-Z. The list pages by offset. */
export async function searchGroups(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	return pageByOffset(await getGroupEntries(sp, filter), paginationToken);
}

export async function searchRoleDefinitions(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const roles = await sp.web.roleDefinitions.select('Id', 'Name', 'Hidden', 'Order')();
	const results: INodeListSearchItems[] = sortAndCap(
		roles
			.filter((role) => !role.Hidden && matches(role.Name, filter))
			.map((role) => ({ name: role.Name, value: String(role.Id) })),
	);
	return { results };
}

interface SiteUser {
	Id: number;
	Title: string;
	LoginName: string;
	Email?: string;
	IsHiddenInUI?: boolean;
}

/** The value of a user entry. An email address or a login name works in every operation. */
function userValue(user: { Email?: string; LoginName: string }): string {
	return user.Email?.trim() || user.LoginName;
}

function userLabel(title: string | undefined, value: string, email?: string): string {
	const name = title?.trim() || value;
	return email && name !== email ? `${name} (${email})` : name;
}

function dedupeResults(items: INodeListSearchItems[]): INodeListSearchItems[] {
	const seen = new Set<string>();
	return items.filter((item) => {
		const key = String(item.value).toLowerCase();
		if (seen.has(key)) return false;
		seen.add(key);
		return true;
	});
}

async function getSiteUserEntries(sp: SPFI, filter?: string): Promise<INodeListSearchItems[]> {
	const users = (await sp.web.siteUsers
		.filter('PrincipalType eq 1')
		.select('Id', 'Title', 'LoginName', 'Email', 'IsHiddenInUI')
		.top(5000)()) as unknown as SiteUser[];
	return users
		.filter((user) => !isSystemAccount(user))
		.map((user) => {
			const value = userValue(user);
			return { name: userLabel(user.Title, value, user.Email), value };
		})
		.filter((entry) => matches(entry.name, filter) || matches(String(entry.value), filter))
		.sort((a, b) => a.name.localeCompare(b.name));
}

async function getDirectoryEntries(sp: SPFI, filter: string): Promise<INodeListSearchItems[]> {
	try {
		const people = await sp.profiles.clientPeoplePickerSearchUser({
			QueryString: filter,
			MaximumEntitySuggestions: 20,
			PrincipalSource: 15,
			PrincipalType: 1,
		});
		return people
			.filter((person) => person.IsResolved !== false && (person.EntityData?.Email || person.Key))
			.map((person) => {
				const email = person.EntityData?.Email;
				const value = email || person.Key;
				return { name: userLabel(person.DisplayText, value, email), value };
			});
	} catch {
		return [];
	}
}

/**
 * Site users sorted A-Z, paged by offset. SharePoint has no stable sort and paging for site users
 * that also filters by text, so the list is read once per page and cut here. The first page also
 * has the people of the directory that match the search text.
 */
export async function searchUsers(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const siteUsers = await getSiteUserEntries(sp, filter);
	const page = pageByOffset(siteUsers, paginationToken);
	if (parseOffset(paginationToken) > 0 || !filter?.trim()) return page;
	const directory = await getDirectoryEntries(sp, filter.trim());
	return { ...page, results: dedupeResults([...page.results, ...directory]) };
}

/** Site groups first, then site users, each sorted A-Z. The combined list pages by offset. */
export async function searchPrincipals(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const [groupEntries, users] = await Promise.all([
		getGroupEntries(sp, filter, 'Group: '),
		getSiteUserEntries(sp, filter),
	]);
	const userEntries = users.map((user) => ({ ...user, name: `User: ${user.name}` }));
	return pageByOffset(dedupeResults([...groupEntries, ...userEntries]), paginationToken);
}

/** Members of the group in the "group" parameter. Falls back to all site users when no group is set. */
export async function searchGroupMembers(
	this: ILoadOptionsFunctions,
	filter?: string,
	paginationToken?: string,
): Promise<INodeListSearchResult> {
	const sp = createSp(this, await getSiteUrl(this));
	const group = locatorValue(this.getNodeParameter('group', ''));
	if (!group) return pageByOffset(await getSiteUserEntries(sp, filter), paginationToken);
	const query = /^\d+$/.test(group)
		? sp.web.siteGroups.getById(Number(group))
		: sp.web.siteGroups.getByName(group);
	const members = (await query.users.select(
		'Id',
		'Title',
		'LoginName',
		'Email',
		'IsHiddenInUI',
		'PrincipalType',
	)()) as unknown as Array<SiteUser & { PrincipalType?: number }>;
	const results = members
		.filter((user) => user.PrincipalType === 1 && !isSystemAccount(user))
		.map((user) => {
			const value = userValue(user);
			return { name: userLabel(user.Title, value, user.Email), value };
		})
		.filter((entry) => matches(entry.name, filter))
		.sort((a, b) => a.name.localeCompare(b.name));
	return pageByOffset(dedupeResults(results), paginationToken);
}
