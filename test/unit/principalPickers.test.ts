import type { IExecuteFunctions, ILoadOptionsFunctions } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
	sp: {} as Record<string, unknown>,
	params: {} as Record<string, unknown>,
}));

vi.mock('../../nodes/SharePointPnp/transport', () => ({
	createSp: () => state.sp,
}));
vi.mock('../../nodes/SharePointPnp/shared/descriptions', async (importOriginal) => ({
	...(await importOriginal<typeof import('../../nodes/SharePointPnp/shared/descriptions')>()),
	getSiteUrl: async () => 'https://contoso.sharepoint.com/sites/Example',
}));

import {
	searchGroupMembers,
	searchPrincipals,
	searchUsers,
} from '../../nodes/SharePointPnp/methods/listSearchPrincipals';
import { searchListItems } from '../../nodes/SharePointPnp/methods/listSearchItems';
import { permissionResource } from '../../nodes/SharePointPnp/resources/permission';
import { locatorValue } from '../../nodes/SharePointPnp/shared/pickers';
import { readPrincipalInput, readTarget } from '../../nodes/SharePointPnp/shared/securables';

const siteUsers = [
	{
		Id: 1,
		Title: 'Ada Example',
		LoginName: 'i:0#.f|membership|ada@contoso.com',
		Email: 'ada@contoso.com',
	},
	{ Id: 2, Title: 'System Account', LoginName: 'SHAREPOINT\\system', Email: '' },
	{ Id: 3, Title: 'Hidden', LoginName: 'i:0#.f|membership|hidden@contoso.com', IsHiddenInUI: true },
	{ Id: 4, Title: 'No Mail', LoginName: 'i:0#.f|membership|nomail', Email: '' },
];

function queryOf<T>(result: T) {
	const query: Record<string, unknown> & (() => Promise<T>) = Object.assign(
		() => Promise.resolve(result),
		{} as Record<string, unknown>,
	);
	for (const name of ['filter', 'select', 'top', 'orderBy', 'expand']) query[name] = () => query;
	return query;
}

function loadContext(): ILoadOptionsFunctions {
	return {
		getNodeParameter: (name: string) => state.params[name] ?? '',
	} as unknown as ILoadOptionsFunctions;
}

describe('principal value parsing', () => {
	it('reads text from a locator or a plain string', () => {
		expect(locatorValue({ mode: 'list', value: ' a@contoso.com ' })).toBe('a@contoso.com');
		expect(locatorValue('Example Members')).toBe('Example Members');
		expect(locatorValue(undefined)).toBe('');
	});

	it('maps the locator mode to a principal type', () => {
		expect(readPrincipalInput({ mode: 'user', value: 'jdoe' })).toEqual({
			text: 'jdoe',
			type: 'user',
		});
		expect(readPrincipalInput({ mode: 'name', value: '12' })).toEqual({
			text: '12',
			type: 'group',
		});
		expect(readPrincipalInput({ mode: 'list', value: '12' })).toEqual({ text: '12' });
		expect(readPrincipalInput('Example Members')).toEqual({ text: 'Example Members' });
	});

	it('reads the item ID from a locator or a number', () => {
		const context = (itemId: unknown) =>
			({
				getNode: () => ({ name: 'SharePoint PnP' }),
				getNodeParameter: (name: string, _i: number, fallback?: unknown) =>
					name === 'itemId' ? itemId : name === 'target' ? 'listItem' : fallback,
			}) as unknown as IExecuteFunctions;
		const siteUrl = 'https://contoso.sharepoint.com/sites/Example';
		expect(readTarget.call(context({ mode: 'id', value: '7' }), siteUrl, 0).itemId).toBe(7);
		expect(readTarget.call(context(8), siteUrl, 0).itemId).toBe(8);
		expect(() => readTarget.call(context({ mode: 'id', value: 'x' }), siteUrl, 0)).toThrow();
	});
});

describe('principal search methods', () => {
	beforeEach(() => {
		state.params = {};
		state.sp = {
			web: {
				siteUsers: queryOf(siteUsers),
				siteGroups: Object.assign(queryOf([{ Id: 5, Title: 'Example Members' }]), {
					getById: () => ({ users: queryOf([{ ...siteUsers[0], PrincipalType: 1 }]) }),
					getByName: () => ({ users: queryOf([{ ...siteUsers[0], PrincipalType: 1 }]) }),
				}),
			},
			profiles: {
				clientPeoplePickerSearchUser: vi.fn().mockResolvedValue([
					{
						DisplayText: 'Ada Example',
						Key: 'i:0#.f|membership|ada@contoso.com',
						IsResolved: true,
						EntityData: { Email: 'ada@contoso.com' },
					},
					{
						DisplayText: 'Dir Person',
						Key: 'i:0#.f|membership|dir@contoso.com',
						IsResolved: true,
						EntityData: { Email: 'dir@contoso.com' },
					},
				]),
			},
		};
	});

	it('lists site users without system or hidden accounts and skips the directory without a filter', async () => {
		const { results } = await searchUsers.call(loadContext());
		expect(results).toEqual([
			{ name: 'Ada Example (ada@contoso.com)', value: 'ada@contoso.com' },
			{ name: 'No Mail', value: 'i:0#.f|membership|nomail' },
		]);
		const profiles = state.sp.profiles as {
			clientPeoplePickerSearchUser: ReturnType<typeof vi.fn>;
		};
		expect(profiles.clientPeoplePickerSearchUser).not.toHaveBeenCalled();
	});

	it('adds directory people for a filter and removes duplicates', async () => {
		const { results } = await searchUsers.call(loadContext(), 'example');
		expect(results.map((entry) => entry.value)).toEqual(['ada@contoso.com', 'dir@contoso.com']);
		const profiles = state.sp.profiles as {
			clientPeoplePickerSearchUser: ReturnType<typeof vi.fn>;
		};
		expect(profiles.clientPeoplePickerSearchUser).toHaveBeenCalledWith({
			QueryString: 'example',
			MaximumEntitySuggestions: 20,
			PrincipalSource: 15,
			PrincipalType: 1,
		});
	});

	it('still returns site users when the directory search fails', async () => {
		(
			state.sp.profiles as { clientPeoplePickerSearchUser: ReturnType<typeof vi.fn> }
		).clientPeoplePickerSearchUser.mockRejectedValue(new Error('denied'));
		const { results } = await searchUsers.call(loadContext(), 'ada');
		expect(results.map((entry) => entry.value)).toEqual(['ada@contoso.com']);
	});

	it('lists groups and users with prefixes', async () => {
		const { results } = await searchPrincipals.call(loadContext());
		expect(results[0]).toEqual({ name: 'Group: Example Members', value: '5' });
		expect(results.slice(1).map((entry) => entry.name)).toEqual([
			'User: Ada Example (ada@contoso.com)',
			'User: No Mail',
		]);
	});

	it('lists the members of the selected group', async () => {
		state.params.group = { mode: 'list', value: '5' };
		const { results } = await searchGroupMembers.call(loadContext());
		expect(results).toEqual([{ name: 'Ada Example (ada@contoso.com)', value: 'ada@contoso.com' }]);
	});

	it('finds list items by ID or title', async () => {
		state.params.list = { mode: 'id', value: '00000000-0000-0000-0000-000000000000' };
		const filter = vi.fn();
		const query = queryOf([
			{ Id: 3, Title: 'Contract' },
			{ Id: 4, FileLeafRef: 'a.docx' },
		]);
		(query as unknown as Record<string, unknown>).filter = (value: string) => {
			filter(value);
			return query;
		};
		state.sp = { web: { lists: { getById: () => ({ items: query }) } } };
		const { results } = await searchListItems.call(loadContext(), "o'k");
		expect(filter).toHaveBeenCalledWith(
			"(substringof('o''k', Title) or substringof('o''k', FileLeafRef))",
		);
		expect(results).toEqual([
			{ name: 'Contract (ID 3)', value: '3' },
			{ name: 'a.docx (ID 4)', value: '4' },
		]);
		await searchListItems.call(loadContext(), '12');
		expect(filter).toHaveBeenCalledWith('Id eq 12');
	});

	it('matches the newest items by name when the text filter is above the view threshold', async () => {
		state.params.list = { mode: 'id', value: '00000000-0000-0000-0000-000000000000' };
		const top = vi.fn();
		const query = queryOf([
			{ Id: 9, Title: 'Zeta report' },
			{ Id: 8, FileLeafRef: 'alpha report.docx' },
			{ Id: 7, Title: 'Other' },
		]);
		(query as unknown as Record<string, unknown>).filter = () =>
			Object.assign(() => Promise.reject(new Error('The attempted operation is prohibited')), {});
		(query as unknown as Record<string, unknown>).top = (value: number) => {
			top(value);
			return query;
		};
		state.sp = { web: { lists: { getById: () => ({ items: query }) } } };
		const { results, paginationToken } = await searchListItems.call(loadContext(), 'report');
		expect(top).toHaveBeenLastCalledWith(300);
		expect(results).toEqual([
			{ name: 'alpha report.docx (ID 8)', value: '8' },
			{ name: 'Zeta report (ID 9)', value: '9' },
		]);
		expect(paginationToken).toBeUndefined();
	});
});

describe('permission Set options', () => {
	const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

	function setup(parameters: Record<string, unknown>, unique: boolean) {
		const breakRoleInheritance = vi.fn().mockResolvedValue(undefined);
		const add = vi.fn().mockResolvedValue(undefined);
		const remove = vi.fn().mockResolvedValue(undefined);
		const notFound = () => Promise.reject(new Error('Not found'));
		const roleAssignments = Object.assign(queryOf([]), { add, remove });
		const web: Record<string, unknown> = {
			select: () => () => Promise.resolve({ HasUniqueRoleAssignments: unique }),
			breakRoleInheritance,
			roleAssignments,
			siteGroups: {
				getByName: (name: string) => ({
					select: () => (name === 'Example Members' ? () => Promise.resolve({ Id: 5 }) : notFound),
				}),
			},
			roleDefinitions: {
				getByName: (name: string) => ({
					select: () => (name === 'Read' ? () => Promise.resolve({ Id: 1073741826 }) : notFound),
				}),
			},
		};
		const sp = { web, batched: () => [{ web }, async () => undefined] };
		const context = {
			getNode: () => ({ name: 'SharePoint PnP' }),
			getNodeParameter: (name: string, _i: number, fallback?: unknown) =>
				name in parameters ? parameters[name] : fallback,
		} as unknown as IExecuteFunctions;
		return { sp, context, breakRoleInheritance, add };
	}

	const run = (parameters: Record<string, unknown>, unique: boolean) => {
		const { sp, context, breakRoleInheritance, add } = setup(
			{ inputMode: 'json', target: 'site', ...parameters },
			unique,
		);
		return {
			breakRoleInheritance,
			add,
			result: permissionResource.operations.set.call(context, {
				sp: sp as never,
				siteUrl,
				itemIndex: 0,
			}),
		};
	};

	const json = JSON.stringify([
		{ principal: 'Example Members', role: 'Read' },
		{ principal: 'Missing Group', role: 'Read' },
	]);

	it('fails on an unresolved principal by default', async () => {
		const { result } = run({ roleAssignmentsJson: json }, true);
		await expect(result).rejects.toThrow(/Missing Group/);
	});

	it('skips unresolved entries and reports warnings', async () => {
		const { result, add } = run(
			{ roleAssignmentsJson: json, setOptions: { skipUnresolved: true } },
			true,
		);
		const output = (await result) as { added: number; warnings: string[] };
		expect(output.added).toBe(1);
		expect(output.warnings).toHaveLength(1);
		expect(output.warnings[0]).toMatch(/Entry 2 was skipped/);
		expect(add).toHaveBeenCalledWith(5, 1073741826);
	});

	it('breaks inheritance again with Always Break Inheritance and passes Clear Subscopes', async () => {
		const { result, breakRoleInheritance } = run(
			{
				roleAssignmentsJson: '[]',
				setOptions: { alwaysBreakInheritance: true, clearSubscopes: true },
			},
			true,
		);
		const output = (await result) as { inheritanceBroken: boolean; warnings?: string[] };
		expect(breakRoleInheritance).toHaveBeenCalledWith(false, true);
		expect(output.inheritanceBroken).toBe(true);
		expect(output.warnings).toBeUndefined();
	});

	it('does not break unique permissions by default', async () => {
		const { result, breakRoleInheritance } = run({ roleAssignmentsJson: '[]' }, true);
		await result;
		expect(breakRoleInheritance).not.toHaveBeenCalled();
	});

	it('reads locator rows in Fields mode', async () => {
		const { result, add } = run(
			{
				inputMode: 'fields',
				assignments: {
					entries: [
						{
							principal: { mode: 'name', value: 'Example Members' },
							role: { mode: 'name', value: 'Read' },
						},
					],
				},
			},
			true,
		);
		expect(((await result) as { added: number }).added).toBe(1);
		expect(add).toHaveBeenCalledWith(5, 1073741826);
	});
});
