import type { IDataObject, ILoadOptionsFunctions } from 'n8n-workflow';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { searchListItems } from '../../nodes/SharePointPnp/methods/listSearchItems';
import {
	searchGroupMembers,
	searchGroups,
	searchPrincipals,
	searchRoleDefinitions,
	searchUsers,
} from '../../nodes/SharePointPnp/methods/listSearchPrincipals';
import { isSystemAccount } from '../../nodes/SharePointPnp/shared/securables';
import {
	createExecuteContext,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

interface AssignmentJson {
	hasUniqueRoleAssignments: boolean;
	principalId: number;
	roleDefinitions: Array<{ id: number; name: string }>;
}

describe.runIf(hasConfig)('permission', () => {
	let testFolder: string | undefined;
	let visitorGroupId = 0;
	let visitorGroupTitle = '';
	let readRoleId = 0;
	let contributeRoleId = 0;

	const folderTarget = (path: string) => ({
		target: 'folder',
		folderPath: locator(path),
	});
	const visitorGroup = () => ({
		principalType: 'group',
		group: locator(String(visitorGroupId), 'id'),
	});
	const readRole = () => ({ role: locator('Read', 'name') });

	async function getAssignments(path: string): Promise<AssignmentJson[]> {
		const items = await runNode({
			resource: 'permission',
			operation: 'getAll',
			...folderTarget(path),
		});
		return items.map((item) => item.json as unknown as AssignmentJson);
	}

	async function createSubfolder(name: string): Promise<string> {
		const path = `${testFolder}/${name}`;
		await createTestSp().web.folders.addUsingPath(path);
		return path;
	}

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const sp = createTestSp();
		const visitors = await sp.web.associatedVisitorGroup.select('Id', 'Title')();
		visitorGroupId = visitors.Id;
		visitorGroupTitle = visitors.Title;
		readRoleId = (await sp.web.roleDefinitions.getByName('Read').select('Id')()).Id;
		contributeRoleId = (await sp.web.roleDefinitions.getByName('Contribute').select('Id')()).Id;
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('gets role assignments of an inheriting folder', async () => {
		const items = await getAssignments(testFolder!);
		expect(items.length).toBeGreaterThan(0);
		expect(items[0].hasUniqueRoleAssignments).toBe(false);
		expect(items[0]).not.toHaveProperty('properties');

		const detailed = await runNode({
			resource: 'permission',
			operation: 'getAll',
			...folderTarget(testFolder!),
			getAllOptions: { additionalProperties: 'Member/IsHiddenInUI' },
		});
		expect(detailed[0].json.roleDefinitions).toBeDefined();
		const properties = detailed[0].json.properties as { Member: { IsHiddenInUI?: boolean } };
		expect(properties.Member.IsHiddenInUI).toBeDefined();
	});

	it('breaks inheritance, adds, removes and resets on a subfolder', async () => {
		const path = await createSubfolder('break-add-remove');

		const broken = await runNode({
			resource: 'permission',
			operation: 'breakInheritance',
			...folderTarget(path),
			breakOptions: { copyRoleAssignments: false, clearSubscopes: true },
		});
		expect(broken[0].json.hasUniqueRoleAssignments).toBe(true);

		await runNode({
			resource: 'permission',
			operation: 'add',
			...folderTarget(path),
			...visitorGroup(),
			...readRole(),
			addOptions: { breakInheritanceFirst: false },
		});
		let assignments = await getAssignments(path);
		expect(assignments.every((a) => a.hasUniqueRoleAssignments)).toBe(true);
		const visitor = assignments.find((a) => a.principalId === visitorGroupId);
		expect(visitor?.roleDefinitions.map((r) => r.id)).toEqual([readRoleId]);

		await runNode({
			resource: 'permission',
			operation: 'remove',
			...folderTarget(path),
			...visitorGroup(),
			...readRole(),
		});
		assignments = await getAssignments(path);
		expect(assignments.find((a) => a.principalId === visitorGroupId)).toBeUndefined();

		await runNode({
			resource: 'permission',
			operation: 'add',
			...folderTarget(path),
			...visitorGroup(),
			...readRole(),
		});
		await runNode({
			resource: 'permission',
			operation: 'remove',
			...folderTarget(path),
			...visitorGroup(),
			role: locator(''),
		});
		assignments = await getAssignments(path);
		expect(assignments.find((a) => a.principalId === visitorGroupId)).toBeUndefined();

		const reset = await runNode({
			resource: 'permission',
			operation: 'resetInheritance',
			...folderTarget(path),
		});
		expect(reset[0].json.hasUniqueRoleAssignments).toBe(false);
		assignments = await getAssignments(path);
		expect(assignments[0].hasUniqueRoleAssignments).toBe(false);
	});

	it('breaks inheritance with a copy when adding to an inheriting folder', async () => {
		const path = await createSubfolder('add-breaks');

		const result = await runNode({
			resource: 'permission',
			operation: 'add',
			...folderTarget(path),
			...visitorGroup(),
			...readRole(),
		});
		expect(result[0].json.inheritanceBroken).toBe(true);

		const after = await getAssignments(path);
		expect(after[0].hasUniqueRoleAssignments).toBe(true);
		expect(after.find((a) => a.principalId === visitorGroupId)).toBeDefined();
	});

	it('sets exactly the given role assignments', async () => {
		const path = await createSubfolder('set');

		const result = await runNode({
			resource: 'permission',
			operation: 'set',
			...folderTarget(path),
			assignments: {
				entries: [{ principal: String(visitorGroupId), role: 'Read' }],
			},
			setOptions: { copyCurrentFirst: true },
		});
		expect(result[0].json.inheritanceBroken).toBe(true);

		let assignments = await getAssignments(path);
		expect(assignments).toHaveLength(1);
		expect(assignments[0].principalId).toBe(visitorGroupId);
		expect(assignments[0].roleDefinitions.map((r) => r.id)).toEqual([readRoleId]);

		const again = await runNode({
			resource: 'permission',
			operation: 'set',
			...folderTarget(path),
			assignments: {
				entries: [{ principal: String(visitorGroupId), role: String(readRoleId) }],
			},
		});
		expect(again[0].json).toMatchObject({ added: 0, removed: 0, unchanged: 1 });

		await runNode({
			resource: 'permission',
			operation: 'set',
			...folderTarget(path),
			assignments: { entries: [] },
		});
		assignments = await getAssignments(path);
		expect(assignments).toHaveLength(0);
	});

	it('adds, removes and sets role assignments from JSON', async () => {
		const path = await createSubfolder('json');
		const json = (value: unknown) => ({
			inputMode: 'json',
			roleAssignmentsJson: JSON.stringify(value),
		});
		const roleIds = async () =>
			(await getAssignments(path))
				.find((a) => a.principalId === visitorGroupId)
				?.roleDefinitions.map((r) => r.id)
				.sort();

		const added = await runNode({
			resource: 'permission',
			operation: 'add',
			...folderTarget(path),
			...json([
				{ principal: visitorGroupTitle, principalType: 'group', roles: ['Read', 'Contribute'] },
			]),
		});
		expect(added[0].json).toMatchObject({ inheritanceBroken: true, added: 2 });
		expect(await roleIds()).toEqual([readRoleId, contributeRoleId].sort());

		await runNode({
			resource: 'permission',
			operation: 'remove',
			...folderTarget(path),
			...json([{ principal: String(visitorGroupId), role: 'Contribute' }]),
		});
		expect(await roleIds()).toEqual([readRoleId]);

		const removed = await runNode({
			resource: 'permission',
			operation: 'remove',
			...folderTarget(path),
			...json([{ principal: visitorGroupTitle }]),
		});
		expect(removed[0].json.removed).toEqual([
			{ principalId: visitorGroupId, allRolesRemoved: true },
		]);
		expect(await roleIds()).toBeUndefined();

		const set = await runNode({
			resource: 'permission',
			operation: 'set',
			...folderTarget(path),
			...json([{ principal: visitorGroupTitle, role: String(readRoleId) }]),
		});
		expect(set[0].json).toMatchObject({ added: 1 });
		expect(await roleIds()).toEqual([readRoleId]);

		await runNode({
			resource: 'permission',
			operation: 'set',
			...folderTarget(path),
			...json([]),
		});
		expect(await getAssignments(path)).toHaveLength(0);
	});

	it('rejects invalid or unknown JSON before it changes anything', async () => {
		const path = await createSubfolder('json-invalid');
		await expect(
			runNode({
				resource: 'permission',
				operation: 'add',
				...folderTarget(path),
				inputMode: 'json',
				roleAssignmentsJson: '[{ "principal": "x" }]',
			}),
		).rejects.toThrow(/entry 1 needs a "role" or "roles"/);
		await expect(
			runNode({
				resource: 'permission',
				operation: 'add',
				...folderTarget(path),
				inputMode: 'json',
				roleAssignmentsJson: JSON.stringify([
					{ principal: visitorGroupTitle, role: 'Read' },
					{ principal: '__pnp-test-missing-group', role: 'Read' },
				]),
			}),
		).rejects.toThrow(/entry 2/);
		expect((await getAssignments(path))[0].hasUniqueRoleAssignments).toBe(false);
	});

	it('rejects a remove on an inheriting folder', async () => {
		const path = await createSubfolder('inherit-remove');
		await expect(
			runNode({
				resource: 'permission',
				operation: 'remove',
				...folderTarget(path),
				...visitorGroup(),
				...readRole(),
			}),
		).rejects.toThrow(/inherits/);
	});

	it('gets effective permissions of an existing user', async () => {
		const users = (await createTestSp()
			.web.siteUsers.filter('PrincipalType eq 1')
			.select('Id', 'LoginName', 'Email', 'IsHiddenInUI')()) as unknown as IDataObject[];
		const user = users.find((u) => !isSystemAccount(u) && u.Email);
		if (!user) return;

		const items = await runNode({
			resource: 'permission',
			operation: 'getEffectivePermissions',
			target: 'site',
			effectiveUser: String(user.Email),
		});
		expect(Array.isArray(items[0].json.permissions)).toBe(true);
	});
});

describe.runIf(hasConfig)('group and user', () => {
	let groupId: number | undefined;
	let groupName = '';
	let member: { Id: number; Email: string } | undefined;

	beforeAll(async () => {
		groupName = `__pnp-test-group-${Date.now()}`;
		const users = (await createTestSp()
			.web.siteUsers.filter('PrincipalType eq 1')
			.select('Id', 'LoginName', 'Email', 'IsHiddenInUI')()) as unknown as Array<{
			Id: number;
			Email: string;
			LoginName: string;
			IsHiddenInUI: boolean;
		}>;
		member = users.find((u) => !isSystemAccount(u) && u.Email);
	});

	afterAll(async () => {
		if (groupId && groupName.startsWith('__pnp-test-group-')) {
			await createTestSp().web.siteGroups.removeById(groupId);
		}
	});

	it('creates, gets and updates a group', async () => {
		const created = await runNode({
			resource: 'group',
			operation: 'create',
			title: groupName,
			additionalFields: { description: 'Temporary test group' },
		});
		groupId = Number(created[0].json.Id);
		expect(created[0].json.Title).toBe(groupName);

		const fetched = await runNode({
			resource: 'group',
			operation: 'get',
			group: locator(groupName, 'name'),
		});
		expect(fetched[0].json.Id).toBe(groupId);
		expect(fetched[0].json).toHaveProperty('OwnerTitle');

		const withOwner = await runNode({
			resource: 'group',
			operation: 'get',
			group: locator(groupName, 'name'),
			getOptions: { additionalProperties: 'Owner/Title' },
		});
		expect(withOwner[0].json.Owner).toBeDefined();
		const everything = await runNode({
			resource: 'group',
			operation: 'get',
			group: locator(groupName, 'name'),
			getOptions: { returnAllProperties: true },
		});
		expect(everything[0].json.Id).toBe(groupId);

		const updated = await runNode({
			resource: 'group',
			operation: 'update',
			group: locator(String(groupId), 'id'),
			updateFields: { description: 'Updated description', allowMembersEditMembership: true },
		});
		expect(updated[0].json.Description).toBe('Updated description');
		expect(updated[0].json.AllowMembersEditMembership).toBe(true);
	});

	it('finds the group in Get Many', async () => {
		const items = await runNode({ resource: 'group', operation: 'getAll', returnAll: true });
		expect(items.some((item) => item.json.Id === groupId)).toBe(true);

		const withOwner = await runNode({
			resource: 'group',
			operation: 'getAll',
			returnAll: true,
			getAllOptions: { additionalProperties: 'Owner/Title' },
		});
		expect(withOwner.find((item) => item.json.Id === groupId)?.json.Owner).toBeDefined();
	});

	it('adds and removes a member', async () => {
		if (!member) return;
		const group = locator(String(groupId), 'id');

		await runNode({ resource: 'group', operation: 'addMember', group, user: member.Email });
		let members = await runNode({
			resource: 'group',
			operation: 'getMembers',
			group,
			returnAll: true,
		});
		expect(members.some((item) => item.json.Id === member!.Id)).toBe(true);

		const detailed = await runNode({
			resource: 'group',
			operation: 'getMembers',
			group,
			returnAll: true,
			getMembersOptions: { additionalProperties: 'UserPrincipalName' },
		});
		expect(detailed.find((item) => item.json.Id === member!.Id)?.json).toHaveProperty(
			'UserPrincipalName',
		);

		await runNode({ resource: 'group', operation: 'removeMember', group, user: member.Email });
		members = await runNode({
			resource: 'group',
			operation: 'getMembers',
			group,
			returnAll: true,
		});
		expect(members.some((item) => item.json.Id === member!.Id)).toBe(false);
	});

	it('deletes the group', async () => {
		const result = await runNode({
			resource: 'group',
			operation: 'delete',
			group: locator(String(groupId), 'id'),
		});
		expect(result[0].json.deleted).toBe(true);
		groupId = undefined;
	});

	it('gets, ensures and lists users', async () => {
		if (!member) return;
		const byEmail = await runNode({
			resource: 'user',
			operation: 'get',
			identifyBy: 'email',
			email: member.Email,
		});
		expect(byEmail[0].json.Id).toBe(member.Id);

		const everything = await runNode({
			resource: 'user',
			operation: 'get',
			identifyBy: 'email',
			email: member.Email,
			getOptions: { returnAllProperties: true },
		});
		expect(everything[0].json.Id).toBe(member.Id);
		expect(everything[0].json).toHaveProperty('Expiration');

		const byId = await runNode({
			resource: 'user',
			operation: 'get',
			identifyBy: 'id',
			userId: member.Id,
		});
		expect(byId[0].json.Email).toBe(member.Email);

		const ensured = await runNode({ resource: 'user', operation: 'ensure', user: member.Email });
		expect(ensured[0].json.Id).toBe(member.Id);

		const many = await runNode({
			resource: 'user',
			operation: 'getAll',
			returnAll: true,
			filters: { principalType: 1, excludeSystemAccounts: true },
		});
		expect(many.length).toBeGreaterThan(0);
		expect(many.every((item) => item.json.PrincipalType === 1)).toBe(true);

		const slim = await runNode({
			resource: 'user',
			operation: 'getAll',
			returnAll: false,
			limit: 1,
			filters: {},
			getAllOptions: { additionalProperties: 'Expiration' },
		});
		expect(slim[0].json).toHaveProperty('Expiration');
	});
});

describe.runIf(hasConfig)('permission pickers and Set options', () => {
	let testFolder: string | undefined;
	let visitorGroupId = 0;

	const loadContext = (extra: Record<string, unknown> = {}) => {
		const execute = createExecuteContext(extra);
		return {
			...execute,
			getNodeParameter: (name: string, _fallback?: unknown, options?: { extractValue?: boolean }) =>
				execute.getNodeParameter(name, 0, '', options),
		} as unknown as ILoadOptionsFunctions;
	};

	const fileTarget = (path: string) => ({ target: 'file', filePath: locator(path) });

	async function createFile(name: string): Promise<string> {
		const path = `${testFolder}/${name}`;
		await createTestSp()
			.web.getFolderByServerRelativePath(testFolder!)
			.files.addUsingPath(name, 'test', { Overwrite: true });
		return path;
	}

	beforeAll(async () => {
		testFolder = await createTestFolder();
		visitorGroupId = (await createTestSp().web.associatedVisitorGroup.select('Id')()).Id;
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('lists users, principals, groups, roles and list items without changing anything', async () => {
		const users = await searchUsers.call(loadContext());
		expect(users.results.every((entry) => entry.value && entry.name)).toBe(true);
		expect(users.results.some((entry) => /^SHAREPOINT\system$/i.test(String(entry.value)))).toBe(
			false,
		);

		const principals = await searchPrincipals.call(loadContext());
		expect(principals.results.some((entry) => entry.name.startsWith('Group: '))).toBe(true);
		expect(
			principals.results.every((entry) => /^(Group|User): /.test(entry.name) && entry.value),
		).toBe(true);

		const groups = await searchGroups.call(loadContext());
		expect(groups.results.some((entry) => entry.value === String(visitorGroupId))).toBe(true);

		const roles = await searchRoleDefinitions.call(loadContext());
		expect(roles.results.some((entry) => entry.name === 'Read')).toBe(true);

		const members = await searchGroupMembers.call(
			loadContext({ group: locator(String(visitorGroupId), 'list') }),
		);
		expect(Array.isArray(members.results)).toBe(true);

		const items = await searchListItems
			.call(loadContext({ list: locator('Documents', 'title') }))
			.catch(() => ({ results: [] }));
		expect(Array.isArray(items.results)).toBe(true);
	});

	it('skips unresolved entries, breaks inheritance again and clears subscopes', async () => {
		const path = await createFile('picker-set.txt');
		const entries = JSON.stringify([
			{ principal: String(visitorGroupId), role: 'Read' },
			{ principal: '__pnp-missing-principal', role: 'Read' },
		]);

		await expect(
			runNode({
				resource: 'permission',
				operation: 'set',
				...fileTarget(path),
				inputMode: 'json',
				roleAssignmentsJson: entries,
			}),
		).rejects.toThrow(/__pnp-missing-principal/);

		const [first] = await runNode({
			resource: 'permission',
			operation: 'set',
			...fileTarget(path),
			inputMode: 'json',
			roleAssignmentsJson: entries,
			setOptions: { skipUnresolved: true, clearSubscopes: true },
		});
		expect(first.json).toMatchObject({ inheritanceBroken: true, added: 1 });
		expect(first.json.warnings).toHaveLength(1);

		const [again] = await runNode({
			resource: 'permission',
			operation: 'set',
			...fileTarget(path),
			assignments: {
				entries: [
					{
						principal: locator(String(visitorGroupId), 'id'),
						role: locator('Read', 'name'),
					},
				],
			},
			setOptions: { alwaysBreakInheritance: true, copyCurrentFirst: false },
		});
		expect(again.json).toMatchObject({ inheritanceBroken: true, removed: 0 });

		const assignments = await runNode({
			resource: 'permission',
			operation: 'getAll',
			...fileTarget(path),
		});
		expect(assignments).toHaveLength(1);
		expect(assignments[0].json.principalId).toBe(visitorGroupId);
	});
});

describe.runIf(hasConfig)('permission Apply Mapping', () => {
	let testFolder: string | undefined;
	let visitorGroupId = 0;
	let visitorGroupTitle = '';
	let readRoleId = 0;
	let contributeRoleId = 0;

	interface FolderJson {
		path: string;
		inheritanceBroken: boolean;
		added: number;
		removed: number;
		unchanged: number;
		created: boolean;
	}

	async function getAssignments(path: string): Promise<AssignmentJson[]> {
		const items = await runNode({
			resource: 'permission',
			operation: 'getAll',
			target: 'folder',
			folderPath: locator(path),
		});
		return items.map((item) => item.json as unknown as AssignmentJson);
	}

	beforeAll(async () => {
		testFolder = await createTestFolder();
		const sp = createTestSp();
		for (const path of ['A', 'A/X', 'A/X/Contracts', 'A/Y', 'A/Y/Contracts', 'B']) {
			await sp.web.folders.addUsingPath(`${testFolder}/${path}`);
		}
		const visitors = await sp.web.associatedVisitorGroup.select('Id', 'Title')();
		visitorGroupId = visitors.Id;
		visitorGroupTitle = visitors.Title;
		readRoleId = (await sp.web.roleDefinitions.getByName('Read').select('Id')()).Id;
		contributeRoleId = (await sp.web.roleDefinitions.getByName('Contribute').select('Id')()).Id;
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('applies a wildcard mapping, then changes nothing on a second run', async () => {
		const [first] = await runNode({
			resource: 'permission',
			operation: 'applyMapping',
			baseFolder: locator(testFolder!),
			inputMode: 'json',
			mappingsJson: JSON.stringify([
				{
					path: '*/*/Contracts',
					principal: visitorGroupTitle,
					principalType: 'group',
					role: 'Read',
				},
				{ path: 'B', principal: String(visitorGroupId), role: 'Read' },
				{ path: 'C', principal: String(visitorGroupId), role: 'Read' },
			]),
		});
		const folders = first.json.folders as unknown as FolderJson[];
		expect(folders.map((folder) => folder.path).sort()).toEqual([
			'A/X/Contracts',
			'A/Y/Contracts',
			'B',
		]);
		expect(folders.every((folder) => folder.inheritanceBroken && folder.added === 1)).toBe(true);
		expect(first.json.missing).toEqual([
			expect.objectContaining({ path: 'C', serverRelativeUrl: `${testFolder}/C` }),
		]);

		for (const path of ['A/X/Contracts', 'A/Y/Contracts', 'B']) {
			const assignments = await getAssignments(`${testFolder}/${path}`);
			expect(assignments).toHaveLength(1);
			expect(assignments[0]).toMatchObject({
				hasUniqueRoleAssignments: true,
				principalId: visitorGroupId,
			});
			expect(assignments[0].roleDefinitions.map((role) => role.id)).toEqual([readRoleId]);
		}
		expect((await getAssignments(`${testFolder}/A/X`))[0].hasUniqueRoleAssignments).toBe(false);

		const again = await runNode({
			resource: 'permission',
			operation: 'applyMapping',
			baseFolder: locator(testFolder!),
			inputMode: 'fields',
			mappings: {
				entries: [
					{
						path: '*/*/contracts',
						principal: locator(String(visitorGroupId), 'id'),
						role: locator('Read', 'name'),
					},
				],
			},
			mappingOptions: { output: 'perFolder' },
		});
		expect(again).toHaveLength(2);
		for (const item of again) {
			expect(item.json).toMatchObject({
				inheritanceBroken: false,
				added: 0,
				removed: 0,
				unchanged: 1,
				missing: false,
			});
		}
	});

	it('creates a missing folder and keeps the copied roles in Add mode', async () => {
		const [result] = await runNode({
			resource: 'permission',
			operation: 'applyMapping',
			baseFolder: locator(testFolder!),
			inputMode: 'json',
			mappingsJson: JSON.stringify([
				{ path: 'B/New', principal: visitorGroupTitle, roles: ['Read', 'Contribute'] },
			]),
			mappingMode: 'add',
			mappingOptions: { createMissingFolders: true },
		});
		// B/New inherits Read from B. Add mode copies it when it breaks inheritance.
		expect(result.json.folders).toEqual([
			expect.objectContaining({
				path: 'B/New',
				created: true,
				inheritanceBroken: true,
				added: 1,
				removed: 0,
				unchanged: 1,
			}),
		]);
		const assignments = await getAssignments(`${testFolder}/B/New`);
		expect(
			assignments
				.find((a) => a.principalId === visitorGroupId)
				?.roleDefinitions.map((r) => r.id)
				.sort(),
		).toEqual([readRoleId, contributeRoleId].sort());
	});

	it('fails before any change when a principal cannot be resolved', async () => {
		await expect(
			runNode({
				resource: 'permission',
				operation: 'applyMapping',
				baseFolder: locator(testFolder!),
				inputMode: 'json',
				mappingsJson: JSON.stringify([
					{ path: 'A', principal: String(visitorGroupId), role: 'Read' },
					{ path: 'A', principal: '__pnp-test-missing-group', role: 'Read' },
				]),
			}),
		).rejects.toThrow(/Mapping 2 "__pnp-test-missing-group"/);
		expect((await getAssignments(`${testFolder}/A`))[0].hasUniqueRoleAssignments).toBe(false);
	});
});
