import type { IExecuteFunctions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import {
	readMappingOptions,
	readMappings,
	toMappingOutput,
} from '../../nodes/SharePointPnp/resources/permission';
import {
	applyPermissionMapping,
	parseMappingPath,
	parseMappingsJson,
	type MappingGateway,
	type MappingInput,
	type MappingOptions,
} from '../../nodes/SharePointPnp/resources/permission/mapping';
import type { RoleAssignmentEntry } from '../../nodes/SharePointPnp/shared/securables';

const BASE = '/sites/Example/Shared Documents/Archive';
const READ = 1073741826;
const CONTRIBUTE = 1073741827;
const FULL = 1073741829;
const OWNERS = 3;
const MEMBERS = 5;
const VISITORS = 6;

const PRINCIPALS: Record<string, number> = {
	'example members': MEMBERS,
	'example visitors': VISITORS,
	'example owners': OWNERS,
};
const ROLES: Record<string, number> = {
	read: READ,
	contribute: CONTRIBUTE,
	'full control': FULL,
};

const INHERITED: RoleAssignmentEntry[] = [{ principalId: OWNERS, roleDefinitionId: FULL }];

const lower = (value: string) => value.toLowerCase();

/** An in-memory folder tree below BASE. Folders inherit until the test breaks inheritance. */
function createFake(relativeFolders: string[]) {
	const folders = new Map<string, string>();
	const addFolder = (path: string) => folders.set(lower(path), path);
	addFolder(BASE);
	for (const relative of relativeFolders) {
		const parts = relative.split('/');
		parts.forEach((_, index) => addFolder(`${BASE}/${parts.slice(0, index + 1).join('/')}`));
	}
	const unique = new Map<string, boolean>();
	const assignments = new Map<string, RoleAssignmentEntry[]>();
	const changes: string[] = [];

	const gateway = {
		getFolder: vi.fn(async (path: string) => {
			const canonical = folders.get(lower(path));
			return canonical ? { serverRelativeUrl: canonical } : undefined;
		}),
		listChildFolders: vi.fn(async (path: string) => {
			if (!folders.has(lower(path))) return undefined;
			return [...folders.values()]
				.filter((folder) => {
					const index = folder.lastIndexOf('/');
					return lower(folder.slice(0, index)) === lower(path);
				})
				.map((folder) => ({
					name: folder.slice(folder.lastIndexOf('/') + 1),
					serverRelativeUrl: folder,
				}));
		}),
		ensureFolder: vi.fn(async (path: string) => {
			if (folders.has(lower(path))) return false;
			changes.push(`create ${path}`);
			const relative = path.slice(BASE.length + 1).split('/');
			relative.forEach((_, index) =>
				addFolder(`${BASE}/${relative.slice(0, index + 1).join('/')}`),
			);
			return true;
		}),
		resolvePrincipal: vi.fn(async (principal: string) => {
			const id = PRINCIPALS[lower(principal)];
			if (id === undefined) throw new Error('Group cannot be found');
			return id;
		}),
		resolveRole: vi.fn(async (role: string | number) => {
			const id = ROLES[lower(String(role))];
			if (id === undefined) throw new Error('Role cannot be found');
			return id;
		}),
		hasUniqueRoleAssignments: vi.fn(async (path: string) => unique.get(lower(path)) ?? false),
		breakInheritance: vi.fn(async (path: string, copy: boolean) => {
			changes.push(`break ${path} copy=${copy}`);
			unique.set(lower(path), true);
			assignments.set(lower(path), copy ? [...INHERITED] : []);
		}),
		getAssignments: vi.fn(async (path: string) => [...(assignments.get(lower(path)) ?? INHERITED)]),
		applyChanges: vi.fn(
			async (path: string, toAdd: RoleAssignmentEntry[], toRemove: RoleAssignmentEntry[]) => {
				changes.push(`apply ${path} +${toAdd.length} -${toRemove.length}`);
				const key = (entry: RoleAssignmentEntry) =>
					`${entry.principalId}:${entry.roleDefinitionId}`;
				const removed = new Set(toRemove.map(key));
				assignments.set(lower(path), [
					...(assignments.get(lower(path)) ?? []).filter((entry) => !removed.has(key(entry))),
					...toAdd,
				]);
			},
		),
	} satisfies MappingGateway;

	return {
		gateway,
		changes,
		assignmentsOf: (relative: string) => assignments.get(lower(`${BASE}/${relative}`)),
		setUnique: (relative: string, entries: RoleAssignmentEntry[]) => {
			unique.set(lower(`${BASE}/${relative}`), true);
			assignments.set(lower(`${BASE}/${relative}`), entries);
		},
	};
}

const options = (overrides: Partial<MappingOptions> = {}): MappingOptions => ({
	mode: 'set',
	breakInheritance: 'ifInheriting',
	copyCurrentFirst: false,
	clearSubscopes: false,
	createMissingFolders: false,
	ifFolderMissing: 'skip',
	skipUnresolved: false,
	...overrides,
});

const mapping = (path: string, principal: string, ...roles: string[]): MappingInput => ({
	path,
	principal,
	roles,
});

const TREE = [
	'Reports/2024/Drafts',
	'Reports/2024/Contracts',
	'Reports/2025/drafts',
	'Reports/2026',
	'Contracts',
];

const sorted = (entries: RoleAssignmentEntry[] | undefined) =>
	[...(entries ?? [])].sort(
		(a, b) => a.principalId - b.principalId || a.roleDefinitionId - b.roleDefinitionId,
	);

describe('parseMappingPath', () => {
	it('splits a path and keeps "*" segments', () => {
		expect(parseMappingPath(' Reports/*/Drafts/ ', 'Mapping 1')).toEqual([
			'Reports',
			'*',
			'Drafts',
		]);
		expect(parseMappingPath('Reports\\2024', 'Mapping 1')).toEqual(['Reports', '2024']);
	});

	it('reads "." as the base folder', () => {
		expect(parseMappingPath('.', 'Mapping 1')).toEqual([]);
		expect(parseMappingPath('./Contracts', 'Mapping 1')).toEqual(['Contracts']);
	});

	it('rejects "..", a partial wildcard and an empty path', () => {
		expect(() => parseMappingPath('../Other', 'Mapping 2')).toThrow(/Mapping 2.*\.\./);
		expect(() => parseMappingPath('Rep*/Drafts', 'Mapping 1')).toThrow(/whole segments/);
		expect(() => parseMappingPath('Reports/***', 'Mapping 1')).toThrow(/whole segments/);
		expect(() => parseMappingPath(' ', 'Mapping 1')).toThrow(/needs a path/);
	});

	it('keeps "**" segments and reads repeated "**" as one', () => {
		expect(parseMappingPath('Reports/**/Drafts', 'Mapping 1')).toEqual(['Reports', '**', 'Drafts']);
		expect(parseMappingPath('Reports/**/**/Drafts', 'Mapping 1')).toEqual([
			'Reports',
			'**',
			'Drafts',
		]);
	});
});

describe('parseMappingsJson', () => {
	it('reads path, principal, principal type and roles', () => {
		expect(
			parseMappingsJson([
				{ path: 'Contracts', principal: 'Example Members', principalType: 'group', role: 'Read' },
				{ path: 'Reports/*', principal: 'user@contoso.com', roles: ['Read', 'Contribute'] },
			]),
		).toEqual([
			{ path: 'Contracts', principal: 'Example Members', principalType: 'group', roles: ['Read'] },
			{ path: 'Reports/*', principal: 'user@contoso.com', roles: ['Read', 'Contribute'] },
		]);
	});

	it('names the parameter in errors', () => {
		expect(() => parseMappingsJson('[{ "principal": "Example Members", "role": "Read" }]')).toThrow(
			/Mappings \(JSON\): entry 1 needs a "path"/,
		);
		expect(() =>
			parseMappingsJson([{ path: 'A', principal: 'Example Members', role: 'Read', extra: 1 }]),
		).toThrow(/Mappings \(JSON\): entry 1 has unknown property "extra"/);
		expect(() => parseMappingsJson('[]')).toThrow(/Mappings \(JSON\): the array has no entries/);
		expect(() => parseMappingsJson([{ path: 'A', principal: 'Example Members' }])).toThrow(
			/needs a "role" or "roles"/,
		);
	});
});

describe('applyPermissionMapping', () => {
	it('expands "*" against existing child folders without case', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('reports/*/DRAFTS', 'Example Members', 'Read')],
			options(),
		);
		expect(result.folders.map((folder) => folder.path)).toEqual([
			'Reports/2024/Drafts',
			'Reports/2025/drafts',
		]);
		expect(result.folders[0]).toMatchObject({
			serverRelativeUrl: `${BASE}/Reports/2024/Drafts`,
			matchedFrom: ['reports/*/DRAFTS'],
			created: false,
		});
		expect(result.missing).toEqual([]);
		expect(result.warnings).toEqual([]);
	});

	it('matches "*/*/Contracts" and a final "*"', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('*/*/Contracts', 'Example Members', 'Read'),
				mapping('Reports/*', 'Example Visitors', 'Read'),
			],
			options(),
		);
		expect(result.folders.map((folder) => folder.path)).toEqual([
			'Reports/2024/Contracts',
			'Reports/2024',
			'Reports/2025',
			'Reports/2026',
		]);
	});

	it('matches "**" at any depth, also with zero folders between', async () => {
		const fake = createFake([...TREE, 'Reports/Drafts', 'Reports/2024/Q1/Drafts']);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Reports/**/drafts', 'Example Members', 'Read')],
			options(),
		);
		expect(result.folders.map((folder) => folder.path).sort()).toEqual([
			'Reports/2024/Drafts',
			'Reports/2024/Q1/Drafts',
			'Reports/2025/drafts',
			'Reports/Drafts',
		]);
		expect(result.warnings).toEqual([]);
	});

	it('matches all folders below with a final "**", but not the folder itself', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Reports/**', 'Example Members', 'Read')],
			options(),
		);
		expect(result.folders.map((folder) => folder.path).sort()).toEqual([
			'Reports/2024',
			'Reports/2024/Contracts',
			'Reports/2024/Drafts',
			'Reports/2025',
			'Reports/2025/drafts',
			'Reports/2026',
		]);
	});

	it('combines "**" with "*" and lists each folder once', async () => {
		const fake = createFake([...TREE, 'Reports/2024/Q1/Contracts']);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('**/*/Contracts', 'Example Members', 'Read')],
			options(),
		);
		expect(result.folders.map((folder) => folder.path).sort()).toEqual([
			'Reports/2024/Contracts',
			'Reports/2024/Q1/Contracts',
		]);
		const listed = fake.gateway.listChildFolders.mock.calls.map(([path]) => lower(path));
		expect(listed.length).toBe(new Set(listed).size);
	});

	it('warns when a wildcard path matches no folders', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Archive/*', 'Example Members', 'Read')],
			options(),
		);
		expect(result.folders).toEqual([]);
		expect(result.warnings).toEqual(['Mapping 1: the path "Archive/*" matched no folders.']);
	});

	it('groups the entries of several mappings by folder', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Reports/*/Drafts', 'Example Members', 'Contribute'),
				mapping('Reports/2024/Drafts', 'Example Visitors', 'Read'),
				mapping('reports/2024/drafts', 'Example Members', 'Contribute'),
			],
			options(),
		);
		const year2024 = result.folders.find((folder) => folder.path === 'Reports/2024/Drafts');
		expect(result.folders).toHaveLength(2);
		expect(year2024).toMatchObject({
			matchedFrom: ['Reports/*/Drafts', 'Reports/2024/Drafts', 'reports/2024/drafts'],
			added: 2,
		});
		expect(fake.gateway.applyChanges).toHaveBeenCalledTimes(2);
		expect(sorted(fake.assignmentsOf('Reports/2024/Drafts'))).toEqual([
			{ principalId: MEMBERS, roleDefinitionId: CONTRIBUTE },
			{ principalId: VISITORS, roleDefinitionId: READ },
		]);
	});

	it('Set mode breaks inheritance without a copy and gives exactly the mapped assignments', async () => {
		const fake = createFake(TREE);
		fake.setUnique('Reports/2024', [
			{ principalId: OWNERS, roleDefinitionId: FULL },
			{ principalId: MEMBERS, roleDefinitionId: READ },
		]);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Contracts', 'Example Members', 'Read'),
				mapping('Reports/2024', 'Example Members', 'Read'),
			],
			options(),
		);
		expect(result.folders).toEqual([
			expect.objectContaining({ path: 'Contracts', inheritanceBroken: true, added: 1, removed: 0 }),
			expect.objectContaining({
				path: 'Reports/2024',
				inheritanceBroken: false,
				added: 0,
				removed: 1,
				unchanged: 1,
			}),
		]);
		expect(fake.gateway.breakInheritance).toHaveBeenCalledWith(`${BASE}/Contracts`, false, false);
		expect(fake.assignmentsOf('Reports/2024')).toEqual([
			{ principalId: MEMBERS, roleDefinitionId: READ },
		]);
	});

	it('Set mode copies the current assignments first when asked', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Contracts', 'Example Members', 'Read')],
			options({ copyCurrentFirst: true, clearSubscopes: true }),
		);
		expect(fake.gateway.breakInheritance).toHaveBeenCalledWith(`${BASE}/Contracts`, true, true);
		expect(result.folders[0]).toMatchObject({ added: 1, removed: 1, unchanged: 0 });
	});

	it('Add mode keeps other assignments and copies when it breaks inheritance', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Contracts', 'Example Members', 'Read', 'Contribute')],
			options({ mode: 'add' }),
		);
		expect(fake.gateway.breakInheritance).toHaveBeenCalledWith(`${BASE}/Contracts`, true, false);
		expect(result.folders[0]).toMatchObject({ inheritanceBroken: true, added: 2, removed: 0 });
		expect(sorted(fake.assignmentsOf('Contracts'))).toEqual([
			{ principalId: OWNERS, roleDefinitionId: FULL },
			{ principalId: MEMBERS, roleDefinitionId: READ },
			{ principalId: MEMBERS, roleDefinitionId: CONTRIBUTE },
		]);
	});

	it('breaks inheritance on unique folders too with Always', async () => {
		const fake = createFake(TREE);
		fake.setUnique('Contracts', [{ principalId: MEMBERS, roleDefinitionId: READ }]);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[mapping('Contracts', 'Example Members', 'Read')],
			options({ breakInheritance: 'always', mode: 'add' }),
		);
		expect(fake.gateway.breakInheritance).toHaveBeenCalledWith(`${BASE}/Contracts`, true, false);
		expect(fake.gateway.hasUniqueRoleAssignments).not.toHaveBeenCalled();
		expect(result.folders[0].inheritanceBroken).toBe(true);
	});

	it('changes nothing on a second run', async () => {
		const fake = createFake(TREE);
		const mappings = [mapping('Reports/*/Drafts', 'Example Members', 'Read')];
		await applyPermissionMapping(fake.gateway, BASE, mappings, options());
		fake.gateway.applyChanges.mockClear();
		fake.gateway.breakInheritance.mockClear();

		const again = await applyPermissionMapping(fake.gateway, BASE, mappings, options());
		expect(
			again.folders.map(({ inheritanceBroken, added, removed, unchanged }) => ({
				inheritanceBroken,
				added,
				removed,
				unchanged,
			})),
		).toEqual([
			{ inheritanceBroken: false, added: 0, removed: 0, unchanged: 1 },
			{ inheritanceBroken: false, added: 0, removed: 0, unchanged: 1 },
		]);
		expect(fake.gateway.applyChanges).not.toHaveBeenCalled();
		expect(fake.gateway.breakInheritance).not.toHaveBeenCalled();
	});

	it('reports a missing folder and skips it by default', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Contracts', 'Example Members', 'Read'),
				mapping('Invoices/2026', 'Example Members', 'Read'),
			],
			options(),
		);
		expect(result.missing).toEqual([
			{
				path: 'Invoices/2026',
				serverRelativeUrl: `${BASE}/Invoices/2026`,
				matchedFrom: ['Invoices/2026'],
			},
		]);
		expect(result.folders.map((folder) => folder.path)).toEqual(['Contracts']);
		expect(fake.gateway.ensureFolder).not.toHaveBeenCalled();
	});

	it('fails on a missing folder before any change when If Folder Is Missing is Fail', async () => {
		const fake = createFake(TREE);
		await expect(
			applyPermissionMapping(
				fake.gateway,
				BASE,
				[
					mapping('Contracts', 'Example Members', 'Read'),
					mapping('Invoices', 'Example Members', 'Read'),
				],
				options({ ifFolderMissing: 'fail' }),
			),
		).rejects.toThrow(/Mapping 2: the folder "Invoices" does not exist/);
		expect(fake.changes).toEqual([]);
	});

	it('creates a missing folder with Create Missing Folders', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Invoices/2026', 'Example Members', 'Read'),
				mapping('Invoices/2026', 'Example Visitors', 'Read'),
				mapping('Contracts', 'Example Visitors', 'Read'),
			],
			options({ createMissingFolders: true, ifFolderMissing: 'fail' }),
		);
		expect(fake.gateway.ensureFolder).toHaveBeenCalledTimes(1);
		expect(result.folders).toEqual([
			expect.objectContaining({ path: 'Invoices/2026', created: true, added: 2 }),
			expect.objectContaining({ path: 'Contracts', created: false }),
		]);
		expect(result.missing).toEqual([]);
	});

	it('fails before any read or change when a principal cannot be resolved', async () => {
		const fake = createFake(TREE);
		await expect(
			applyPermissionMapping(
				fake.gateway,
				BASE,
				[
					mapping('Contracts', 'Example Members', 'Read'),
					mapping('Reports/*', 'Missing Group', 'Read'),
				],
				options(),
			),
		).rejects.toThrow(/Mapping 2 "Missing Group" was not found/);
		expect(fake.gateway.getFolder).not.toHaveBeenCalled();
		expect(fake.gateway.listChildFolders).not.toHaveBeenCalled();
		expect(fake.changes).toEqual([]);
	});

	it('names the role when a role cannot be resolved', async () => {
		const fake = createFake(TREE);
		await expect(
			applyPermissionMapping(
				fake.gateway,
				BASE,
				[mapping('Contracts', 'Example Members', 'Read', 'Approve')],
				options(),
			),
		).rejects.toThrow(/Mapping 1 "Example Members, role Approve"/);
	});

	it('skips unresolved mappings with warnings and applies the others', async () => {
		const fake = createFake(TREE);
		const result = await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Contracts', 'Example Members', 'Read'),
				mapping('Contracts', 'Missing Group', 'Read'),
				mapping('Reports/*', 'Missing Group', 'Read'),
			],
			options({ skipUnresolved: true }),
		);
		expect(result.warnings).toHaveLength(2);
		expect(result.warnings[0]).toMatch(/^Mapping 2 "Missing Group" was not found.*skipped\.$/);
		expect(result.folders.map((folder) => folder.path)).toEqual(['Contracts']);
		expect(fake.gateway.listChildFolders).not.toHaveBeenCalled();
	});

	it('resolves each principal and role once and lists each folder once', async () => {
		const fake = createFake(TREE);
		await applyPermissionMapping(
			fake.gateway,
			BASE,
			[
				mapping('Reports/*/Drafts', 'Example Members', 'Read'),
				mapping('Reports/*/Contracts', 'example members', 'read', 'Contribute'),
				mapping('Reports/*', 'Example Members', 'Contribute'),
				mapping('Contracts', 'Example Visitors', 'Read'),
			],
			options(),
		);
		expect(fake.gateway.resolvePrincipal).toHaveBeenCalledTimes(2);
		expect(fake.gateway.resolveRole).toHaveBeenCalledTimes(2);
		const listed = fake.gateway.listChildFolders.mock.calls.map(([path]) => path);
		expect(new Set(listed.map(lower)).size).toBe(listed.length);
	});
});

describe('Apply Mapping parameters', () => {
	const context = (parameters: Record<string, unknown>) =>
		({
			getNode: () => ({ name: 'SharePoint PnP' }),
			getNodeParameter: (name: string, _i: number, fallback?: unknown) =>
				name in parameters ? parameters[name] : fallback,
		}) as unknown as IExecuteFunctions;

	it('reads the mapping rows of Fields mode', () => {
		const rows = readMappings.call(
			context({
				mappings: {
					entries: [
						{
							path: ' Reports/*/Drafts ',
							principal: { mode: 'name', value: 'Example Members' },
							role: { mode: 'name', value: 'Read' },
						},
						{ path: 'Contracts', principal: 'user@contoso.com', role: '1073741826' },
					],
				},
			}),
			0,
		);
		expect(rows).toEqual([
			{
				path: 'Reports/*/Drafts',
				principal: 'Example Members',
				principalType: 'group',
				roles: ['Read'],
			},
			{ path: 'Contracts', principal: 'user@contoso.com', roles: ['1073741826'] },
		]);
	});

	it('rejects Fields mode without rows and bad JSON', () => {
		expect(() => readMappings.call(context({}), 0)).toThrow(/at least one mapping/);
		expect(() => readMappings.call(context({ inputMode: 'json', mappingsJson: '{' }), 0)).toThrow(
			/Mappings \(JSON\): the value is not valid JSON/,
		);
	});

	it('uses Set mode, Only If Inheriting and Skip by default', () => {
		expect(readMappingOptions.call(context({}), 0)).toEqual({
			options: options(),
			output: 'single',
		});
	});

	it('builds one item per folder and missing folder', () => {
		const result = {
			folders: [
				{
					path: 'Contracts',
					serverRelativeUrl: `${BASE}/Contracts`,
					matchedFrom: ['Contracts'],
					inheritanceBroken: true,
					added: 1,
					removed: 0,
					unchanged: 0,
					created: false,
				},
			],
			warnings: ['w'],
			missing: [{ path: 'B', serverRelativeUrl: `${BASE}/B`, matchedFrom: ['B'] }],
		};
		const items = toMappingOutput(result, BASE, 'set', 'perFolder') as Array<
			Record<string, unknown>
		>;
		expect(items).toHaveLength(2);
		expect(items[0]).toMatchObject({
			baseFolder: BASE,
			path: 'Contracts',
			missing: false,
			warnings: ['w'],
		});
		expect(items[1]).toMatchObject({ path: 'B', missing: true });
		expect(
			toMappingOutput({ folders: [], warnings: [], missing: [] }, BASE, 'add', 'perFolder'),
		).toEqual({ baseFolder: BASE, mode: 'add', folders: [], warnings: [], missing: [] });
		expect(toMappingOutput(result, BASE, 'set', 'single')).toMatchObject({
			baseFolder: BASE,
			folders: result.folders,
			missing: result.missing,
		});
	});
});
