import { describe, expect, it, vi } from 'vitest';

import {
	desiredListProperties,
	diffProperties,
	ensureList,
	findListContentType,
	type EnsureListGateway,
	type EnsureListSpec,
	type ListContentType,
} from '../../nodes/SharePointPnp/resources/list/ensure';
import { listResource } from '../../nodes/SharePointPnp/resources/list';
import type { RoleAssignmentEntry } from '../../nodes/SharePointPnp/shared/securables';

const SITE = '/sites/Example';
const DOCUMENT = { id: '0x0101009A', name: 'Document' };
const SITE_TYPE = '0x010100ABCDEF';
const LIST_COPY = `${SITE_TYPE}00${'1'.repeat(32)}`;

interface FakeList {
	id: string;
	title: string;
	url: string;
	properties: Record<string, unknown>;
	contentTypes: ListContentType[];
	order: string[];
	views: string[];
	unique: boolean;
	assignments: RoleAssignmentEntry[];
}

function fakeGateway(existing?: Partial<FakeList>) {
	const lists: FakeList[] = [];
	if (existing) {
		lists.push({
			id: 'existing-id',
			title: 'Reports',
			url: `${SITE}/Reports`,
			properties: { OnQuickLaunch: false, EnableVersioning: false, ContentTypesEnabled: false },
			contentTypes: [DOCUMENT],
			order: [DOCUMENT.id],
			views: ['All Documents'],
			unique: false,
			assignments: [],
			...existing,
		});
	}
	const get = (id: string) => lists.find((list) => list.id === id)!;
	const summary = (list: FakeList) => ({
		id: list.id,
		title: list.title,
		serverRelativeUrl: list.url,
	});
	const gateway = {
		findByUrl: vi.fn(async (url: string) => {
			const list = lists.find((entry) => entry.url.toLowerCase() === url.toLowerCase());
			return list && summary(list);
		}),
		findByTitle: vi.fn(async (title: string) => {
			const list = lists.find((entry) => entry.title.toLowerCase() === title.toLowerCase());
			return list && summary(list);
		}),
		getSummary: vi.fn(async (id: string) => summary(get(id))),
		create: vi.fn(async (title: string, _description: string, template: number) => {
			const id = `new-${lists.length}`;
			lists.push({
				id,
				title,
				url: `${SITE}/${template === 100 ? 'Lists/' : ''}${title}`,
				properties: { OnQuickLaunch: false, EnableVersioning: true, ContentTypesEnabled: false },
				contentTypes: [DOCUMENT],
				order: [],
				views: ['All Documents'],
				unique: false,
				assignments: [{ principalId: 3, roleDefinitionId: 1073741829 }],
			});
			return id;
		}),
		getProperties: vi.fn(async (id: string, names: string[]) =>
			Object.fromEntries(names.map((name) => [name, get(id).properties[name]])),
		),
		update: vi.fn(async (id: string, properties: Record<string, unknown>) => {
			const list = get(id);
			if (typeof properties.Title === 'string') list.title = properties.Title;
			Object.assign(list.properties, properties);
		}),
		getContentTypes: vi.fn(async (id: string) => [...get(id).contentTypes]),
		addContentType: vi.fn(async (id: string, siteId: string) => {
			get(id).contentTypes.push({ id: `${siteId}00${'1'.repeat(32)}`, name: 'Report Document' });
			return 'Report Document';
		}),
		getContentTypeOrder: vi.fn(async (id: string) => {
			const list = get(id);
			return list.order.length ? [...list.order] : list.contentTypes.map((type) => type.id);
		}),
		setContentTypeOrder: vi.fn(async (id: string, order: string[]) => {
			get(id).order = order;
		}),
		getViewTitles: vi.fn(async (id: string) => [...get(id).views]),
		getSourceViewTitle: vi.fn(async () => 'Report View'),
		copyView: vi.fn(async (_source, _view, targetId: string) => {
			get(targetId).views.push('Report View');
			return ['MissingField'];
		}),
		hasUniqueRoleAssignments: vi.fn(async (id: string) => get(id).unique),
		breakInheritance: vi.fn(async (id: string, copy: boolean) => {
			const list = get(id);
			list.unique = true;
			if (!copy) list.assignments = [{ principalId: 9, roleDefinitionId: 1073741829 }];
		}),
		getRoleAssignments: vi.fn(async (id: string) => [...get(id).assignments]),
		addRoleAssignment: vi.fn(async (id: string, entry: RoleAssignmentEntry) => {
			get(id).assignments.push(entry);
		}),
		removeRoleAssignment: vi.fn(async (id: string, entry: RoleAssignmentEntry) => {
			const list = get(id);
			list.assignments = list.assignments.filter(
				(item) =>
					item.principalId !== entry.principalId ||
					item.roleDefinitionId !== entry.roleDefinitionId,
			);
		}),
	} satisfies EnsureListGateway;
	return { gateway, lists };
}

const spec = (overrides: Partial<EnsureListSpec> = {}): EnsureListSpec => ({
	title: 'Reports',
	template: 101,
	applyToExisting: false,
	settings: {},
	contentTypes: [],
	...overrides,
});

const writeCalls = (gateway: ReturnType<typeof fakeGateway>['gateway']) =>
	[
		gateway.create,
		gateway.update,
		gateway.addContentType,
		gateway.setContentTypeOrder,
		gateway.copyView,
		gateway.breakInheritance,
		gateway.addRoleAssignment,
		gateway.removeRoleAssignment,
	].reduce((count, fn) => count + fn.mock.calls.length, 0);

describe('List Ensure: list exists', () => {
	it('returns the list by title and changes nothing when Apply Settings is off', async () => {
		const { gateway } = fakeGateway({});
		const resolveAssignments = vi.fn(async () => []);
		const result = await ensureList(
			spec({
				settings: { onQuickLaunch: true, enableVersioning: true },
				contentTypes: [SITE_TYPE],
				permissions: {
					breakInheritance: true,
					copyRoleAssignments: false,
					clearSubscopes: false,
					mode: 'replace',
					resolveAssignments,
				},
			}),
			gateway,
		);
		expect(result).toEqual({
			id: 'existing-id',
			title: 'Reports',
			serverRelativeUrl: `${SITE}/Reports`,
			created: false,
			changes: [],
		});
		expect(writeCalls(gateway)).toBe(0);
		expect(resolveAssignments).not.toHaveBeenCalled();
	});

	it('finds the list by URL', async () => {
		const { gateway } = fakeGateway({ title: 'Other Title' });
		const result = await ensureList(spec({ url: `${SITE}/reports` }), gateway);
		expect(result).toMatchObject({ id: 'existing-id', created: false });
		expect(gateway.findByTitle).not.toHaveBeenCalled();
	});

	it('fails when the title exists at another URL', async () => {
		const { gateway } = fakeGateway({});
		await expect(ensureList(spec({ url: `${SITE}/Archive` }), gateway)).rejects.toThrow(
			/exists at "\/sites\/Example\/Reports", not at "\/sites\/Example\/Archive"/,
		);
		expect(gateway.create).not.toHaveBeenCalled();
	});

	it('applies only the changed settings when Apply Settings is on', async () => {
		const { gateway } = fakeGateway({
			properties: { OnQuickLaunch: true, EnableVersioning: false },
		});
		const result = await ensureList(
			spec({
				applyToExisting: true,
				description: 'New text',
				settings: { onQuickLaunch: true, enableVersioning: true, majorVersionLimit: 50 },
			}),
			gateway,
		);
		expect(gateway.update).toHaveBeenCalledWith('existing-id', {
			EnableVersioning: true,
			MajorVersionLimit: 50,
			Description: 'New text',
		});
		expect(result.created).toBe(false);
		expect(result.changes).toEqual([
			'Set EnableVersioning to true',
			'Set MajorVersionLimit to 50',
			'Set Description to "New text"',
		]);
	});

	it('reports no changes when an existing list already matches', async () => {
		const { gateway } = fakeGateway({
			properties: { OnQuickLaunch: true, ContentTypesEnabled: true },
			contentTypes: [DOCUMENT, { id: LIST_COPY, name: 'Report Document' }],
			order: [LIST_COPY, DOCUMENT.id],
			views: ['All Documents', 'Report View'],
			unique: true,
			assignments: [{ principalId: 7, roleDefinitionId: 1073741827 }],
		});
		const result = await ensureList(
			spec({
				applyToExisting: true,
				settings: { onQuickLaunch: true },
				contentTypes: [SITE_TYPE],
				defaultContentType: SITE_TYPE,
				copyView: { sourceList: { mode: 'title', value: 'Template' }, view: 'Report View' },
				permissions: {
					breakInheritance: true,
					copyRoleAssignments: false,
					clearSubscopes: false,
					mode: 'replace',
					resolveAssignments: async () => [{ principalId: 7, roleDefinitionId: 1073741827 }],
				},
			}),
			gateway,
		);
		expect(result.changes).toEqual([]);
		expect(result.warnings).toEqual([
			'The list has a view "Report View". The view was not copied.',
		]);
		expect(writeCalls(gateway)).toBe(0);
	});
});

describe('List Ensure: list is missing', () => {
	it('creates the list with the title, description and template', async () => {
		const { gateway } = fakeGateway();
		const result = await ensureList(spec({ description: 'Report files' }), gateway);
		expect(gateway.create).toHaveBeenCalledWith('Reports', 'Report files', 101);
		expect(result).toEqual({
			id: 'new-0',
			title: 'Reports',
			serverRelativeUrl: `${SITE}/Reports`,
			created: true,
			changes: ['Created the list "Reports" with template 101'],
		});
	});

	it('creates the list with the last part of the URL and then sets the title', async () => {
		const { gateway } = fakeGateway();
		const result = await ensureList(
			spec({ title: 'Report Files', url: `${SITE}/ReportFiles` }),
			gateway,
		);
		expect(gateway.create).toHaveBeenCalledWith('ReportFiles', '', 101);
		expect(gateway.update).toHaveBeenCalledWith('new-0', { Title: 'Report Files' });
		expect(result).toMatchObject({
			title: 'Report Files',
			serverRelativeUrl: `${SITE}/ReportFiles`,
			created: true,
		});
		expect(result.warnings).toBeUndefined();
	});

	it('warns when SharePoint puts the list at another URL', async () => {
		const { gateway } = fakeGateway();
		const result = await ensureList(
			spec({ title: 'Tasks', url: `${SITE}/Tasks`, template: 100 }),
			gateway,
		);
		expect(result.serverRelativeUrl).toBe(`${SITE}/Lists/Tasks`);
		expect(result.warnings?.[0]).toMatch(/created the list at "\/sites\/Example\/Lists\/Tasks"/);
	});

	it('sets Quick Launch and versioning', async () => {
		const { gateway } = fakeGateway();
		const result = await ensureList(
			spec({
				settings: {
					onQuickLaunch: true,
					enableVersioning: true,
					majorVersionLimit: 100,
					enableMinorVersions: true,
					minorVersionLimit: 5,
				},
			}),
			gateway,
		);
		// EnableVersioning is already true on the new list, so it is not sent.
		expect(gateway.update).toHaveBeenCalledWith('new-0', {
			OnQuickLaunch: true,
			MajorVersionLimit: 100,
			EnableMinorVersions: true,
			MajorWithMinorVersionsLimit: 5,
		});
		expect(result.changes).toContain('Set OnQuickLaunch to true');
	});

	it('enables content types, adds them and sets the default', async () => {
		const { gateway, lists } = fakeGateway();
		const result = await ensureList(
			spec({ contentTypes: [SITE_TYPE, DOCUMENT.id], defaultContentType: SITE_TYPE }),
			gateway,
		);
		expect(gateway.update).toHaveBeenCalledWith('new-0', { ContentTypesEnabled: true });
		expect(gateway.addContentType).toHaveBeenCalledTimes(1);
		expect(gateway.addContentType).toHaveBeenCalledWith('new-0', SITE_TYPE);
		expect(lists[0].order).toEqual([LIST_COPY, DOCUMENT.id]);
		expect(result.changes).toEqual([
			'Created the list "Reports" with template 101',
			'Set ContentTypesEnabled to true',
			'Added the content type "Report Document"',
			'Set the default content type to "Report Document"',
		]);
	});

	it('enables content types without adding any', async () => {
		const { gateway } = fakeGateway();
		await ensureList(spec({ settings: { enableContentTypes: true } }), gateway);
		expect(gateway.update).toHaveBeenCalledWith('new-0', { ContentTypesEnabled: true });
		expect(gateway.addContentType).not.toHaveBeenCalled();
	});

	it('fails when the default content type is not in the list', async () => {
		const { gateway } = fakeGateway();
		await expect(ensureList(spec({ defaultContentType: '0x0120' }), gateway)).rejects.toThrow(
			/"0x0120" is not in the list/,
		);
	});

	it('copies a view and warns about fields that the list does not have', async () => {
		const { gateway } = fakeGateway();
		const sourceList = { mode: 'title', value: 'Template' };
		const result = await ensureList(
			spec({ copyView: { sourceList, view: 'Report View' } }),
			gateway,
		);
		expect(gateway.copyView).toHaveBeenCalledWith(sourceList, 'Report View', 'new-0');
		expect(result.changes).toContain('Copied the view "Report View"');
		expect(result.warnings).toEqual([
			'The view "Report View" was copied without these fields that the list does not have: MissingField',
		]);
	});

	it('breaks inheritance with the given options', async () => {
		const { gateway } = fakeGateway();
		const result = await ensureList(
			spec({
				permissions: {
					breakInheritance: true,
					copyRoleAssignments: true,
					clearSubscopes: true,
					mode: 'replace',
				},
			}),
			gateway,
		);
		expect(gateway.breakInheritance).toHaveBeenCalledWith('new-0', true, true);
		expect(gateway.getRoleAssignments).not.toHaveBeenCalled();
		expect(result.changes).toContain(
			'Broke permission inheritance (copy role assignments: true, clear subscopes: true)',
		);
	});

	it('replaces the role assignments in Replace mode', async () => {
		const { gateway, lists } = fakeGateway();
		const result = await ensureList(
			spec({
				permissions: {
					breakInheritance: false,
					copyRoleAssignments: false,
					clearSubscopes: false,
					mode: 'replace',
					resolveAssignments: async () => [{ principalId: 7, roleDefinitionId: 1073741827 }],
				},
			}),
			gateway,
		);
		// Role assignments always break inheritance first.
		expect(gateway.breakInheritance).toHaveBeenCalledWith('new-0', false, false);
		expect(lists[0].assignments).toEqual([{ principalId: 7, roleDefinitionId: 1073741827 }]);
		expect(result.changes).toEqual([
			'Created the list "Reports" with template 101',
			'Broke permission inheritance (copy role assignments: false, clear subscopes: false)',
			'Added role 1073741827 for principal 7',
			'Removed role 1073741829 from principal 9',
		]);
	});

	it('keeps the other role assignments in Add mode', async () => {
		const { gateway, lists } = fakeGateway();
		await ensureList(
			spec({
				permissions: {
					breakInheritance: true,
					copyRoleAssignments: false,
					clearSubscopes: false,
					mode: 'add',
					resolveAssignments: async () => [{ principalId: 7, roleDefinitionId: 1073741827 }],
				},
			}),
			gateway,
		);
		expect(gateway.removeRoleAssignment).not.toHaveBeenCalled();
		expect(lists[0].assignments).toEqual([
			{ principalId: 9, roleDefinitionId: 1073741829 },
			{ principalId: 7, roleDefinitionId: 1073741827 },
		]);
	});

	it('resolves the principals before it creates the list', async () => {
		const { gateway } = fakeGateway();
		await expect(
			ensureList(
				spec({
					permissions: {
						breakInheritance: true,
						copyRoleAssignments: false,
						clearSubscopes: false,
						mode: 'replace',
						resolveAssignments: async () => {
							throw new Error('Principal not found');
						},
					},
				}),
				gateway,
			),
		).rejects.toThrow('Principal not found');
		expect(gateway.create).not.toHaveBeenCalled();
	});
});

describe('findListContentType', () => {
	const types = [DOCUMENT, { id: LIST_COPY, name: 'Report Document' }];

	it('finds by list ID, by site content type ID and by name', () => {
		expect(findListContentType(types, LIST_COPY.toUpperCase())?.name).toBe('Report Document');
		expect(findListContentType(types, SITE_TYPE)?.name).toBe('Report Document');
		expect(findListContentType(types, 'report document')?.id).toBe(LIST_COPY);
	});

	it('does not match a content type that only shares the prefix', () => {
		expect(findListContentType(types, '0x010100AB')).toBeUndefined();
	});
});

describe('desiredListProperties and diffProperties', () => {
	it('sends the description only for an existing list', () => {
		const value = spec({ description: 'Text' });
		expect(desiredListProperties(value, true)).toEqual({});
		expect(desiredListProperties(value, false)).toEqual({ Description: 'Text' });
	});

	it('turns off content types only when asked', () => {
		expect(desiredListProperties(spec({ settings: { enableContentTypes: false } }), true)).toEqual({
			ContentTypesEnabled: false,
		});
		expect(
			desiredListProperties(
				spec({ settings: { enableContentTypes: false }, contentTypes: [SITE_TYPE] }),
				true,
			),
		).toEqual({ ContentTypesEnabled: true });
	});

	it('keeps only the changed values', () => {
		expect(diffProperties({ A: 1, B: true }, { A: 1, B: false, C: 'x' })).toEqual({
			B: false,
			C: 'x',
		});
	});
});

describe('List Ensure description', () => {
	it('offers the operation and its options', () => {
		const operation = listResource.description[0];
		expect(
			(operation.options as Array<{ value: string }>).some((option) => option.value === 'ensure'),
		).toBe(true);
		expect(listResource.operations.ensure).toBeTypeOf('function');
		const options = listResource.description.find((entry) => entry.name === 'ensureOptions');
		expect((options?.options as Array<{ name: string }>).map((option) => option.name)).toEqual(
			expect.arrayContaining([
				'applyToExisting',
				'contentTypes',
				'copyViewSourceList',
				'copyViewName',
				'defaultContentType',
				'enableContentTypes',
				'enableVersioning',
				'majorVersionLimit',
				'minorVersionLimit',
				'onQuickLaunch',
				'permissions',
			]),
		);
	});
});
