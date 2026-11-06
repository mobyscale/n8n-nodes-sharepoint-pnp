import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import { joinPath, toServerRelativePath } from '../../nodes/SharePointPnp/shared/utils';
import {
	config,
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
} from './harness';

interface Group {
	id: string;
	name: string;
}

interface SetInfo {
	id: string;
	name: string;
	groupId: string;
}

interface TermInfo {
	id: string;
	name: string;
	path: string;
	termSetId: string;
	childCount: number;
	labels: Array<{ value: string; language: number; isDefault: boolean }>;
}

const run = async (operation: string, parameters: Record<string, unknown> = {}) =>
	(await runNode({ resource: 'taxonomy', operation, ...parameters })).map((item) => item.json);

describe.runIf(hasConfig)('taxonomy (read only)', () => {
	it('gets term groups', async () => {
		const groups = (await run('getTermGroups')) as unknown as Group[];
		expect(groups.length).toBeGreaterThan(0);
		expect(groups[0].id).toMatch(/^[0-9a-f-]{36}$/);
		expect(typeof groups[0].name).toBe('string');
	});

	it('gets term sets of all groups and of one group', async () => {
		const all = (await run('getTermSets', { returnAll: true })) as unknown as SetInfo[];
		expect(all.length).toBeGreaterThan(0);
		const first = all[0];
		const inGroup = (await run('getTermSets', {
			termGroup: locator(first.groupId, 'id'),
			returnAll: true,
		})) as unknown as SetInfo[];
		expect(inGroup.map((set) => set.id)).toContain(first.id);
		expect(inGroup.every((set) => set.groupId === first.groupId)).toBe(true);
	});

	it('gets a term set, its terms, one term and searches by label', async () => {
		const sets = (await run('getTermSets', { returnAll: true })) as unknown as SetInfo[];
		let found: { set: SetInfo; terms: TermInfo[] } | undefined;
		for (const set of sets) {
			const terms = (await run('getTerms', {
				termSet: locator(set.id, 'id'),
				getTermsOptions: { recursive: true },
				returnAll: true,
			})) as unknown as TermInfo[];
			if (terms.length) {
				found = { set, terms };
				break;
			}
		}
		if (!found) return;
		const { set, terms } = found;

		const [setResult] = (await run('getTermSet', {
			termSet: locator(set.id, 'id'),
		})) as unknown as SetInfo[];
		expect(setResult.id).toBe(set.id);

		const term = terms[0];
		expect(term.labels.length).toBeGreaterThan(0);
		expect(term.termSetId).toBe(set.id);

		const [single] = (await run('getTerm', { termId: term.id })) as unknown as TermInfo[];
		expect(single.id).toBe(term.id);
		expect(single.termSetId).toBe(set.id);
		expect(single.labels.length).toBeGreaterThan(0);

		const matches = (await run('searchTerms', {
			termSet: locator(set.id, 'id'),
			searchText: term.labels[0].value,
			matchType: 'exact',
			returnAll: true,
		})) as unknown as TermInfo[];
		expect(matches.map((match) => match.id)).toContain(term.id);

		const limited = await run('getTerms', {
			termSet: locator(set.id, 'id'),
			returnAll: false,
			limit: 1,
		});
		expect(limited).toHaveLength(1);
	});

	it('gets the child terms of a term that has children', async () => {
		const sets = (await run('getTermSets', { returnAll: true })) as unknown as SetInfo[];
		for (const set of sets) {
			const roots = (await run('getTerms', {
				termSet: locator(set.id, 'id'),
				returnAll: true,
			})) as unknown as TermInfo[];
			const parent = roots.find((term) => term.childCount > 0);
			if (!parent) continue;
			const children = (await run('getTerms', {
				termSet: locator(set.id, 'id'),
				getTermsOptions: { parentTermId: parent.id },
				returnAll: true,
			})) as unknown as TermInfo[];
			expect(children).toHaveLength(parent.childCount);
			const descendants = await run('getTerms', {
				termSet: locator(set.id, 'id'),
				getTermsOptions: { parentTermId: parent.id, recursive: true },
				returnAll: true,
			});
			expect(descendants.length).toBeGreaterThanOrEqual(children.length);
			return;
		}
	});

	it.skip('creates, updates and deletes a term (term store writes are tenant wide)', async () => {
		const sets = (await run('getTermSets', { returnAll: true })) as unknown as SetInfo[];
		const setId = sets[0].id;
		const [created] = (await run('createTerm', {
			termSet: locator(setId, 'id'),
			termName: `__pnp-test-${Date.now()}`,
			createOptions: { description: 'Test term' },
		})) as unknown as TermInfo[];
		try {
			await run('updateTerm', { termId: created.id, updateOptions: { description: 'Changed' } });
		} finally {
			await run('deleteTerm', { termId: created.id });
		}
	});
});

interface TagField {
	InternalName: string;
	TypeAsString: string;
	TermSetId: string;
	SspId: string;
	AnchorId: string;
}

describe.runIf(hasConfig)('tags on list items', () => {
	let testFolder: string | undefined;
	let field: TagField | undefined;
	let filePath = '';
	let candidates: TermInfo[] = [];

	const readTags = async () => {
		const item = (await createTestSp()
			.web.getFileByServerRelativePath(filePath)
			.listItemAllFields.select(field!.InternalName)()) as Record<string, unknown>;
		const value = item[field!.InternalName];
		const results = (value as { results?: unknown[] } | null)?.results;
		const list = Array.isArray(results) ? results : Array.isArray(value) ? value : [value];
		return list
			.filter((entry): entry is { TermGuid: string } => Boolean(entry))
			.map((entry) => entry.TermGuid.toLowerCase());
	};

	beforeAll(async () => {
		const sp = createTestSp();
		const libraryRoot = toServerRelativePath(config!.siteUrl, config!.libraryPath);
		const lists = (await sp.web.lists
			.select('Id', 'RootFolder/ServerRelativeUrl')
			.expand('RootFolder')()) as Array<{
			Id: string;
			RootFolder: { ServerRelativeUrl: string };
		}>;
		const library = lists.find(
			(list) => list.RootFolder.ServerRelativeUrl.toLowerCase() === libraryRoot.toLowerCase(),
		);
		if (!library) return;
		const fields = (await sp.web.lists
			.getById(library.Id)
			.fields.filter(
				"TypeAsString eq 'TaxonomyFieldType' or TypeAsString eq 'TaxonomyFieldTypeMulti'",
			)
			.select(
				'InternalName',
				'TypeAsString',
				'Hidden',
				'ReadOnlyField',
				'TermSetId',
				'SspId',
				'AnchorId',
			)()) as Array<TagField & { Hidden: boolean; ReadOnlyField: boolean }>;
		field = fields.find((entry) => !entry.Hidden && !entry.ReadOnlyField);
		if (!field) return;

		const terms = (await run('getTerms', {
			termSet: locator(field.TermSetId, 'id'),
			getTermsOptions: { recursive: true },
			returnAll: true,
		})) as unknown as TermInfo[];
		const anchor = terms.find((term) => term.id === field!.AnchorId.toLowerCase());
		const labelCount = new Map<string, number>();
		for (const term of terms) {
			const key = term.name.toLowerCase();
			labelCount.set(key, (labelCount.get(key) ?? 0) + 1);
		}
		candidates = terms.filter(
			(term) =>
				labelCount.get(term.name.toLowerCase()) === 1 &&
				!/[;|"<>]/.test(term.name) &&
				(!anchor || term.path.startsWith(anchor.path + ';')),
		);
		if (!candidates.length) return;

		testFolder = await createTestFolder();
		filePath = joinPath(testFolder, 'tagged.txt');
		await sp.web
			.getFolderByServerRelativePath(testFolder)
			.files.addUsingPath('tagged.txt', 'tag test');
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	const update = async (value: unknown, updateOptions: Record<string, unknown> = {}) =>
		await runNode({
			resource: 'listItem',
			operation: 'update',
			identifyBy: 'path',
			itemPath: locator(filePath),
			fieldsJson: JSON.stringify({ [field!.InternalName]: value }),
			updateOptions,
		});

	it('sets a tag by label, then by label with ID, and appends another', async (context) => {
		if (!field || !candidates.length || !testFolder) {
			context.skip();
			return;
		}
		const [first, second] = candidates;

		await update(first.name);
		expect(await readTags()).toEqual([first.id]);

		await update(`${first.name}|${first.id}`);
		expect(await readTags()).toEqual([first.id]);

		if (field.TypeAsString === 'TaxonomyFieldTypeMulti' && second) {
			await update([second.name], { appendToMultiValue: true });
			expect((await readTags()).sort()).toEqual([first.id, second.id].sort());

			await update([first.id, second.name]);
			expect((await readTags()).sort()).toEqual([first.id, second.id].sort());
		}

		await update(null);
		expect(await readTags()).toEqual([]);
	});

	it('fails for an unknown label when Create Missing Terms is off', async (context) => {
		if (!field || !candidates.length || !testFolder) {
			context.skip();
			return;
		}
		await expect(update(`__pnp-missing-${Date.now()}`)).rejects.toThrow(/does not exist/);
	});

	it.skip('creates a missing term (term store writes are tenant wide)', async () => {
		await update(`__pnp-test-${Date.now()}`, { createMissingTerms: true });
	});
});
