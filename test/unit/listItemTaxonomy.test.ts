import { describe, expect, it } from 'vitest';

import { buildFormValues } from '../../nodes/SharePointPnp/resources/listItem/helpers';
import {
	formatTagValue,
	parseTagInput,
	resolveTags,
	resolveTaxonomyValues,
	type TagClient,
	type TaxonomyFieldInfo,
} from '../../nodes/SharePointPnp/resources/listItem/taxonomy';
import type { Term, TermSet } from '../../nodes/SharePointPnp/resources/taxonomy/service';

const SET_ID = '11111111-1111-1111-1111-111111111111';
const STORE_ID = '22222222-2222-2222-2222-222222222222';
const A = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const B = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const C = 'cccccccc-cccc-cccc-cccc-cccccccccccc';

function term(id: string, name: string, extra: Partial<Term> = {}): Term {
	return {
		id,
		name,
		labels: [{ value: name, language: 1033, isDefault: true }],
		description: '',
		path: name,
		created: null,
		modified: null,
		childCount: 0,
		isRoot: true,
		isDeprecated: false,
		isAvailableForTagging: true,
		termSetId: SET_ID,
		groupId: null,
		...extra,
	};
}

function mockClient(initial: Term[], open = true) {
	let terms = initial;
	const calls = { getTerms: 0, created: [] as string[] };
	const client: TagClient = {
		getTerms: async () => {
			calls.getTerms += 1;
			return terms;
		},
		getTermSet: async () => ({ name: 'Set', isOpenForTermCreation: open }) as TermSet,
		createTerm: async (_store, _set, label, parent) => {
			calls.created.push(`${label}@${parent ?? ''}`);
			const created = term(C, label);
			terms = [...terms, created];
			return created;
		},
	};
	return { client, calls };
}

const multi: TaxonomyFieldInfo = {
	InternalName: 'Tags',
	TypeAsString: 'TaxonomyFieldTypeMulti',
	TermSetId: SET_ID,
	SspId: STORE_ID,
};
const single: TaxonomyFieldInfo = {
	...multi,
	InternalName: 'Topic',
	TypeAsString: 'TaxonomyFieldType',
};

describe('parseTagInput', () => {
	it('reads labels, GUIDs and Label|GUID', () => {
		expect(parseTagInput('Red')).toEqual([{ label: 'Red' }]);
		expect(parseTagInput(A.toUpperCase())).toEqual([{ id: A }]);
		expect(parseTagInput(`Red|${A}`)).toEqual([{ label: 'Red', id: A }]);
		expect(parseTagInput(`Red | Blue|${A}`)).toEqual([{ label: 'Red | Blue', id: A }]);
	});

	it('reads lists as text, arrays and JSON', () => {
		expect(parseTagInput(`Red;Blue|${B}; `)).toEqual([{ label: 'Red' }, { label: 'Blue', id: B }]);
		expect(parseTagInput(['Red', `Blue|${B}`, A])).toEqual([
			{ label: 'Red' },
			{ label: 'Blue', id: B },
			{ id: A },
		]);
		expect(parseTagInput('["Red","Blue, Light"]')).toEqual([
			{ label: 'Red' },
			{ label: 'Blue, Light' },
		]);
	});

	it('keeps commas inside labels', () => {
		expect(parseTagInput('Blue, Light')).toEqual([{ label: 'Blue, Light' }]);
	});

	it('reads term objects as SharePoint returns them', () => {
		expect(parseTagInput([{ Label: 'Red', TermGuid: A, WssId: 3 }])).toEqual([
			{ label: 'Red', id: A },
		]);
		expect(parseTagInput(null)).toEqual([]);
		expect(parseTagInput('')).toEqual([]);
		expect(parseTagInput([])).toEqual([]);
	});
});

describe('formatTagValue', () => {
	it('joins terms with a semicolon', () => {
		expect(formatTagValue([{ label: 'Red', id: A }])).toBe(`Red|${A}`);
		expect(
			formatTagValue([
				{ label: 'Red', id: A },
				{ label: 'Blue', id: B },
			]),
		).toBe(`Red|${A};Blue|${B}`);
		expect(formatTagValue([])).toBe('');
	});
});

describe('resolveTags', () => {
	const terms = [
		term(A, 'Red'),
		term(B, 'Blue', {
			labels: [
				{ value: 'Blue', language: 1033, isDefault: true },
				{ value: 'Azul', language: 3082, isDefault: false },
			],
		}),
	];

	it('resolves labels without case, synonyms and IDs', async () => {
		const { client } = mockClient(terms);
		const result = await resolveTags(parseTagInput(['red', 'Azul', B]), multi, client);
		expect(result).toEqual([
			{ id: A, label: 'Red' },
			{ id: B, label: 'Blue' },
		]);
	});

	it('passes Label|GUID without a term store call', async () => {
		const { client, calls } = mockClient(terms);
		const result = await resolveTags(parseTagInput(`Anything|${C}`), multi, client);
		expect(result).toEqual([{ id: C, label: 'Anything' }]);
		expect(calls.getTerms).toBe(0);
	});

	it('removes duplicates and loads the terms once', async () => {
		const { client, calls } = mockClient(terms);
		const result = await resolveTags(parseTagInput(['Red', A, 'red']), multi, client);
		expect(result).toHaveLength(1);
		expect(calls.getTerms).toBe(1);
	});

	it('fails for an unknown label with a clear message', async () => {
		const { client, calls } = mockClient(terms);
		await expect(resolveTags(parseTagInput('Green'), multi, client)).rejects.toThrow(
			/"Green" does not exist.*Create Missing Terms/,
		);
		expect(calls.created).toEqual([]);
	});

	it('fails for an unknown term ID', async () => {
		const { client } = mockClient(terms);
		await expect(resolveTags(parseTagInput(C), multi, client)).rejects.toThrow(
			/not in the term set/,
		);
	});

	it('fails when a single field gets many terms', async () => {
		const { client } = mockClient(terms);
		await expect(resolveTags(parseTagInput(['Red', 'Blue']), single, client)).rejects.toThrow(
			/accepts one term/,
		);
	});

	it('fails for a label that matches many terms, and for deprecated terms', async () => {
		const many = [term(A, 'Red', { path: 'X;Red' }), term(B, 'Red', { path: 'Y;Red' })];
		await expect(resolveTags(parseTagInput('Red'), multi, mockClient(many).client)).rejects.toThrow(
			/matches 2 terms.*X;Red, Y;Red/,
		);
		const old = [term(A, 'Red', { isDeprecated: true })];
		await expect(resolveTags(parseTagInput('Red'), multi, mockClient(old).client)).rejects.toThrow(
			/deprecated/,
		);
	});

	it('fails when the field has no term set', async () => {
		const { client } = mockClient(terms);
		await expect(
			resolveTags(parseTagInput('Red'), { ...multi, TermSetId: undefined }, client),
		).rejects.toThrow(/not bound/);
	});

	it('creates missing terms only when allowed', async () => {
		const { client, calls } = mockClient(terms);
		const result = await resolveTags(parseTagInput(['Red', 'Green', 'green']), multi, client, {
			createMissing: true,
		});
		expect(result.map((entry) => entry.label)).toEqual(['Red', 'Green']);
		expect(calls.created).toEqual(['Green@']);
	});

	it('does not create terms in a closed term set or with invalid characters', async () => {
		await expect(
			resolveTags(parseTagInput('Green'), multi, mockClient(terms, false).client, {
				createMissing: true,
			}),
		).rejects.toThrow(/closed for new terms/);
		await expect(
			resolveTags(parseTagInput('Gre"en'), multi, mockClient(terms).client, {
				createMissing: true,
			}),
		).rejects.toThrow(/must not contain/);
	});

	it('limits label search to the anchor term and creates under it', async () => {
		const anchor = '99999999-9999-9999-9999-999999999999';
		const tree = [
			term(anchor, 'Root'),
			term(A, 'Red', { path: 'Root;Red' }),
			term(B, 'Red', { path: 'Other;Red' }),
		];
		const field = { ...multi, AnchorId: anchor };
		const { client, calls } = mockClient(tree);
		expect(await resolveTags(parseTagInput('Red'), field, client)).toEqual([
			{ id: A, label: 'Red' },
		]);
		await resolveTags(parseTagInput('Green'), field, client, { createMissing: true });
		expect(calls.created).toEqual([`Green@${anchor}`]);
	});
});

describe('resolveTaxonomyValues', () => {
	const terms = [term(A, 'Red'), term(B, 'Blue')];
	const infos = [multi, single, { InternalName: 'Title', TypeAsString: 'Text' }];

	it('formats tag fields and leaves other fields alone', async () => {
		const { client } = mockClient(terms);
		const result = await resolveTaxonomyValues(
			{ Tags: ['Red', 'Blue'], Topic: 'Blue', Title: 'a, b', Other: 1 },
			infos,
			client,
		);
		expect(result).toEqual({
			Tags: `Red|${A};Blue|${B}`,
			Topic: `Blue|${B}`,
			Title: 'a, b',
			Other: 1,
		});
		expect(buildFormValues(result, infos)).toContainEqual({
			FieldName: 'Tags',
			FieldValue: `Red|${A};Blue|${B}`,
		});
	});

	it('clears a tag field with an empty value', async () => {
		const { client, calls } = mockClient(terms);
		const result = await resolveTaxonomyValues({ Tags: [], Topic: null }, infos, client);
		expect(result).toEqual({ Tags: '', Topic: '' });
		expect(calls.getTerms).toBe(0);
	});

	it('appends to the current tags of a multi-value field', async () => {
		const { client } = mockClient(terms);
		const result = await resolveTaxonomyValues({ Tags: ['Blue', 'Red'] }, infos, client, {
			current: { Tags: [{ Label: 'Red', TermGuid: A, WssId: 4 }] },
		});
		expect(result.Tags).toBe(`Red|${A};Blue|${B}`);
	});

	it('does not append to a single-value field', async () => {
		const { client } = mockClient(terms);
		const result = await resolveTaxonomyValues({ Topic: 'Blue' }, infos, client, {
			current: { Topic: { Label: 'Red', TermGuid: A } },
		});
		expect(result.Topic).toBe(`Blue|${B}`);
	});
});
