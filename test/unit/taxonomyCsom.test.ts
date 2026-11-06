import { describe, expect, it } from 'vitest';

import {
	CsomRequest,
	csomGuid,
	csomString,
	escapeXml,
	isGuid,
	normalizeGuid,
	parseCsomDate,
	parseCsomGuid,
	parseCsomResponse,
} from '../../nodes/SharePointPnp/resources/taxonomy/csom';
import {
	mapGroup,
	mapTerm,
	mapTermSet,
	matchesLabel,
	type Term,
} from '../../nodes/SharePointPnp/resources/taxonomy/service';

const ID = '12345678-1234-1234-1234-123456789abc';

describe('csom values', () => {
	it('reads guids and dates', () => {
		expect(parseCsomGuid(`/Guid(${ID.toUpperCase()})/`)).toBe(ID);
		expect(parseCsomGuid(`{${ID}}`)).toBe(ID);
		expect(parseCsomDate('/Date(1700000000000)/')).toBe('2023-11-14T22:13:20.000Z');
		expect(parseCsomDate(null)).toBeNull();
	});

	it('checks and normalizes guids', () => {
		expect(isGuid(`{${ID}}`)).toBe(true);
		expect(isGuid('Marketing')).toBe(false);
		expect(normalizeGuid(` {${ID.toUpperCase()}} `)).toBe(ID);
	});

	it('escapes XML', () => {
		expect(escapeXml(`a&b<c>"d"'e'`)).toBe('a&amp;b&lt;c&gt;&quot;d&quot;&apos;e&apos;');
	});
});

describe('CsomRequest', () => {
	it('builds object paths, instantiates parents and escapes parameters', () => {
		const request = new CsomRequest();
		const { store } = request.termStore();
		const set = request.method(store, 'GetTermSet', [csomGuid(`{${ID.toUpperCase()}}`)]);
		const term = request.method(set, 'CreateTerm', [csomString('R&D <x>')]);
		request.run(term);
		request.call(store, 'CommitAll');
		const key = request.query(term, { nested: ['Labels'] });
		const xml = request.toXml();

		expect(xml).toContain('<StaticMethod Id="1" Name="GetTaxonomySession"');
		expect(xml).toContain('Name="GetDefaultSiteCollectionTermStore"');
		expect(xml).toContain(`<Parameter Type="Guid">{${ID}}</Parameter>`);
		expect(xml).toContain('<Parameter Type="String">R&amp;D &lt;x&gt;</Parameter>');
		expect(xml).toContain('<Method Name="CommitAll"');
		expect(xml).toContain('<Property Name="Labels" SelectAll="true">');
		expect(xml.match(/<ObjectPath Id=/g)).toHaveLength(4);
		expect(key).toBeGreaterThan(term);
		const ids = xml.match(/ Id="\d+"/g) ?? [];
		expect(new Set(ids).size).toBe(ids.length);
	});

	it('uses a given term store', () => {
		const request = new CsomRequest();
		request.termStore(ID);
		expect(request.toXml()).toContain(`Name="GetById"><Parameters><Parameter Type="Guid">{${ID}}`);
	});

	it('uses a collection query for lists', () => {
		const request = new CsomRequest();
		const { store } = request.termStore();
		request.query(request.property(store, 'Groups'), { collection: true });
		expect(request.toXml()).toContain('<ChildItemQuery SelectAllProperties="true">');
	});
});

describe('parseCsomResponse', () => {
	it('maps action IDs to results', () => {
		const result = parseCsomResponse([
			{ ErrorInfo: null },
			3,
			{ IsNull: false },
			5,
			{ _Child_Items_: [{ Name: 'A' }, { Name: 'B' }] },
		]);
		expect(result.items(5)).toHaveLength(2);
		expect(result.record(99)).toEqual({});
		expect(result.items(99)).toEqual([]);
	});

	it('throws the SharePoint error message', () => {
		expect(() =>
			parseCsomResponse([{ ErrorInfo: { ErrorMessage: 'Term set not found' } }]),
		).toThrow('Term set not found');
		expect(() => parseCsomResponse({})).toThrow();
	});
});

describe('mappers', () => {
	it('maps a term with labels', () => {
		const term = mapTerm(
			{
				Id: `/Guid(${ID})/`,
				Name: 'Marketing',
				Description: 'Text',
				PathOfTerm: 'Marketing',
				CreatedDate: '/Date(1700000000000)/',
				LastModifiedDate: '/Date(1700000000000)/',
				TermsCount: 2,
				IsRoot: true,
				IsDeprecated: false,
				IsAvailableForTagging: true,
				Labels: {
					_Child_Items_: [
						{ Value: 'Marketing', Language: 1033, IsDefaultForLanguage: true },
						{ Value: 'Mercadeo', Language: 3082, IsDefaultForLanguage: true },
					],
				},
			},
			'set-id',
			'group-id',
		);
		expect(term).toMatchObject({
			id: ID,
			name: 'Marketing',
			childCount: 2,
			termSetId: 'set-id',
			groupId: 'group-id',
			labels: [
				{ value: 'Marketing', language: 1033, isDefault: true },
				{ value: 'Mercadeo', language: 3082, isDefault: true },
			],
		});
	});

	it('maps groups and term sets', () => {
		const group = mapGroup({ Id: `/Guid(${ID})/`, Name: 'Group', IsSystemGroup: true });
		expect(group).toMatchObject({ id: ID, name: 'Group', isSystemGroup: true, description: '' });
		const set = mapTermSet(
			{ Id: `/Guid(${ID})/`, Name: 'Set', IsOpenForTermCreation: true },
			group,
		);
		expect(set).toMatchObject({
			name: 'Set',
			groupId: ID,
			groupName: 'Group',
			isOpenForTermCreation: true,
		});
	});
});

describe('matchesLabel', () => {
	const term = {
		name: 'Marketing',
		labels: [
			{ value: 'Marketing', language: 1033, isDefault: true },
			{ value: 'Mercadeo', language: 3082, isDefault: false },
		],
	} as Pick<Term, 'name' | 'labels'>;

	it('matches exact, starts with and contains', () => {
		expect(matchesLabel(term, 'marketing', 'exact')).toBe(true);
		expect(matchesLabel(term, 'market', 'exact')).toBe(false);
		expect(matchesLabel(term, 'Mark', 'startsWith')).toBe(true);
		expect(matchesLabel(term, 'keting', 'startsWith')).toBe(false);
		expect(matchesLabel(term, 'keting', 'contains')).toBe(true);
	});

	it('limits the match to one language', () => {
		expect(matchesLabel(term, 'Mercadeo', 'exact', 1033)).toBe(false);
		expect(matchesLabel(term, 'Mercadeo', 'exact', 3082)).toBe(true);
	});
});
