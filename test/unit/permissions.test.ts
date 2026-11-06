import { describe, expect, it } from 'vitest';

import {
	decodePermissions,
	diffAssignments,
	flattenAssignments,
	isSystemAccount,
	parsePrincipal,
	parseRoleId,
	toLoginName,
} from '../../nodes/SharePointPnp/shared/securables';

describe('toLoginName', () => {
	it('builds a claims login name from an email address', () => {
		expect(toLoginName(' user@contoso.com ')).toBe('i:0#.f|membership|user@contoso.com');
	});

	it('keeps a login name that has a claims prefix', () => {
		expect(toLoginName('c:0t.c|tenant|abc')).toBe('c:0t.c|tenant|abc');
	});

	it('rejects empty input', () => {
		expect(() => toLoginName('  ')).toThrow();
	});
});

describe('parsePrincipal', () => {
	it('reads digits as an ID', () => {
		expect(parsePrincipal('12')).toEqual({ kind: 'id', id: 12 });
	});

	it('reads an email address as a user', () => {
		expect(parsePrincipal('user@contoso.com')).toEqual({
			kind: 'user',
			loginName: 'i:0#.f|membership|user@contoso.com',
		});
	});

	it('reads other text as a group name', () => {
		expect(parsePrincipal('Example Members')).toEqual({ kind: 'group', name: 'Example Members' });
	});

	it('rejects empty input', () => {
		expect(() => parsePrincipal('')).toThrow();
	});
});

describe('parseRoleId', () => {
	it('returns a number for digits and the name otherwise', () => {
		expect(parseRoleId('1073741827')).toBe(1073741827);
		expect(parseRoleId('Contribute')).toBe('Contribute');
	});
});

describe('diffAssignments', () => {
	const entry = (principalId: number, roleDefinitionId: number) => ({
		principalId,
		roleDefinitionId,
	});

	it('adds missing pairs and removes extra pairs', () => {
		const result = diffAssignments(
			[entry(1, 10), entry(2, 10), entry(2, 11)],
			[entry(1, 10), entry(3, 12)],
		);
		expect(result.toAdd).toEqual([entry(3, 12)]);
		expect(result.toRemove).toEqual([entry(2, 10), entry(2, 11)]);
	});

	it('does nothing when both lists are equal', () => {
		const result = diffAssignments([entry(1, 10)], [entry(1, 10)]);
		expect(result).toEqual({ toAdd: [], toRemove: [] });
	});

	it('ignores duplicate desired entries', () => {
		const result = diffAssignments([], [entry(1, 10), entry(1, 10)]);
		expect(result.toAdd).toEqual([entry(1, 10)]);
	});

	it('removes everything for an empty desired list', () => {
		const result = diffAssignments([entry(1, 10)], []);
		expect(result.toRemove).toEqual([entry(1, 10)]);
	});
});

describe('flattenAssignments', () => {
	it('creates one pair for each role binding', () => {
		expect(
			flattenAssignments([
				{ PrincipalId: 1, RoleDefinitionBindings: [{ Id: 10 }, { Id: 11 }] },
				{ PrincipalId: 2 },
			]),
		).toEqual([
			{ principalId: 1, roleDefinitionId: 10 },
			{ principalId: 1, roleDefinitionId: 11 },
		]);
	});
});

describe('decodePermissions', () => {
	it('lists the set permissions', () => {
		// ViewListItems (bit 0) and OpenItems (bit 5)
		expect(decodePermissions(0b100001, 0)).toEqual(['ViewListItems', 'OpenItems']);
	});

	it('reads the high word', () => {
		// ManageWeb is kind 31 (bit 30); kind 37 UseClientIntegration is bit 36 (high bit 4)
		expect(decodePermissions(0, 1 << 4)).toEqual(['UseClientIntegration']);
	});

	it('returns nothing for an empty mask', () => {
		expect(decodePermissions(0, 0)).toEqual([]);
	});
});

describe('isSystemAccount', () => {
	it('detects built-in accounts', () => {
		expect(isSystemAccount({ LoginName: 'SHAREPOINT\\system' })).toBe(true);
		expect(isSystemAccount({ LoginName: 'c:0(.s|true' })).toBe(true);
		expect(
			isSystemAccount({ LoginName: 'i:0#.f|membership|user@contoso.com', IsHiddenInUI: true }),
		).toBe(true);
	});

	it('accepts a normal user', () => {
		expect(isSystemAccount({ LoginName: 'i:0#.f|membership|user@contoso.com' })).toBe(false);
	});
});
