import { RoleType, SharingLinkKind, SharingRole } from '@pnp/sp/sharing/index.js';
import { describe, expect, it } from 'vitest';

import {
	buildShareObjectBody,
	explainSharingError,
	isAnyoneLink,
	linkKindName,
	parseRecipients,
	toExpiration,
	toLinkKind,
	toRoleTypeKind,
	toSharingRole,
} from '../../nodes/SharePointPnp/resources/sharing/helpers';

describe('toLinkKind', () => {
	it('maps scope and type to the PnPjs link kind', () => {
		expect(toLinkKind('view', 'anyone')).toBe(SharingLinkKind.AnonymousView);
		expect(toLinkKind('edit', 'anyone')).toBe(SharingLinkKind.AnonymousEdit);
		expect(toLinkKind('view', 'organization')).toBe(SharingLinkKind.OrganizationView);
		expect(toLinkKind('edit', 'organization')).toBe(SharingLinkKind.OrganizationEdit);
		expect(toLinkKind('view', 'specificPeople')).toBe(SharingLinkKind.Direct);
		expect(toLinkKind('edit', 'specificPeople')).toBe(SharingLinkKind.Direct);
	});

	it('rejects an unknown scope', () => {
		expect(() => toLinkKind('view', 'nobody' as never)).toThrow('not valid');
	});

	it('finds Anyone links and names kinds', () => {
		expect(isAnyoneLink(SharingLinkKind.AnonymousEdit)).toBe(true);
		expect(isAnyoneLink(SharingLinkKind.OrganizationView)).toBe(false);
		expect(linkKindName(SharingLinkKind.OrganizationView)).toBe('OrganizationView');
	});
});

describe('roles', () => {
	it('maps the role names', () => {
		expect(toSharingRole('view')).toBe(SharingRole.View);
		expect(toSharingRole('edit')).toBe(SharingRole.Edit);
		expect(toSharingRole('owner')).toBe(SharingRole.Owner);
		expect(() => toSharingRole('other' as never)).toThrow('not valid');
	});

	it('maps a sharing role to a permission level kind', () => {
		expect(toRoleTypeKind(SharingRole.View)).toBe(RoleType.Reader);
		expect(toRoleTypeKind(SharingRole.Edit)).toBe(RoleType.Contributor);
		expect(toRoleTypeKind(SharingRole.Owner)).toBe(RoleType.Administrator);
	});
});

describe('parseRecipients', () => {
	it('splits, trims and removes duplicates', () => {
		expect(parseRecipients(' a@contoso.com; b@contoso.com,\na@contoso.com ,, ')).toEqual([
			'a@contoso.com',
			'b@contoso.com',
		]);
	});

	it('accepts a list', () => {
		expect(parseRecipients(['a@contoso.com', ' '])).toEqual(['a@contoso.com']);
	});
});

describe('toExpiration', () => {
	it('returns null for empty input', () => {
		expect(toExpiration('')).toBeNull();
		expect(toExpiration(undefined)).toBeNull();
	});

	it('parses a date', () => {
		expect(toExpiration('2030-01-02T00:00:00.000Z')?.toISOString()).toBe(
			'2030-01-02T00:00:00.000Z',
		);
	});

	it('rejects an invalid date', () => {
		expect(() => toExpiration('tomorrow-ish')).toThrow('not a valid expiration date');
	});
});

describe('buildShareObjectBody', () => {
	it('builds a body without email', () => {
		expect(
			buildShareObjectBody({
				loginNames: ['user@contoso.com'],
				roleId: 1073741826,
				requireSignIn: true,
				propagateAcl: false,
			}),
		).toEqual({
			includeAnonymousLinkInEmail: false,
			peoplePickerInput: '[{"Key":"user@contoso.com"}]',
			propagateAcl: false,
			roleValue: 'role:1073741826',
			useSimplifiedRoles: true,
		});
	});

	it('adds the email fields and propagation', () => {
		const body = buildShareObjectBody({
			loginNames: ['a', 'b'],
			roleId: 5,
			requireSignIn: false,
			propagateAcl: true,
			email: { subject: 'Hello', body: 'Text' },
		});
		expect(body).toMatchObject({
			sendEmail: true,
			emailSubject: 'Hello',
			emailBody: 'Text',
			propagateAcl: true,
			includeAnonymousLinkInEmail: true,
		});
	});
});

describe('explainSharingError', () => {
	it('explains a blocked Anyone link', () => {
		const text = explainSharingError(
			'Anonymous links are not allowed',
			SharingLinkKind.AnonymousView,
		);
		expect(text).toContain('Anyone links are probably not allowed');
		expect(text).toContain('Anonymous links are not allowed');
	});

	it('explains a generic error of an Anyone link', () => {
		expect(explainSharingError('Operation is not valid', SharingLinkKind.AnonymousEdit)).toContain(
			'Anyone links are probably not allowed',
		);
	});

	it('ignores unrelated errors', () => {
		expect(explainSharingError('File not found')).toBeUndefined();
	});
});
