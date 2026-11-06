import type { IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import { permissionResource } from '../../nodes/SharePointPnp/resources/permission';
import {
	mergeByPrincipal,
	parseRoleAssignmentsJson,
	type ParseOptions,
} from '../../nodes/SharePointPnp/resources/permission/json';
import { parsePrincipal } from '../../nodes/SharePointPnp/shared/securables';

const required: ParseOptions = { rolesOptional: false, allowEmpty: false };
const optionalRoles: ParseOptions = { rolesOptional: true, allowEmpty: false };

describe('parseRoleAssignmentsJson', () => {
	it('reads role and roles', () => {
		const result = parseRoleAssignmentsJson(
			'[{ "principal": "Example Members", "role": "Edit" }, { "principal": "user@contoso.com", "roles": ["Read", 1073741827] }]',
			required,
		);
		expect(result).toEqual([
			{ principal: 'Example Members', roles: ['Edit'] },
			{ principal: 'user@contoso.com', roles: ['Read', 1073741827] },
		]);
	});

	it('accepts an array that an expression returned', () => {
		expect(parseRoleAssignmentsJson([{ principal: 5, role: 'Read' }], required)).toEqual([
			{ principal: '5', roles: ['Read'] },
		]);
	});

	it('reads the principal type in any case', () => {
		const [entry] = parseRoleAssignmentsJson(
			[{ principal: 'jdoe', principalType: 'User', role: 'Read' }],
			required,
		);
		expect(entry.principalType).toBe('user');
	});

	it('allows a missing role only when roles are optional', () => {
		expect(parseRoleAssignmentsJson([{ principal: 'Example Members' }], optionalRoles)).toEqual([
			{ principal: 'Example Members', roles: [] },
		]);
		expect(() => parseRoleAssignmentsJson([{ principal: 'Example Members' }], required)).toThrow(
			/entry 1 needs a "role" or "roles"/,
		);
	});

	it('allows an empty array only when it is allowed', () => {
		expect(parseRoleAssignmentsJson('[]', { ...required, allowEmpty: true })).toEqual([]);
		expect(() => parseRoleAssignmentsJson('[]', required)).toThrow(/no entries/);
	});

	it.each([
		['', /empty/],
		['{ not json', /not valid JSON/],
		['{"principal":"a","role":"Read"}', /must be an array/],
		['[1]', /entry 1 must be an object/],
		['[[]]', /entry 1 must be an object/],
		['[{"role":"Read"}]', /entry 1 needs a "principal"/],
		['[{"principal":" ","role":"Read"}]', /entry 1 needs a "principal"/],
		['[{"principal":"a","role":"Read","extra":1}]', /unknown property "extra"/],
		['[{"principal":"a","role":"Read","roles":["Edit"]}]', /both "role" and "roles"/],
		['[{"principal":"a","roles":[]}]', /at least one role/],
		['[{"principal":"a","roles":"Read"}]', /at least one role/],
		['[{"principal":"a","roles":["Read",""]}]', /every item of "roles"/],
		['[{"principal":"a","role":{}}]', /"role" must be a role name or ID/],
		['[{"principal":"a","principalType":"team","role":"Read"}]', /"principalType" must be/],
		['[{"principal":"a","role":"Read"},{"role":"Read"}]', /entry 2 needs a "principal"/],
	])('rejects %j', (input, message) => {
		expect(() => parseRoleAssignmentsJson(input, required)).toThrow(message);
	});
});

describe('mergeByPrincipal', () => {
	it('merges roles of the same principal without duplicates', () => {
		expect(
			mergeByPrincipal([
				{ principalId: 1, roleDefinitionIds: [10, 11] },
				{ principalId: 2, roleDefinitionIds: [10] },
				{ principalId: 1, roleDefinitionIds: [11, 12] },
			]),
		).toEqual([
			{ principalId: 1, roleDefinitionIds: [10, 11, 12] },
			{ principalId: 2, roleDefinitionIds: [10] },
		]);
	});

	it('treats an entry without roles as all roles', () => {
		expect(
			mergeByPrincipal([
				{ principalId: 1, roleDefinitionIds: [10] },
				{ principalId: 1, roleDefinitionIds: [] },
				{ principalId: 1, roleDefinitionIds: [11] },
			]),
		).toEqual([{ principalId: 1, roleDefinitionIds: [] }]);
	});
});

describe('parsePrincipal with a principal type', () => {
	it('keeps the guess when there is no type', () => {
		expect(parsePrincipal('jdoe')).toEqual({ kind: 'group', name: 'jdoe' });
	});

	it('reads text as a user when the type is user', () => {
		expect(parsePrincipal('jdoe', 'user')).toEqual({ kind: 'user', loginName: 'jdoe' });
		expect(parsePrincipal('user@contoso.com', 'user')).toEqual({
			kind: 'user',
			loginName: 'i:0#.f|membership|user@contoso.com',
		});
	});

	it('reads text as a group name when the type is group', () => {
		expect(parsePrincipal('team@contoso.com', 'group')).toEqual({
			kind: 'group',
			name: 'team@contoso.com',
		});
	});

	it('reads digits as an ID for both types', () => {
		expect(parsePrincipal('12', 'group')).toEqual({ kind: 'id', id: 12 });
		expect(parsePrincipal('12', 'user')).toEqual({ kind: 'id', id: 12 });
	});
});

describe('permission input mode', () => {
	const property = (name: string, operation: string): INodeProperties | undefined =>
		permissionResource.description.find(
			(entry) =>
				entry.name === name &&
				(entry.displayOptions?.show?.operation as string[] | undefined)?.includes(operation),
		);

	it.each(['add', 'remove', 'set'])('offers the input mode and the JSON for %s', (operation) => {
		expect(property('inputMode', operation)?.default).toBe('fields');
		expect(property('roleAssignmentsJson', operation)?.displayOptions?.show?.inputMode).toEqual([
			'json',
		]);
	});

	it('hides the field inputs in JSON mode', () => {
		expect(property('principalType', 'add')?.displayOptions?.show?.inputMode).toEqual(['fields']);
		expect(property('role', 'remove')?.displayOptions?.show?.inputMode).toEqual(['fields']);
		expect(property('assignments', 'set')?.displayOptions?.show?.inputMode).toEqual(['fields']);
	});

	it.each(['add', 'remove', 'set'])('rejects invalid JSON for %s before it calls SharePoint', async (operation) => {
		const parameters: Record<string, unknown> = {
			inputMode: 'json',
			roleAssignmentsJson: '[{ "role": "Read" }]',
		};
		const context = {
			getNode: () => ({ name: 'SharePoint PnP' }),
			getNodeParameter: (name: string, _index: number, fallback?: unknown) =>
				name in parameters ? parameters[name] : fallback,
		} as unknown as IExecuteFunctions;
		const sp = { web: {} } as never;
		const error = await permissionResource.operations[operation]
			.call(context, { sp, siteUrl: 'https://contoso.sharepoint.com/sites/Example', itemIndex: 0 })
			.catch((reason: unknown) => reason);
		expect(error).toBeInstanceOf(NodeOperationError);
		expect((error as Error).message).toContain('Role Assignments (JSON): entry 1 needs a "principal"');
	});
});
