import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import type { SecurableQueryable } from '@pnp/sp/security/index.js';

import { runInBatches } from '../../shared/batch';
import {
	buildSecurable,
	type RoleAssignmentEntry,
	type RoleAssignmentWithBindings,
	type TargetDescriptor,
} from '../../shared/securables';
import { applySelection, type SelectionValues } from '../../shared/selection';

export const ASSIGNMENT_SELECT = [
	'PrincipalId',
	'Member/Id',
	'Member/Title',
	'Member/LoginName',
	'Member/PrincipalType',
	'Member/Email',
	'RoleDefinitionBindings/Id',
	'RoleDefinitionBindings/Name',
];

const ASSIGNMENT_EXPAND = ['Member', 'RoleDefinitionBindings'];

export async function getAssignments(
	securable: SecurableQueryable,
	options: SelectionValues = {},
): Promise<RoleAssignmentWithBindings[]> {
	return (await applySelection(
		securable.roleAssignments,
		ASSIGNMENT_SELECT,
		options,
		ASSIGNMENT_EXPAND,
	)()) as unknown as RoleAssignmentWithBindings[];
}

/** Adds and removes role assignments of a target in SharePoint $batch requests. Throws the first error. */
export async function applyAssignmentChanges(
	sp: SPFI,
	target: TargetDescriptor,
	toAdd: RoleAssignmentEntry[],
	toRemove: RoleAssignmentEntry[],
): Promise<void> {
	const changes = [
		...toAdd.map((entry) => ({ entry, add: true })),
		...toRemove.map((entry) => ({ entry, add: false })),
	];
	if (!changes.length) return;
	const results = await runInBatches(sp, changes, (batchedSp, { entry, add }) => {
		const assignments = buildSecurable(batchedSp, target).roleAssignments;
		return add
			? assignments.add(entry.principalId, entry.roleDefinitionId)
			: assignments.remove(entry.principalId, entry.roleDefinitionId);
	});
	const failed = results.find((result) => !result.ok);
	if (failed) throw failed.error;
}

interface RawAssignment {
	PrincipalId?: number;
	Member?: IDataObject;
	RoleDefinitionBindings?: Array<{ Id: number; Name?: string }>;
}

/** True when the user changed the property selection. Then the raw SharePoint object is returned too. */
export function hasCustomSelection(options: SelectionValues): boolean {
	return Boolean(
		options.returnAllProperties || options.additionalProperties?.length || options.expand?.length,
	);
}

/** Maps a role assignment of SharePoint to the output of Get Many. */
export function toAssignmentOutput(
	assignment: RawAssignment,
	hasUniqueRoleAssignments: boolean,
	options: SelectionValues = {},
): IDataObject {
	const { Member, RoleDefinitionBindings, PrincipalId } = assignment;
	return {
		hasUniqueRoleAssignments,
		principalId: PrincipalId,
		member: {
			id: Member?.Id,
			title: Member?.Title,
			loginName: Member?.LoginName,
			principalType: Member?.PrincipalType,
			email: Member?.Email,
		},
		roleDefinitions: (RoleDefinitionBindings ?? []).map((role) => ({
			id: role.Id,
			name: role.Name,
		})),
		...(hasCustomSelection(options) ? { properties: assignment as unknown as IDataObject } : {}),
	};
}
