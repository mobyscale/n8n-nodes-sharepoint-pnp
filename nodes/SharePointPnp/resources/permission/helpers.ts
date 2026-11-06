import type { IDataObject } from 'n8n-workflow';

import type { SelectionValues } from '../../shared/selection';

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
