const PARAMETER = 'Role Assignments (JSON)';
const ALLOWED_KEYS = ['principal', 'principalType', 'role', 'roles'];

export type PrincipalType = 'user' | 'group';

export interface RoleAssignmentInput {
	principal: string;
	principalType?: PrincipalType;
	/** Role names or IDs. Empty only when roles are optional (Remove). */
	roles: Array<string | number>;
}

export interface ParseOptions {
	/** Whether an entry can have no role. Remove uses this to mean "all roles of the principal". */
	rolesOptional: boolean;
	/** Whether an empty array is valid. Set uses this to clear all role assignments. */
	allowEmpty: boolean;
}

function fail(message: string): never {
	throw new Error(`${PARAMETER}: ${message}`);
}

function isText(value: unknown): value is string | number {
	return (
		(typeof value === 'string' && value.trim() !== '') ||
		(typeof value === 'number' && Number.isFinite(value))
	);
}

function parseRoles(entry: Record<string, unknown>, label: string): Array<string | number> {
	const hasRole = entry.role !== undefined && entry.role !== null && entry.role !== '';
	const hasRoles = entry.roles !== undefined && entry.roles !== null;
	if (hasRole && hasRoles) fail(`${label} has both "role" and "roles". Use only one of them.`);
	if (hasRole) {
		if (!isText(entry.role)) fail(`${label}: "role" must be a role name or ID.`);
		return [typeof entry.role === 'string' ? entry.role.trim() : entry.role];
	}
	if (hasRoles) {
		if (!Array.isArray(entry.roles) || entry.roles.length === 0) {
			fail(`${label}: "roles" must be a list with at least one role name or ID.`);
		}
		const roles = entry.roles as unknown[];
		if (!roles.every(isText)) fail(`${label}: every item of "roles" must be a role name or ID.`);
		return roles.map((role) => (typeof role === 'string' ? role.trim() : (role as number)));
	}
	return [];
}

function parseEntry(value: unknown, index: number, options: ParseOptions): RoleAssignmentInput {
	const label = `entry ${index + 1}`;
	if (value === null || typeof value !== 'object' || Array.isArray(value)) {
		fail(`${label} must be an object, for example { "principal": "Example Members", "role": "Read" }.`);
	}
	const entry = value as Record<string, unknown>;
	const unknownKeys = Object.keys(entry).filter((key) => !ALLOWED_KEYS.includes(key));
	if (unknownKeys.length) {
		fail(
			`${label} has unknown ${unknownKeys.length > 1 ? 'properties' : 'property'} "${unknownKeys.join('", "')}". Allowed properties are ${ALLOWED_KEYS.join(', ')}.`,
		);
	}
	if (!isText(entry.principal)) fail(`${label} needs a "principal" (a name, an ID or an email address).`);
	const result: RoleAssignmentInput = {
		principal: String(entry.principal).trim(),
		roles: parseRoles(entry, label),
	};
	if (entry.principalType !== undefined && entry.principalType !== null && entry.principalType !== '') {
		const type = typeof entry.principalType === 'string' ? entry.principalType.trim().toLowerCase() : '';
		if (type !== 'user' && type !== 'group') {
			fail(`${label}: "principalType" must be "user" or "group".`);
		}
		result.principalType = type;
	}
	if (!options.rolesOptional && result.roles.length === 0) {
		fail(`${label} needs a "role" or "roles".`);
	}
	return result;
}

/** Reads and checks the JSON of the permission operations. Accepts JSON text or an array. */
export function parseRoleAssignmentsJson(
	raw: unknown,
	options: ParseOptions,
): RoleAssignmentInput[] {
	let parsed = raw;
	if (typeof raw === 'string') {
		if (raw.trim() === '') fail('the value is empty. Add a JSON array.');
		try {
			parsed = JSON.parse(raw);
		} catch (error) {
			fail(`the value is not valid JSON (${(error as Error).message}).`);
		}
	}
	if (!Array.isArray(parsed)) {
		fail('the value must be an array, for example [{ "principal": "Example Members", "role": "Read" }].');
	}
	const list = parsed as unknown[];
	if (list.length === 0 && !options.allowEmpty) fail('the array has no entries. Add at least one.');
	return list.map((entry, index) => parseEntry(entry, index, options));
}

export interface ResolvedAssignment {
	principalId: number;
	/** Empty means all roles of the principal. */
	roleDefinitionIds: number[];
}

/** Merges entries of the same principal. An entry without roles wins, because it means all roles. */
export function mergeByPrincipal(entries: ResolvedAssignment[]): ResolvedAssignment[] {
	const merged = new Map<number, ResolvedAssignment>();
	for (const entry of entries) {
		const current = merged.get(entry.principalId);
		if (!current) {
			merged.set(entry.principalId, {
				principalId: entry.principalId,
				roleDefinitionIds: [...new Set(entry.roleDefinitionIds)],
			});
		} else if (current.roleDefinitionIds.length === 0 || entry.roleDefinitionIds.length === 0) {
			current.roleDefinitionIds = [];
		} else {
			current.roleDefinitionIds = [
				...new Set([...current.roleDefinitionIds, ...entry.roleDefinitionIds]),
			];
		}
	}
	return [...merged.values()];
}
