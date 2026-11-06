import { RoleType, SharingLinkKind, SharingRole } from '@pnp/sp/sharing/index.js';

export type LinkType = 'view' | 'edit';
export type LinkScope = 'anyone' | 'organization' | 'specificPeople';
export type ShareRole = 'view' | 'edit' | 'owner';

const LINK_KINDS: Record<LinkScope, Record<LinkType, SharingLinkKind>> = {
	anyone: { view: SharingLinkKind.AnonymousView, edit: SharingLinkKind.AnonymousEdit },
	organization: { view: SharingLinkKind.OrganizationView, edit: SharingLinkKind.OrganizationEdit },
	specificPeople: { view: SharingLinkKind.Direct, edit: SharingLinkKind.Direct },
};

export function toLinkKind(type: LinkType, scope: LinkScope): SharingLinkKind {
	const kind = LINK_KINDS[scope]?.[type];
	if (kind === undefined) {
		throw new Error(`The link scope "${scope}" or type "${type}" is not valid`);
	}
	return kind;
}

export function isAnyoneLink(kind: SharingLinkKind): boolean {
	return kind === SharingLinkKind.AnonymousView || kind === SharingLinkKind.AnonymousEdit;
}

export function linkKindName(kind: SharingLinkKind): string {
	return SharingLinkKind[kind] ?? String(kind);
}

export function toSharingRole(role: ShareRole): SharingRole {
	switch (role) {
		case 'view':
			return SharingRole.View;
		case 'edit':
			return SharingRole.Edit;
		case 'owner':
			return SharingRole.Owner;
		default:
			throw new Error(`The role "${String(role)}" is not valid`);
	}
}

/** The site permission level (RoleTypeKind) that backs each sharing role. */
export function toRoleTypeKind(role: SharingRole): RoleType {
	switch (role) {
		case SharingRole.Edit:
			return RoleType.Contributor;
		case SharingRole.Owner:
			return RoleType.Administrator;
		default:
			return RoleType.Reader;
	}
}

/** Splits comma, semicolon or line separated text, or accepts a list, and removes duplicates. */
export function parseRecipients(input: string | string[]): string[] {
	const entries = Array.isArray(input) ? input : String(input ?? '').split(/[,;\n]/);
	return [...new Set(entries.map((entry) => String(entry).trim()).filter(Boolean))];
}

/** Returns a date for a date input, or null when the input is empty. Throws for an invalid date. */
export function toExpiration(input: string | undefined | null): Date | null {
	const value = String(input ?? '').trim();
	if (!value) return null;
	const date = new Date(value);
	if (Number.isNaN(date.getTime())) throw new Error(`"${value}" is not a valid expiration date`);
	return date;
}

export function buildShareObjectBody(args: {
	loginNames: string[];
	roleId: number;
	requireSignIn: boolean;
	propagateAcl: boolean;
	email?: { subject?: string; body: string };
}): Record<string, unknown> {
	const body: Record<string, unknown> = {
		includeAnonymousLinkInEmail: !args.requireSignIn,
		peoplePickerInput: JSON.stringify(args.loginNames.map((Key) => ({ Key }))),
		propagateAcl: args.propagateAcl,
		roleValue: `role:${args.roleId}`,
		useSimplifiedRoles: true,
	};
	if (args.email) {
		body.emailBody = args.email.body;
		body.emailSubject = args.email.subject ?? '';
		body.sendEmail = true;
	}
	return body;
}

const POLICY_PATTERN =
	/anonymous|anyone|external|guest|not allowed|disabled|policy|sharing is|expiration/i;

/** Adds a clear hint to a SharePoint error that a tenant or site sharing policy can cause. */
export function explainSharingError(message: string, kind?: SharingLinkKind): string | undefined {
	const anyone = kind !== undefined && isAnyoneLink(kind);
	if (!anyone && !POLICY_PATTERN.test(message)) return undefined;
	const subject = anyone
		? 'Anyone links are probably not allowed'
		: 'A sharing policy probably blocks this action';
	return `${subject}. Check the external sharing setting of the tenant and the site, and the link expiration rules. SharePoint said: ${message}`;
}
