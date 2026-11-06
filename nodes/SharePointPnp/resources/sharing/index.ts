import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { body } from '@pnp/queryable/index.js';
import { SPInstance, spPost, type ISPQueryable, type SPFI } from '@pnp/sp/index.js';
import {
	SharingLinkKind,
	SharingOperationStatusCode,
	type ISharingResult,
} from '@pnp/sp/sharing/index.js';
import type { IItem } from '@pnp/sp/items/index.js';
import '@pnp/sp/sharing/index.js';
import '@pnp/sp/security/index.js';

import { buildSecurable, readTarget, targetProperties } from '../../shared/securables';
import type { TargetDescriptor } from '../../shared/securables';
import { buildSelection, selectionOptions, type SelectionValues } from '../../shared/selection';
import { getSharePointErrorMessage, toNodeError } from '../../shared/utils';
import type { OperationContext, ResourceModule } from '../types';
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
	type LinkScope,
	type LinkType,
	type ShareRole,
} from './helpers';

const RESOURCE = 'sharing';
const TARGET_OPERATIONS = [
	'createLink',
	'shareWithUsers',
	'getSharingInformation',
	'unshare',
	'deleteLink',
	'checkPermissions',
];
const show = (operations: string[], extra: IDataObject = {}) => ({
	show: { resource: [RESOURCE], operation: operations, ...extra },
});

const INFORMATION_SELECT = [
	'canAddExternalPrincipal',
	'canAddInternalPrincipal',
	'canSendEmail',
	'canUseSimplifiedRoles',
	'hasUniquePermissions',
	'currentRole',
	'requiresAccessApproval',
	'sharedObjectType',
	'directUrl',
	'webUrl',
	'defaultLinkKind',
	'domainRestrictionMode',
	'anonymousLinkExpirationRestrictionDays',
	'permissionsInformation',
];

const targetPicker = targetProperties(RESOURCE, TARGET_OPERATIONS).map((property) =>
	property.name === 'target'
		? {
				...property,
				default: 'file',
				description: 'The object to share',
				options: (property.options as Array<{ name: string; value: string }>).filter((option) =>
					['file', 'folder', 'listItem'].includes(option.value),
				),
			}
		: property,
);

const description: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: { show: { resource: [RESOURCE] } },
		options: [
			{
				name: 'Check Permissions',
				value: 'checkPermissions',
				description: 'Check which role users have on a target through sharing',
				action: 'Check sharing permissions of users',
			},
			{
				name: 'Create Link',
				value: 'createLink',
				description: 'Create or get a sharing link',
				action: 'Create a sharing link',
			},
			{
				name: 'Delete Link',
				value: 'deleteLink',
				description: 'Delete a sharing link by its kind or share ID',
				action: 'Delete a sharing link',
			},
			{
				name: 'Get Sharing Information',
				value: 'getSharingInformation',
				description: 'Get the sharing links and the people with access',
				action: 'Get sharing information',
			},
			{
				name: 'Share With Users',
				value: 'shareWithUsers',
				description: 'Give users access, with an optional email',
				action: 'Share with users',
			},
			{
				name: 'Unshare',
				value: 'unshare',
				description: 'Remove all sharing from a target',
				action: 'Remove all sharing',
			},
		],
		default: 'getSharingInformation',
	},
	...targetPicker,
	{
		displayName: 'Link Type',
		name: 'linkType',
		type: 'options',
		default: 'view',
		options: [
			{ name: 'Edit', value: 'edit' },
			{ name: 'View', value: 'view' },
		],
		description: 'Whether the link lets people view or edit',
		displayOptions: show(['createLink']),
	},
	{
		displayName: 'Link Scope',
		name: 'linkScope',
		type: 'options',
		default: 'organization',
		options: [
			{ name: 'Anyone', value: 'anyone', description: 'Anyone with the link, without sign-in' },
			{ name: 'Organization', value: 'organization', description: 'People in your organization' },
			{
				name: 'Specific People',
				value: 'specificPeople',
				description: 'The direct link of the target. Add people with Share With Users.',
			},
		],
		description: 'Who can use the link',
		displayOptions: show(['createLink']),
	},
	{
		displayName: 'Options',
		name: 'linkOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['createLink']),
		options: [
			{
				displayName: 'Expiration Date',
				name: 'expiration',
				type: 'dateTime',
				default: '',
				description:
					'When the link stops to work. Required by some tenants for Anyone links. The REST API does not support link passwords.',
			},
		],
	},
	{
		displayName: 'Users',
		name: 'users',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'user@contoso.com, other@contoso.com',
		description: 'Email addresses or login names, separated by commas',
		displayOptions: show(['shareWithUsers', 'checkPermissions']),
	},
	{
		displayName: 'Role',
		name: 'shareRole',
		type: 'options',
		default: 'view',
		options: [
			{ name: 'Edit', value: 'edit' },
			{ name: 'Owner', value: 'owner' },
			{ name: 'View', value: 'view' },
		],
		description: 'The access to give',
		displayOptions: show(['shareWithUsers']),
	},
	{
		displayName: 'Require Sign-In',
		name: 'requireSignIn',
		type: 'boolean',
		default: true,
		description: 'Whether users must sign in. If off, an Anyone link can go in the email.',
		displayOptions: show(['shareWithUsers']),
	},
	{
		displayName: 'Send Email',
		name: 'sendEmail',
		type: 'boolean',
		default: false,
		description: 'Whether SharePoint sends an invitation email to the users',
		displayOptions: show(['shareWithUsers']),
	},
	{
		displayName: 'Email Subject',
		name: 'emailSubject',
		type: 'string',
		default: '',
		displayOptions: show(['shareWithUsers'], { sendEmail: [true] }),
	},
	{
		displayName: 'Email Message',
		name: 'emailBody',
		type: 'string',
		typeOptions: { rows: 3 },
		default: '',
		displayOptions: show(['shareWithUsers'], { sendEmail: [true] }),
	},
	{
		displayName: 'Options',
		name: 'shareOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['shareWithUsers']),
		options: [
			{
				displayName: 'Propagate ACL',
				name: 'propagateAcl',
				type: 'boolean',
				default: false,
				description:
					'Whether to apply the sharing to all children of a folder, also items with unique permissions',
			},
		],
	},
	{
		displayName: 'Link Kind',
		name: 'deleteKind',
		type: 'options',
		default: 'organizationView',
		options: [
			{ name: 'Anyone Edit', value: 'anyoneEdit' },
			{ name: 'Anyone View', value: 'anyoneView' },
			{ name: 'Organization Edit', value: 'organizationEdit' },
			{ name: 'Organization View', value: 'organizationView' },
			{ name: 'Specific People', value: 'specificPeople' },
		],
		description: 'The kind of link to delete',
		displayOptions: show(['deleteLink']),
	},
	{
		displayName: 'Share ID',
		name: 'shareId',
		type: 'string',
		default: '',
		description:
			'The share ID of one link, from Create Link or Get Sharing Information. Leave empty to delete the link of the chosen kind.',
		displayOptions: show(['deleteLink']),
	},
	{
		displayName: 'Options',
		name: 'informationOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(['getSharingInformation']),
		options: [
			{
				displayName: 'Max Principals',
				name: 'maxPrincipalsToReturn',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 50,
				description: 'The most people to return',
			},
			{
				displayName: 'Include Inherited Links',
				name: 'populateInheritedLinks',
				type: 'boolean',
				default: false,
				description: 'Whether to include links that the target gets from its parent',
			},
			...selectionOptions(INFORMATION_SELECT),
		],
	},
];

type Shareable = Pick<
	IItem,
	| 'getShareLink'
	| 'getSharingInformation'
	| 'unshare'
	| 'unshareLink'
	| 'deleteSharingLinkByKind'
	| 'checkSharingPermissions'
> &
	ISPQueryable;

function describeTarget(target: TargetDescriptor): IDataObject {
	return {
		target: target.kind,
		...(target.path ? { path: target.path } : {}),
		...(target.itemId ? { itemId: target.itemId } : {}),
	};
}

function getShareable(
	this: IExecuteFunctions,
	sp: SPFI,
	siteUrl: string,
	itemIndex: number,
): { shareable: Shareable; target: TargetDescriptor } {
	const target = readTarget.call(this, siteUrl, itemIndex);
	if (!['file', 'folder', 'listItem'].includes(target.kind)) {
		throw new NodeOperationError(
			this.getNode(),
			'Sharing works on a file, a folder or a list item',
			{ itemIndex },
		);
	}
	return { shareable: buildSecurable(sp, target) as unknown as Shareable, target };
}

async function withPolicyHint<T>(
	context: IExecuteFunctions,
	itemIndex: number,
	kind: SharingLinkKind | undefined,
	action: () => Promise<T>,
): Promise<T> {
	try {
		return await action();
	} catch (error) {
		const httpError = error as { isHttpRequestError?: boolean; message?: string; status?: number };
		if (httpError?.isHttpRequestError) {
			const hint = explainSharingError(
				getSharePointErrorMessage(httpError as { message: string }),
				kind,
			);
			if (hint) throw new NodeOperationError(context.getNode(), hint, { itemIndex });
		}
		const nodeError = toNodeError(context.getNode(), error, itemIndex);
		throw nodeError;
	}
}

async function createLink(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const kind = toLinkKind(
		this.getNodeParameter('linkType', itemIndex) as LinkType,
		this.getNodeParameter('linkScope', itemIndex) as LinkScope,
	);
	const options = this.getNodeParameter('linkOptions', itemIndex, {}) as { expiration?: string };
	let expiration: Date | null;
	try {
		expiration = toExpiration(options.expiration);
	} catch (error) {
		throw new NodeOperationError(this.getNode(), (error as Error).message, { itemIndex });
	}
	const response = await withPolicyHint(this, itemIndex, kind, () =>
		shareable.getShareLink(kind, expiration ?? undefined),
	);
	const info = response.sharingLinkInfo;
	return {
		...describeTarget(target),
		url: info?.Url,
		shareId: info?.ShareId,
		linkKind: linkKindName(info?.LinkKind ?? kind),
		isEditLink: info?.IsEditLink,
		allowsAnonymousAccess: info?.AllowsAnonymousAccess,
		isActive: info?.IsActive,
		expiration: info?.Expiration,
		created: info?.Created,
		details: info as unknown as IDataObject,
	};
}

function throwOnFailure(
	context: IExecuteFunctions,
	result: ISharingResult,
	itemIndex: number,
): void {
	const code = result.StatusCode ?? SharingOperationStatusCode.CompletedSuccessfully;
	if (code < 0) {
		const name = SharingOperationStatusCode[code] ?? String(code);
		const hint = explainSharingError(result.ErrorMessage ?? '');
		throw new NodeOperationError(
			context.getNode(),
			hint ?? `Sharing failed (${name}). ${result.ErrorMessage ?? ''}`.trim(),
			{ itemIndex },
		);
	}
}

async function shareWithUsers(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const users = parseRecipients(this.getNodeParameter('users', itemIndex, '') as string);
	if (!users.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one user', { itemIndex });
	}
	const role = toSharingRole(this.getNodeParameter('shareRole', itemIndex) as ShareRole);
	const sendEmail = this.getNodeParameter('sendEmail', itemIndex, false) as boolean;
	const options = this.getNodeParameter('shareOptions', itemIndex, {}) as {
		propagateAcl?: boolean;
	};

	const definitions = await sp.web.roleDefinitions
		.select('Id')
		.filter(`RoleTypeKind eq ${toRoleTypeKind(role)}`)
		.top(1)();
	if (!definitions.length) {
		throw new NodeOperationError(
			this.getNode(),
			'The site has no permission level for this role. Use another role.',
			{ itemIndex },
		);
	}

	const payload = buildShareObjectBody({
		loginNames: users,
		roleId: definitions[0].Id,
		requireSignIn: this.getNodeParameter('requireSignIn', itemIndex, true) as boolean,
		propagateAcl: options.propagateAcl ?? false,
		email: sendEmail
			? {
					subject: this.getNodeParameter('emailSubject', itemIndex, '') as string,
					body: this.getNodeParameter('emailBody', itemIndex, '') as string,
				}
			: undefined,
	});

	const result = await withPolicyHint(this, itemIndex, undefined, () =>
		spPost<ISharingResult>(SPInstance(shareable, 'shareObject'), body(payload)),
	);
	throwOnFailure(this, result, itemIndex);
	return {
		...describeTarget(target),
		role: this.getNodeParameter('shareRole', itemIndex) as string,
		emailSent: sendEmail,
		statusCode: result.StatusCode,
		users: (result.UniquelyPermissionedUsers ?? []) as unknown as IDataObject[],
		invitedUsers: (result.InvitedUsers ?? []) as unknown as IDataObject[],
		groupsSharedWith: (result.GroupsSharedWith ?? []) as IDataObject[],
	};
}

async function getSharingInformation(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const options = this.getNodeParameter('informationOptions', itemIndex, {}) as SelectionValues & {
		maxPrincipalsToReturn?: number;
		populateInheritedLinks?: boolean;
	};
	const { select, expand } = buildSelection(INFORMATION_SELECT, options);
	const information = await shareable.getSharingInformation(
		{
			maxPrincipalsToReturn: options.maxPrincipalsToReturn ?? 50,
			clientSupportedFeatures: '12',
			populateInheritedLinks: options.populateInheritedLinks ?? false,
		},
		expand,
		select.length ? select : ['*'],
	);
	return { ...describeTarget(target), ...(information as unknown as IDataObject) };
}

async function unshare(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const result = await shareable.unshare();
	throwOnFailure(this, result, itemIndex);
	return { ...describeTarget(target), unshared: true, statusCode: result?.StatusCode };
}

const DELETE_KINDS: Record<string, SharingLinkKind> = {
	anyoneEdit: SharingLinkKind.AnonymousEdit,
	anyoneView: SharingLinkKind.AnonymousView,
	organizationEdit: SharingLinkKind.OrganizationEdit,
	organizationView: SharingLinkKind.OrganizationView,
	specificPeople: SharingLinkKind.Direct,
};

async function deleteLink(this: IExecuteFunctions, { sp, siteUrl, itemIndex }: OperationContext) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const kind = DELETE_KINDS[this.getNodeParameter('deleteKind', itemIndex) as string];
	if (kind === undefined) {
		throw new NodeOperationError(this.getNode(), 'The link kind is not valid', { itemIndex });
	}
	const shareId = String(this.getNodeParameter('shareId', itemIndex, '') ?? '').trim();
	if (shareId) await shareable.unshareLink(kind, shareId);
	else await shareable.deleteSharingLinkByKind(kind);
	return {
		...describeTarget(target),
		deleted: true,
		linkKind: linkKindName(kind),
		anyoneLink: isAnyoneLink(kind),
		...(shareId ? { shareId } : {}),
	};
}

async function checkPermissions(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
) {
	const { shareable, target } = getShareable.call(this, sp, siteUrl, itemIndex);
	const users = parseRecipients(this.getNodeParameter('users', itemIndex, '') as string);
	if (!users.length) {
		throw new NodeOperationError(this.getNode(), 'Add at least one user', { itemIndex });
	}
	const recipients = users.map((user) =>
		user.includes('@') && !user.includes('|') ? { email: user } : { alias: user },
	);
	const result = await shareable.checkSharingPermissions(recipients);
	return (result ?? []).map(
		(entry) =>
			({
				...describeTarget(target),
				inputEntity: entry.inputEntity,
				resolvedEntity: entry.resolvedEntity,
				hasAccess: entry.hasAccess,
				role: entry.role,
			}) as IDataObject,
	);
}

export const sharingResource: ResourceModule = {
	option: {
		name: 'Sharing',
		value: RESOURCE,
		description: 'Create sharing links and share files, folders and items',
	},
	description,
	operations: {
		createLink,
		shareWithUsers,
		getSharingInformation,
		unshare,
		deleteLink,
		checkPermissions,
	},
};
