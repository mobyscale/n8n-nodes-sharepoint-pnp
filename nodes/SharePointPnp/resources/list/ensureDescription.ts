import type { INodeProperties } from 'n8n-workflow';

import { listRLC } from '../../shared/descriptions';
import { LIST_TEMPLATES } from './helpers';

const show = { show: { resource: ['list'], operation: ['ensure'] } };

export const ensureDescription: INodeProperties[] = [
	{
		displayName: 'Title',
		name: 'title',
		type: 'string',
		default: '',
		required: true,
		description:
			'The title of the list. The node finds the list by this title when the URL is empty, and uses it when it creates the list.',
		displayOptions: show,
	},
	{
		displayName: 'URL',
		name: 'listUrl',
		type: 'string',
		default: '',
		placeholder: 'Shared Documents',
		description:
			'The URL of the list root folder: site-relative (Archive, Lists/Tasks), server-relative (/sites/Example/Archive) or absolute. If set, the node finds the list by this URL and creates the list with the last part of the URL. Empty uses the title.',
		displayOptions: show,
	},
	{
		displayName: 'Template',
		name: 'template',
		type: 'options',
		options: LIST_TEMPLATES,
		default: 101,
		description:
			'The kind of list to create. The node does not change the template of a list that exists.',
		displayOptions: show,
	},
	{
		displayName: 'Description',
		name: 'listDescription',
		type: 'string',
		default: '',
		description:
			'The description of a new list. With "Apply Settings to Existing List", it also replaces the description of a list that exists.',
		displayOptions: show,
	},
	{
		displayName: 'Options',
		name: 'ensureOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show,
		options: [
			{
				displayName: 'Apply Settings to Existing List',
				name: 'applyToExisting',
				type: 'boolean',
				default: false,
				description:
					'Whether to also apply these options when the list exists. If off, the node only returns a list that exists and does not change it.',
			},
			{
				displayName: 'Content Types to Add Names or IDs',
				name: 'contentTypes',
				type: 'multiOptions',
				typeOptions: { loadOptionsMethod: 'getSiteContentTypeOptions' },
				default: [],
				description:
					'Site content types to add to the list. The node turns on content types and skips a content type that the list has. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				...listRLC,
				displayName: 'Copy View From List',
				name: 'copyViewSourceList',
				required: false,
				description:
					'The list that has the view to copy. It must be in the same site. Also set "View to Copy".',
			},
			{
				displayName: 'Default Content Type Name or ID',
				name: 'defaultContentType',
				type: 'options',
				typeOptions: { loadOptionsMethod: 'getSiteContentTypeOptions' },
				default: '',
				description:
					'The content type to make the default of the list. Give a site content type, a list content type ID or a content type name. The list must have it, or it must be in "Content Types to Add". Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Enable Content Types',
				name: 'enableContentTypes',
				type: 'boolean',
				default: true,
				description: 'Whether the list allows the management of content types',
			},
			{
				displayName: 'Enable Minor Versions',
				name: 'enableMinorVersions',
				type: 'boolean',
				default: true,
				description: 'Whether to keep minor (draft) versions. Only for libraries.',
			},
			{
				displayName: 'Enable Versioning',
				name: 'enableVersioning',
				type: 'boolean',
				default: true,
				description: 'Whether to keep versions of items and files',
			},
			{
				displayName: 'Major Version Limit',
				name: 'majorVersionLimit',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 500,
				description: 'The number of major versions to keep',
			},
			{
				displayName: 'Minor Version Limit',
				name: 'minorVersionLimit',
				type: 'number',
				typeOptions: { minValue: 0 },
				default: 10,
				description:
					'The number of major versions that keep their minor versions. Only for libraries.',
			},
			{
				displayName: 'Permissions',
				name: 'permissions',
				type: 'fixedCollection',
				default: {},
				placeholder: 'Add Permissions',
				description: 'Give the list its own permissions',
				options: [
					{
						displayName: 'Settings',
						name: 'settings',
						values: [
							{
								displayName: 'Break Inheritance',
								name: 'breakInheritance',
								type: 'boolean',
								default: true,
								description:
									'Whether to stop the list from inheriting permissions from the site. Role assignments always break inheritance first.',
							},
							{
								displayName: 'Clear Subscopes',
								name: 'clearSubscopes',
								type: 'boolean',
								default: false,
								description:
									'Whether to remove unique permissions of items in the list when the node breaks inheritance',
							},
							{
								displayName: 'Copy Role Assignments',
								name: 'copyRoleAssignments',
								type: 'boolean',
								default: false,
								description:
									'Whether to copy the role assignments of the site when the node breaks inheritance',
							},
							{
								displayName: 'Role Assignment Mode',
								name: 'roleAssignmentMode',
								type: 'options',
								default: 'replace',
								options: [
									{
										name: 'Replace',
										value: 'replace',
										description:
											'The list has only the role assignments of the JSON, like Permission > Set',
									},
									{
										name: 'Add',
										value: 'add',
										description: 'Add the role assignments of the JSON and keep the others',
									},
								],
							},
							{
								displayName: 'Role Assignments (JSON)',
								name: 'roleAssignmentsJson',
								type: 'json',
								default: '',
								placeholder: '[{ "principal": "Example Members", "role": "Edit" }]',
								description:
									'An array of objects with "principal" and "role" (or "roles"), the same as Permission > Set in JSON mode. Empty changes no role assignments.',
							},
						],
					},
				],
			},
			{
				displayName: 'Show on Quick Launch',
				name: 'onQuickLaunch',
				type: 'boolean',
				default: false,
				description: 'Whether to show the list in the site navigation (Quick Launch)',
			},
			{
				displayName: 'View to Copy',
				name: 'copyViewName',
				type: 'string',
				default: '',
				placeholder: 'All Documents',
				description:
					'The title or ID of the view in "Copy View From List". If the list has a view with the same title, the view is not copied.',
			},
		],
	},
];
