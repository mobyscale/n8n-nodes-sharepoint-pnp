import type { INodeProperties } from 'n8n-workflow';

import { pathRLC } from '../../shared/descriptions';
import { principalRLC, roleRLC } from '../../shared/securables';

const show = (extra: Record<string, string[]> = {}) => ({
	show: { resource: ['permission'], operation: ['applyMapping'], ...extra },
});

/** The properties of Permission > Apply Mapping. The Input Mode property is shared with Add, Remove and Set. */
export const mappingProperties: INodeProperties[] = [
	pathRLC({
		displayName: 'Base Folder',
		name: 'baseFolder',
		description:
			'The folder that the mapping paths are relative to. A server-relative path (/sites/Example/Shared Documents/Archive), a site-relative path (Shared Documents/Archive), or an absolute URL.',
		displayOptions: show(),
	}),
	{
		displayName: 'Mappings',
		name: 'mappings',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true, sortable: true },
		placeholder: 'Add Mapping',
		default: { entries: [] },
		description: 'The folders, principals and roles to apply',
		displayOptions: show({ inputMode: ['fields'] }),
		options: [
			{
				displayName: 'Entries',
				name: 'entries',
				values: [
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						default: '',
						placeholder: 'Reports/*/Drafts',
						description:
							'The folder path relative to the base folder. "*" matches all child folders. "**" matches folders at any depth, and at the end of a path all folders below. Use "." for the base folder.',
					},
					{ ...principalRLC, required: false },
					{
						...roleRLC,
						required: false,
						description: 'A permission level of the site',
						typeOptions: { loadOptionsDependsOn: ['site.value'] },
					},
				],
			},
		],
	},
	{
		displayName: 'Mappings (JSON)',
		name: 'mappingsJson',
		type: 'json',
		required: true,
		default: '[]',
		description:
			'An array of objects with "path", "principal" and "role" (or "roles"). "principalType" is optional. See the documentation for an example.',
		displayOptions: show({ inputMode: ['json'] }),
	},
	{
		displayName: 'Mode',
		name: 'mappingMode',
		type: 'options',
		default: 'set',
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Add the mapped role assignments. Other role assignments stay.',
			},
			{
				name: 'Set',
				value: 'set',
				description:
					'Each folder gets exactly the mapped role assignments. Other role assignments are removed.',
			},
		],
		displayOptions: show(),
	},
	{
		displayName: 'Options',
		name: 'mappingOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show(),
		options: [
			{
				displayName: 'Break Inheritance',
				name: 'breakInheritance',
				type: 'options',
				default: 'ifInheriting',
				options: [
					{
						name: 'Always',
						value: 'always',
						description: 'Break inheritance on each folder, also when it has unique permissions',
					},
					{
						name: 'Only If Inheriting',
						value: 'ifInheriting',
						description: 'Break inheritance only on folders that inherit permissions',
					},
				],
			},
			{
				displayName: 'Clear Subscopes',
				name: 'clearSubscopes',
				type: 'boolean',
				default: false,
				description:
					'Whether to remove unique permissions of child objects when the operation breaks inheritance',
			},
			{
				displayName: 'Copy Current Assignments First',
				name: 'copyCurrentFirst',
				type: 'boolean',
				default: false,
				description:
					'Whether to copy the current role assignments when Set mode breaks inheritance. Add mode always copies them.',
			},
			{
				displayName: 'Create Missing Folders',
				name: 'createMissingFolders',
				type: 'boolean',
				default: false,
				description:
					'Whether to create a folder of a path without "*" or "**" when it does not exist. Paths with wildcards only match folders that exist.',
			},
			{
				displayName: 'If Folder Is Missing',
				name: 'ifFolderMissing',
				type: 'options',
				default: 'skip',
				description: 'What to do when a folder of a path without "*" or "**" does not exist',
				options: [
					{
						name: 'Skip',
						value: 'skip',
						description: 'Skip the folder and report it in "missing"',
					},
					{ name: 'Fail', value: 'fail', description: 'Stop before any permission change' },
				],
				displayOptions: { hide: { createMissingFolders: [true] } },
			},
			{
				displayName: 'Output',
				name: 'output',
				type: 'options',
				default: 'single',
				options: [
					{
						name: 'One Item',
						value: 'single',
						description: 'One item with a "folders" list for each input item',
					},
					{
						name: 'One Item Per Folder',
						value: 'perFolder',
						description: 'One item for each folder and each missing folder',
					},
				],
			},
			{
				displayName: 'Skip Unresolved Principals',
				name: 'skipUnresolved',
				type: 'boolean',
				default: false,
				description:
					'Whether to skip a mapping whose principal or role cannot be found, and report it in "warnings", instead of failing before any change',
			},
		],
	},
];
