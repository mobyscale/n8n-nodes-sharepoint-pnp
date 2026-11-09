import type { INodeProperties } from 'n8n-workflow';

import { pathRLC } from '../../shared/descriptions';

const show = (operation: string, extra: Record<string, string[]> = {}) => ({
	show: { resource: ['folder'], operation: [operation], ...extra },
});

const ensureFolders: INodeProperties[] = [
	{
		displayName: 'Folders Input',
		name: 'foldersInputMode',
		type: 'options',
		default: 'rows',
		options: [
			{ name: 'Rows', value: 'rows', description: 'Add one row for each folder' },
			{ name: 'JSON', value: 'json', description: 'Give the folders as a JSON array' },
		],
		displayOptions: show('ensureFolders'),
	},
	{
		displayName: 'Folders',
		name: 'folders',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true, sortable: true },
		placeholder: 'Add Folder',
		default: {},
		description:
			'The folders to make below the folder. A parent that two folders share is created once.',
		displayOptions: show('ensureFolders', { foldersInputMode: ['rows'] }),
		options: [
			{
				displayName: 'Folder',
				name: 'values',
				values: [
					{
						displayName: 'Append To Multi-Choice Field',
						name: 'appendField',
						type: 'string',
						default: '',
						description: 'The internal name of a multi-choice field to add values to',
					},
					{
						displayName: 'Append Values',
						name: 'appendValues',
						type: 'string',
						default: '',
						description:
							'The choices to add, separated by commas or line breaks. The current choices stay.',
					},
					{
						displayName: 'Fields Always (JSON)',
						name: 'fieldsAlways',
						type: 'json',
						default: '',
						description:
							'Field values to set on each run. The node does not write a value that is already equal.',
					},
					{
						displayName: 'Fields When Created (JSON)',
						name: 'fieldsWhenCreated',
						type: 'json',
						default: '',
						description:
							'Field values to set only when this run creates the folder. An object with internal field names as keys.',
					},
					{
						displayName: 'Key',
						name: 'key',
						type: 'string',
						default: '',
						placeholder: 'quarter',
						description:
							'A unique name for the folder. The output has the folder under this key, for example folders.quarter.serverRelativeUrl.',
					},
					{
						displayName: 'Path',
						name: 'path',
						type: 'string',
						default: '',
						placeholder: 'Reports/2024/Q1',
						description:
							'The path below the folder. It can have more than one level, separated by "/".',
					},
				],
			},
		],
	},
	{
		displayName: 'Folders (JSON)',
		name: 'foldersJson',
		type: 'json',
		required: true,
		default: '[]',
		description:
			'An array of objects with "key" and "path", and optional "fieldsWhenCreated", "fieldsAlways" and "appendToMultiChoice"',
		displayOptions: show('ensureFolders', { foldersInputMode: ['json'] }),
	},
	{
		displayName: 'Output',
		name: 'foldersOutput',
		type: 'options',
		default: 'single',
		options: [
			{
				name: 'One Item With All Folders',
				value: 'single',
				description: 'One item with a "folders" object that has an entry for each key',
			},
			{
				name: 'One Item Per Folder',
				value: 'perFolder',
				description: 'One item for each folder, with its key',
			},
		],
		displayOptions: show('ensureFolders'),
	},
];

const ensurePath: INodeProperties[] = [
	{
		displayName: 'Segments Input',
		name: 'segmentsInputMode',
		type: 'options',
		default: 'rows',
		options: [
			{ name: 'Rows', value: 'rows', description: 'Add one row for each folder level' },
			{ name: 'JSON', value: 'json', description: 'Give the segments as a JSON array' },
		],
		displayOptions: show('ensurePath'),
	},
	{
		displayName: 'Segments',
		name: 'segments',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true, sortable: true },
		placeholder: 'Add Segment',
		default: {},
		description: 'The folder levels below the folder, from the top level down',
		displayOptions: show('ensurePath', { segmentsInputMode: ['rows'] }),
		options: [
			{
				displayName: 'Segment',
				name: 'values',
				values: [
					{
						displayName: 'Append To Multi-Choice Field',
						name: 'appendField',
						type: 'string',
						default: '',
						description: 'The internal name of a multi-choice field to add values to',
					},
					{
						displayName: 'Append Values',
						name: 'appendValues',
						type: 'string',
						default: '',
						description:
							'The choices to add, separated by commas or line breaks. The current choices stay.',
					},
					{
						displayName: 'Fields Always (JSON)',
						name: 'fieldsAlways',
						type: 'json',
						default: '',
						description:
							'Field values to set on each run. The node does not write a value that is already equal.',
					},
					{
						displayName: 'Fields When Created (JSON)',
						name: 'fieldsWhenCreated',
						type: 'json',
						default: '',
						description:
							'Field values to set only when this run creates the folder. An object with internal field names as keys.',
					},
					{
						displayName: 'Name',
						name: 'name',
						type: 'string',
						default: '',
						placeholder: 'Reports',
						description: "The folder name. It must not have a '/'.",
					},
				],
			},
		],
	},
	{
		displayName: 'Segments (JSON)',
		name: 'segmentsJson',
		type: 'json',
		required: true,
		default: '[]',
		description:
			'An array of folder names or segment objects. See the documentation for the format.',
		displayOptions: show('ensurePath', { segmentsInputMode: ['json'] }),
	},
];

const copyChildren: INodeProperties[] = [
	{
		...pathRLC({
			displayName: 'Destination Folder',
			name: 'childrenDestinationFolder',
			description: 'The folder to copy the children into',
		}),
		displayOptions: show('copyChildrenIfMissing'),
	},
	{
		displayName: 'Options',
		name: 'copyChildrenOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show('copyChildrenIfMissing'),
		options: [
			{
				displayName: 'Continue on Job Errors',
				name: 'continueOnJobErrors',
				type: 'boolean',
				default: false,
				description:
					'Whether to return the result when a copy job reports errors. If off, the item fails. The errors are in the output in both cases.',
			},
			{
				displayName: 'Copy Method',
				name: 'copyMethod',
				type: 'options',
				default: 'copyJob',
				options: [
					{
						name: 'Copy Job',
						value: 'copyJob',
						description:
							'Use the migration API of SharePoint. The node starts one job for each child and waits for them.',
					},
					{
						name: 'Direct',
						value: 'direct',
						description: 'Copy each child in one request. Best for small content.',
					},
				],
			},
			{
				displayName: 'Exclude Names',
				name: 'excludeNames',
				type: 'string',
				typeOptions: { rows: 2 },
				default: '',
				placeholder: 'Archive, *.tmp',
				description:
					'Names to leave out, separated by commas or line breaks. "*" matches any text. Not case-sensitive.',
			},
			{
				displayName: 'Ignore Version History',
				name: 'ignoreVersionHistory',
				type: 'boolean',
				default: false,
				description: 'Whether a copy job keeps only the current version of each file',
			},
			{
				displayName: 'Include',
				name: 'include',
				type: 'options',
				default: 'folders',
				description: 'Which children to copy',
				options: [
					{ name: 'Both', value: 'both' },
					{ name: 'Files', value: 'files' },
					{ name: 'Folders', value: 'folders' },
				],
			},
			{
				displayName: 'Include Names',
				name: 'includeNames',
				type: 'string',
				typeOptions: { rows: 2 },
				default: '',
				placeholder: '01 *, Images',
				description:
					'Copy only these names, separated by commas or line breaks. "*" matches any text. Empty copies all. Not case-sensitive.',
			},
			{
				displayName: 'Max Wait (Seconds)',
				name: 'maxWaitSeconds',
				type: 'number',
				typeOptions: { minValue: 1 },
				default: 300,
				description:
					'How long to wait for all copy jobs. A job that runs longer continues in SharePoint and has the status "running".',
			},
			{
				displayName: 'Name Conflict Behavior',
				name: 'nameConflictBehavior',
				type: 'options',
				default: 'skip',
				description: 'What to do with a child whose name exists in the destination',
				options: [
					{ name: 'Fail', value: 'fail', description: 'The item fails and nothing is copied' },
					{
						name: 'Keep Both',
						value: 'keepBoth',
						description: 'Copy it, and SharePoint adds a number to the new name',
					},
					{
						name: 'Replace',
						value: 'replace',
						description: 'Copy it and replace the existing child',
					},
					{ name: 'Skip', value: 'skip', description: 'Do not copy it' },
				],
			},
			{
				displayName: 'Only If Destination Exists',
				name: 'onlyIfDestinationExists',
				type: 'boolean',
				default: false,
				description:
					'Whether to skip the copy when the destination folder does not exist. If off, a missing destination fails.',
			},
			{
				displayName: 'Reset Author and Created',
				name: 'resetAuthorAndCreated',
				type: 'boolean',
				default: false,
				description:
					'Whether to set the author to the current user and the created date to now. Only the Direct method supports this. A copy job keeps the source values.',
			},
		],
	},
];

/**
 * The options of Set Metadata Recursive. Inherit Metadata From Parent uses the same options, with
 * the folder itself included by default and the option to search up for values.
 */
function metadataOptionList(inherit: boolean): INodeProperties[] {
	const options: INodeProperties[] = [
		{
			displayName: 'Batch Size',
			name: 'batchSize',
			type: 'number',
			typeOptions: { minValue: 1, maxValue: 100 },
			default: 100,
			description: 'How many updates to send in one SharePoint batch request (1 to 100)',
		},
		{
			displayName: 'Continue On Item Error',
			name: 'continueOnItemError',
			type: 'boolean',
			default: false,
			description:
				'Whether to continue when the update of an item fails. The errors are in the output. If off, the run stops after the batch with the first error.',
		},
		{
			displayName: 'Ignore Empty Values',
			name: 'ignoreEmptyValues',
			type: 'boolean',
			default: false,
			description: 'Whether to leave out empty values, so that they do not clear current values',
		},
		{
			displayName: 'Include',
			name: 'include',
			type: 'options',
			default: 'both',
			description: 'Which items below the folder to update',
			options: [
				{ name: 'Both', value: 'both' },
				{ name: 'Files', value: 'files' },
				{ name: 'Folders', value: 'folders' },
			],
		},
		{
			displayName: 'Include The Folder Itself',
			name: 'includeFolderItself',
			type: 'boolean',
			default: false,
			description: 'Whether to update the folder too. A library root folder has no fields.',
		},
		{
			displayName: 'Max Depth',
			name: 'maxDepth',
			type: 'number',
			typeOptions: { minValue: 0 },
			default: 0,
			description:
				'How many folder levels to go down. 1 updates only the direct children. 0 has no limit.',
		},
		{
			displayName: 'Only Fields In Content Type',
			name: 'onlyFieldsInContentType',
			type: 'boolean',
			default: false,
			description:
				'Whether to set on each item only the fields that its content type has. An item without such fields is skipped.',
		},
		{
			displayName: 'Skip Unchanged',
			name: 'skipUnchanged',
			type: 'boolean',
			default: true,
			description:
				'Whether to skip a text or choice value that is already equal. Other field types are always written.',
		},
		{
			displayName: 'System Update',
			name: 'systemUpdate',
			type: 'boolean',
			default: false,
			description:
				'Whether to update without a new version and keep Modified and Modified By. This needs one more request for each batch.',
		},
	];
	if (!inherit) return options;
	const searchUp: INodeProperties = {
		displayName: 'Search Up To Nearest Folder With Values',
		name: 'searchUp',
		type: 'boolean',
		default: false,
		description:
			'Whether to go up from the parent folder, one level at a time, until a folder has a value for each field. The search stops below the library root.',
	};
	return [
		...options.map((option): INodeProperties => {
			if (option.name === 'includeFolderItself') {
				return {
					...option,
					default: true,
					description:
						'Whether to update the selected folder too. On by default for this operation.',
				};
			}
			if (option.name === 'ignoreEmptyValues') {
				return {
					...option,
					description:
						'Whether to leave out fields that have no value in the source folder, so that they do not clear current values',
				};
			}
			return option;
		}),
		searchUp,
	].sort((a, b) => a.displayName.localeCompare(b.displayName));
}

const setMetadata: INodeProperties[] = [
	{
		displayName: 'Fields Input',
		name: 'metadataInputMode',
		type: 'options',
		default: 'keyValue',
		options: [
			{ name: 'Key-Value Pairs', value: 'keyValue' },
			{ name: 'JSON', value: 'json' },
		],
		displayOptions: show('setMetadataRecursive'),
	},
	{
		displayName: 'Fields',
		name: 'metadataFields',
		type: 'fixedCollection',
		typeOptions: { multipleValues: true },
		placeholder: 'Add Field',
		default: {},
		displayOptions: show('setMetadataRecursive', { metadataInputMode: ['keyValue'] }),
		options: [
			{
				displayName: 'Field',
				name: 'values',
				values: [
					{
						displayName: 'Name',
						name: 'name',
						type: 'string',
						default: '',
						description: 'The internal name of the field',
					},
					{
						displayName: 'Value',
						name: 'value',
						type: 'string',
						default: '',
						description:
							'The value. Use commas for multi-choice values and an ISO date for date fields.',
					},
				],
			},
		],
	},
	{
		displayName: 'Fields (JSON)',
		name: 'metadataFieldsJson',
		type: 'json',
		required: true,
		default: '{}',
		description: 'An object with internal field names as keys',
		displayOptions: show('setMetadataRecursive', { metadataInputMode: ['json'] }),
	},
	{
		displayName: 'Options',
		name: 'metadataOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show('setMetadataRecursive'),
		options: metadataOptionList(false),
	},
];

const inheritMetadata: INodeProperties[] = [
	{
		displayName: 'Fields To Inherit Names or IDs',
		name: 'inheritFieldsPicker',
		type: 'multiOptions',
		typeOptions: {
			loadOptionsMethod: 'getFolderListFieldOptions',
			loadOptionsDependsOn: ['folderPath.value'],
		},
		default: [],
		description:
			'The fields to copy from the parent folder. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		displayOptions: show('inheritMetadataFromParent'),
	},
	{
		displayName: 'Other Fields To Inherit',
		name: 'inheritFields',
		type: 'string',
		default: '',
		placeholder: 'Category, Status',
		description:
			'More internal field names to copy, separated by commas or line breaks. They are added to the fields above. Give at least one field in one of the two parameters.',
		displayOptions: show('inheritMetadataFromParent'),
	},
	{
		displayName: 'Options',
		name: 'inheritOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: show('inheritMetadataFromParent'),
		options: metadataOptionList(true),
	},
];

export const provisionProperties: INodeProperties[] = [
	...ensureFolders,
	...ensurePath,
	...copyChildren,
	...setMetadata,
	...inheritMetadata,
];
