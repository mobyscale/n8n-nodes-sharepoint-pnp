import type { IDataObject, IExecuteFunctions, INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError, sleep } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import { SPQueryable, spPost } from '@pnp/sp/index.js';
import { Item } from '@pnp/sp/items/index.js';

import type { OperationContext } from '../resources/types';
import { getDestinationPath, undefinedIfNotFound } from './filesFolders';
import { toServerRelativePath } from './utils';

export type NameConflictBehavior = 'fail' | 'replace' | 'keepBoth';

/** The values of SPMigrationNameConflictBehavior in PnP Core. */
const NAME_CONFLICT_BEHAVIOR: Record<NameConflictBehavior, number> = {
	fail: 0,
	replace: 1,
	keepBoth: 2,
};

/** The values of MigrationJobState in PnP Core. */
const JOB_STATE_NAMES: Record<number, string> = { 0: 'Finished', 2: 'Queued', 4: 'Processing' };

const FATAL_EVENT = 'JobFatalError';
const ERROR_EVENTS = new Set(['JobError', FATAL_EVENT]);
const MAX_LISTED_ISSUES = 100;

export interface CopyJobSettings {
	move: boolean;
	nameConflictBehavior: NameConflictBehavior;
	newName?: string;
	allowSchemaMismatch: boolean;
	allowSmallerVersionLimitOnDestination: boolean;
	bypassSharedLock: boolean;
	clientEtags: string[];
	excludeChildren: boolean;
	ignoreVersionHistory: boolean;
	moveButKeepSource: boolean;
}

/** The "Add Option" fields of the Copy Job method, as n8n gives them. */
export interface CopyJobOptionFields {
	allowSchemaMismatch?: boolean;
	allowSmallerVersionLimitOnDestination?: boolean;
	bypassSharedLock?: boolean;
	clientEtags?: string;
	continueOnJobErrors?: boolean;
	deleteEmptySourceFolder?: boolean;
	excludeChildren?: boolean;
	ignoreVersionHistory?: boolean;
	moveButKeepSource?: boolean;
	resetRoleInheritanceAfterMove?: boolean;
}

export interface CopyJobInfo {
	jobId: string;
	jobQueueUri: string;
	encryptionKey: string;
}

export function parseList(value: string | undefined): string[] {
	return (value ?? '')
		.split(/[\r\n,]+/)
		.map((entry) => entry.trim())
		.filter(Boolean);
}

/** Maps the n8n parameters to the settings of the copy job. */
export function mapCopyJobSettings(
	move: boolean,
	fields: { nameConflictBehavior?: NameConflictBehavior; newName?: string },
	options: CopyJobOptionFields,
): CopyJobSettings {
	return {
		move,
		nameConflictBehavior: fields.nameConflictBehavior ?? 'fail',
		newName: fields.newName?.trim() || undefined,
		allowSchemaMismatch: options.allowSchemaMismatch ?? true,
		allowSmallerVersionLimitOnDestination: options.allowSmallerVersionLimitOnDestination ?? false,
		bypassSharedLock: options.bypassSharedLock ?? false,
		clientEtags: parseList(options.clientEtags),
		excludeChildren: options.excludeChildren ?? false,
		ignoreVersionHistory: options.ignoreVersionHistory ?? false,
		moveButKeepSource: move && (options.moveButKeepSource ?? false),
	};
}

/** Builds the body of the CreateCopyJobs request. The option names are the ones of CopyMigrationOptions. */
export function buildCopyJobPayload(
	sourceUrl: string,
	destinationFolderUrl: string,
	settings: CopyJobSettings,
): IDataObject {
	const options: IDataObject = {
		AllowSchemaMismatch: settings.allowSchemaMismatch,
		AllowSmallerVersionLimitOnDestination: settings.allowSmallerVersionLimitOnDestination,
		BypassSharedLock: settings.bypassSharedLock,
		ExcludeChildren: settings.excludeChildren,
		IgnoreVersionHistory: settings.ignoreVersionHistory,
		IsMoveMode: settings.move,
		MoveButKeepSource: settings.moveButKeepSource,
		NameConflictBehavior: NAME_CONFLICT_BEHAVIOR[settings.nameConflictBehavior],
	};
	if (settings.clientEtags.length) options.ClientEtags = settings.clientEtags;
	if (settings.newName) options.CustomizedItemName = [settings.newName];
	return { exportObjectUris: [sourceUrl], destinationUri: destinationFolderUrl, options };
}

export function toAbsoluteUrl(siteUrl: string, serverRelativePath: string): string {
	return `${new URL(siteUrl).origin}${encodeURI(serverRelativePath)}`;
}

function text(value: unknown): string | undefined {
	return typeof value === 'string' && value !== '' ? value : undefined;
}

/** Reads the job info from an object or a JSON string. Accepts camelCase and PascalCase names. */
export function parseJobInfo(
	node: INode,
	value: unknown,
	itemIndex: number,
): CopyJobInfo & IDataObject {
	const fail = (message: string) => new NodeOperationError(node, message, { itemIndex });
	let data = value;
	if (typeof data === 'string') {
		try {
			data = JSON.parse(data);
		} catch {
			throw fail('The job info is not valid JSON');
		}
	}
	if (!data || typeof data !== 'object' || Array.isArray(data)) {
		throw fail('The job info must be an object with jobId, jobQueueUri and encryptionKey');
	}
	const source = data as IDataObject;
	const jobId = text(source.jobId ?? source.JobId);
	const jobQueueUri = text(source.jobQueueUri ?? source.JobQueueUri);
	const encryptionKey = text(source.encryptionKey ?? source.EncryptionKey);
	if (!jobId || !jobQueueUri || !encryptionKey) {
		throw fail('The job info must have jobId, jobQueueUri and encryptionKey');
	}
	const result: CopyJobInfo & IDataObject = { jobId, jobQueueUri, encryptionKey };
	for (const key of ['sourcePath', 'destinationPath'] as const) {
		if (text(source[key])) result[key] = source[key];
	}
	return result;
}

export function parseLogs(logs: unknown): IDataObject[] {
	if (!Array.isArray(logs)) return [];
	return logs.map((entry) => {
		if (typeof entry !== 'string') return entry as IDataObject;
		try {
			return JSON.parse(entry) as IDataObject;
		} catch {
			return { Message: entry };
		}
	});
}

function toNumber(value: unknown): number {
	const number = Number(value);
	return Number.isFinite(number) ? number : 0;
}

function toIssue(log: IDataObject): IDataObject {
	return {
		event: log.Event ?? null,
		message: typeof log.Message === 'string' ? log.Message.trim() : null,
		objectType: log.ObjectType ?? null,
		url: log.Url ?? null,
		errorCode: log.ErrorCode ?? null,
	};
}

function countFolders(statsByType: unknown): number {
	if (typeof statsByType !== 'string') return 0;
	try {
		const stats = JSON.parse(statsByType) as { SPFolder?: { Count?: number } };
		return toNumber(stats.SPFolder?.Count);
	} catch {
		return 0;
	}
}

export interface JobSummary extends IDataObject {
	files: number;
	folders: number;
	bytes: number;
	objectsProcessed: number;
	errors: number;
	warnings: number;
}

export interface LogDigest {
	summary: JobSummary;
	errors: IDataObject[];
	warnings: IDataObject[];
	fatal: boolean;
	targetUrl?: string;
}

/**
 * Reads the job logs. The counters come from the last progress or end entry of the import phase,
 * because the export phase only reads the source. SharePoint returns each log entry once.
 */
export function digestLogs(logs: IDataObject[]): LogDigest {
	const errors = logs.filter((log) => ERROR_EVENTS.has(String(log.Event))).map(toIssue);
	const warnings = logs.filter((log) => log.Event === 'JobWarning').map(toIssue);
	const counters = logs.filter(
		(log) =>
			(log.Event === 'JobEnd' || log.Event === 'JobProgress') &&
			log.MigrationDirection === 'Import',
	);
	const last = counters[counters.length - 1] ?? {};
	const endErrors = logs
		.filter((log) => log.Event === 'JobEnd')
		.map((log) => toNumber(log.TotalErrors));
	const targets = logs.filter((log) => log.Event === 'JobFinishedObjectInfo');
	const target = targets[targets.length - 1]?.TargetObjectSiteRelativeUrl;
	return {
		summary: {
			files: toNumber(last.FilesCreated),
			folders: countFolders(last.ObjectsStatsByType),
			bytes: toNumber(last.BytesProcessed),
			objectsProcessed: toNumber(last.ObjectsProcessed),
			errors: Math.max(errors.length, ...endErrors, toNumber(last.TotalErrors)),
			warnings: Math.max(warnings.length, toNumber(last.TotalWarnings)),
		},
		errors: errors.slice(0, MAX_LISTED_ISSUES),
		warnings: warnings.slice(0, MAX_LISTED_ISSUES),
		fatal: logs.some((log) => log.Event === FATAL_EVENT),
		targetUrl: typeof target === 'string' && target ? target : undefined,
	};
}

export type JobStatus = 'running' | 'completed' | 'failed';

/** Builds the output of a copy job: its info, state, summary and the problems that were logged. */
export function buildJobStatus(
	info: CopyJobInfo,
	jobState: number | undefined,
	logs: IDataObject[],
): IDataObject & { status: JobStatus; finished: boolean } {
	const digest = digestLogs(logs);
	const finished = jobState === 0 || digest.fatal;
	const failed = digest.summary.errors > 0 && finished;
	const status: JobStatus = failed ? 'failed' : finished ? 'completed' : 'running';
	return {
		status,
		finished,
		completed: status === 'completed',
		jobId: info.jobId,
		jobQueueUri: info.jobQueueUri,
		encryptionKey: info.encryptionKey,
		jobState: jobState === undefined ? null : (JOB_STATE_NAMES[jobState] ?? String(jobState)),
		summary: digest.summary,
		errors: digest.errors,
		warnings: digest.warnings,
	};
}

export interface CopyJobProgress {
	JobState?: number;
	Logs?: unknown;
}

export type ProgressReader = (info: CopyJobInfo) => Promise<CopyJobProgress>;

export interface WaitOptions {
	maxWaitMs: number;
	intervalMs?: number;
	now?: () => number;
	pause?: (ms: number) => Promise<void>;
}

/** Polls the job until it finishes or the maximum wait time has passed. Collects all logs. */
export async function waitForCopyJob(
	read: ProgressReader,
	info: CopyJobInfo,
	options: WaitOptions,
): Promise<{ jobState: number | undefined; logs: IDataObject[] }> {
	const { intervalMs = 2000, now = Date.now, pause = sleep } = options;
	const deadline = now() + options.maxWaitMs;
	const logs: IDataObject[] = [];
	let polls = 0;
	for (;;) {
		const progress = await read(info);
		polls++;
		logs.push(...parseLogs(progress.Logs));
		const fatal = logs.some((log) => log.Event === FATAL_EVENT);
		// A job that was just created can report "finished" before its first log entry exists.
		const finished = progress.JobState === 0 && (logs.length > 0 || polls > 2);
		if (fatal || finished || now() >= deadline) return { jobState: progress.JobState, logs };
		await pause(intervalMs);
	}
}

interface RawJobInfo {
	EncryptionKey?: string;
	JobId?: string;
	JobQueueUri?: string;
}

function unwrapCollection(response: unknown): RawJobInfo[] {
	if (Array.isArray(response)) return response as RawJobInfo[];
	const data = response as { value?: RawJobInfo[]; d?: { CreateCopyJobs?: unknown } };
	if (Array.isArray(data.value)) return data.value;
	const wrapped = data.d?.CreateCopyJobs as { results?: RawJobInfo[] } | RawJobInfo[] | undefined;
	if (Array.isArray(wrapped)) return wrapped;
	return wrapped?.results ?? [];
}

function postSite(sp: SPFI, siteUrl: string, path: string, body: IDataObject) {
	return spPost(SPQueryable([sp.web, `${siteUrl}/_api/site/${path}`]), {
		body: JSON.stringify(body),
		headers: { 'Content-Type': 'application/json;odata=nometadata' },
	});
}

export async function createCopyJob(
	node: INode,
	sp: SPFI,
	siteUrl: string,
	payload: IDataObject,
	itemIndex: number,
): Promise<CopyJobInfo> {
	const [job] = unwrapCollection(await postSite(sp, siteUrl, 'CreateCopyJobs', payload));
	if (!job?.JobId || !job.JobQueueUri || !job.EncryptionKey) {
		throw new NodeOperationError(node, 'SharePoint did not return a copy job', { itemIndex });
	}
	return { jobId: job.JobId, jobQueueUri: job.JobQueueUri, encryptionKey: job.EncryptionKey };
}

export function progressReader(sp: SPFI, siteUrl: string): ProgressReader {
	return async (info) =>
		(await postSite(sp, siteUrl, 'GetCopyJobProgress', {
			copyJobInfo: {
				EncryptionKey: info.encryptionKey,
				JobId: info.jobId,
				JobQueueUri: info.jobQueueUri,
			},
		})) as CopyJobProgress;
}

/** Reads the current progress of a job once. */
export async function getCopyJobStatus(
	sp: SPFI,
	siteUrl: string,
	info: CopyJobInfo,
): Promise<IDataObject> {
	const progress = await progressReader(sp, siteUrl)(info);
	return buildJobStatus(info, progress.JobState, parseLogs(progress.Logs));
}

export function describeIssues(issues: IDataObject[], limit = 3): string {
	return issues
		.slice(0, limit)
		.map((issue) => {
			const message = String(issue.message ?? issue.event ?? 'Unknown error');
			return message.length > 300 ? `${message.slice(0, 300)}...` : message;
		})
		.join(' | ');
}

type CopyJobResource = 'folder' | 'file';

const jobDisplay = (resource: CopyJobResource, kind: 'move' | 'copy', extra: IDataObject = {}) => ({
	show: { resource: [resource], operation: [kind], method: ['copyJob'], ...extra },
});

export function methodProperty(resource: CopyJobResource): INodeProperties {
	return {
		displayName: 'Method',
		name: 'method',
		type: 'options',
		default: 'copyJob',
		description: `How SharePoint copies or moves the ${resource}`,
		displayOptions: { show: { resource: [resource], operation: ['copy', 'move'] } },
		options: [
			{
				name: 'Copy Job',
				value: 'copyJob',
				description:
					'Use the asynchronous migration API of SharePoint. It scales to large content and is the recommended method. A move removes the source in SharePoint.',
			},
			{
				name: 'Direct',
				value: 'direct',
				description: 'Run the operation in one request. Best for small content.',
			},
		],
	};
}

/** The fields of the Copy Job method. They show only when the method is "Copy Job". */
export function copyJobProperties(
	resource: CopyJobResource,
	kind: 'move' | 'copy',
): INodeProperties[] {
	const verb = kind === 'move' ? 'move' : 'copy';
	const isFolder = resource === 'folder';
	const options: INodeProperties[] = [
		{
			displayName: 'Allow Schema Mismatch',
			name: 'allowSchemaMismatch',
			type: 'boolean',
			default: true,
			description:
				'Whether to allow the operation when the destination list has a different schema than the source list. This is on when you do not add the option.',
		},
		{
			displayName: 'Allow Smaller Version Limit On Destination',
			name: 'allowSmallerVersionLimitOnDestination',
			type: 'boolean',
			default: false,
			description:
				'Whether to allow the operation when the destination keeps fewer file versions than the source. By default this is off to prevent loss of versions.',
		},
		{
			displayName: 'Bypass Shared Lock',
			name: 'bypassSharedLock',
			type: 'boolean',
			default: false,
			description: `Whether to ${verb} files that have a co-authoring shared lock`,
		},
		{
			displayName: 'Client ETags',
			name: 'clientEtags',
			type: 'string',
			default: '',
			description:
				'Comma-separated list of ETags. If the ETag of the source does not match, SharePoint does not do the operation. Leave empty for no check.',
		},
		{
			displayName: 'Continue on Job Errors',
			name: 'continueOnJobErrors',
			type: 'boolean',
			default: false,
			description:
				'Whether to return the job result when the job reports errors. If off, the item fails when the job reports errors. The errors are in the output in both cases.',
		},
		...(isFolder
			? [
					{
						displayName: 'Exclude Children',
						name: 'excludeChildren',
						type: 'boolean',
						default: false,
						description: `Whether to ${verb} only the folder itself, without its files and subfolders`,
					} as INodeProperties,
				]
			: []),
		{
			displayName: 'Ignore Version History',
			name: 'ignoreVersionHistory',
			type: 'boolean',
			default: false,
			description: 'Whether to keep only the current version of each file at the destination',
		},
	];

	if (kind === 'move') {
		options.push(
			{
				displayName: 'Move But Keep Source',
				name: 'moveButKeepSource',
				type: 'boolean',
				default: false,
				description:
					'Whether to keep the source after the move. This works like a copy that also keeps the version history and the metadata.',
			},
			{
				displayName: 'Reset Role Inheritance After Move',
				name: 'resetRoleInheritanceAfterMove',
				type: 'boolean',
				default: false,
				description: `Whether to reset the role inheritance of the moved ${resource} at the destination, so that it uses the permissions of its new parent. This runs only after the job finished without errors, and only when the node waits for the job.`,
			},
		);
		if (isFolder) {
			options.splice(5, 0, {
				displayName: 'Delete Empty Source Folder',
				name: 'deleteEmptySourceFolder',
				type: 'boolean',
				default: true,
				description:
					'Whether to delete the source folder after a successful move when the folder is empty. SharePoint removes the source itself at the end of the job, so this is a safety net. It never deletes a folder that has content. This is on when you do not add the option, and it runs only when the node waits for the job.',
			});
		}
	}

	return [
		{
			displayName: 'Name Conflict Behavior',
			name: 'nameConflictBehavior',
			type: 'options',
			default: 'fail',
			description: `What to do when a ${resource} with the same name exists in the destination`,
			displayOptions: jobDisplay(resource, kind),
			options: [
				{ name: 'Fail', value: 'fail', description: 'The job reports an error' },
				{
					name: 'Keep Both',
					value: 'keepBoth',
					description: 'SharePoint adds a number to the name of the new item',
				},
				{ name: 'Replace', value: 'replace', description: 'Replace the existing item' },
			],
		},
		{
			displayName: 'New Name',
			name: 'copyJobNewName',
			type: 'string',
			default: '',
			description: `A new ${resource} name for the destination. Leave empty to keep the name.`,
			displayOptions: jobDisplay(resource, kind),
		},
		{
			displayName: 'Wait for Completion',
			name: 'waitForCompletion',
			type: 'boolean',
			default: true,
			description: `Whether to wait until SharePoint has finished the job. If off, the node returns at once with the job info (job ID, job queue URI, encryption key, source and destination). The ${verb} then continues in SharePoint after the node finishes. To check it, use the operation "Get Copy Job Status" of the File/Folder resource.`,
			displayOptions: jobDisplay(resource, kind),
		},
		{
			displayName: 'Max Wait (Seconds)',
			name: 'maxWaitSeconds',
			type: 'number',
			typeOptions: { minValue: 1 },
			default: 300,
			description: `How long to wait for the job to finish. If the job is still running after this time, the node does not fail. It returns the job info with the status "running", and the ${verb} continues in SharePoint. To check it later, use the operation "Get Copy Job Status".`,
			displayOptions: jobDisplay(resource, kind, { waitForCompletion: [true] }),
		},
		{
			displayName: 'Copy Job Options',
			name: 'copyJobOptions',
			type: 'collection',
			placeholder: 'Add Option',
			default: {},
			displayOptions: jobDisplay(resource, kind),
			options,
		},
	];
}

async function readTarget(
	sp: SPFI,
	resource: CopyJobResource,
	path: string,
): Promise<IDataObject | undefined> {
	const query =
		resource === 'folder'
			? sp.web.getFolderByServerRelativePath(path)
			: sp.web.getFileByServerRelativePath(path);
	return (await undefinedIfNotFound(query())) as IDataObject | undefined;
}

export async function deleteSourceFolderIfEmpty(sp: SPFI, path: string): Promise<boolean> {
	const folder = sp.web.getFolderByServerRelativePath(path);
	const info = await undefinedIfNotFound(folder.select('ItemCount')());
	if (!info || Number(info.ItemCount) !== 0) return false;
	await folder.delete();
	return true;
}

async function resetRoleInheritance(
	sp: SPFI,
	resource: CopyJobResource,
	path: string,
): Promise<void> {
	const target =
		resource === 'folder'
			? sp.web.getFolderByServerRelativePath(path).listItemAllFields
			: sp.web.getFileByServerRelativePath(path).listItemAllFields;
	await Item(target).resetRoleInheritance();
}

/** Starts a copy or move job for a folder or a file, and waits for it when the node is set to wait. */
export async function runMoveCopyJob(
	this: IExecuteFunctions,
	{ sp, siteUrl, itemIndex }: OperationContext,
	spec: {
		resource: CopyJobResource;
		kind: 'move' | 'copy';
		sourcePath: string;
		destinationFolder: string;
	},
): Promise<IDataObject> {
	const { resource, kind, sourcePath, destinationFolder } = spec;
	const newName = this.getNodeParameter('copyJobNewName', itemIndex, '') as string;
	const optionFields = this.getNodeParameter(
		'copyJobOptions',
		itemIndex,
		{},
	) as CopyJobOptionFields;
	const settings = mapCopyJobSettings(
		kind === 'move',
		{
			nameConflictBehavior: this.getNodeParameter(
				'nameConflictBehavior',
				itemIndex,
				'fail',
			) as NameConflictBehavior,
			newName,
		},
		optionFields,
	);
	const wait = this.getNodeParameter('waitForCompletion', itemIndex, true) as boolean;
	const maxWaitSeconds = this.getNodeParameter('maxWaitSeconds', itemIndex, 300) as number;
	let destinationPath = getDestinationPath(sourcePath, destinationFolder, settings.newName);

	const payload = buildCopyJobPayload(
		toAbsoluteUrl(siteUrl, sourcePath),
		toAbsoluteUrl(siteUrl, destinationFolder),
		settings,
	);
	const info = await createCopyJob(this.getNode(), sp, siteUrl, payload, itemIndex);
	const base: IDataObject = { sourcePath, destinationPath, method: 'copyJob' };

	if (!wait) {
		return { ...buildJobStatus(info, undefined, []), ...base };
	}

	const { jobState, logs } = await waitForCopyJob(progressReader(sp, siteUrl), info, {
		maxWaitMs: maxWaitSeconds * 1000,
	});
	const status = buildJobStatus(info, jobState, logs);
	const { targetUrl } = digestLogs(logs);
	if (targetUrl) destinationPath = toServerRelativePath(siteUrl, targetUrl);
	const result: IDataObject = { ...status, ...base, destinationPath };

	if (status.status === 'failed') {
		if (!optionFields.continueOnJobErrors) {
			throw new NodeOperationError(
				this.getNode(),
				`The copy job ${info.jobId} reported ${(status.summary as IDataObject).errors} error(s): ${describeIssues(status.errors as IDataObject[])}`,
				{ itemIndex },
			);
		}
		return result;
	}
	if (status.status !== 'completed') return result;

	if (kind === 'move' && !settings.moveButKeepSource) {
		if (resource === 'folder' && (optionFields.deleteEmptySourceFolder ?? true)) {
			result.sourceFolderDeleted = await deleteSourceFolderIfEmpty(sp, sourcePath);
		}
	}
	if (kind === 'move' && optionFields.resetRoleInheritanceAfterMove) {
		await resetRoleInheritance(sp, resource, destinationPath);
		result.roleInheritanceReset = true;
	}
	const target = await readTarget(sp, resource, destinationPath);
	return target ? { ...result, ...target } : result;
}

export function getJobInfoParameter(
	this: IExecuteFunctions,
	itemIndex: number,
): CopyJobInfo & IDataObject {
	return parseJobInfo(this.getNode(), this.getNodeParameter('jobInfo', itemIndex), itemIndex);
}
