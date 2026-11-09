import type { IDataObject, INode } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

import {
	buildCopyJobPayload,
	buildJobStatus,
	createCopyJob,
	describeIssues,
	digestLogs,
	mapCopyJobSettings,
	parseList,
	progressReader,
	toAbsoluteUrl,
	waitForCopyJob,
	type CopyJobInfo,
	type NameConflictBehavior,
	type ProgressReader,
} from '../../shared/copyJob';
import { getDestinationPath, undefinedIfNotFound } from '../../shared/filesFolders';
import { toServerRelativePath } from '../../shared/utils';

export type ChildType = 'file' | 'folder';
export type IncludeChildren = 'folders' | 'files' | 'both';
export type ExistingBehavior = 'skip' | 'fail' | 'replace' | 'keepBoth';
export type SkipReason = 'exists' | 'notIncluded' | 'excluded' | 'type';

export interface Child {
	name: string;
	type: ChildType;
	serverRelativeUrl: string;
}

export interface CopyChildrenOptions {
	includeNames?: string;
	excludeNames?: string;
	include?: IncludeChildren;
	nameConflictBehavior?: ExistingBehavior;
	copyMethod?: 'copyJob' | 'direct';
	resetAuthorAndCreated?: boolean;
	onlyIfDestinationExists?: boolean;
	maxWaitSeconds?: number;
	ignoreVersionHistory?: boolean;
	continueOnJobErrors?: boolean;
}

export interface PlannedCopy extends Child {
	exists: boolean;
}

export interface CopyPlan {
	toCopy: PlannedCopy[];
	skipped: Array<Child & { reason: SkipReason }>;
	conflicts: Child[];
}

function escapeRegExp(value: string): string {
	return value.replace(/[.+^${}()|[\]\\]/g, '\\$&');
}

/** Converts a name pattern to a regular expression. "*" matches any text, "?" one character. */
export function wildcardToRegExp(pattern: string): RegExp {
	const source = escapeRegExp(pattern.trim()).replace(/\*/g, '.*').replace(/\?/g, '.');
	return new RegExp(`^${source}$`, 'i');
}

/** Reads a comma or line separated list of name patterns. */
export function parseNamePatterns(value: string | undefined): RegExp[] {
	return parseList(value).map(wildcardToRegExp);
}

export function matchesAny(name: string, patterns: RegExp[]): boolean {
	return patterns.some((pattern) => pattern.test(name));
}

/** Sorts the source children into the ones to copy and the ones to skip. Names are not case-sensitive. */
export function planCopy(
	children: Child[],
	destinationNames: Set<string>,
	options: CopyChildrenOptions,
): CopyPlan {
	const include = parseNamePatterns(options.includeNames);
	const exclude = parseNamePatterns(options.excludeNames);
	const kinds = options.include ?? 'folders';
	const behavior = options.nameConflictBehavior ?? 'skip';
	const plan: CopyPlan = { toCopy: [], skipped: [], conflicts: [] };
	for (const child of children) {
		if (kinds !== 'both' && `${child.type}s` !== kinds) {
			plan.skipped.push({ ...child, reason: 'type' });
		} else if (include.length && !matchesAny(child.name, include)) {
			plan.skipped.push({ ...child, reason: 'notIncluded' });
		} else if (matchesAny(child.name, exclude)) {
			plan.skipped.push({ ...child, reason: 'excluded' });
		} else if (destinationNames.has(child.name.toLowerCase())) {
			if (behavior === 'skip') plan.skipped.push({ ...child, reason: 'exists' });
			else if (behavior === 'fail') plan.conflicts.push(child);
			else plan.toCopy.push({ ...child, exists: true });
		} else {
			plan.toCopy.push({ ...child, exists: false });
		}
	}
	return plan;
}

/** The SharePoint calls that Copy Children If Missing needs. */
export interface CopyChildrenGateway {
	/** Lists the files and subfolders of a folder. Undefined when the folder does not exist. */
	listChildren(path: string): Promise<Child[] | undefined>;
	copyDirect(
		child: PlannedCopy,
		destinationPath: string,
		options: CopyChildrenOptions,
	): Promise<void>;
	startCopyJob(
		child: PlannedCopy,
		destinationFolder: string,
		options: CopyChildrenOptions,
	): Promise<CopyJobInfo>;
	readProgress: ProgressReader;
	/** Converts the site-relative target URL of a job log to a server-relative path. */
	toPath(siteRelativeUrl: string): string;
	now?: () => number;
	pause?: (ms: number) => Promise<void>;
}

function toEntry(child: Child, extra: IDataObject = {}): IDataObject {
	return { name: child.name, type: child.type, sourcePath: child.serverRelativeUrl, ...extra };
}

export async function copyChildrenIfMissing(
	sourcePath: string,
	destinationFolder: string,
	options: CopyChildrenOptions,
	gateway: CopyChildrenGateway,
): Promise<IDataObject> {
	const base: IDataObject = { sourcePath, destinationPath: destinationFolder };
	const destination = await gateway.listChildren(destinationFolder);
	if (!destination) {
		if (options.onlyIfDestinationExists) {
			return { ...base, destinationExists: false, changed: false, copied: [], skipped: [] };
		}
		throw new Error(`The destination folder "${destinationFolder}" does not exist`);
	}
	const source = await gateway.listChildren(sourcePath);
	if (!source) throw new Error(`The source folder "${sourcePath}" does not exist`);

	const plan = planCopy(
		source,
		new Set(destination.map((child) => child.name.toLowerCase())),
		options,
	);
	if (plan.conflicts.length) {
		throw new Error(
			`These names exist in the destination: ${plan.conflicts.map((child) => child.name).join(', ')}`,
		);
	}
	const method = options.copyMethod ?? 'copyJob';
	if (method === 'direct' && options.nameConflictBehavior === 'replace') {
		const folder = plan.toCopy.find((child) => child.type === 'folder' && child.exists);
		if (folder) {
			throw new Error(
				`The Direct method cannot replace the folder "${folder.name}". Use the Copy Job method.`,
			);
		}
	}

	const copied: IDataObject[] = [];
	if (method === 'direct') {
		for (const child of plan.toCopy) {
			const destinationPath = getDestinationPath(child.serverRelativeUrl, destinationFolder);
			await gateway.copyDirect(child, destinationPath, options);
			copied.push(toEntry(child, { destinationPath, status: 'completed' }));
		}
	} else {
		copied.push(...(await runCopyJobs(plan.toCopy, destinationFolder, options, gateway)));
	}

	return {
		...base,
		destinationExists: true,
		method,
		changed: copied.length > 0,
		copied,
		skipped: plan.skipped.map((child) => toEntry(child, { reason: child.reason })),
	};
}

/** Starts one copy job for each child, so the jobs run side by side, then waits for all of them. */
async function runCopyJobs(
	children: PlannedCopy[],
	destinationFolder: string,
	options: CopyChildrenOptions,
	gateway: CopyChildrenGateway,
): Promise<IDataObject[]> {
	const now = gateway.now ?? Date.now;
	const jobs: Array<{ child: PlannedCopy; info: CopyJobInfo }> = [];
	for (const child of children) {
		jobs.push({ child, info: await gateway.startCopyJob(child, destinationFolder, options) });
	}
	const deadline = now() + (options.maxWaitSeconds ?? 300) * 1000;
	const entries: IDataObject[] = [];
	const failures: string[] = [];
	for (const { child, info } of jobs) {
		const { jobState, logs } = await waitForCopyJob(gateway.readProgress, info, {
			maxWaitMs: Math.max(0, deadline - now()),
			now,
			pause: gateway.pause,
		});
		const status = buildJobStatus(info, jobState, logs);
		const target = digestLogs(logs).targetUrl;
		const destinationPath = target
			? gateway.toPath(target)
			: getDestinationPath(child.serverRelativeUrl, destinationFolder);
		entries.push(
			toEntry(child, {
				destinationPath,
				status: status.status,
				jobId: info.jobId,
				summary: status.summary,
				errors: status.errors,
			}),
		);
		if (status.status === 'failed') {
			failures.push(`${child.name}: ${describeIssues(status.errors as IDataObject[])}`);
		}
	}
	if (failures.length && !options.continueOnJobErrors) {
		throw new Error(`The copy job reported errors. ${failures.join(' | ')}`);
	}
	return entries;
}

interface RawChild {
	Name?: string;
	ServerRelativeUrl?: string;
	ListItemAllFields?: { Id?: number } | null;
}

/** The copy job uses "Fail" for Skip, so a name that appears after the check is not overwritten. */
function jobConflictBehavior(behavior: ExistingBehavior | undefined): NameConflictBehavior {
	return behavior === 'replace' || behavior === 'keepBoth' ? behavior : 'fail';
}

/** Connects Copy Children If Missing to SharePoint through PnPjs. */
export function createCopyChildrenGateway(
	sp: SPFI,
	siteUrl: string,
	node: INode,
	itemIndex: number,
): CopyChildrenGateway {
	return {
		async listChildren(path) {
			const folder = () => sp.web.getFolderByServerRelativePath(path);
			const info = await undefinedIfNotFound(folder().select('Exists')());
			if (!info?.Exists) return undefined;
			const [folders, files] = await Promise.all([
				folder()
					.folders.select('Name', 'ServerRelativeUrl', 'ListItemAllFields/Id')
					.expand('ListItemAllFields')() as Promise<RawChild[]>,
				folder().files.select('Name', 'ServerRelativeUrl')() as Promise<RawChild[]>,
			]);
			// Folders without a list item are system folders, such as "Forms" in a library root.
			return [
				...folders
					.filter((entry) => entry.ListItemAllFields?.Id)
					.map((entry) => ({ type: 'folder' as const, entry })),
				...files.map((entry) => ({ type: 'file' as const, entry })),
			].map(({ type, entry }) => ({
				name: String(entry.Name ?? ''),
				type,
				serverRelativeUrl: String(entry.ServerRelativeUrl ?? ''),
			}));
		},
		async copyDirect(child, destinationPath, options) {
			const moveOptions = {
				KeepBoth: options.nameConflictBehavior === 'keepBoth',
				ResetAuthorAndCreatedOnCopy: options.resetAuthorAndCreated ?? false,
				ShouldBypassSharedLocks: false,
			};
			if (child.type === 'folder') {
				await sp.web
					.getFolderByServerRelativePath(child.serverRelativeUrl)
					.copyByPath(destinationPath, moveOptions);
			} else {
				await sp.web
					.getFileByServerRelativePath(child.serverRelativeUrl)
					.copyByPath(destinationPath, options.nameConflictBehavior === 'replace', moveOptions);
			}
		},
		async startCopyJob(child, destinationFolder, options) {
			const settings = mapCopyJobSettings(
				false,
				{ nameConflictBehavior: jobConflictBehavior(options.nameConflictBehavior) },
				{ ignoreVersionHistory: options.ignoreVersionHistory },
			);
			const payload = buildCopyJobPayload(
				toAbsoluteUrl(siteUrl, child.serverRelativeUrl),
				toAbsoluteUrl(siteUrl, destinationFolder),
				settings,
			);
			return await createCopyJob(node, sp, siteUrl, payload, itemIndex);
		},
		readProgress: progressReader(sp, siteUrl),
		toPath: (siteRelativeUrl) => toServerRelativePath(siteUrl, siteRelativeUrl),
	};
}
