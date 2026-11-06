import type { IDataObject } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import {
	buildCopyJobPayload,
	buildJobStatus,
	copyJobProperties,
	digestLogs,
	mapCopyJobSettings,
	parseJobInfo,
	parseLogs,
	toAbsoluteUrl,
	waitForCopyJob,
	type CopyJobInfo,
	type CopyJobProgress,
} from '../../nodes/SharePointPnp/shared/copyJob';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';
const info: CopyJobInfo = {
	jobId: 'job-1',
	jobQueueUri: 'https://queue.example/q',
	encryptionKey: 'key',
};
const node = { name: 'Test' } as never;

const entry = (data: IDataObject) => JSON.stringify(data, null, 2).replace(/\n/g, '\r\n');

const importEnd = {
	Event: 'JobEnd',
	MigrationDirection: 'Import',
	FilesCreated: '2',
	BytesProcessed: '11',
	ObjectsProcessed: '8',
	TotalErrors: '0',
	TotalWarnings: '2',
	ObjectsStatsByType: '{"SPFolder":{"Count":2},"SPListItem":{"Count":4}}',
};

describe('copy job settings', () => {
	it('uses safe defaults', () => {
		expect(mapCopyJobSettings(false, {}, {})).toEqual({
			move: false,
			nameConflictBehavior: 'fail',
			newName: undefined,
			allowSchemaMismatch: true,
			allowSmallerVersionLimitOnDestination: false,
			bypassSharedLock: false,
			clientEtags: [],
			excludeChildren: false,
			ignoreVersionHistory: false,
			moveButKeepSource: false,
		});
	});

	it('maps all options', () => {
		const settings = mapCopyJobSettings(
			true,
			{ nameConflictBehavior: 'keepBoth', newName: ' renamed ' },
			{
				allowSchemaMismatch: false,
				allowSmallerVersionLimitOnDestination: true,
				bypassSharedLock: true,
				clientEtags: 'a, b\nc',
				excludeChildren: true,
				ignoreVersionHistory: true,
				moveButKeepSource: true,
			},
		);
		expect(settings).toEqual({
			move: true,
			nameConflictBehavior: 'keepBoth',
			newName: 'renamed',
			allowSchemaMismatch: false,
			allowSmallerVersionLimitOnDestination: true,
			bypassSharedLock: true,
			clientEtags: ['a', 'b', 'c'],
			excludeChildren: true,
			ignoreVersionHistory: true,
			moveButKeepSource: true,
		});
	});

	it('ignores "move but keep source" for a copy', () => {
		expect(mapCopyJobSettings(false, {}, { moveButKeepSource: true }).moveButKeepSource).toBe(
			false,
		);
	});
});

describe('copy job payload', () => {
	it('builds the request body with the names of CopyMigrationOptions', () => {
		const payload = buildCopyJobPayload(
			'https://contoso.sharepoint.com/sites/Example/Lib/a',
			'https://contoso.sharepoint.com/sites/Example/Lib/b',
			mapCopyJobSettings(
				true,
				{ nameConflictBehavior: 'replace', newName: 'c' },
				{ bypassSharedLock: true, clientEtags: 'x' },
			),
		);
		expect(payload).toEqual({
			exportObjectUris: ['https://contoso.sharepoint.com/sites/Example/Lib/a'],
			destinationUri: 'https://contoso.sharepoint.com/sites/Example/Lib/b',
			options: {
				AllowSchemaMismatch: true,
				AllowSmallerVersionLimitOnDestination: false,
				BypassSharedLock: true,
				ExcludeChildren: false,
				IgnoreVersionHistory: false,
				IsMoveMode: true,
				MoveButKeepSource: false,
				NameConflictBehavior: 1,
				ClientEtags: ['x'],
				CustomizedItemName: ['c'],
			},
		});
	});

	it('maps every name conflict behavior to its PnP Core value', () => {
		const value = (behavior: 'fail' | 'replace' | 'keepBoth') =>
			(
				buildCopyJobPayload(
					'a',
					'b',
					mapCopyJobSettings(false, { nameConflictBehavior: behavior }, {}),
				).options as IDataObject
			).NameConflictBehavior;
		expect([value('fail'), value('replace'), value('keepBoth')]).toEqual([0, 1, 2]);
	});

	it('builds absolute URLs with encoding', () => {
		expect(toAbsoluteUrl(siteUrl, '/sites/Example/My Lib/a b')).toBe(
			'https://contoso.sharepoint.com/sites/Example/My%20Lib/a%20b',
		);
	});
});

describe('copy job fields', () => {
	const names = (properties: ReturnType<typeof copyJobProperties>) =>
		(properties.find((property) => property.name === 'copyJobOptions')?.options ?? []).map(
			(option) => option.name,
		);

	it('has all PnP Core options for a folder move', () => {
		expect(names(copyJobProperties('folder', 'move'))).toEqual([
			'allowSchemaMismatch',
			'allowSmallerVersionLimitOnDestination',
			'bypassSharedLock',
			'clientEtags',
			'continueOnJobErrors',
			'deleteEmptySourceFolder',
			'excludeChildren',
			'ignoreVersionHistory',
			'moveButKeepSource',
			'resetRoleInheritanceAfterMove',
		]);
	});

	it('hides move and folder options where they do not apply', () => {
		expect(names(copyJobProperties('folder', 'copy'))).not.toContain('moveButKeepSource');
		expect(names(copyJobProperties('file', 'move'))).not.toContain('excludeChildren');
		expect(names(copyJobProperties('file', 'move'))).not.toContain('deleteEmptySourceFolder');
	});

	it('shows the fields only for the Copy Job method', () => {
		for (const property of copyJobProperties('file', 'copy')) {
			expect(property.displayOptions?.show?.method).toEqual(['copyJob']);
		}
	});
});

describe('job info', () => {
	it('reads camelCase, PascalCase and JSON text', () => {
		const expected = { jobId: 'a', jobQueueUri: 'b', encryptionKey: 'c' };
		expect(parseJobInfo(node, { jobId: 'a', jobQueueUri: 'b', encryptionKey: 'c' }, 0)).toEqual(
			expected,
		);
		expect(parseJobInfo(node, { JobId: 'a', JobQueueUri: 'b', EncryptionKey: 'c' }, 0)).toEqual(
			expected,
		);
		expect(parseJobInfo(node, JSON.stringify(expected), 0)).toEqual(expected);
	});

	it('keeps the source and destination for the output', () => {
		const result = parseJobInfo(
			node,
			{ ...info, sourcePath: '/s', destinationPath: '/d', other: 1 },
			0,
		);
		expect(result.sourcePath).toBe('/s');
		expect(result.destinationPath).toBe('/d');
		expect(result.other).toBeUndefined();
	});

	it('rejects incomplete or invalid info', () => {
		expect(() => parseJobInfo(node, { jobId: 'a' }, 0)).toThrow('jobQueueUri');
		expect(() => parseJobInfo(node, 'not json', 0)).toThrow('JSON');
		expect(() => parseJobInfo(node, undefined, 0)).toThrow('object');
	});
});

describe('job logs', () => {
	it('parses JSON entries and keeps plain text', () => {
		expect(parseLogs([entry({ Event: 'JobStart' }), 'plain', { Event: 'x' }])).toEqual([
			{ Event: 'JobStart' },
			{ Message: 'plain' },
			{ Event: 'x' },
		]);
		expect(parseLogs(undefined)).toEqual([]);
	});

	it('counts files, folders and bytes from the import phase', () => {
		const digest = digestLogs([
			{
				Event: 'JobEnd',
				MigrationDirection: 'Export',
				FilesCreated: '9',
				BytesProcessed: '999',
				TotalErrors: '0',
			},
			{
				Event: 'JobProgress',
				MigrationDirection: 'Import',
				FilesCreated: '1',
				BytesProcessed: '5',
			},
			importEnd,
		]);
		expect(digest.summary).toEqual({
			files: 2,
			folders: 2,
			bytes: 11,
			objectsProcessed: 8,
			errors: 0,
			warnings: 2,
		});
	});

	it('lists errors and warnings with their messages', () => {
		const digest = digestLogs([
			{
				Event: 'JobWarning',
				Message: ' Tag problem ',
				ObjectType: 'ListItem',
				Url: 'Lib/a',
				ErrorCode: '1',
			},
			{ Event: 'JobError', Message: 'Locked', ObjectType: 'File', Url: 'Lib/b' },
		]);
		expect(digest.warnings).toEqual([
			{
				event: 'JobWarning',
				message: 'Tag problem',
				objectType: 'ListItem',
				url: 'Lib/a',
				errorCode: '1',
			},
		]);
		expect(digest.errors[0]).toMatchObject({ event: 'JobError', message: 'Locked', url: 'Lib/b' });
		expect(digest.summary.errors).toBe(1);
		expect(digest.summary.warnings).toBe(1);
	});

	it('counts errors that only the end entry reports', () => {
		const digest = digestLogs([
			{ Event: 'JobEnd', MigrationDirection: 'Import', TotalErrors: '3' },
		]);
		expect(digest.summary.errors).toBe(3);
		expect(digest.errors).toEqual([]);
	});

	it('finds the final target of the job', () => {
		expect(
			digestLogs([{ Event: 'JobFinishedObjectInfo', TargetObjectSiteRelativeUrl: 'Lib/dst/a' }])
				.targetUrl,
		).toBe('Lib/dst/a');
		expect(digestLogs([]).targetUrl).toBeUndefined();
	});
});

describe('job status', () => {
	it('is running while the job processes', () => {
		const status = buildJobStatus(info, 4, []);
		expect(status).toMatchObject({
			status: 'running',
			finished: false,
			completed: false,
			jobId: 'job-1',
			jobQueueUri: info.jobQueueUri,
			encryptionKey: 'key',
			jobState: 'Processing',
		});
	});

	it('is completed when the job finished without errors', () => {
		const status = buildJobStatus(info, 0, [importEnd]);
		expect(status).toMatchObject({
			status: 'completed',
			finished: true,
			completed: true,
			jobState: 'Finished',
		});
	});

	it('fails when the finished job has errors', () => {
		expect(buildJobStatus(info, 0, [{ Event: 'JobError', Message: 'x' }]).status).toBe('failed');
	});

	it('does not fail for a non-fatal error while the job runs', () => {
		expect(buildJobStatus(info, 4, [{ Event: 'JobError', Message: 'x' }]).status).toBe('running');
	});

	it('fails at once after a fatal error', () => {
		expect(buildJobStatus(info, 4, [{ Event: 'JobFatalError', Message: 'x' }])).toMatchObject({
			status: 'failed',
			finished: true,
		});
	});

	it('has no job state before the first check', () => {
		expect(buildJobStatus(info, undefined, []).jobState).toBeNull();
	});
});

describe('waiting for a job', () => {
	function scripted(progress: CopyJobProgress[]) {
		let calls = 0;
		let time = 0;
		return {
			read: async () => progress[Math.min(calls++, progress.length - 1)],
			calls: () => calls,
			options: (maxWaitMs: number) => ({
				maxWaitMs,
				intervalMs: 1000,
				now: () => time,
				pause: async (ms: number) => {
					time += ms;
				},
			}),
		};
	}

	it('collects the logs of all polls until the job is finished', async () => {
		const run = scripted([
			{ JobState: 4, Logs: [entry({ Event: 'JobQueued' })] },
			{ JobState: 4, Logs: [] },
			{ JobState: 0, Logs: [entry(importEnd)] },
		]);
		const result = await waitForCopyJob(run.read, info, run.options(60_000));
		expect(result.jobState).toBe(0);
		expect(result.logs.map((log) => log.Event)).toEqual(['JobQueued', 'JobEnd']);
		expect(run.calls()).toBe(3);
	});

	it('does not stop at the end of the import phase of a move', async () => {
		const run = scripted([
			{ JobState: 4, Logs: [entry(importEnd)] },
			{ JobState: 4, Logs: [entry({ Event: 'JobStart', MigrationDirection: 'MoveCleanup' })] },
			{ JobState: 0, Logs: [entry({ Event: 'JobEnd', MigrationDirection: 'MoveCleanup' })] },
		]);
		const result = await waitForCopyJob(run.read, info, run.options(60_000));
		expect(result.jobState).toBe(0);
		expect(run.calls()).toBe(3);
	});

	it('returns a running job when the maximum wait passes', async () => {
		const run = scripted([{ JobState: 4, Logs: [] }]);
		const result = await waitForCopyJob(run.read, info, run.options(3000));
		expect(result.jobState).toBe(4);
		expect(run.calls()).toBe(4);
		expect(buildJobStatus(info, result.jobState, result.logs).status).toBe('running');
	});

	it('polls at least once when the maximum wait is zero', async () => {
		const run = scripted([{ JobState: 4, Logs: [] }]);
		await waitForCopyJob(run.read, info, run.options(0));
		expect(run.calls()).toBe(1);
	});

	it('stops after a fatal error', async () => {
		const run = scripted([
			{ JobState: 4, Logs: [entry({ Event: 'JobFatalError', Message: 'x' })] },
		]);
		const result = await waitForCopyJob(run.read, info, run.options(60_000));
		expect(run.calls()).toBe(1);
		expect(buildJobStatus(info, result.jobState, result.logs).status).toBe('failed');
	});

	it('waits for the first log when a new job reports "finished" too early', async () => {
		const run = scripted([
			{ JobState: 0, Logs: [] },
			{ JobState: 4, Logs: [entry({ Event: 'JobQueued' })] },
			{ JobState: 0, Logs: [entry(importEnd)] },
		]);
		const result = await waitForCopyJob(run.read, info, run.options(60_000));
		expect(run.calls()).toBe(3);
		expect(result.logs).toHaveLength(2);
	});
});
