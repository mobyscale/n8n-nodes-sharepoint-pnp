import type { IDataObject, INode, IPollFunctions } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import * as changes from '../../nodes/SharePointPnp/shared/changes';
import {
	MANUAL_LIMIT,
	SharePointPnpTrigger,
} from '../../nodes/SharePointPnpTrigger/SharePointPnpTrigger.node';

vi.mock('../../nodes/SharePointPnp/shared/changes', async (importOriginal) => {
	const original = await importOriginal<typeof changes>();
	return { ...original, getChangeTarget: vi.fn(), runChangeQuery: vi.fn() };
});

const LIST_ID = '11111111-2222-3333-4444-555555555555';
const node: INode = {
	id: 'trigger',
	name: 'SharePoint PnP Trigger',
	type: 'n8n-nodes-sharepoint-pnp.sharePointPnpTrigger',
	typeVersion: 1,
	position: [0, 0],
	parameters: {},
};

const target = {
	scope: 'list' as const,
	getChanges: vi.fn(),
	getCurrentToken: vi.fn(async () => 'current-token'),
	getId: vi.fn(async () => LIST_ID),
};

const getChangeTarget = vi.mocked(changes.getChangeTarget);
const runChangeQuery = vi.mocked(changes.runChangeQuery);

function createContext(
	parameters: IDataObject,
	staticData: IDataObject,
	mode: 'trigger' | 'manual' = 'trigger',
): IPollFunctions {
	const all: IDataObject = {
		authentication: 'certificate',
		site: { __rl: true, mode: 'url', value: '/sites/Example' },
		scope: 'list',
		list: { __rl: true, mode: 'id', value: LIST_ID },
		events: ['fileCreated'],
		folderPathFilter: '',
		options: {},
		...parameters,
	};
	return {
		getNode: () => node,
		getMode: () => mode,
		getWorkflowStaticData: () => staticData,
		getCredentials: async () => ({ subdomain: 'contoso' }),
		getNodeParameter: (name: string, fallback?: unknown, options?: { extractValue?: boolean }) => {
			const value = name in all ? all[name] : fallback;
			if (options?.extractValue && value && typeof value === 'object' && '__rl' in value) {
				return (value as IDataObject).value;
			}
			return value;
		},
		helpers: { returnJsonArray: (data: IDataObject[]) => data.map((json) => ({ json })) },
	} as unknown as IPollFunctions;
}

async function poll(ctx: IPollFunctions) {
	return await new SharePointPnpTrigger().poll.call(ctx);
}

describe('SharePointPnpTrigger.poll', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		getChangeTarget.mockReturnValue(target);
	});

	it('stores the current token on the first poll and emits nothing', async () => {
		const staticData: IDataObject = {};
		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('current-token');
		expect(staticData.scopeKey).toContain(LIST_ID);
		expect(runChangeQuery).not.toHaveBeenCalled();
	});

	it('reads changes after the stored token and saves the next token', async () => {
		const staticData: IDataObject = {};
		await poll(createContext({}, staticData));
		runChangeQuery.mockResolvedValueOnce({
			changes: [{ changeType: 'Add' }],
			nextChangeToken: 'next-token',
			hasMore: false,
		});

		const result = await poll(createContext({}, staticData));

		expect(result).toEqual([[{ json: { changeType: 'Add' } }]]);
		expect(staticData.changeToken).toBe('next-token');
		const [, , query, options] = runChangeQuery.mock.calls[0];
		expect(query.ChangeTokenStart).toEqual({ StringValue: 'current-token' });
		expect(query.File).toBe(true);
		expect(query.Add).toBe(true);
		expect(options).toMatchObject({ events: ['fileCreated'], returnAll: true });
	});

	it('returns null and keeps the token when there are no changes', async () => {
		const staticData: IDataObject = {};
		await poll(createContext({}, staticData));
		runChangeQuery.mockResolvedValueOnce({
			changes: [],
			nextChangeToken: undefined,
			hasMore: false,
		});

		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('current-token');
	});

	it('moves the token past changes that the filters skipped', async () => {
		const staticData: IDataObject = {};
		await poll(createContext({}, staticData));
		runChangeQuery.mockResolvedValueOnce({
			changes: [],
			nextChangeToken: 'skipped',
			hasMore: false,
		});

		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('skipped');
	});

	it('starts from a date on the first poll when the option is set', async () => {
		const staticData: IDataObject = {};
		runChangeQuery.mockResolvedValueOnce({
			changes: [{ changeType: 'Add' }],
			nextChangeToken: 'next-token',
			hasMore: false,
		});

		const result = await poll(
			createContext({ options: { startFromDate: '2024-01-01T00:00:00Z' } }, staticData),
		);

		expect(result).toHaveLength(1);
		const [, , query] = runChangeQuery.mock.calls[0];
		expect(query.ChangeTokenStart).toEqual({
			StringValue: `1;3;${LIST_ID};638396640000000000;-1`,
		});
		expect(staticData.changeToken).toBe('next-token');
	});

	it('starts again when the scope changes', async () => {
		const staticData: IDataObject = { changeToken: 'old-token', scopeKey: 'another scope' };
		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('current-token');
		expect(runChangeQuery).not.toHaveBeenCalled();
	});

	it('returns the latest matching changes in manual mode and does not change the token', async () => {
		const staticData: IDataObject = { changeToken: 'kept' };
		const all = Array.from({ length: MANUAL_LIMIT + 5 }, (_, index) => ({ index }));
		runChangeQuery.mockResolvedValueOnce({ changes: all, nextChangeToken: 'x', hasMore: false });

		const result = await poll(createContext({}, staticData, 'manual'));

		expect(result?.[0]).toHaveLength(MANUAL_LIMIT);
		expect(result?.[0][0].json).toEqual({ index: 5 });
		expect(staticData).toEqual({ changeToken: 'kept' });
		const [, , query] = runChangeQuery.mock.calls[0];
		expect(query.ChangeTokenStart?.StringValue).toMatch(new RegExp(`^1;3;${LIST_ID};\\d+;-1$`));
	});

	it('returns null in manual mode when there are no changes', async () => {
		runChangeQuery.mockResolvedValueOnce({ changes: [], hasMore: false });
		expect(await poll(createContext({}, {}, 'manual'))).toBeNull();
	});

	it('rejects an empty event list', async () => {
		await expect(poll(createContext({ events: [] }, {}))).rejects.toBeInstanceOf(
			NodeOperationError,
		);
	});

	it('converts SharePoint errors to node errors', async () => {
		target.getCurrentToken.mockRejectedValueOnce(
			Object.assign(
				new Error(
					'Error making HttpClient request ::> {"odata.error":{"message":{"value":"Denied"}}}',
				),
				{
					isHttpRequestError: true,
					status: 403,
				},
			),
		);
		const error = await poll(createContext({}, {})).catch((caught: unknown) => caught);
		expect(error).toBeInstanceOf(NodeApiError);
		expect((error as NodeApiError).message).toBe('Denied');
	});
});
