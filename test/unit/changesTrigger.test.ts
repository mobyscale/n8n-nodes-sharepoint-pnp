import type { IDataObject, INode, IPollFunctions } from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import * as changes from '../../nodes/SharePointPnp/shared/changes';
import {
	MANUAL_LIMIT,
	SharePointPnpTrigger,
} from '../../nodes/SharePointPnpTrigger/SharePointPnpTrigger.node';

const original = await vi.importActual<typeof changes>('../../nodes/SharePointPnp/shared/changes');

vi.mock('../../nodes/SharePointPnp/shared/changes', async (importOriginal) => {
	const actual = await importOriginal<typeof changes>();
	return {
		...actual,
		getChangeTarget: vi.fn(),
		runChangeQuery: vi.fn(),
		runChangeQueryWithExpiry: vi.fn(),
	};
});

const LIST_ID = '11111111-2222-3333-4444-555555555555';
const node: INode = {
	id: 'trigger',
	name: 'SharePoint PnP Trigger',
	type: '@mobyscale/n8n-nodes-sharepoint-pnp.sharePointPnpTrigger',
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
const runWithExpiry = vi.mocked(changes.runChangeQueryWithExpiry);

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
		expect(runWithExpiry).not.toHaveBeenCalled();
	});

	it('reads changes after the stored token and saves the next token', async () => {
		const staticData: IDataObject = {};
		await poll(createContext({}, staticData));
		runWithExpiry.mockResolvedValueOnce({
			tokenExpired: false,
			changes: [{ changeType: 'Add' }],
			nextChangeToken: 'next-token',
			hasMore: false,
		});

		const result = await poll(createContext({}, staticData));

		expect(result).toEqual([[{ json: { changeType: 'Add' } }]]);
		expect(staticData.changeToken).toBe('next-token');
		const [, , query, options] = runWithExpiry.mock.calls[0];
		expect(query.ChangeTokenStart).toEqual({ StringValue: 'current-token' });
		expect(query.File).toBe(true);
		expect(query.Add).toBe(true);
		expect(options).toMatchObject({ events: ['fileCreated'], returnAll: true });
	});

	it('returns null and keeps the token when there are no changes', async () => {
		const staticData: IDataObject = {};
		await poll(createContext({}, staticData));
		runWithExpiry.mockResolvedValueOnce({
			tokenExpired: false,
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
		runWithExpiry.mockResolvedValueOnce({
			tokenExpired: false,
			changes: [],
			nextChangeToken: 'skipped',
			hasMore: false,
		});

		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('skipped');
	});

	it('starts from a date on the first poll when the option is set', async () => {
		const staticData: IDataObject = {};
		runWithExpiry.mockResolvedValueOnce({
			tokenExpired: false,
			changes: [{ changeType: 'Add' }],
			nextChangeToken: 'next-token',
			hasMore: false,
		});

		const result = await poll(
			createContext({ options: { startFromDate: '2024-01-01T00:00:00Z' } }, staticData),
		);

		expect(result).toHaveLength(1);
		const [, , query] = runWithExpiry.mock.calls[0];
		expect(query.ChangeTokenStart).toEqual({
			StringValue: `1;3;${LIST_ID};638396640000000000;-1`,
		});
		expect(staticData.changeToken).toBe('next-token');
	});

	it('passes the filter options to the change query', async () => {
		const staticData: IDataObject = { changeToken: 'saved' };
		await poll(createContext({}, staticData));
		const fieldFilters = { filter: [{ field: 'Status', operator: 'equals', value: 'Open' }] };
		runWithExpiry.mockResolvedValueOnce({
			tokenExpired: false,
			changes: [],
			nextChangeToken: undefined,
			hasMore: false,
		});

		await poll(
			createContext(
				{
					options: { changeMaxDepth: 2, excludeEditors: 'Robot', changeFieldFilters: fieldFilters },
				},
				staticData,
			),
		);

		const [, , , options] = runWithExpiry.mock.calls[0];
		expect(options).toMatchObject({ maxDepth: 2, excludeEditors: 'Robot', fieldFilters });
	});

	it('has the filter options, but not Item Kind, because the events set the kinds', () => {
		const names = (original.triggerOptionsProperty.options as Array<{ name: string }>).map(
			(option) => option.name,
		);
		expect(names).toEqual(
			expect.arrayContaining(['changeMaxDepth', 'excludeEditors', 'changeFieldFilters']),
		);
		expect(names).not.toContain('changeItemKinds');
		const changeNames = (original.changeOptionsProperty.options as Array<{ name: string }>).map(
			(option) => option.name,
		);
		expect(changeNames).toEqual(
			expect.arrayContaining([
				'changeItemKinds',
				'changeMaxDepth',
				'excludeEditors',
				'changeFieldFilters',
			]),
		);
	});

	it('starts again when the scope changes', async () => {
		const staticData: IDataObject = { changeToken: 'old-token', scopeKey: 'another scope' };
		expect(await poll(createContext({}, staticData))).toBeNull();
		expect(staticData.changeToken).toBe('current-token');
		expect(runWithExpiry).not.toHaveBeenCalled();
	});

	describe('expired token', () => {
		const oldToken = () =>
			changes.createChangeTokenFromDate('list', LIST_ID, new Date(Date.now() - 75 * 864e5));
		const itemAdd = {
			'odata.type': 'SP.ChangeItem',
			ChangeType: 1,
			ChangeToken: { StringValue: 't1' },
		};

		async function expiredState(): Promise<IDataObject> {
			const staticData: IDataObject = {};
			await poll(createContext({}, staticData));
			staticData.changeToken = oldToken();
			return staticData;
		}

		beforeEach(() => {
			runWithExpiry.mockImplementation(original.runChangeQueryWithExpiry);
		});

		it('emits the oldest available changes with tokenExpired by default', async () => {
			const staticData = await expiredState();
			target.getChanges.mockResolvedValueOnce([itemAdd]).mockResolvedValueOnce([]);

			const result = await poll(createContext({ events: ['itemCreated'] }, staticData));

			expect(result?.[0].map((item) => item.json.tokenExpired)).toEqual([true]);
			expect(staticData.changeToken).toBe('t1');
			const [query] = target.getChanges.mock.calls[0] as [changes.ChangeQuery];
			expect(query.IgnoreStartTokenNotFoundError).toBe(true);
		});

		it('saves the current token and emits nothing with "Start From Now"', async () => {
			const staticData = await expiredState();

			const result = await poll(
				createContext({ events: ['itemCreated'], options: { onExpiredToken: 'now' } }, staticData),
			);

			expect(result).toBeNull();
			expect(staticData.changeToken).toBe('current-token');
			expect(target.getChanges).not.toHaveBeenCalled();
		});

		it('fails and keeps the token with "Fail"', async () => {
			const staticData = await expiredState();
			const token = staticData.changeToken;

			await expect(
				poll(createContext({ options: { onExpiredToken: 'fail' } }, staticData)),
			).rejects.toThrow('The change token has expired');
			expect(staticData.changeToken).toBe(token);
		});
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
