import type {
	IDataObject,
	IExecuteFunctions,
	INode,
	INodeExecutionData,
	INodeTypes,
} from 'n8n-workflow';
import { NodeHelpers, Workflow } from 'n8n-workflow';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { resources } from '../../nodes/SharePointPnp/resources';
import {
	fileFolderResource,
	summarizeExists,
} from '../../nodes/SharePointPnp/resources/fileFolder';
import {
	configuredOutputs,
	getRoutedOutputs,
	outputsExpression,
} from '../../nodes/SharePointPnp/shared/outputs';
import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';

const found = (path: string) => ({ path, serverRelativeUrl: path, exists: true, type: 'file' });
const missing = (path: string) => ({ path, serverRelativeUrl: path, exists: false, type: null });

/** Resolves the outputs of a node with the n8n expression engine, as the editor and the backend do. */
function evaluateOutputs(parameters: IDataObject) {
	const nodeType = new SharePointPnp();
	const nodeTypes: INodeTypes = {
		getByName: () => nodeType,
		getByNameAndVersion: () => nodeType,
		getKnownTypes: () => ({}),
	};
	const node: INode = {
		id: 'test',
		name: 'SharePoint PnP',
		type: 'n8n-nodes-sharepoint-pnp.sharePointPnp',
		typeVersion: 1,
		position: [0, 0],
		parameters: parameters as INode['parameters'],
	};
	const workflow = new Workflow({ nodes: [node], connections: {}, active: false, nodeTypes });
	return NodeHelpers.getNodeOutputs(workflow, node, nodeType.description);
}

describe('summarizeExists', () => {
	const results = [found('a'), missing('b')];

	it('uses AND for All', () => {
		expect(summarizeExists(results, 'all')).toEqual({
			exists: false,
			allExist: false,
			anyExists: true,
			combine: 'all',
			results,
		});
	});

	it('uses OR for Any', () => {
		expect(summarizeExists(results, 'any')).toMatchObject({ exists: true, combine: 'any' });
	});

	it('is true for All when every path exists', () => {
		expect(summarizeExists([found('a'), found('b')], 'all').exists).toBe(true);
	});

	it('is false for Any when no path exists', () => {
		expect(summarizeExists([missing('a')], 'any')).toMatchObject({
			exists: false,
			allExist: false,
			anyExists: false,
		});
	});
});

describe('Check Exists routing', () => {
	const routing = fileFolderResource.routing!.checkExists;

	it('has a True and a False output and puts errors on False', () => {
		expect(routing.outputs).toEqual(['True', 'False']);
		expect(routing.errorOutput).toBe(1);
	});

	it('routes by the combined result', () => {
		expect(routing.route(summarizeExists([found('a')], 'all'))).toBe(0);
		expect(routing.route(summarizeExists([missing('a')], 'all'))).toBe(1);
	});

	it('has a Combine option with All as the default', () => {
		const combine = fileFolderResource.description.find((entry) => entry.name === 'combine');
		expect(combine?.default).toBe('all');
		expect((combine?.options as Array<{ value: string }>).map((option) => option.value)).toEqual([
			'all',
			'any',
		]);
	});
});

describe('node outputs', () => {
	it('maps only the routed operations', () => {
		expect(getRoutedOutputs(resources)).toEqual({ 'fileFolder.checkExists': ['True', 'False'] });
	});

	it('returns two named outputs for Check Exists and one output otherwise', () => {
		const routed = getRoutedOutputs(resources);
		expect(configuredOutputs({ resource: 'fileFolder', operation: 'checkExists' }, routed)).toEqual(
			[
				{ type: 'main', displayName: 'True' },
				{ type: 'main', displayName: 'False' },
			],
		);
		expect(configuredOutputs({ resource: 'fileFolder', operation: 'getAll' }, routed)).toEqual([
			{ type: 'main' },
		]);
		expect(configuredOutputs({ resource: 'file', operation: 'checkExists' }, routed)).toEqual([
			{ type: 'main' },
		]);
		expect(configuredOutputs({}, routed)).toEqual([{ type: 'main' }]);
	});

	it('is an expression that evaluates to the same outputs', () => {
		const description = new SharePointPnp().description;
		expect(description.outputs).toBe(outputsExpression(resources));
		expect(evaluateOutputs({ resource: 'fileFolder', operation: 'checkExists' })).toEqual([
			{ type: 'main', displayName: 'True' },
			{ type: 'main', displayName: 'False' },
		]);
		expect(evaluateOutputs({ resource: 'list', operation: 'get' })).toEqual([{ type: 'main' }]);
	});

	it('has no "}}" inside the expression, so n8n finds the right end', () => {
		const expression = outputsExpression(resources);
		expect(expression.startsWith('={{')).toBe(true);
		expect(expression.indexOf('}}')).toBe(expression.length - 2);
	});
});

describe('execute with routed outputs', () => {
	const original = fileFolderResource.operations.checkExists;

	afterEach(() => {
		fileFolderResource.operations.checkExists = original;
	});

	function context(
		parameters: IDataObject,
		items: INodeExecutionData[],
		continueOnFail: boolean,
	): IExecuteFunctions {
		const all: IDataObject = {
			authentication: 'certificate',
			site: { __rl: true, mode: 'url', value: 'https://contoso.sharepoint.com/sites/Example' },
			...parameters,
		};
		return {
			getInputData: () => items,
			getNode: () => ({ name: 'SharePoint PnP', type: 'sharePointPnp', typeVersion: 1 }),
			continueOnFail: () => continueOnFail,
			getCredentials: async () => ({ subdomain: 'contoso' }),
			getNodeParameter: (
				name: string,
				_index: number,
				fallback?: unknown,
				options?: { extractValue?: boolean },
			) => {
				const value = name in all ? all[name] : fallback;
				if (options?.extractValue && value && typeof value === 'object') {
					return (value as IDataObject).value;
				}
				return value;
			},
			helpers: {},
		} as unknown as IExecuteFunctions;
	}

	const answers: Array<boolean | Error> = [true, false, new Error('Access denied')];

	function stubHandler() {
		fileFolderResource.operations.checkExists = vi.fn(async ({ itemIndex }) => {
			const answer = answers[itemIndex];
			if (answer instanceof Error) throw answer;
			return summarizeExists([answer ? found('a') : missing('a')], 'all');
		});
	}

	const items = answers.map(() => ({ json: {} }));

	it('puts each input item on True or False and errors on False', async () => {
		stubHandler();
		const result = await new SharePointPnp().execute.call(
			context({ resource: 'fileFolder', operation: 'checkExists' }, items, true),
		);
		expect(result).toHaveLength(2);
		expect(result[0].map((item) => item.pairedItem)).toEqual([{ item: 0 }]);
		expect(result[1].map((item) => item.pairedItem)).toEqual([{ item: 1 }, { item: 2 }]);
		expect(result[1][1].json.error).toBe('Access denied');
	});

	it('throws the error when Continue On Fail is off', async () => {
		stubHandler();
		await expect(
			new SharePointPnp().execute.call(
				context({ resource: 'fileFolder', operation: 'checkExists' }, items, false),
			),
		).rejects.toThrow('Access denied');
	});

	it('returns one output for other operations', async () => {
		const result = await new SharePointPnp().execute.call(
			context({ resource: 'fileFolder', operation: 'unknown' }, [{ json: {} }], true),
		);
		expect(result).toHaveLength(1);
		expect(result[0][0].json.error).toContain('not supported');
	});

	it('returns empty outputs for no input items', async () => {
		const result = await new SharePointPnp().execute.call(
			context({ resource: 'fileFolder', operation: 'checkExists' }, [], false),
		);
		expect(result).toEqual([[]]);
	});
});
