import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { fileResource } from '../../nodes/SharePointPnp/resources/file';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

function setup(parameters: Record<string, unknown>, moderation: unknown = undefined) {
	const calls: Array<[string, string]> = [];
	const record = (name: string) =>
		vi.fn(async (comment: string) => {
			calls.push([name, comment]);
		});
	const file = Object.assign(
		async () => ({ Name: 'a.txt', Level: 1 }),
		{
			publish: record('publish'),
			unpublish: record('unpublish'),
			approve: record('approve'),
			deny: record('deny'),
			listItemAllFields: {
				select: () => async () => {
					if (moderation === undefined) throw new Error('no moderation');
					return moderation;
				},
			},
		},
	);
	const sp = { web: { getFileByServerRelativePath: vi.fn(() => file) } };
	const context = {
		getNode: () => ({ name: 'SharePoint PnP' }),
		getNodeParameter: (
			name: string,
			_index: number,
			fallback?: unknown,
			options?: { extractValue?: boolean },
		) => {
			const value = name in parameters ? parameters[name] : fallback;
			if (options?.extractValue && typeof value === 'object' && value && 'value' in value) {
				return (value as { value: unknown }).value;
			}
			return value;
		},
	} as unknown as IExecuteFunctions;
	return { calls, sp, context };
}

const filePath = { mode: 'path', value: 'Shared Documents/a.txt' };

describe('file publishing operations', () => {
	it('offers Publish, Unpublish, Approve and Deny with a comment', () => {
		const operation = fileResource.description[0];
		const values = (operation.options as Array<{ value: string }>).map((option) => option.value);
		expect(values).toEqual(expect.arrayContaining(['publish', 'unpublish', 'approve', 'deny']));
		const comment = fileResource.description.filter((entry) => entry.name === 'comment');
		const operations = comment.flatMap((entry) => entry.displayOptions?.show?.operation as string[]);
		expect(operations).toEqual(
			expect.arrayContaining(['checkIn', 'publish', 'unpublish', 'approve', 'deny']),
		);
	});

	it.each(['publish', 'unpublish', 'approve', 'deny'])('%s calls the file method with the comment', async (operation) => {
		const { calls, sp, context } = setup({ filePath, comment: 'Ready' });
		const result = await fileResource.operations[operation].call(context, {
			sp: sp as never,
			siteUrl,
			itemIndex: 0,
		});
		expect(calls).toEqual([[operation, 'Ready']]);
		expect(sp.web.getFileByServerRelativePath).toHaveBeenCalledWith(
			'/sites/Example/Shared Documents/a.txt',
		);
		expect(result).toMatchObject({ Name: 'a.txt' });
	});

	it('adds the approval status when the library has content approval', async () => {
		const { sp, context } = setup(
			{ filePath },
			{ OData__ModerationStatus: 0, OData__ModerationComments: 'ok' },
		);
		const result = await fileResource.operations.approve.call(context, {
			sp: sp as never,
			siteUrl,
			itemIndex: 0,
		});
		expect(result).toMatchObject({ moderationStatus: 0, moderationComments: 'ok' });
	});

	it('rejects a comment that is longer than SharePoint accepts', async () => {
		const { calls, sp, context } = setup({ filePath, comment: 'x'.repeat(1024) });
		const error = await fileResource.operations.publish
			.call(context, { sp: sp as never, siteUrl, itemIndex: 0 })
			.catch((reason: unknown) => reason);
		expect(error).toBeInstanceOf(NodeOperationError);
		expect((error as Error).message).toContain('at most 1023');
		expect(calls).toEqual([]);
	});
});
