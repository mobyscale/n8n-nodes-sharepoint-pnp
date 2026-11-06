import type { IExecuteFunctions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { attachmentResource } from '../../nodes/SharePointPnp/resources/attachment';
import {
	getAttachmentSupportError,
	parseItemId,
	toArrayBuffer,
} from '../../nodes/SharePointPnp/resources/attachment/helpers';

const siteUrl = 'https://contoso.sharepoint.com/sites/Example';

describe('attachment helpers', () => {
	it('parses item IDs', () => {
		expect(parseItemId('12')).toBe(12);
		expect(parseItemId(3)).toBe(3);
		expect(() => parseItemId('0')).toThrow('not a valid item ID');
		expect(() => parseItemId('abc')).toThrow('not a valid item ID');
	});

	it('rejects document libraries and lists without attachments', () => {
		expect(getAttachmentSupportError({ BaseType: 1 })).toContain('Document libraries');
		expect(getAttachmentSupportError({ BaseType: 0, EnableAttachments: false })).toContain(
			'turned off',
		);
		expect(getAttachmentSupportError({ BaseType: 0, EnableAttachments: true })).toBeUndefined();
	});

	it('copies a buffer slice to an array buffer', () => {
		const source = Buffer.from('abcdef').subarray(1, 4);
		expect(Buffer.from(toArrayBuffer(source)).toString()).toBe('bcd');
	});
});

function notFound() {
	return Object.assign(new Error('missing'), { isHttpRequestError: true, status: 404 });
}

function setup(
	params: Record<string, unknown>,
	listInfo: Record<string, unknown> = { BaseType: 0 },
) {
	const attachment = {
		select: vi
			.fn()
			.mockReturnValue(
				vi.fn().mockResolvedValue({ FileName: 'a.txt', ServerRelativeUrl: '/x/a.txt' }),
			),
		setContent: vi.fn().mockResolvedValue(undefined),
		recycle: vi.fn().mockResolvedValue(undefined),
		delete: vi.fn().mockResolvedValue(undefined),
		getBuffer: vi.fn().mockResolvedValue(new TextEncoder().encode('hi').buffer),
	};
	const attachments = {
		getByName: vi.fn().mockReturnValue(attachment),
		add: vi.fn().mockResolvedValue({}),
	};
	const list = {
		select: vi.fn().mockReturnValue(vi.fn().mockResolvedValue(listInfo)),
		items: { getById: vi.fn().mockReturnValue({ attachmentFiles: attachments }) },
	};
	const sp = { web: { lists: { getById: vi.fn().mockReturnValue(list) } } };
	const context = {
		getNode: () => ({ name: 'SharePoint PnP' }),
		getInputData: () => [
			{ json: {}, binary: { data: { fileName: 'in.txt', mimeType: 'text/plain' } } },
		],
		getNodeParameter: (name: string, _i: number, fallback?: unknown) =>
			name in params ? params[name] : fallback,
		helpers: {
			getBinaryDataBuffer: vi.fn().mockResolvedValue(Buffer.from('content')),
			prepareBinaryData: vi.fn().mockResolvedValue({ fileName: 'a.txt' }),
		},
	} as unknown as IExecuteFunctions;
	const run = (operation: string) =>
		attachmentResource.operations[operation].call(context, {
			sp: sp as never,
			siteUrl,
			itemIndex: 0,
		});
	return { attachment, attachments, list, run };
}

const base = { list: { mode: 'id', value: 'guid' }, itemId: 5 };

describe('attachment operations', () => {
	it('fails for a document library', async () => {
		const { run, attachments } = setup({ ...base }, { BaseType: 1 });
		await expect(run('getAll')).rejects.toThrow('Document libraries do not support attachments');
		expect(attachments.add).not.toHaveBeenCalled();
	});

	it('adds an attachment with the binary file name', async () => {
		const { run, attachments } = setup({ ...base, binaryPropertyName: 'data' });
		const result = await run('add');
		expect(attachments.add).toHaveBeenCalledWith('in.txt', expect.any(ArrayBuffer));
		expect(result).toMatchObject({ itemId: 5, overwritten: false, FileName: 'a.txt' });
	});

	it('overwrites an existing attachment', async () => {
		const { run, attachments, attachment } = setup({
			...base,
			binaryPropertyName: 'data',
			fileName: 'a.txt',
			addOptions: { overwrite: true },
		});
		const result = await run('add');
		expect(attachment.setContent).toHaveBeenCalled();
		expect(attachments.add).not.toHaveBeenCalled();
		expect(result).toMatchObject({ overwritten: true });
	});

	it('adds a new attachment when overwrite finds none', async () => {
		const { run, attachments, attachment } = setup({
			...base,
			binaryPropertyName: 'data',
			fileName: 'new.txt',
			addOptions: { overwrite: true },
		});
		attachment.select.mockReturnValueOnce(vi.fn().mockRejectedValue(notFound()));
		await run('add');
		expect(attachments.add).toHaveBeenCalledWith('new.txt', expect.any(ArrayBuffer));
	});

	it('recycles by default and deletes permanently on request', async () => {
		const first = setup({ ...base, attachmentName: 'a.txt' });
		await first.run('delete');
		expect(first.attachment.recycle).toHaveBeenCalled();
		expect(first.attachment.delete).not.toHaveBeenCalled();

		const second = setup({
			...base,
			attachmentName: 'a.txt',
			deleteOptions: { permanent: true },
		});
		await second.run('delete');
		expect(second.attachment.delete).toHaveBeenCalled();
		expect(second.attachment.recycle).not.toHaveBeenCalled();
	});

	it('downloads to a binary field', async () => {
		const { run } = setup({ ...base, attachmentName: 'a.txt', binaryPropertyName: 'file' });
		const result = (await run('download')) as Array<{ binary: Record<string, unknown> }>;
		expect(Object.keys(result[0].binary)).toEqual(['file']);
	});

	it('updates the content', async () => {
		const { run, attachment } = setup({
			...base,
			attachmentName: 'a.txt',
			binaryPropertyName: 'data',
		});
		const result = await run('updateContent');
		expect(attachment.setContent).toHaveBeenCalled();
		expect(result).toMatchObject({ updated: true });
	});
});
