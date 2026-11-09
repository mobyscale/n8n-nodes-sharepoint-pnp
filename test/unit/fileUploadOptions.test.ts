import type { INodeProperties } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { fileResource } from '../../nodes/SharePointPnp/resources/file';
import { folderResource } from '../../nodes/SharePointPnp/resources/folder';
import {
	finishUpload,
	hasPostUploadSteps,
	parseUploadFields,
} from '../../nodes/SharePointPnp/resources/file/upload';
import { withListItemFields } from '../../nodes/SharePointPnp/shared/filesFolders';
import { buildSelection } from '../../nodes/SharePointPnp/shared/selection';

const path = '/sites/Example/Shared Documents/a.txt';

interface FakeState {
	CheckOutType: number;
	Level: number;
	fieldErrors?: boolean;
}

function fakeSp(state: FakeState) {
	const calls: string[] = [];
	const validateUpdateListItem = vi.fn(
		async (values: Array<{ FieldName: string; FieldValue: string }>, newDocument: boolean) => {
			calls.push(`validate:${String(newDocument)}`);
			return values.map((entry) => ({
				FieldName: entry.FieldName,
				FieldValue: entry.FieldValue,
				HasException: state.fieldErrors === true,
				ErrorMessage: state.fieldErrors ? 'bad value' : null,
			}));
		},
	);
	const file = Object.assign(async () => ({ Name: 'a.txt', Level: state.Level }), {
		select: () => async () => ({ CheckOutType: state.CheckOutType, Level: state.Level }),
		checkin: vi.fn(async (comment: string, type: number) => {
			calls.push(`checkin:${comment}:${type}`);
			state.CheckOutType = 2;
			if (type === 1) state.Level = 1;
			else state.Level = 2;
		}),
		publish: vi.fn(async (comment: string) => {
			calls.push(`publish:${comment}`);
			state.Level = 1;
		}),
		listItemAllFields: {
			select: () => ({ expand: () => async () => ({ Id: 7, ParentList: { Id: 'list-id' } }) }),
		},
	});
	const list = {
		fields: {
			select: () => async () => [
				{ InternalName: 'Title', TypeAsString: 'Text' },
				{ InternalName: 'Done', TypeAsString: 'Boolean' },
			],
		},
		items: { getById: () => ({ validateUpdateListItem }) },
	};
	const sp = {
		web: {
			getFileByServerRelativePath: vi.fn(() => file),
			lists: { getById: vi.fn(() => list) },
		},
	};
	return { sp: sp as never, calls, validateUpdateListItem };
}

describe('upload fields', () => {
	it('merges Fields (JSON) and the Fields entries, with the entries last', () => {
		expect(
			parseUploadFields({
				fieldsJson: '{"Title": "From JSON", "Done": true}',
				fields: {
					values: [
						{ name: 'Title', value: 'From entry' },
						{ name: ' ', value: 'x' },
					],
				},
			}),
		).toEqual({ Title: 'From entry', Done: true });
	});

	it('accepts an object and an empty text', () => {
		expect(parseUploadFields({ fieldsJson: { Title: 'A' } })).toEqual({ Title: 'A' });
		expect(parseUploadFields({ fieldsJson: '' })).toEqual({});
		expect(parseUploadFields({})).toEqual({});
	});

	it('rejects JSON that is not an object', () => {
		expect(() => parseUploadFields({ fieldsJson: '{bad' })).toThrow('not valid JSON');
		expect(() => parseUploadFields({ fieldsJson: '[1]' })).toThrow('must be an object');
	});

	it('only adds requests when an option needs them', () => {
		expect(hasPostUploadSteps({ overwrite: true }, {})).toBe(false);
		expect(hasPostUploadSteps({}, { Title: 'A' })).toBe(true);
		expect(hasPostUploadSteps({ checkIn: true }, {})).toBe(true);
		expect(hasPostUploadSteps({ publish: true }, {})).toBe(true);
	});
});

describe('finishUpload', () => {
	it('sets the fields in one request without a new version and returns the stored values', async () => {
		const { sp, calls, validateUpdateListItem } = fakeSp({ CheckOutType: 2, Level: 1 });
		const result = await finishUpload(sp, path, {}, { Title: 'Report', Done: true });
		expect(validateUpdateListItem).toHaveBeenCalledTimes(1);
		expect(validateUpdateListItem.mock.calls[0][0]).toEqual([
			{ FieldName: 'Title', FieldValue: 'Report' },
			{ FieldName: 'Done', FieldValue: '1' },
		]);
		expect(calls).toEqual(['validate:true']);
		expect(result).toEqual({ Name: 'a.txt', Level: 1, fields: { Title: 'Report', Done: '1' } });
	});

	it('fails with the field errors', async () => {
		const { sp } = fakeSp({ CheckOutType: 2, Level: 1, fieldErrors: true });
		await expect(finishUpload(sp, path, {}, { Title: 'x' })).rejects.toThrow(
			'The file was uploaded, but SharePoint rejected the field values. Title: bad value',
		);
	});

	it('checks in a checked-out file and publishes a draft', async () => {
		const { sp, calls } = fakeSp({ CheckOutType: 0, Level: 255 });
		const result = await finishUpload(
			sp,
			path,
			{
				checkIn: true,
				checkInComment: 'In',
				checkInType: 'minor',
				publish: true,
				publishComment: 'Out',
			},
			{},
		);
		expect(calls).toEqual(['checkin:In:0', 'publish:Out']);
		expect(result).toMatchObject({ checkedIn: true, published: true });
	});

	it('skips the check in and the publish when they do not apply', async () => {
		const { sp, calls } = fakeSp({ CheckOutType: 2, Level: 1 });
		const result = await finishUpload(sp, path, { checkIn: true, publish: true }, {});
		expect(calls).toEqual([]);
		expect(result).toMatchObject({ checkedIn: false, published: false });
	});

	it('rejects a publish comment that is too long', async () => {
		const { sp } = fakeSp({ CheckOutType: 2, Level: 2 });
		await expect(
			finishUpload(sp, path, { publish: true, publishComment: 'x'.repeat(1024) }, {}),
		).rejects.toThrow('at most 1023');
	});
});

describe('upload description', () => {
	const options = fileResource.description.find(
		(entry) =>
			entry.name === 'options' && entry.displayOptions?.show?.operation?.includes('upload'),
	);
	const names = (options?.options as INodeProperties[]).map((option) => option.name);

	it('has the new upload options', () => {
		expect(names).toEqual(
			expect.arrayContaining([
				'createMissingFolders',
				'fields',
				'fieldsJson',
				'checkIn',
				'checkInComment',
				'checkInType',
				'publish',
				'publishComment',
				'overwrite',
			]),
		);
	});
});

describe('include list item fields', () => {
	const optionNames = (resource: typeof fileResource) =>
		(
			resource.description.find(
				(entry) =>
					entry.name === 'options' && entry.displayOptions?.show?.operation?.includes('get'),
			)?.options as INodeProperties[]
		).map((option) => option.name);

	it('is an option of File Get and Folder Get', () => {
		expect(optionNames(fileResource)).toContain('includeListItem');
		expect(optionNames(folderResource)).toContain('includeListItem');
	});

	it('expands ListItemAllFields and keeps the default selection', () => {
		const defaults = ['Name', 'ServerRelativeUrl'];
		expect(buildSelection(defaults, withListItemFields({ includeListItem: true }))).toEqual({
			select: defaults,
			expand: ['ListItemAllFields'],
		});
		expect(buildSelection(defaults, withListItemFields({}))).toEqual({
			select: defaults,
			expand: [],
		});
	});
});
