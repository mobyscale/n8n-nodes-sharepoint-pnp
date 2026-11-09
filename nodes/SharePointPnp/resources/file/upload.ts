import type { IDataObject } from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';
import { CheckinType } from '@pnp/sp/files/index.js';

import { buildFormValues, type FieldInfo } from '../listItem/helpers';

export interface UploadOptions {
	overwrite?: boolean;
	createMissingFolders?: boolean;
	fields?: { values?: Array<{ name?: string; value?: unknown }> };
	fieldsJson?: string | IDataObject;
	checkIn?: boolean;
	checkInComment?: string;
	checkInType?: 'major' | 'minor' | 'overwrite';
	publish?: boolean;
	publishComment?: string;
}

/** SPFile.CheckOutType: 2 means that the file is not checked out. */
const NOT_CHECKED_OUT = 2;
/** SPFile.Level: 2 means that the latest version is a draft (minor version). */
const DRAFT_LEVEL = 2;
const COMMENT_LIMIT = 1023;

const CHECKIN_TYPES: Record<string, CheckinType> = {
	minor: CheckinType.Minor,
	major: CheckinType.Major,
	overwrite: CheckinType.Overwrite,
};

function parseJson(text: string): unknown {
	try {
		return JSON.parse(text) as unknown;
	} catch {
		return undefined;
	}
}

/** Reads the fields to set: first Fields (JSON), then the Fields entries, which replace the same keys. */
export function parseUploadFields(options: UploadOptions): Record<string, unknown> {
	let fromJson: unknown = options.fieldsJson ?? {};
	if (typeof fromJson === 'string') {
		const text = fromJson.trim();
		fromJson = text ? parseJson(text) : {};
		if (fromJson === undefined) throw new Error('The "Fields (JSON)" value is not valid JSON');
	}
	if (fromJson === null || typeof fromJson !== 'object' || Array.isArray(fromJson)) {
		throw new Error('The "Fields (JSON)" value must be an object');
	}
	const values: Record<string, unknown> = { ...(fromJson as Record<string, unknown>) };
	for (const entry of options.fields?.values ?? []) {
		const name = String(entry.name ?? '').trim();
		if (name) values[name] = entry.value;
	}
	return values;
}

/** Whether the upload needs more requests after the file is added. */
export function hasPostUploadSteps(
	options: UploadOptions,
	fields: Record<string, unknown>,
): boolean {
	return Object.keys(fields).length > 0 || options.checkIn === true || options.publish === true;
}

/**
 * Sets list item fields of a file in one validateUpdateListItem request. bNewDocumentUpdate is true,
 * so SharePoint changes the version that the upload made and does not make a new version.
 * Returns the values that SharePoint stored, by field name.
 */
export async function setFileFields(
	sp: SPFI,
	path: string,
	values: Record<string, unknown>,
): Promise<IDataObject> {
	const file = sp.web.getFileByServerRelativePath(path);
	const item = (await file.listItemAllFields
		.select('Id', 'ParentList/Id')
		.expand('ParentList')()) as {
		Id?: number;
		ParentList?: { Id: string };
	};
	if (!item.Id || !item.ParentList) throw new Error(`No list item found for the file "${path}"`);
	const list = sp.web.lists.getById(item.ParentList.Id);
	const fieldInfos = (await list.fields.select('InternalName', 'TypeAsString')()) as FieldInfo[];
	const results = await list.items
		.getById(item.Id)
		.validateUpdateListItem(buildFormValues(values, fieldInfos), true);
	const failed = results.filter((entry) => entry.HasException);
	if (failed.length) {
		const message = failed
			.map((entry) => `${entry.FieldName}: ${entry.ErrorMessage ?? 'invalid value'}`)
			.join('; ');
		throw new Error(`The file was uploaded, but SharePoint rejected the field values. ${message}`);
	}
	return Object.fromEntries(results.map((entry) => [entry.FieldName, entry.FieldValue]));
}

async function readState(sp: SPFI, path: string) {
	return (await sp.web.getFileByServerRelativePath(path).select('CheckOutType', 'Level')()) as {
		CheckOutType?: number;
		Level?: number;
	};
}

/**
 * Runs the steps after the upload: set fields, check in, publish. Check in and publish are only
 * done when they apply, so a library without check out or minor versions does not fail.
 */
export async function finishUpload(
	sp: SPFI,
	path: string,
	options: UploadOptions,
	fields: Record<string, unknown>,
): Promise<IDataObject> {
	const publishComment = options.publishComment ?? '';
	if (options.publish && publishComment.length > COMMENT_LIMIT) {
		throw new Error(
			`The file was uploaded, but the publish comment has ${publishComment.length} characters. SharePoint accepts at most ${COMMENT_LIMIT}.`,
		);
	}
	const file = () => sp.web.getFileByServerRelativePath(path);
	const output: IDataObject = {};
	if (Object.keys(fields).length) output.fields = await setFileFields(sp, path, fields);
	if (options.checkIn) {
		const state = await readState(sp, path);
		output.checkedIn = state.CheckOutType !== undefined && state.CheckOutType !== NOT_CHECKED_OUT;
		if (output.checkedIn) {
			await file().checkin(
				options.checkInComment ?? '',
				CHECKIN_TYPES[options.checkInType ?? 'major'] ?? CheckinType.Major,
			);
		}
	}
	if (options.publish) {
		const state = await readState(sp, path);
		output.published = state.Level === DRAFT_LEVEL;
		if (output.published) await file().publish(publishComment);
	}
	return { ...((await file()()) as unknown as IDataObject), ...output };
}
