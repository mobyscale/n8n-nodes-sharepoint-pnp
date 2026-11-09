import type { IList } from '@pnp/sp/lists/index.js';
import type { IView } from '@pnp/sp/views/index.js';

import { splitViewFields } from './helpers';

const COPIED_VIEW_SETTINGS = [
	'ViewQuery',
	'RowLimit',
	'Paged',
	'Scope',
	'Aggregations',
	'AggregationsStatus',
	'ViewData',
	'MobileView',
	'TabularView',
	'IncludeRootFolder',
	'OrderedView',
	'Threaded',
] as const;

const GUID_PATTERN = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;

/** PnPjs objects keep query state, so callers pass factories that return a new list each time. */
export type ListFactory = () => IList;

/** Finds a view by ID or by title. */
export function getViewByIdOrTitle(list: IList, view: string): IView {
	const value = view.trim();
	return GUID_PATTERN.test(value)
		? list.views.getById(value.replace(/^\{|\}$/g, ''))
		: list.views.getByTitle(value);
}

export async function applyViewFields(
	list: IList,
	viewId: string,
	fields: string[],
): Promise<void> {
	const viewFields = list.views.getById(viewId).fields;
	await viewFields.removeAll();
	for (const field of fields) await viewFields.add(field);
}

export async function getViewTitle(source: ListFactory, view: string): Promise<string> {
	const info = (await getViewByIdOrTitle(source(), view).select('Title')()) as { Title: string };
	return info.Title;
}

export interface CopiedView {
	id: string;
	title: string;
	/** The view fields that exist in the target list. */
	kept: string[];
	/** The view fields that the target list does not have. */
	skipped: string[];
}

/** Copies a view (query, settings and the fields that exist in the target) to another list. */
export async function copyListView(
	source: ListFactory,
	view: string,
	target: ListFactory,
	newTitle?: string,
): Promise<CopiedView> {
	const info = (await getViewByIdOrTitle(source(), view).select(
		'Title',
		'PersonalView',
		...COPIED_VIEW_SETTINGS,
	)()) as Record<string, unknown>;
	const sourceFields =
		((await getViewByIdOrTitle(source(), view).fields()) as { Items?: string[] }).Items ?? [];
	const targetFields = (await target().fields.select('InternalName')()) as Array<{
		InternalName: string;
	}>;
	const { kept, skipped } = splitViewFields(
		sourceFields,
		targetFields.map((field) => field.InternalName),
	);

	const settings: Record<string, unknown> = {};
	for (const key of COPIED_VIEW_SETTINGS) {
		if (info[key] !== null && info[key] !== undefined) settings[key] = info[key];
	}

	const title = newTitle?.trim() || String(info.Title);
	const created = (await target().views.add(title, info.PersonalView === true, settings)) as {
		Id: string;
	};
	await applyViewFields(target(), created.Id, kept);
	return { id: created.Id, title, kept, skipped };
}
