import type { IDataObject } from 'n8n-workflow';

export const MAX_ROWS_PER_PAGE = 500;
export const DEFAULT_ROWS_PER_PAGE = 100;

export const DEFAULT_SELECT_PROPERTIES = [
	'Title',
	'Path',
	'FileType',
	'LastModifiedTime',
	'Author',
	'UniqueId',
	'SiteId',
	'WebId',
	'ListId',
	'ListItemID',
	'IsContainer',
];

export interface SearchCell {
	Key: string;
	Value: unknown;
	ValueType?: string;
}

export interface SearchRow {
	Cells: SearchCell[];
}

export interface SearchSort {
	property: string;
	direction: 'ascending' | 'descending';
}

/** Adds a site scope to a KQL query. The scope matches every item below the site URL. */
export function buildQueryText(query: string, siteUrl?: string): string {
	const text = query.trim();
	if (!siteUrl) return text;
	const scope = `Path:"${siteUrl.replace(/\/+$/, '')}/*"`;
	return text ? `(${text}) ${scope}` : scope;
}

/** Splits a comma separated string or an array into trimmed, non-empty values. */
export function parseList(value: unknown): string[] {
	const parts = Array.isArray(value) ? value : String(value ?? '').split(',');
	return parts.map((part) => String(part).trim()).filter((part) => part !== '');
}

export function toSortList(sorts: SearchSort[]): Array<{ Property: string; Direction: number }> {
	return sorts
		.filter((sort) => sort.property?.trim())
		.map((sort) => ({
			Property: sort.property.trim(),
			Direction: sort.direction === 'descending' ? 1 : 0,
		}));
}

/** Turns the Cells of a search row into an object (property name to value). */
export function flattenCells(row: SearchRow): IDataObject {
	const result: IDataObject = {};
	for (const cell of row.Cells ?? []) {
		result[cell.Key] = cell.Value as IDataObject[string];
	}
	return result;
}

export function clampRowLimit(value: number | undefined): number {
	const number = Number(value);
	if (!Number.isFinite(number) || number < 1) return DEFAULT_ROWS_PER_PAGE;
	return Math.min(Math.floor(number), MAX_ROWS_PER_PAGE);
}

/** The size of the next page, or 0 when no more rows are needed. */
export function nextPageSize(
	pageSize: number,
	startRow: number,
	totalRows: number,
	wanted: number | undefined,
): number {
	if (startRow >= totalRows) return 0;
	const remaining = wanted === undefined ? Infinity : wanted - startRow;
	return Math.max(0, Math.min(pageSize, remaining));
}
