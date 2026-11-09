import type { INodeListSearchItems } from 'n8n-workflow';

/** Dropdown and multi-select pickers cannot page in n8n, so they load at most this many entries. */
export const MAX_OPTIONS = 500;

/** Entries per page of a "From List" picker that pages. */
export const PICKER_PAGE_SIZE = 50;

export function byName(a: { name: string }, b: { name: string }): number {
	return a.name.localeCompare(b.name);
}

/** A copy of the entries sorted A-Z by name. */
export function sortByName<T extends { name: string }>(entries: T[]): T[] {
	return [...entries].sort(byName);
}

/** Sorts A-Z by name and keeps at most MAX_OPTIONS entries. */
export function sortAndCap<T extends { name: string }>(entries: T[], max = MAX_OPTIONS): T[] {
	return sortByName(entries).slice(0, max);
}

/** The offset that a pagination token holds. A missing or invalid token means the first page. */
export function parseOffset(token?: string): number {
	const offset = Number(token ?? 0);
	return Number.isInteger(offset) && offset > 0 ? offset : 0;
}

/**
 * Returns one page of a list that is fully loaded. The pagination token is the offset of the next page.
 * The entries must already have their final order.
 */
export function pageByOffset<T extends INodeListSearchItems>(
	entries: T[],
	token?: string,
	size = PICKER_PAGE_SIZE,
): { results: T[]; paginationToken?: string } {
	const start = parseOffset(token);
	const next = start + size;
	return {
		results: entries.slice(start, next),
		paginationToken: next < entries.length ? String(next) : undefined,
	};
}
