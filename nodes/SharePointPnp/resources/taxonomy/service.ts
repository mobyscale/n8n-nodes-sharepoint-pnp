import { randomUUID } from 'node:crypto';

import type { SharePointContext } from '../../transport';
import {
	CsomRequest,
	csomBoolean,
	csomGuid,
	csomNumber,
	csomString,
	normalizeGuid,
	parseCsomDate,
	parseCsomGuid,
	sendCsom,
	type CsomRecord,
} from './csom';

export interface TermLabel {
	value: string;
	language: number;
	isDefault: boolean;
}

export interface TermGroup {
	id: string;
	name: string;
	description: string;
	created: string | null;
	modified: string | null;
	isSiteCollectionGroup: boolean;
	isSystemGroup: boolean;
}

export interface TermSet {
	id: string;
	name: string;
	description: string;
	created: string | null;
	modified: string | null;
	isOpenForTermCreation: boolean;
	isAvailableForTagging: boolean;
	owner: string;
	contact: string;
	names: Record<string, string>;
	groupId: string | null;
	groupName: string | null;
}

export interface Term {
	id: string;
	name: string;
	labels: TermLabel[];
	description: string;
	path: string;
	created: string | null;
	modified: string | null;
	childCount: number;
	isRoot: boolean;
	isDeprecated: boolean;
	isAvailableForTagging: boolean;
	termSetId: string | null;
	groupId: string | null;
}

export interface StoreInfo {
	id: string;
	name: string;
	defaultLanguage: number;
	languages: number[];
}

export interface TermLocation {
	siteUrl: string;
	storeId?: string;
}

export type TermMatchType = 'exact' | 'startsWith' | 'contains';

const text = (value: unknown): string => (typeof value === 'string' ? value : '');

export function mapLabels(raw: CsomRecord): TermLabel[] {
	const labels = raw.Labels as { _Child_Items_?: CsomRecord[] } | undefined;
	return (labels?._Child_Items_ ?? []).map((label) => ({
		value: text(label.Value),
		language: Number(label.Language ?? 0),
		isDefault: label.IsDefaultForLanguage === true,
	}));
}

export function mapGroup(raw: CsomRecord): TermGroup {
	return {
		id: parseCsomGuid(raw.Id),
		name: text(raw.Name),
		description: text(raw.Description),
		created: parseCsomDate(raw.CreatedDate),
		modified: parseCsomDate(raw.LastModifiedDate),
		isSiteCollectionGroup: raw.IsSiteCollectionGroup === true,
		isSystemGroup: raw.IsSystemGroup === true,
	};
}

export function mapTermSet(raw: CsomRecord, group?: { id: string; name: string }): TermSet {
	return {
		id: parseCsomGuid(raw.Id),
		name: text(raw.Name),
		description: text(raw.Description),
		created: parseCsomDate(raw.CreatedDate),
		modified: parseCsomDate(raw.LastModifiedDate),
		isOpenForTermCreation: raw.IsOpenForTermCreation === true,
		isAvailableForTagging: raw.IsAvailableForTagging !== false,
		owner: text(raw.Owner),
		contact: text(raw.Contact),
		names: (raw.Names as Record<string, string> | undefined) ?? {},
		groupId: group?.id ?? null,
		groupName: group?.name ?? null,
	};
}

export function mapTerm(raw: CsomRecord, termSetId?: string | null, groupId?: string | null): Term {
	return {
		id: parseCsomGuid(raw.Id),
		name: text(raw.Name),
		labels: mapLabels(raw),
		description: text(raw.Description),
		path: text(raw.PathOfTerm),
		created: parseCsomDate(raw.CreatedDate),
		modified: parseCsomDate(raw.LastModifiedDate),
		childCount: Number(raw.TermsCount ?? 0),
		isRoot: raw.IsRoot === true,
		isDeprecated: raw.IsDeprecated === true,
		isAvailableForTagging: raw.IsAvailableForTagging !== false,
		termSetId: termSetId ?? null,
		groupId: groupId ?? null,
	};
}

export async function getStoreInfo(
	ctx: SharePointContext,
	location: TermLocation,
): Promise<StoreInfo> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const key = request.query(store);
	const raw = (await sendCsom(ctx, location.siteUrl, request)).record(key);
	return {
		id: parseCsomGuid(raw.Id),
		name: text(raw.Name),
		defaultLanguage: Number(raw.DefaultLanguage ?? 1033),
		languages: Array.isArray(raw.Languages) ? (raw.Languages as number[]) : [],
	};
}

export async function getTermGroups(
	ctx: SharePointContext,
	location: TermLocation,
): Promise<TermGroup[]> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const key = request.query(request.property(store, 'Groups'), { collection: true });
	const result = await sendCsom(ctx, location.siteUrl, request);
	return result.items(key).map(mapGroup);
}

export async function getTermSets(
	ctx: SharePointContext,
	location: TermLocation,
	groupId?: string,
): Promise<TermSet[]> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const groups = groupId
		? request.method(store, 'GetGroup', [csomGuid(groupId)])
		: request.property(store, 'Groups');
	if (groupId) {
		const groupKey = request.query(groups);
		const setsKey = request.query(request.property(groups, 'TermSets'), { collection: true });
		const result = await sendCsom(ctx, location.siteUrl, request);
		const group = mapGroup(result.record(groupKey));
		return result.items(setsKey).map((raw) => mapTermSet(raw, group));
	}
	const key = request.query(groups, { collection: true, nested: ['TermSets'] });
	const result = await sendCsom(ctx, location.siteUrl, request);
	return result.items(key).flatMap((rawGroup) => {
		const group = mapGroup(rawGroup);
		const sets = rawGroup.TermSets as { _Child_Items_?: CsomRecord[] } | undefined;
		return (sets?._Child_Items_ ?? []).map((raw) => mapTermSet(raw, group));
	});
}

export async function getTermSet(
	ctx: SharePointContext,
	location: TermLocation,
	termSetId: string,
): Promise<TermSet> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const set = request.method(store, 'GetTermSet', [csomGuid(termSetId)]);
	const setKey = request.query(set);
	const groupKey = request.query(request.property(set, 'Group'));
	const result = await sendCsom(ctx, location.siteUrl, request);
	const group = mapGroup(result.record(groupKey));
	return mapTermSet(result.record(setKey), group);
}

/**
 * Loads the terms of a term set. Without a parent, it returns the root terms, or every term of the
 * set when recursive. With a parent term, it returns the children, or all descendants when recursive.
 */
export async function getTerms(
	ctx: SharePointContext,
	location: TermLocation,
	termSetId: string,
	options: { parentTermId?: string; recursive?: boolean } = {},
): Promise<Term[]> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const set = request.method(store, 'GetTermSet', [csomGuid(termSetId)]);
	const groupKey = request.query(request.property(set, 'Group'));
	const parent = options.parentTermId
		? request.method(store, 'GetTerm', [csomGuid(options.parentTermId)])
		: set;
	const collection = options.recursive
		? request.method(parent, 'GetAllTerms')
		: request.property(parent, 'Terms');
	const key = request.query(collection, { collection: true, nested: ['Labels'] });
	const result = await sendCsom(ctx, location.siteUrl, request);
	const groupId = parseCsomGuid(result.record(groupKey).Id);
	const setId = normalizeGuid(termSetId);
	return result.items(key).map((raw) => mapTerm(raw, setId, groupId));
}

export async function getTerm(
	ctx: SharePointContext,
	location: TermLocation,
	termId: string,
): Promise<Term> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const term = request.method(store, 'GetTerm', [csomGuid(termId)]);
	const termKey = request.query(term, { nested: ['Labels'] });
	const set = request.property(term, 'TermSet');
	const setKey = request.query(set);
	const groupKey = request.query(request.property(set, 'Group'));
	const result = await sendCsom(ctx, location.siteUrl, request);
	return mapTerm(
		result.record(termKey),
		parseCsomGuid(result.record(setKey).Id),
		parseCsomGuid(result.record(groupKey).Id),
	);
}

export function matchesLabel(
	term: Pick<Term, 'labels' | 'name'>,
	query: string,
	matchType: TermMatchType,
	language?: number,
): boolean {
	const needle = query.trim().toLowerCase();
	const labels = term.labels.length
		? term.labels
		: [{ value: term.name, language: 0, isDefault: true }];
	return labels
		.filter((label) => !language || label.language === language)
		.some((label) => {
			const value = label.value.toLowerCase();
			if (matchType === 'exact') return value === needle;
			if (matchType === 'startsWith') return value.startsWith(needle);
			return value.includes(needle);
		});
}

export interface NewLabel {
	value: string;
	language: number;
}

export interface CreateTermInput {
	name: string;
	language?: number;
	description?: string;
	parentTermId?: string;
	labels?: NewLabel[];
}

/** Creates a term in a term set, or under a parent term, and commits the change. */
export async function createTerm(
	ctx: SharePointContext,
	location: TermLocation,
	termSetId: string,
	input: CreateTermInput,
): Promise<Term> {
	const language = input.language || (await getStoreInfo(ctx, location)).defaultLanguage;
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const set = request.method(store, 'GetTermSet', [csomGuid(termSetId)]);
	const parent = input.parentTermId
		? request.method(store, 'GetTerm', [csomGuid(input.parentTermId)])
		: set;
	const term = request.method(parent, 'CreateTerm', [
		csomString(input.name),
		csomNumber(language),
		csomGuid(randomUUID()),
	]);
	request.run(term);
	if (input.description) {
		request.call(term, 'SetDescription', [csomString(input.description), csomNumber(language)]);
	}
	for (const label of input.labels ?? []) {
		request.run(
			request.method(term, 'CreateLabel', [
				csomString(label.value),
				csomNumber(label.language || language),
				csomBoolean(false),
			]),
		);
	}
	request.call(store, 'CommitAll');
	const key = request.query(term, { nested: ['Labels'] });
	const result = await sendCsom(ctx, location.siteUrl, request);
	return mapTerm(result.record(key), normalizeGuid(termSetId));
}

export interface UpdateTermInput {
	name?: string;
	language?: number;
	description?: string;
	labels?: NewLabel[];
}

export async function updateTerm(
	ctx: SharePointContext,
	location: TermLocation,
	termId: string,
	input: UpdateTermInput,
): Promise<Term> {
	const language = input.language || (await getStoreInfo(ctx, location)).defaultLanguage;
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const term = request.method(store, 'GetTerm', [csomGuid(termId)]);
	if (input.name !== undefined) request.setProperty(term, 'Name', csomString(input.name));
	if (input.description !== undefined) {
		request.call(term, 'SetDescription', [csomString(input.description), csomNumber(language)]);
	}
	for (const label of input.labels ?? []) {
		request.run(
			request.method(term, 'CreateLabel', [
				csomString(label.value),
				csomNumber(label.language || language),
				csomBoolean(false),
			]),
		);
	}
	request.call(store, 'CommitAll');
	await sendCsom(ctx, location.siteUrl, request);
	return await getTerm(ctx, location, termId);
}

export async function deleteTerm(
	ctx: SharePointContext,
	location: TermLocation,
	termId: string,
): Promise<void> {
	const request = new CsomRequest();
	const { store } = request.termStore(location.storeId);
	const term = request.method(store, 'GetTerm', [csomGuid(termId)]);
	request.call(term, 'DeleteObject');
	request.call(store, 'CommitAll');
	await sendCsom(ctx, location.siteUrl, request);
}
