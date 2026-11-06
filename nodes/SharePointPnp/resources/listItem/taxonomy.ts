import type { SharePointContext } from '../../transport';
import { isEmptyGuid, isGuid, normalizeGuid } from '../taxonomy/csom';
import { createTerm, getTermSet, getTerms, type Term, type TermSet } from '../taxonomy/service';

export const TAXONOMY_TYPES = new Set(['TaxonomyFieldType', 'TaxonomyFieldTypeMulti']);

export interface TaxonomyFieldInfo {
	InternalName: string;
	TypeAsString: string;
	TermSetId?: string;
	SspId?: string;
	AnchorId?: string;
}

/** One tag given by the user: a label, a term ID, or both. */
export interface TagRef {
	label?: string;
	id?: string;
}

export interface TagTerm {
	id: string;
	label: string;
}

export interface TagClient {
	getTerms(storeId: string | undefined, termSetId: string): Promise<Term[]>;
	getTermSet(storeId: string | undefined, termSetId: string): Promise<TermSet>;
	createTerm(
		storeId: string | undefined,
		termSetId: string,
		label: string,
		parentTermId?: string,
	): Promise<Term>;
}

export interface ResolveOptions {
	createMissing?: boolean;
	/** Values that the field has now. They come first, so "append" keeps them. */
	current?: Record<string, unknown>;
}

const INVALID_TERM_NAME = /[;"<>|\t]/;

function parseSingle(text: string): TagRef | undefined {
	const value = text.trim();
	if (!value) return undefined;
	const bar = value.lastIndexOf('|');
	if (bar > 0 && isGuid(value.slice(bar + 1))) {
		const label = value.slice(0, bar).trim();
		const id = normalizeGuid(value.slice(bar + 1));
		return label ? { label, id } : { id };
	}
	if (isGuid(value)) return { id: normalizeGuid(value) };
	return { label: value };
}

/**
 * Reads a tag value: a label, a term ID, "Label|ID", a ";" separated list of these, an array of
 * these, or term objects such as { Label, TermGuid } (the shape SharePoint returns).
 */
export function parseTagInput(value: unknown): TagRef[] {
	if (value === undefined || value === null) return [];
	if (Array.isArray(value)) return value.flatMap(parseTagInput);
	if (typeof value === 'object') {
		const entry = value as Record<string, unknown>;
		const rawId = entry.TermGuid ?? entry.termGuid ?? entry.id ?? entry.Id;
		const rawLabel = entry.Label ?? entry.label ?? entry.name ?? entry.Name;
		const id = typeof rawId === 'string' && isGuid(rawId) ? normalizeGuid(rawId) : undefined;
		const label = typeof rawLabel === 'string' && rawLabel.trim() ? rawLabel.trim() : undefined;
		return id || label ? [{ ...(label ? { label } : {}), ...(id ? { id } : {}) }] : [];
	}
	const text = String(value).trim();
	if (text.startsWith('[')) {
		try {
			const parsed: unknown = JSON.parse(text);
			if (Array.isArray(parsed)) return parseTagInput(parsed);
		} catch {
			// not JSON, read it as a list
		}
	}
	return text
		.split(';')
		.map(parseSingle)
		.filter((ref): ref is TagRef => ref !== undefined);
}

/** The format of validateUpdateListItem for taxonomy fields: "Label|ID;Label|ID". */
export function formatTagValue(terms: TagTerm[]): string {
	return terms.map((term) => `${term.label}|${term.id}`).join(';');
}

function sameText(left: string, right: string): boolean {
	return left.trim().toLowerCase() === right.trim().toLowerCase();
}

function usable(term: Term): boolean {
	return term.isAvailableForTagging && !term.isDeprecated;
}

export function findTermsByLabel(terms: Term[], label: string): Term[] {
	const byName = terms.filter(
		(term) =>
			sameText(term.name, label) ||
			term.labels.some((entry) => entry.isDefault && sameText(entry.value, label)),
	);
	if (byName.length) return byName;
	return terms.filter((term) => term.labels.some((entry) => sameText(entry.value, label)));
}

function termDescription(term: Term): string {
	return term.path || term.name;
}

function restrictToAnchor(terms: Term[], anchorId: string | undefined): Term[] {
	if (isEmptyGuid(anchorId)) return terms;
	const anchor = terms.find((term) => term.id === normalizeGuid(anchorId!));
	if (!anchor?.path) return terms;
	const prefix = `${anchor.path};`;
	return terms.filter((term) => term.path.startsWith(prefix));
}

/** Resolves the tags of one field to term IDs and default labels. Creates missing terms when allowed. */
export async function resolveTags(
	refs: TagRef[],
	field: TaxonomyFieldInfo,
	client: TagClient,
	options: ResolveOptions = {},
): Promise<TagTerm[]> {
	const multi = field.TypeAsString === 'TaxonomyFieldTypeMulti';
	const result = new Map<string, TagTerm>();
	if (!refs.length) return [];
	if (!multi && refs.length > 1) {
		throw new Error(
			`The field "${field.InternalName}" accepts one term, but ${refs.length} were given`,
		);
	}
	if (!field.TermSetId || isEmptyGuid(field.TermSetId)) {
		throw new Error(`The field "${field.InternalName}" is not bound to a term set`);
	}
	const termSetId = normalizeGuid(field.TermSetId);
	const storeId = field.SspId && !isEmptyGuid(field.SspId) ? normalizeGuid(field.SspId) : undefined;

	let terms: Term[] | undefined;
	const load = async (): Promise<Term[]> => {
		terms ??= await client.getTerms(storeId, termSetId);
		return terms;
	};

	for (const ref of refs) {
		if (ref.label && ref.id) {
			result.set(ref.id, { id: ref.id, label: ref.label });
			continue;
		}
		if (ref.id) {
			const term = (await load()).find((entry) => entry.id === ref.id);
			if (!term) {
				throw new Error(
					`The term "${ref.id}" is not in the term set of the field "${field.InternalName}"`,
				);
			}
			result.set(term.id, { id: term.id, label: term.name });
			continue;
		}

		const label = ref.label ?? '';
		const candidates = restrictToAnchor(await load(), field.AnchorId);
		const found = findTermsByLabel(candidates, label);
		const available = found.filter(usable);
		if (available.length > 1) {
			throw new Error(
				`The label "${label}" matches ${available.length} terms in the field "${field.InternalName}" (${available
					.map(termDescription)
					.join(', ')}). Use "Label|ID" or the term ID.`,
			);
		}
		if (available.length === 1) {
			result.set(available[0].id, { id: available[0].id, label: available[0].name });
			continue;
		}
		if (found.length) {
			throw new Error(
				`The term "${label}" is deprecated or not available for tagging in the field "${field.InternalName}"`,
			);
		}

		if (!options.createMissing) {
			throw new Error(
				`The term "${label}" does not exist in the term set of the field "${field.InternalName}". Use an existing term, or turn on "Create Missing Terms".`,
			);
		}
		if (INVALID_TERM_NAME.test(label)) {
			throw new Error(
				`The term "${label}" cannot be created: a term name must not contain ; " < > | or tab characters`,
			);
		}
		const termSet = await client.getTermSet(storeId, termSetId);
		if (!termSet.isOpenForTermCreation) {
			throw new Error(
				`The term "${label}" does not exist, and the term set "${termSet.name}" is closed for new terms`,
			);
		}
		const anchorId = isEmptyGuid(field.AnchorId) ? undefined : normalizeGuid(field.AnchorId!);
		const created = await client.createTerm(storeId, termSetId, label, anchorId);
		terms = [...(await load()), created];
		result.set(created.id, { id: created.id, label: created.name || label });
	}
	return [...result.values()];
}

/**
 * Replaces the taxonomy values in a field value object with "Label|ID" text. Other values stay as
 * they are. A cleared value (null, empty text or an empty array) stays empty.
 */
export async function resolveTaxonomyValues(
	values: Record<string, unknown>,
	fieldInfos: TaxonomyFieldInfo[],
	client: TagClient,
	options: ResolveOptions = {},
): Promise<Record<string, unknown>> {
	const fields = new Map(fieldInfos.map((field) => [field.InternalName, field]));
	const result: Record<string, unknown> = { ...values };
	for (const [name, value] of Object.entries(values)) {
		const field = fields.get(name);
		if (!field || !TAXONOMY_TYPES.has(field.TypeAsString) || value === undefined) continue;
		const refs = parseTagInput(value);
		if (!refs.length) {
			result[name] = '';
			continue;
		}
		const existing =
			field.TypeAsString === 'TaxonomyFieldTypeMulti' && options.current
				? parseTagInput(options.current[name])
				: [];
		const tags = await resolveTags([...existing, ...refs], field, client, options);
		result[name] = formatTagValue(tags);
	}
	return result;
}

export function hasTaxonomyValues(
	values: Record<string, unknown>,
	fieldInfos: TaxonomyFieldInfo[],
): boolean {
	const types = new Map(fieldInfos.map((field) => [field.InternalName, field.TypeAsString]));
	return Object.keys(values).some((name) => TAXONOMY_TYPES.has(types.get(name) ?? ''));
}

export function createTagClient(ctx: SharePointContext, siteUrl: string): TagClient {
	return {
		getTerms: async (storeId, termSetId) =>
			await getTerms(ctx, { siteUrl, storeId }, termSetId, { recursive: true }),
		getTermSet: async (storeId, termSetId) =>
			await getTermSet(ctx, { siteUrl, storeId }, termSetId),
		createTerm: async (storeId, termSetId, label, parentTermId) =>
			await createTerm(ctx, { siteUrl, storeId }, termSetId, { name: label, parentTermId }),
	};
}
