import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import '@pnp/sp/regional-settings/index.js';
import { getStoreInfo } from '../resources/taxonomy/service';
import { getSiteUrl } from '../shared/descriptions';
import { MAX_OPTIONS, sortAndCap } from '../shared/paging';
import { createSp } from '../transport';

async function getLanguageNames(
	ctx: ILoadOptionsFunctions,
	siteUrl: string,
): Promise<Map<number, string>> {
	try {
		const languages = await createSp(ctx, siteUrl).web.regionalSettings.getInstalledLanguages();
		return new Map(languages.map((language) => [language.Lcid, language.DisplayName]));
	} catch {
		return new Map();
	}
}

/** Working languages of the term store. The value 0 means the default language of the store. */
export async function getTermStoreLanguageOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const siteUrl = await getSiteUrl(this);
	const [store, names] = await Promise.all([
		getStoreInfo(this, { siteUrl }),
		getLanguageNames(this, siteUrl),
	]);
	const languages = [...new Set([store.defaultLanguage, ...store.languages])];
	return [
		{ name: 'Default Language of the Term Store', value: 0 },
		...sortAndCap(
			languages.map((lcid) => ({
				name: `${names.get(lcid) ?? 'Language'} (${lcid})`,
				value: lcid,
			})),
			MAX_OPTIONS - 1,
		),
	];
}
