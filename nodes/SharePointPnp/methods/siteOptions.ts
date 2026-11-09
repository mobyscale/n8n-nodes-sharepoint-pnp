import type { ILoadOptionsFunctions, INodePropertyOptions } from 'n8n-workflow';

import '@pnp/sp/regional-settings/index.js';
import '@pnp/sp/site-designs/index.js';
import { cleanPropertyBag } from '../resources/site/helpers';
import { getSiteUrl } from '../shared/descriptions';
import { sortAndCap } from '../shared/paging';
import { createSp } from '../transport';

const COMMUNICATION_DESIGNS: INodePropertyOptions[] = [
	{ name: 'Topic (Communication Site)', value: '00000000-0000-0000-0000-000000000000' },
	{ name: 'Showcase (Communication Site)', value: '6142d2a0-63a5-4ba0-aede-d9fefca2c767' },
	{ name: 'Blank (Communication Site)', value: 'f6cc5403-0d63-442e-96c0-285923709ffc' },
];

/** Built-in communication site designs and the custom site designs of the tenant, sorted A-Z. */
export async function getSiteDesignOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const designs = await sp.siteDesigns.getSiteDesigns();
	return sortAndCap([
		...COMMUNICATION_DESIGNS,
		...designs.map((design) => ({
			name: design.Title,
			value: design.Id,
			description: design.Description || undefined,
		})),
	]);
}

/** Keys of the property bag of the selected site. */
export async function getPropertyBagKeyOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const bag = cleanPropertyBag((await sp.web.allProperties()) as Record<string, unknown>);
	return sortAndCap(Object.keys(bag).map((key) => ({ name: key, value: key })));
}

/** Languages that are installed on the site. The value is the LCID. */
export async function getInstalledLanguageOptions(
	this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
	const sp = createSp(this, await getSiteUrl(this));
	const languages = await sp.web.regionalSettings.getInstalledLanguages();
	return sortAndCap(
		languages.map((language) => ({
			name: `${language.DisplayName} (${language.Lcid})`,
			value: language.Lcid,
		})),
	);
}
