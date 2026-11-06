export interface DefaultColumnValue {
	name: string;
	value: unknown;
	path: string;
}

export function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

export function normalizeFolderPath(path: string): string {
	return decodeURIComponent(path).replace(/\/+$/, '').toLowerCase();
}

/**
 * Replaces the defaults of one folder. Defaults of other folders stay as they are. With
 * "removeOthers", the defaults of the folder that are not in "updates" are removed.
 */
export function mergeFolderDefaults(
	existing: DefaultColumnValue[],
	folderPath: string,
	updates: Array<{ name: string; value: unknown }>,
	removeOthers: boolean,
): DefaultColumnValue[] {
	const target = normalizeFolderPath(folderPath);
	const names = new Set(updates.map((update) => update.name));
	const kept = existing.filter((entry) => {
		if (normalizeFolderPath(entry.path) !== target) return true;
		return !removeOthers && !names.has(entry.name);
	});
	return [...kept, ...updates.map((update) => ({ ...update, path: folderPath }))];
}

export function filterDefaultsByPath(
	defaults: DefaultColumnValue[],
	folderPath: string | undefined,
): DefaultColumnValue[] {
	if (!folderPath) return defaults;
	const target = normalizeFolderPath(folderPath);
	return defaults.filter((entry) => normalizeFolderPath(entry.path) === target);
}

/** Keeps the view fields that exist in the target list, and reports the others. */
export function splitViewFields(
	sourceFields: string[],
	targetInternalNames: string[],
): { kept: string[]; skipped: string[] } {
	const available = new Set(targetInternalNames);
	return {
		kept: sourceFields.filter((name) => available.has(name)),
		skipped: sourceFields.filter((name) => !available.has(name)),
	};
}
