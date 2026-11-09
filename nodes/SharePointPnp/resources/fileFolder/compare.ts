export type CompareInclude = 'folders' | 'files' | 'both';

export interface FolderChild {
	name: string;
	type: 'folder' | 'file';
	serverRelativeUrl: string;
}

export interface TreeEntry {
	/** The path below the compared folder, for example "A/B". */
	path: string;
	type: 'folder' | 'file';
}

export interface TreeOptions {
	include: CompareInclude;
	/** 1 reads only the direct children. 0 reads all levels. */
	maxDepth: number;
	caseSensitive: boolean;
	/** Names or relative paths, with "*" and "?" wildcards. */
	exclude: string[];
}

/** Reads the direct children of one folder. "foldersOnly" lets the reader skip the files. */
export type FolderReader = (
	serverRelativeUrl: string,
	foldersOnly: boolean,
) => Promise<FolderChild[]>;

export interface CompareResult {
	isComplete: boolean;
	missing: string[];
	extra: string[];
	matched: number;
}

export function parsePatterns(value: string | undefined): string[] {
	return String(value ?? '')
		.split(/[,\r\n]+/)
		.map((entry) => entry.trim())
		.filter(Boolean);
}

function patternToRegExp(pattern: string, caseSensitive: boolean): RegExp {
	const source = pattern
		.replace(/^\/+|\/+$/g, '')
		.replace(/[.+^${}()|[\]\\]/g, '\\$&')
		.replace(/\*/g, '.*')
		.replace(/\?/g, '.');
	return new RegExp(`^${source}$`, caseSensitive ? '' : 'i');
}

/** A pattern with "/" matches the relative path. Other patterns match the name. */
export function isExcluded(
	name: string,
	path: string,
	patterns: string[],
	caseSensitive: boolean,
): boolean {
	return patterns.some((pattern) =>
		patternToRegExp(pattern, caseSensitive).test(pattern.includes('/') ? path : name),
	);
}

/** Reads a folder tree level by level. An excluded folder is not read. */
export async function readTree(
	root: string,
	read: FolderReader,
	options: TreeOptions,
): Promise<TreeEntry[]> {
	const entries: TreeEntry[] = [];
	const foldersOnly = options.include === 'folders';
	const queue: Array<{ url: string; path: string; depth: number }> = [
		{ url: root, path: '', depth: 1 },
	];
	while (queue.length) {
		const { url, path, depth } = queue.shift()!;
		for (const child of await read(url, foldersOnly)) {
			const childPath = path ? `${path}/${child.name}` : child.name;
			if (isExcluded(child.name, childPath, options.exclude, options.caseSensitive)) continue;
			if (options.include === 'both' || options.include === `${child.type}s`) {
				entries.push({ path: childPath, type: child.type });
			}
			if (child.type === 'folder' && (options.maxDepth === 0 || depth < options.maxDepth)) {
				queue.push({ url: child.serverRelativeUrl, path: childPath, depth: depth + 1 });
			}
		}
	}
	return entries;
}

const byPath = (a: string, b: string) => a.localeCompare(b);

/** Compares two trees. A folder and a file with the same path do not match. */
export function compareTrees(
	source: TreeEntry[],
	target: TreeEntry[],
	caseSensitive: boolean,
): CompareResult {
	const key = (entry: TreeEntry) =>
		`${entry.type}:${caseSensitive ? entry.path : entry.path.toLowerCase()}`;
	const sourceKeys = new Set(source.map(key));
	const targetKeys = new Set(target.map(key));
	const missing = source.filter((entry) => !targetKeys.has(key(entry))).map((entry) => entry.path);
	const extra = target.filter((entry) => !sourceKeys.has(key(entry))).map((entry) => entry.path);
	return {
		isComplete: missing.length === 0,
		missing: [...new Set(missing)].sort(byPath),
		extra: [...new Set(extra)].sort(byPath),
		matched: source.length - missing.length,
	};
}
