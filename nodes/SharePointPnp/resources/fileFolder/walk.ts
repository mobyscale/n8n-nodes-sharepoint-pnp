import type { IDataObject } from 'n8n-workflow';

export interface DepthRange {
	min: number;
	max: number;
}

/**
 * Reads Min Depth and Max Depth. The direct children of the folder are depth 1. Without Recursive
 * only depth 1 is read. A Max Depth of 0 or less means no limit.
 */
export function toDepthRange(recursive: boolean, minDepth?: number, maxDepth?: number): DepthRange {
	if (!recursive) return { min: 1, max: 1 };
	const min = Math.max(1, Math.floor(Number(minDepth ?? 1) || 1));
	const maxValue = Math.floor(Number(maxDepth ?? 0) || 0);
	const max = maxValue > 0 ? maxValue : Infinity;
	if (max < min) {
		throw new Error(`Min Depth (${min}) must not be greater than Max Depth (${max})`);
	}
	return { min, max };
}

export interface WalkReaders {
	/** The entries of one folder level that pass the CAML filters, at most "max". */
	readOutput: (path: string, max: number) => Promise<IDataObject[]>;
	/** All subfolders of one folder level. */
	readFolders: (path: string) => Promise<IDataObject[]>;
	/** Whether the output of a level has all its subfolders, so that no second request is needed. */
	outputHasAllFolders: boolean;
	/** A filter that runs after the request, for example the file extension filter. */
	accept?: (entry: IDataObject) => boolean;
}

export interface WalkOptions {
	depth: DepthRange;
	limit: number;
	/** Adds "depth" to each entry. */
	addDepth: boolean;
}

/**
 * Walks the folder tree level by level, one request per folder. Levels above Min Depth are walked
 * but not returned. Folders at Max Depth are returned but not read.
 */
export async function walkFolders(
	root: string,
	readers: WalkReaders,
	options: WalkOptions,
): Promise<IDataObject[]> {
	const { depth, limit } = options;
	const entries: IDataObject[] = [];
	const queue: Array<{ path: string; depth: number }> = [{ path: root, depth: 1 }];
	while (queue.length && entries.length < limit) {
		const current = queue.shift() as { path: string; depth: number };
		const output = current.depth >= depth.min;
		let level: IDataObject[] = [];
		if (output) {
			level = await readers.readOutput(
				current.path,
				readers.accept ? Infinity : limit - entries.length,
			);
			for (const entry of level) {
				if (entries.length >= limit) break;
				if (readers.accept && !readers.accept(entry)) continue;
				entries.push(options.addDepth ? { ...entry, depth: current.depth } : entry);
			}
		}
		if (current.depth < depth.max) {
			const folders =
				output && readers.outputHasAllFolders ? level : await readers.readFolders(current.path);
			queue.push(
				...folders
					.filter((entry) => entry.type === 'folder')
					.map((entry) => ({
						path: String(entry.serverRelativeUrl),
						depth: current.depth + 1,
					})),
			);
		}
	}
	return entries;
}
