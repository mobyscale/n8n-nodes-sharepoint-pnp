export interface BatchResult<T> {
	index: number;
	ok: boolean;
	value?: T;
	error?: unknown;
}

/** An SPFI or an IWeb: both have `batched()` when the PnPjs batching module is loaded. */
export interface Batchable<B> {
	batched(props?: { maxRequests?: number }): [B, () => Promise<void>];
}

/**
 * Runs one request for each input in SharePoint $batch requests of up to `chunkSize` requests.
 * `run` must start its request on `batchedSp` synchronously, before its first `await`: PnPjs only adds
 * requests to the batch that start before `execute` runs. One failed request does not stop the others.
 * PnPjs sends at most `maxRequests` requests in one HTTP request (PnPjs default: 20).
 */
export async function runInBatches<B, I, T>(
	sp: Batchable<B>,
	inputs: I[],
	run: (batchedSp: B, input: I, index: number) => Promise<T>,
	chunkSize = 100,
	maxRequests?: number,
): Promise<Array<BatchResult<T>>> {
	const results: Array<BatchResult<T>> = [];
	for (let start = 0; start < inputs.length; start += chunkSize) {
		const [batchedSp, execute] = maxRequests ? sp.batched({ maxRequests }) : sp.batched();
		const pending = inputs.slice(start, start + chunkSize).map((input, offset) => {
			const index = start + offset;
			return run(batchedSp, input, index).then(
				(value): BatchResult<T> => ({ index, ok: true, value }),
				(error: unknown): BatchResult<T> => ({ index, ok: false, error }),
			);
		});
		await execute();
		results.push(...(await Promise.all(pending)));
	}
	return results;
}
