import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';

import type { BulkItemOutcome, BulkOperationHandler } from '../resources/types';
import { createSp } from '../transport';
import { getSiteUrl } from './descriptions';
import { toExecutionData, toNodeError } from './utils';

/**
 * Runs a bulk operation for all input items. Items are grouped by site, and the handler gets each
 * group in one call. The output has the items in input order, with pairedItem and Continue On Fail
 * handled for each input item.
 */
export async function executeBulk(
	this: IExecuteFunctions,
	handler: BulkOperationHandler,
): Promise<INodeExecutionData[]> {
	const count = this.getInputData().length;
	const outcomes = new Map<number, BulkItemOutcome>();
	const sites = new Map<string, number[]>();

	for (let itemIndex = 0; itemIndex < count; itemIndex++) {
		try {
			const siteUrl = await getSiteUrl(this, itemIndex);
			sites.set(siteUrl, [...(sites.get(siteUrl) ?? []), itemIndex]);
		} catch (error) {
			outcomes.set(itemIndex, { itemIndex, error });
		}
	}

	for (const [siteUrl, itemIndexes] of sites) {
		try {
			const results = await handler.call(this, {
				sp: createSp(this, siteUrl),
				siteUrl,
				itemIndexes,
			});
			for (const outcome of results) outcomes.set(outcome.itemIndex, outcome);
		} catch (error) {
			for (const itemIndex of itemIndexes) {
				if (!outcomes.has(itemIndex)) outcomes.set(itemIndex, { itemIndex, error });
			}
		}
	}

	const returnData: INodeExecutionData[] = [];
	for (let itemIndex = 0; itemIndex < count; itemIndex++) {
		const outcome: BulkItemOutcome = outcomes.get(itemIndex) ?? {
			itemIndex,
			error: new Error('The operation returned no result for this item'),
		};
		if ('error' in outcome) {
			const nodeError = toNodeError(this.getNode(), outcome.error, itemIndex);
			if (this.continueOnFail()) {
				returnData.push({ json: { error: nodeError.message }, pairedItem: { item: itemIndex } });
				continue;
			}
			throw nodeError;
		}
		returnData.push(...toExecutionData(outcome.result, itemIndex));
	}
	return returnData;
}
