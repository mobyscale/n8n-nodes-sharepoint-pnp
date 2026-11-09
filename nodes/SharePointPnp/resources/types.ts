import type {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	INodeProperties,
	INodePropertyOptions,
} from 'n8n-workflow';
import type { SPFI } from '@pnp/sp/index.js';

export interface OperationContext {
	sp: SPFI;
	siteUrl: string;
	itemIndex: number;
}

export type OperationResult = IDataObject | IDataObject[] | INodeExecutionData[];

export type OperationHandler = (
	this: IExecuteFunctions,
	context: OperationContext,
) => Promise<OperationResult>;

export interface BulkOperationContext {
	sp: SPFI;
	siteUrl: string;
	/** The indexes of the input items that use this site, in input order. */
	itemIndexes: number[];
}

/** The result of one input item. An `error` is handled like an error of a per-item operation. */
export type BulkItemOutcome =
	| { itemIndex: number; result: OperationResult }
	| { itemIndex: number; error: unknown };

export type BulkOperationHandler = (
	this: IExecuteFunctions,
	context: BulkOperationContext,
) => Promise<BulkItemOutcome[]>;

/**
 * An operation with more than one output, like the n8n If node. The node shows one output for each
 * name and puts each input item on the output that "route" returns.
 */
export interface OperationRouting {
	outputs: string[];
	/** Returns the index of the output for the result of one input item. */
	route: (result: OperationResult) => number;
	/** The output for an error item when "Continue On Fail" is on. */
	errorOutput: number;
}

export interface ResourceModule {
	option: INodePropertyOptions;
	/** The first entry must be the resource's "operation" property. The node places the site picker after it. */
	description: INodeProperties[];
	operations: Record<string, OperationHandler>;
	/**
	 * Operations that get all input items of the node run in one call, for example to send them in
	 * SharePoint $batch requests. The node calls the handler once for each site.
	 */
	bulkOperations?: Record<string, BulkOperationHandler>;
	/** Operations with more than one output. All other operations have one output. */
	routing?: Record<string, OperationRouting>;
}
