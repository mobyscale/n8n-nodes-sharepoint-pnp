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

export interface ResourceModule {
	option: INodePropertyOptions;
	/** The first entry must be the resource's "operation" property. The node places the site picker after it. */
	description: INodeProperties[];
	operations: Record<string, OperationHandler>;
}
