import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import * as listSearchBase from './methods/listSearch';
import * as listSearchFields from './methods/listSearchFields';
import * as listSearchPrincipals from './methods/listSearchPrincipals';
import * as listSearchSite from './methods/listSearchSite';
import * as listSearchTaxonomy from './methods/listSearchTaxonomy';
import { getListItemFields } from './methods/resourceMapping';
import { resources } from './resources';
import { getSiteUrl, siteRLC } from './shared/descriptions';
import { toExecutionData, toNodeError } from './shared/utils';
import { createSp } from './transport';

const resourceMap = new Map(resources.map((resource) => [resource.option.value, resource]));

export class SharePointPnp implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'SharePoint PnP',
		name: 'sharePointPnp',
		icon: { light: 'file:sharepoint.svg', dark: 'file:sharepoint.svg' },
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Advanced SharePoint operations through the SharePoint REST API (PnPjs)',
		defaults: { name: 'SharePoint PnP' },
		usableAsTool: true,
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'sharePointPnpOAuth2Api',
				required: true,
				displayOptions: { show: { authentication: ['oAuth2'] } },
			},
			{
				name: 'sharePointPnpCertificateApi',
				required: true,
				displayOptions: { show: { authentication: ['certificate'] } },
			},
		],
		properties: [
			{
				displayName: 'Authentication',
				name: 'authentication',
				type: 'options',
				options: [
					{ name: 'OAuth2 (Delegated)', value: 'oAuth2' },
					{ name: 'Certificate (App-Only)', value: 'certificate' },
				],
				default: 'oAuth2',
			},
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: resources.map((resource) => resource.option),
				default: 'file',
			},
			...resources.flatMap((resource) => resource.description.slice(0, 1)),
			siteRLC,
			...resources.flatMap((resource) => resource.description.slice(1)),
		],
	};

	methods = {
		listSearch: { ...listSearchBase, ...listSearchPrincipals, ...listSearchTaxonomy },
		loadOptions: { ...listSearchFields, ...listSearchSite },
		resourceMapping: { getListItemFields },
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const resourceName = this.getNodeParameter('resource', itemIndex) as string;
				const operationName = this.getNodeParameter('operation', itemIndex) as string;
				const handler = resourceMap.get(resourceName)?.operations[operationName];
				if (!handler) {
					throw new NodeOperationError(
						this.getNode(),
						`The operation "${operationName}" is not supported for "${resourceName}"`,
						{ itemIndex },
					);
				}

				const siteUrl = await getSiteUrl(this, itemIndex);
				const sp = createSp(this, siteUrl);
				const result = await handler.call(this, { sp, siteUrl, itemIndex });
				returnData.push(...toExecutionData(result, itemIndex));
			} catch (error) {
				const nodeError = toNodeError(this.getNode(), error, itemIndex);
				if (this.continueOnFail()) {
					returnData.push({ json: { error: nodeError.message }, pairedItem: { item: itemIndex } });
					continue;
				}
				throw nodeError;
			}
		}

		return [returnData];
	}
}
