import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import * as changeOptions from './methods/changeOptions';
import * as listSearchBase from './methods/listSearch';
import * as listSearchFields from './methods/listSearchFields';
import { searchFiles, searchFileVersions } from './methods/listSearchFiles';
import * as listSearchItems from './methods/listSearchItems';
import * as listSearchScoped from './methods/listSearchScoped';
import * as listSearchPrincipals from './methods/listSearchPrincipals';
import * as listSearchPropertyBag from './methods/listSearchPropertyBag';
import * as listSearchSite from './methods/listSearchSite';
import * as listSearchTaxonomy from './methods/listSearchTaxonomy';
import * as listSearchTerms from './methods/listSearchTerms';
import * as recycleBinOptions from './methods/recycleBinOptions';
import * as siteOptions from './methods/siteOptions';
import * as taxonomyOptions from './methods/taxonomyOptions';
import { getListItemFields } from './methods/resourceMapping';
import { resources } from './resources';
import { executeBulk } from './shared/bulk';
import { getSiteUrl, siteRLC } from './shared/descriptions';
import { outputsExpression } from './shared/outputs';
import { toExecutionData, toNodeError } from './shared/utils';
import { createSp } from './transport';

const resourceMap = new Map(resources.map((resource) => [resource.option.value, resource]));

export class SharePointPnp implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'SharePoint PnP',
		name: 'sharePointPnp',
		icon: { light: 'file:sharepoint-pnp.svg', dark: 'file:sharepoint-pnp.svg' },
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Advanced SharePoint operations through the SharePoint REST API (PnPjs)',
		defaults: { name: 'SharePoint PnP' },
		usableAsTool: true,
		inputs: [NodeConnectionTypes.Main],
		outputs: outputsExpression(resources),
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
		listSearch: {
			...listSearchBase,
			...listSearchItems,
			...listSearchScoped,
			...listSearchPrincipals,
			...listSearchTaxonomy,
			...listSearchPropertyBag,
			...listSearchTerms,
			searchFiles,
			searchFileVersions,
		},
		loadOptions: {
			...listSearchFields,
			...listSearchSite,
			...changeOptions,
			...recycleBinOptions,
			...siteOptions,
			...taxonomyOptions,
		},
		resourceMapping: { getListItemFields },
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		// Resource and operation do not accept expressions, so the first item decides for all items.
		const module = items.length
			? resourceMap.get(this.getNodeParameter('resource', 0) as string)
			: undefined;
		const operation = items.length ? (this.getNodeParameter('operation', 0) as string) : '';
		const bulkHandler = module?.bulkOperations?.[operation];
		if (bulkHandler) return [await executeBulk.call(this, bulkHandler)];

		const routing = module?.routing?.[operation];
		const outputs: INodeExecutionData[][] = (routing?.outputs ?? ['main']).map(() => []);

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
				outputs[routing ? routing.route(result) : 0].push(...toExecutionData(result, itemIndex));
			} catch (error) {
				const nodeError = toNodeError(this.getNode(), error, itemIndex);
				if (this.continueOnFail()) {
					outputs[routing?.errorOutput ?? 0].push({
						json: { error: nodeError.message },
						pairedItem: { item: itemIndex },
					});
					continue;
				}
				throw nodeError;
			}
		}

		return outputs;
	}
}
