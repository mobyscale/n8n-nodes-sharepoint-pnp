import type {
	IDataObject,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IPollFunctions,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

import * as listSearch from '../SharePointPnp/methods/listSearch';
import {
	eventsProperty,
	folderPathFilterProperty,
	getTriggerEvents,
	getTriggerSettings,
} from '../SharePointPnp/shared/changeTrigger';
import type { ChangeScope, LocatorValue } from '../SharePointPnp/shared/changes';
import {
	buildChangeQuery,
	changeScopeProperty,
	createChangeTokenFromDate,
	getChangeTarget,
	runChangeQuery,
	runChangeQueryWithExpiry,
	showFor,
	triggerOptionsProperty,
} from '../SharePointPnp/shared/changes';
import { getSiteUrl, listRLC, siteRLC } from '../SharePointPnp/shared/descriptions';
import { toNodeError } from '../SharePointPnp/shared/utils';
import { createSp } from '../SharePointPnp/transport';

/** In manual mode the trigger shows the matching changes of this period. */
export const MANUAL_LOOKBACK_MS = 24 * 60 * 60 * 1000;
export const MANUAL_LIMIT = 10;

export interface TriggerState extends IDataObject {
	changeToken?: string;
	scopeKey?: string;
}

export class SharePointPnpTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'SharePoint PnP Trigger',
		name: 'sharePointPnpTrigger',
		icon: { light: 'file:sharepoint.svg', dark: 'file:sharepoint.svg' },
		group: ['trigger'],
		version: 1,
		subtitle: '={{$parameter["events"].join(", ")}}',
		description: 'Starts the workflow when files, folders or items change in SharePoint',
		defaults: { name: 'SharePoint PnP Trigger' },
		polling: true,
		inputs: [],
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
			siteRLC,
			changeScopeProperty,
			...showFor([listRLC], { scope: ['list'] }),
			eventsProperty,
			folderPathFilterProperty,
			triggerOptionsProperty,
		],
	};

	methods = { listSearch };

	async poll(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
		try {
			return await pollChanges.call(this);
		} catch (error) {
			throw toNodeError(this.getNode(), error, 0);
		}
	}
}

export async function pollChanges(this: IPollFunctions): Promise<INodeExecutionData[][] | null> {
	getTriggerEvents(this);
	const scope = this.getNodeParameter('scope') as ChangeScope;
	const list = scope === 'list' ? (this.getNodeParameter('list') as LocatorValue) : undefined;
	if (scope === 'list' && !String(list?.value ?? '').trim()) {
		throw new NodeOperationError(this.getNode(), 'Select a list');
	}

	const siteUrl = await getSiteUrl(this);
	const sp = createSp(this, siteUrl);
	const { flags, folderPathFilter, options, queryOptions, runOptions } = getTriggerSettings(
		this,
		siteUrl,
	);
	const target = getChangeTarget(sp, scope, list, scope === 'list' ? folderPathFilter : undefined);

	if (this.getMode() === 'manual') {
		const since = new Date(Date.now() - MANUAL_LOOKBACK_MS);
		const startToken = createChangeTokenFromDate(scope, await target.getId(), since);
		const result = await runChangeQuery(
			sp,
			target,
			buildChangeQuery({ ...flags, startToken, queryOptions }),
			{ ...runOptions, deadline: getDeadline(this) },
		);
		const latest = result.changes.slice(-MANUAL_LIMIT);
		return latest.length ? [this.helpers.returnJsonArray(latest)] : null;
	}

	const state = this.getWorkflowStaticData('node') as TriggerState;
	const scopeKey = [siteUrl, scope, list?.mode ?? '', list?.value ?? ''].join('|').toLowerCase();
	let startToken = state.scopeKey === scopeKey ? state.changeToken : undefined;

	if (!startToken) {
		if (!options.startFromDate) {
			state.changeToken = await target.getCurrentToken();
			state.scopeKey = scopeKey;
			return null;
		}
		const date = new Date(options.startFromDate);
		if (!Number.isFinite(date.getTime())) {
			throw new NodeOperationError(this.getNode(), 'The start date is not valid');
		}
		startToken = createChangeTokenFromDate(scope, await target.getId(), date);
	}

	const result = await runChangeQueryWithExpiry(
		sp,
		target,
		buildChangeQuery({ ...flags, startToken, queryOptions }),
		{ ...runOptions, deadline: getDeadline(this) },
	);

	state.changeToken = result.nextChangeToken ?? startToken;
	state.scopeKey = scopeKey;
	return result.changes.length ? [this.helpers.returnJsonArray(result.changes)] : null;
}

/** Stops reading pages before the poll budget ends. Older n8n versions do not have a budget. */
function getDeadline(ctx: IPollFunctions): number | undefined {
	const budget = typeof ctx.getPollBudgetMs === 'function' ? ctx.getPollBudgetMs() : undefined;
	if (!budget || !Number.isFinite(budget)) return undefined;
	return Date.now() + Math.max(budget * 0.8, budget - 10_000);
}
