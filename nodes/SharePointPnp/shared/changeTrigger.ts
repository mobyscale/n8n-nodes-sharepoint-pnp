import type { INode, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import type { ExpiredTokenMode, ExpiredTokenRunOptions } from './changes';
import { pathRLC } from './descriptions';
import {
	CHANGE_EVENTS,
	getEventFlags,
	resolveFolderFilter,
	triggerOptionsProperty,
} from './changes';

export interface TriggerOptions {
	changeFieldFilters?: unknown;
	changeMaxDepth?: number;
	excludeEditors?: string;
	fetchLimit?: number;
	includeItemDetails?: boolean;
	itemFields?: string | string[];
	onExpiredToken?: ExpiredTokenMode;
	resolveDeletedItems?: boolean;
	startFromDate?: string;
}

export const eventsProperty: INodeProperties = {
	displayName: 'Events',
	name: 'events',
	type: 'multiOptions',
	required: true,
	options: CHANGE_EVENTS.map(({ name, value, description }) => ({ name, value, description })),
	default: ['fileCreated', 'fileUpdated'],
	description: 'The changes that start the workflow',
};

export const folderPathFilterProperty: INodeProperties = pathRLC({
	displayName: 'Folder Path Filter',
	name: 'folderPathFilter',
	required: false,
	default: { mode: 'path', value: '' },
	description:
		'Only trigger for files, folders and items inside this folder and its subfolders. Leave empty for all. With the List scope, SharePoint reads the change log of the folder.',
});

/** The options of the webhook trigger. It has no start date, because it saves a token at activation. */
export const webhookTriggerOptionsProperty: INodeProperties = {
	...triggerOptionsProperty,
	options: (triggerOptionsProperty.options as INodeProperties[]).filter(
		(option) => option.name !== 'startFromDate',
	),
};

interface ParameterContext {
	getNode(): INode;
	getNodeParameter(name: string, fallback?: unknown): unknown;
}

export interface TriggerSettings {
	events: string[];
	options: TriggerOptions;
	flags: ReturnType<typeof getEventFlags>;
	folderPathFilter?: string;
	queryOptions: { fetchLimit?: number };
	runOptions: ExpiredTokenRunOptions;
}

/** Reads the events, the folder path filter and the options that the polling and webhook triggers share. */
export function getTriggerSettings(ctx: ParameterContext, siteUrl: string): TriggerSettings {
	const events = getTriggerEvents(ctx);
	const options = (ctx.getNodeParameter('options', {}) ?? {}) as TriggerOptions;
	const folderPathFilter = resolveFolderFilter(
		siteUrl,
		ctx.getNodeParameter('folderPathFilter', ''),
	);
	return {
		events,
		options,
		flags: getEventFlags(events),
		folderPathFilter,
		queryOptions: { fetchLimit: options.fetchLimit },
		runOptions: {
			returnAll: true,
			events,
			folderPathFilter,
			includeItemDetails: options.includeItemDetails,
			itemFields: options.itemFields,
			resolveDeletedItems: options.resolveDeletedItems ?? true,
			onExpiredToken: options.onExpiredToken,
			maxDepth: options.changeMaxDepth,
			excludeEditors: options.excludeEditors,
			fieldFilters: options.changeFieldFilters,
		},
	};
}

export function getTriggerEvents(ctx: ParameterContext): string[] {
	const events = ctx.getNodeParameter('events', []) as string[];
	if (!events.length) {
		throw new NodeOperationError(ctx.getNode(), 'Select at least one event');
	}
	return events;
}
