import type { ResourceModule } from '../resources/types';

export interface OutputDefinition {
	type: 'main';
	displayName?: string;
}

/**
 * Returns the outputs for the selected resource and operation. n8n runs this function as an
 * expression in the editor and in the backend, so it must only use its parameters and plain JavaScript
 * that the expression sandbox allows.
 */
export function configuredOutputs(
	parameters: { resource?: unknown; operation?: unknown },
	routed: Record<string, string[]>,
): OutputDefinition[] {
	const key = String(parameters.resource) + '.' + String(parameters.operation);
	// The n8n expression sandbox blocks "prototype", so Array.isArray checks the entry.
	const names = routed[key];
	if (!Array.isArray(names)) return [{ type: 'main' }];
	return names.map(function (name) {
		return { type: 'main', displayName: name };
	});
}

/** Maps "resource.operation" to the output names, for each operation that has more than one output. */
export function getRoutedOutputs(resources: ResourceModule[]): Record<string, string[]> {
	const routed: Record<string, string[]> = {};
	for (const resource of resources) {
		for (const [operation, routing] of Object.entries(resource.routing ?? {})) {
			routed[`${resource.option.value}.${operation}`] = routing.outputs;
		}
	}
	return routed;
}

/** The value for the "outputs" field of the node description. */
export function outputsExpression(resources: ResourceModule[]): `={{${string}}}` {
	return `={{(${configuredOutputs.toString()})($parameter, ${JSON.stringify(getRoutedOutputs(resources))})}}`;
}
