import type { INodeProperties } from 'n8n-workflow';
import { describe, expect, it } from 'vitest';

import { SharePointPnp } from '../../nodes/SharePointPnp/SharePointPnp.node';

// n8n keeps one value per parameter name, so two resources must not use one name with two types.
function findTypeClashes(properties: INodeProperties[]): string[] {
	const seen = new Map<string, string>();
	const clashes: string[] = [];
	const walk = (list: INodeProperties[], prefix: string) => {
		for (const property of list) {
			const key = prefix + property.name;
			const type = seen.get(key);
			if (type !== undefined && type !== property.type) {
				clashes.push(`${key}: ${type} and ${property.type}`);
			} else {
				seen.set(key, property.type);
			}
			if (property.type === 'collection' && property.options) {
				walk(property.options as INodeProperties[], `${key}.`);
			}
			if (property.type === 'fixedCollection' && property.options) {
				for (const option of property.options as Array<{
					name: string;
					values?: INodeProperties[];
				}>) {
					walk(option.values ?? [], `${key}.${option.name}.`);
				}
			}
		}
	};
	walk(properties, '');
	return clashes;
}

describe('node parameters', () => {
	it('use each parameter name with only one type', () => {
		expect(findTypeClashes(new SharePointPnp().description.properties)).toEqual([]);
	});
});
