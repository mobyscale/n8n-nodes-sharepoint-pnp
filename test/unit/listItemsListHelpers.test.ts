import { describe, expect, it } from 'vitest';

import {
	filterDefaultsByPath,
	mergeFolderDefaults,
	splitViewFields,
} from '../../nodes/SharePointPnp/resources/list/helpers';

const existing = [
	{ name: 'Status', value: 'Open', path: '/sites/Example/Docs' },
	{ name: 'Owner', value: 'A', path: '/sites/Example/Docs/Sub%20Folder' },
	{ name: 'Kind', value: 'X', path: '/sites/Example/Docs/Sub%20Folder' },
];

describe('mergeFolderDefaults', () => {
	it('replaces a default with the same name in the same folder', () => {
		const result = mergeFolderDefaults(
			existing,
			'/sites/Example/Docs/Sub Folder',
			[{ name: 'Owner', value: 'B' }],
			false,
		);
		expect(result).toEqual([
			{ name: 'Status', value: 'Open', path: '/sites/Example/Docs' },
			{ name: 'Kind', value: 'X', path: '/sites/Example/Docs/Sub%20Folder' },
			{ name: 'Owner', value: 'B', path: '/sites/Example/Docs/Sub Folder' },
		]);
	});

	it('removes other defaults of the folder only when asked', () => {
		const result = mergeFolderDefaults(
			existing,
			'/sites/Example/Docs/Sub Folder/',
			[{ name: 'Owner', value: 'B' }],
			true,
		);
		expect(result.map((entry) => entry.name)).toEqual(['Status', 'Owner']);
	});

	it('does not change other folders', () => {
		const result = mergeFolderDefaults(existing, '/sites/Example/Docs/New', [], true);
		expect(result).toEqual(existing);
	});
});

describe('filterDefaultsByPath', () => {
	it('matches paths without case or encoding differences', () => {
		expect(filterDefaultsByPath(existing, '/sites/example/docs/sub folder/')).toHaveLength(2);
		expect(filterDefaultsByPath(existing, undefined)).toHaveLength(3);
	});
});

describe('splitViewFields', () => {
	it('splits fields by availability in the target list', () => {
		expect(splitViewFields(['Title', 'Gone', 'Status'], ['Title', 'Status'])).toEqual({
			kept: ['Title', 'Status'],
			skipped: ['Gone'],
		});
	});
});
