import { describe, expect, it, vi } from 'vitest';

import {
	applySelection,
	buildSelection,
	selectionOptions,
	toList,
} from '../../nodes/SharePointPnp/shared/selection';

describe('property selection', () => {
	it('parses comma-separated text and arrays without duplicates', () => {
		expect(toList(' Id, Title ,,Id')).toEqual(['Id', 'Title']);
		expect(toList(['Id', ' Title'])).toEqual(['Id', 'Title']);
		expect(toList(undefined)).toEqual([]);
	});

	it('adds additional properties to the defaults and expands navigation properties', () => {
		expect(
			buildSelection(['Id', 'Title'], { additionalProperties: 'Author/Title, TimeCreated' }),
		).toEqual({ select: ['Id', 'Title', 'Author/Title', 'TimeCreated'], expand: ['Author'] });
	});

	it('keeps default and requested expands', () => {
		expect(buildSelection(['Id'], { expand: 'ListItemAllFields' }, ['RootFolder'])).toEqual({
			select: ['Id'],
			expand: ['RootFolder', 'ListItemAllFields'],
		});
	});

	it('sends no select when all properties are requested', () => {
		expect(
			buildSelection(['Id'], { returnAllProperties: true, additionalProperties: 'X' }),
		).toEqual({ select: [], expand: [] });
	});

	it('applies select and expand to a query', () => {
		const query = { select: vi.fn(), expand: vi.fn() };
		query.select.mockReturnValue(query);
		query.expand.mockReturnValue(query);
		applySelection(query, ['Id'], { additionalProperties: 'Author/Title' });
		expect(query.select).toHaveBeenCalledWith('Id', 'Author/Title');
		expect(query.expand).toHaveBeenCalledWith('Author');
	});

	it('describes the defaults in the options', () => {
		const [additional] = selectionOptions(['Id', 'Title']);
		expect(additional.description).toContain('Id, Title');
	});
});
