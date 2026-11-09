import type { IExecuteFunctions } from 'n8n-workflow';
import { describe, expect, it, vi } from 'vitest';

import { listItemResource } from '../../nodes/SharePointPnp/resources/listItem';

describe('List Item Get Many with field filters', () => {
	const fields = [
		{ InternalName: 'Title', TypeAsString: 'Text' },
		{ InternalName: 'Owner', TypeAsString: 'User' },
	];

	function setup() {
		const query = {
			select: vi.fn().mockReturnThis(),
			expand: vi.fn().mockReturnThis(),
			filter: vi.fn().mockReturnThis(),
			orderBy: vi.fn().mockReturnThis(),
			top: vi.fn().mockReturnValue(vi.fn().mockResolvedValue([])),
		};
		const render = vi.fn().mockResolvedValue({ Row: [] });
		const fieldQuery = vi.fn().mockResolvedValue(fields);
		const list = {
			items: query,
			renderListDataAsStream: render,
			fields: { select: () => fieldQuery },
		};
		const sp = { web: { lists: { getById: () => list, getByTitle: () => list } } };
		return { sp, query, render, fieldQuery };
	}

	const run = async (sp: unknown, parameters: Record<string, unknown>) => {
		const all: Record<string, unknown> = {
			list: { mode: 'id', value: 'guid' },
			returnAll: false,
			limit: 10,
			getAllOptions: {},
			...parameters,
		};
		const context = {
			getNode: () => ({ name: 'SharePoint PnP' }),
			getNodeParameter: (name: string, _index: number, fallback?: unknown) =>
				name in all ? all[name] : fallback,
		} as unknown as IExecuteFunctions;
		await listItemResource.operations.getAll.call(context, {
			sp: sp as never,
			siteUrl: 'https://contoso.sharepoint.com/sites/Example',
			itemIndex: 0,
		});
	};

	const fieldFilters = {
		conditions: [
			{ field: 'Title', operator: 'beginsWith', value: 'Draft' },
			{ field: 'Owner', operator: 'eq', value: '12' },
		],
	};

	it('adds the field filters to the OData filter with and', async () => {
		const { sp, query } = setup();
		await run(sp, { queryMode: 'filter', odataFilter: 'Id gt 5', fieldValueFilters: fieldFilters });
		expect(query.filter).toHaveBeenCalledWith(
			"(Id gt 5) and ((startswith(Title,'Draft')) and (OwnerId eq 12))",
		);
	});

	it('uses only the field filters when there is no OData filter', async () => {
		const { sp, query } = setup();
		await run(sp, {
			queryMode: 'filter',
			fieldValueFilters: fieldFilters,
			getAllOptions: { fieldFiltersCombine: 'or' },
		});
		expect(query.filter).toHaveBeenCalledWith("((startswith(Title,'Draft')) or (OwnerId eq 12))");
	});

	it('adds the field filters to the Where element of the CAML query', async () => {
		const { sp, render } = setup();
		await run(sp, {
			queryMode: 'caml',
			camlQuery: '<Where><Eq><FieldRef Name="ID"/><Value Type="Counter">1</Value></Eq></Where>',
			fieldValueFilters: { conditions: [{ field: 'Title', operator: 'isNotEmpty' }] },
			getAllOptions: { folderPath: '/sites/Example/Lists/Items/Archive' },
		});
		const [parameters] = render.mock.calls[0] as [
			{ ViewXml: string; FolderServerRelativeUrl: string },
		];
		expect(parameters.ViewXml).toContain(
			'<Where><And><Eq><FieldRef Name="ID"/><Value Type="Counter">1</Value></Eq><IsNotNull><FieldRef Name="Title"/></IsNotNull></And></Where>',
		);
		expect(parameters.FolderServerRelativeUrl).toBe('/sites/Example/Lists/Items/Archive');
	});

	it('does not read the fields when there are no field filters', async () => {
		const { sp, query, fieldQuery } = setup();
		await run(sp, { queryMode: 'filter', odataFilter: '' });
		expect(query.filter).not.toHaveBeenCalled();
		expect(fieldQuery).not.toHaveBeenCalled();
	});
});
