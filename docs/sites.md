# Sites

The **Site** resource of the **SharePoint PnP** node.

PnPjs: [Webs](https://pnp.github.io/pnpjs/sp/webs/), [Sites](https://pnp.github.io/pnpjs/sp/sites/), [Hub sites](https://pnp.github.io/pnpjs/sp/hubsites/).

Create Site and Get Site Creation Status use the SharePoint `SPSiteManager` REST API. See [Manage modern SharePoint sites using REST](https://learn.microsoft.com/sharepoint/dev/apis/site-creation-rest).

| Operation                | Description                                                                   | PnPjs reference                     |
| ------------------------ | ----------------------------------------------------------------------------- | ----------------------------------- |
| Get                      | Get the selected site, with its site collection information                   | `web`                               |
| Get Subsites             | Get the subsites of the selected site                                         | `web.webs`                          |
| Get Property Bag         | Get all property bag values of the site, or only the given keys               | `web.allProperties`                 |
| Set Property Bag Values  | Set property bag values of the site                                           | `web.allProperties`                 |
| Create Site              | Create a modern Communication site or Team site without a Microsoft 365 group | `SPSiteManager/create` (REST)       |
| Create Subsite           | Create a classic subsite in the selected site                                 | `web.webs.add`                      |
| Get Site Creation Status | Get the creation status of a site                                             | `SPSiteManager/status` (REST)       |
| Get Hub Sites            | Get the hub sites of the tenant                                               | `hubSites`                          |
| Join Hub Site            | Associate the selected site with a hub site                                   | `site.joinHubSite`                  |
| Leave Hub Site           | Remove the association of the selected site with its hub site                 | `site.joinHubSite` with an empty ID |

## Limits and known behavior

- **Site creation** with app-only authentication needs an owner. **Wait Until Ready** stops at **Wait Timeout**. The site can still finish later.
- **Property bag writes** can need "custom scripts" on the site.
- **Hub site join** needs site collection admin rights.

## Pickers

- **Hub Site**, **Site Design** and **Language** in the Create Site and Create Subsite options are lists. **Site Design** shows the built-in communication designs and the site designs of the tenant. Leave it empty for the default. **Language** shows the languages that are installed on the site. An expression can give any LCID.
- **Site to Check** in Get Site Creation Status has **From List** and **By URL**. A new site can be missing from the list because the search index is late. Use **By URL** then.
- **Property Keys** in Get Property Bag is a multi-select of the keys of the site. An expression can give an array or a comma-separated list.
- **Key** in Set Property Bag Values has **From List** and **By Name**. Use **By Name** for a new key.
