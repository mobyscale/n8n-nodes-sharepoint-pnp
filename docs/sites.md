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
