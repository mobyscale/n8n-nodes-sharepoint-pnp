# n8n-nodes-sharepoint-pnp

n8n community nodes for SharePoint Online. The nodes wrap [PnPjs](https://pnp.github.io/pnpjs/) (`@pnp/sp`). PnPjs is the community library for the SharePoint REST API. See the [PnPjs repository](https://github.com/pnp/pnpjs) for the source.

Use these nodes for operations that the built-in Microsoft SharePoint node does not have. Examples are permissions, sharing links, the change log, search, the recycle bin, managed metadata, and folder copy and move with the SharePoint copy job.

These nodes use the SharePoint REST API. They do not use Microsoft Graph.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

<img width="1327" height="842" alt="image" src="https://github.com/user-attachments/assets/78db8ceb-93c9-4e34-88a5-e2d3ce88d847" />

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation. The package name is `n8n-nodes-sharepoint-pnp`.

PnPjs is bundled into the nodes when the package is built. The package has no runtime dependencies.

## Credentials

| Credential                     | Use                                                                                 |
| ------------------------------ | ----------------------------------------------------------------------------------- |
| SharePoint PnP OAuth2 API      | Delegated. The node signs in as a user                                              |
| SharePoint PnP Certificate API | App-only. The node signs in as an Entra app with a certificate. No user is involved |

See [Credentials](docs/credentials.md) for the setup steps and for the rights that each feature needs.

## Nodes

| Node                           | Description                                                                                    |
| ------------------------------ | ---------------------------------------------------------------------------------------------- |
| SharePoint PnP                 | Actions. Select a **Resource** and an **Operation**                                            |
| SharePoint PnP Trigger         | Starts a workflow from the change log of a site collection, site or list. Polls. No public URL |
| SharePoint PnP Webhook Trigger | Starts a workflow from a list webhook. SharePoint must reach n8n over HTTPS                    |

## Areas

| Area                                                       | Resources and nodes                                                                                                |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| [Sites](docs/sites.md)                                     | Site: site info, subsites, property bag, site creation, hub sites                                                  |
| [Files and folders](docs/files-and-folders.md)             | File, Folder, File/Folder: upload, download, copy and move with copy jobs, check in and out, publishing, versions  |
| [Lists and items](docs/lists-and-items.md)                 | List, List Item, Attachment: list settings, fields, views, content types, default column values, items, tag values |
| [Permissions and sharing](docs/permissions-and-sharing.md) | Permission, Sharing, Group, User: role assignments, inheritance, sharing links, site groups and users              |
| [Search and recycle bin](docs/search-and-recycle-bin.md)   | Search, Recycle Bin                                                                                                |
| [Taxonomy](docs/taxonomy.md)                               | Taxonomy: term groups, term sets and terms                                                                         |
| [Change log and triggers](docs/changes-and-triggers.md)    | Change, SharePoint PnP Trigger, SharePoint PnP Webhook Trigger                                                     |

[Common concepts](docs/common-concepts.md) explains the rules that all resources share: site selection, path formats and property selection.

## Compatibility

- Package version: 0.1.0 (see `package.json`).
- n8n nodes API version: 1.
- `n8n-workflow` is a peer dependency with any version (`*`). The package is built and tested against `n8n-workflow` 2.46.5 or later in the 2.x line.
- Built with `@n8n/node-cli` 0.55.5. The bundle targets Node.js 20.
- PnPjs `@pnp/sp` 4.22 or later is bundled.

## Development

See [Development](docs/development.md) for the build, the tests and the contribution rules.

## Resources and license

- [PnPjs documentation](https://pnp.github.io/pnpjs/) and [repository](https://github.com/pnp/pnpjs)
- [SharePoint REST API](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [Change log](CHANGELOG.md)

License: [MIT](LICENSE).
