# Search and recycle bin

The **Search** and **Recycle Bin** resources of the **SharePoint PnP** node.

## Search

PnPjs: [Search](https://pnp.github.io/pnpjs/sp/search/).

| Operation | Description                      | PnPjs reference    |
| --------- | -------------------------------- | ------------------ |
| Query     | Run a KQL search query           | `sp.search`        |
| Suggest   | Get query suggestions for a text | `sp.searchSuggest` |

## Recycle Bin

PnPjs: [Recycle bin](https://pnp.github.io/pnpjs/sp/recycle-bin/).

| Operation            | Description                                            | PnPjs reference                                                  |
| -------------------- | ------------------------------------------------------ | ---------------------------------------------------------------- |
| Get Many             | Get many recycle bin items                             | `site.recycleBin`, `web.recycleBin`, `getrecyclebinitems` (REST) |
| Restore              | Restore recycle bin items to their original location   | `recycleBin.getById().restore`                                   |
| Move to Second Stage | Move recycle bin items to the second stage recycle bin | `recycleBin.getById().moveToSecondStage`                         |
| Delete Permanently   | Delete recycle bin items for good                      | `recycleBin.getById().delete`                                    |

## Limits and known behavior

- **Search** reads results in pages of up to 500 rows. SharePoint can repeat a row across pages when the index changes. **Include Raw Result** contains the last page only.
- **Recycle Bin, Site scope** includes items of subsites. Get Many reads the recycle bin of the site collection and filters by path.
