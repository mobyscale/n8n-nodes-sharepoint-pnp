# Taxonomy

The **Taxonomy** resource of the **SharePoint PnP** node. It reads and changes the term store (managed metadata). To set tag values on list items, see [Managed metadata values](lists-and-items.md#managed-metadata-tag-values).

Taxonomy does not use PnPjs. It uses SharePoint CSOM through the `_vti_bin/client.svc/ProcessQuery` endpoint. The REST term store API (`_api/v2.1/termStore`) does not accept the app-only SharePoint token. CSOM accepts it. See [Complete basic operations using SharePoint client library code](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-client-library-code) for the CSOM model.

| Operation       | Description                                                                            | API                         |
| --------------- | -------------------------------------------------------------------------------------- | --------------------------- |
| Get Term Groups | Get the term groups of the term store                                                  | CSOM `TermStore.Groups`     |
| Get Term Sets   | Get the term sets of a group, or of all groups                                         | CSOM `TermGroup.TermSets`   |
| Get Term Set    | Get a term set                                                                         | CSOM `TermStore.GetTermSet` |
| Get Terms       | Get the terms of a term set                                                            | CSOM `TermSet.Terms`        |
| Search Terms    | Find the terms of a term set by label. The node loads the terms and matches the labels | CSOM `TermSet.Terms`        |
| Get Term        | Get a term                                                                             | CSOM `TermStore.GetTerm`    |
| Create Term     | Create a term in a term set or under another term                                      | CSOM `TermSet.CreateTerm`   |
| Update Term     | Change the name, description or labels of a term                                       | CSOM `Term`                 |
| Delete Term     | Delete a term and its child terms                                                      | CSOM `Term.DeleteObject`    |

## Limits and known behavior

- The identity must have access to the term store.
- Term store changes apply to the whole tenant.
- **Delete Term** deletes the child terms too.
