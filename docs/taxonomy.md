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

## Pickers

- **Term** in Get Term, Update Term and Delete Term has **From List** and **By ID**. The list shows the terms of one term set. Select the term set in **Term Set (to Pick From)** first. SharePoint does not search all term sets by label in one call, so this helper is needed. The operation itself does not use the helper.
- **Parent Term** in the Create Term and Get Terms options shows the terms of the selected **Term Set**. Empty means the top of the term set.
- **Language** of the additional labels shows the languages of the term store. **Default Language of the Term Store** is 0. An expression can give any LCID.
- Older workflows with a plain term ID in these fields keep working.
