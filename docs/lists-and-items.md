# Lists and items

The **List**, **List Item** and **Attachment** resources of the **SharePoint PnP** node.

## List

PnPjs: [Lists](https://pnp.github.io/pnpjs/sp/lists/), [Fields](https://pnp.github.io/pnpjs/sp/fields/), [Views](https://pnp.github.io/pnpjs/sp/views/), [Column defaults](https://pnp.github.io/pnpjs/sp/column-defaults/), [Content types](https://pnp.github.io/pnpjs/sp/content-types/).

| Operation                     | Description                                             | PnPjs reference                          |
| ----------------------------- | ------------------------------------------------------- | ---------------------------------------- |
| Get                           | Get a list or library                                   | `lists.getById`                          |
| Get Many                      | Get many lists and libraries in a site                  | `lists`                                  |
| Create                        | Create a list or library                                | `lists.add`                              |
| Update                        | Update the title, description and settings of a list    | `list.update`                            |
| Delete                        | Delete a list or library, or move it to the recycle bin | `list.delete`, `list.recycle`            |
| Get Fields                    | Get the fields (columns) of a list                      | `list.fields`                            |
| Create Field                  | Create a field in a list or in the site                 | `fields.add`, `fields.createFieldAsXml`  |
| Update Field                  | Update a field of a list or of the site                 | `field.update`                           |
| Delete Field                  | Delete a field from a list or from the site             | `field.delete`                           |
| Get Views                     | Get the views of a list                                 | `list.views`                             |
| Create View                   | Create a view in a list                                 | `views.add`                              |
| Copy View                     | Copy a view from another list in the same site          | `views.add`, `view.fields`               |
| Get Content Types             | Get the content types of a list or of the site          | `contentTypes`                           |
| Add Content Type to List      | Add an existing site content type to a list             | `contentTypes.addAvailableContentType`   |
| Remove Content Type From List | Remove a content type from a list                       | `contentType.delete`                     |
| Set Default Content Type      | Make a content type the default of a list               | `list.rootFolder.uniqueContentTypeOrder` |
| Get Default Column Values     | Get the default column values of a library              | `list.getDefaultColumnValues`            |
| Set Default Column Values     | Set the default column values of a folder in a library  | `list.setDefaultColumnValues`            |

## List Item

PnPjs: [Items](https://pnp.github.io/pnpjs/sp/items/).

| Operation | Description                                              | PnPjs reference                                     |
| --------- | -------------------------------------------------------- | --------------------------------------------------- |
| Get       | Get a list item                                          | `items.getById`                                     |
| Get Many  | Get many list items. Use an OData filter or a CAML query | `items`, `list.renderListDataAsStream`              |
| Create    | Create a list item or a folder                           | `items.add`, `items.addValidateUpdateItemUsingPath` |
| Update    | Update the fields of a list item. Can do a system update | `item.update`, `item.validateUpdateListItem`        |
| Delete    | Delete a list item, or move it to the recycle bin        | `item.recycle`, `item.delete`                       |

The fields of Create and Update use a resource mapper. It loads the fields of the list. You can also give the fields as JSON. Managed metadata (tag) fields accept term labels. See [Managed metadata values](#managed-metadata-tag-values).

## Attachment

PnPjs: [Attachments](https://pnp.github.io/pnpjs/sp/attachments/). Document libraries do not support attachments.

| Operation      | Description                                                  | PnPjs reference                           |
| -------------- | ------------------------------------------------------------ | ----------------------------------------- |
| Add            | Add an attachment to a list item                             | `item.attachmentFiles.add`                |
| Get Many       | Get the attachments of a list item                           | `item.attachmentFiles`                    |
| Download       | Download an attachment as binary data                        | `attachment.getBuffer`                    |
| Update Content | Replace the content of an attachment                         | `attachment.setContent`                   |
| Delete         | Move an attachment to the recycle bin, or delete it for good | `attachment.recycle`, `attachment.delete` |

## Managed metadata (tag) values

List Item Create and Update accept these values for a managed metadata field:

| Value                                    | Example                                                                     |
| ---------------------------------------- | --------------------------------------------------------------------------- |
| A label                                  | `Example Term`                                                              |
| A term ID                                | `00000000-0000-0000-0000-000000000000`                                      |
| A label and an ID                        | `Example Term\|00000000-0000-0000-0000-000000000000`                        |
| A list that is separated with semicolons | `Term A;Term B`                                                             |
| An array, or a JSON array text           | `["Term A", "Term B"]`                                                      |
| A term object such as SharePoint returns | `{ "Label": "Term A", "TermGuid": "00000000-0000-0000-0000-000000000000" }` |

Notes:

- The node finds the term set from the field. It looks up labels through the term store. See [Taxonomy](taxonomy.md).
- If a label does not exist, the operation fails. Turn on **Create Missing Terms** to create it. The term set must be open for new terms.
- **Append to Multi-Value Fields** (Update) adds the new values to the current values.
- Default column values do not support taxonomy fields.

## Limits and known behavior

- **Attachments** are not available in document libraries. Use [File](files-and-folders.md#file) operations for libraries.
- **System Update** (List Item Update) uses `validateUpdateListItem` with `bNewDocumentUpdate`. It does not change Modified, Modified By or the version number. It does not start "item updated" workflows or alerts. SharePoint reads the date in the time zone of the site. The editor can only be sent by sign-in name. An app principal cannot be resolved, so the editor can change with app-only authentication.
- **List Item Get Many** with **Recursive** in filter mode uses the `FileDirRef` field. It can fail on lists above the list view threshold. Use CAML mode for those lists.
- **List Item Create** can create folders in a document library. It cannot create files. Use File Upload.
- **Default column values** do not support taxonomy fields.
