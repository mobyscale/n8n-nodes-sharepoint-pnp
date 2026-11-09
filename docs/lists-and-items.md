# Lists and items

The **List**, **List Item** and **Attachment** resources of the **SharePoint PnP** node.

## List

PnPjs: [Lists](https://pnp.github.io/pnpjs/sp/lists/), [Fields](https://pnp.github.io/pnpjs/sp/fields/), [Views](https://pnp.github.io/pnpjs/sp/views/), [Column defaults](https://pnp.github.io/pnpjs/sp/column-defaults/), [Content types](https://pnp.github.io/pnpjs/sp/content-types/).

| Operation                     | Description                                             | PnPjs reference                          |
| ----------------------------- | ------------------------------------------------------- | ---------------------------------------- |
| Get                           | Get a list or library                                   | `lists.getById`                          |
| Get Many                      | Get many lists and libraries in a site                  | `lists`                                  |
| Create                        | Create a list or library                                | `lists.add`                              |
| Create If Missing             | Get a list by URL or title, create it if it is missing  | `web.getList`, `lists.add`               |
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

### Create If Missing

Gets a list or library. If it does not exist, the node creates it and applies the options. Use it to make sure that a library with the correct settings exists before other steps of a workflow.

- **Title**: the title of the list. If **URL** is empty, the node finds the list by this title.
- **URL**: optional. The URL of the list root folder: site-relative (`Archive`, `Lists/Tasks`), server-relative (`/sites/Example/Archive`) or absolute. If set, the node finds the list by this URL. It creates the list with the last part of the URL and then sets the title. The template sets the parent folder: libraries go to the site root and lists go to `Lists/`. If the new list has another URL than the given URL, the output has a warning.
- **Template**: the kind of list to create (default Document Library). The node does not change the template of a list that exists.
- **Description**: the description of a new list.

If **URL** is set and the list is not at that URL, but a list with the same title exists at another URL, the operation fails. SharePoint cannot create two lists with the same title.

Options:

- **Apply Settings to Existing List** (default off): if off, the node only returns a list that exists and does not change it. If on, it applies the options to that list too, and also replaces the description. The node reads the current values first and writes only the values that are different.
- **Show on Quick Launch**, **Enable Versioning**, **Major Version Limit**, **Enable Minor Versions** and **Minor Version Limit**: list settings. Minor versions are only for libraries.
- **Enable Content Types**: allow the management of content types. Content Types to Add and Default Content Type turn it on.
- **Content Types to Add**: site content types to add. A content type that the list has is skipped.
- **Default Content Type**: a site content type, a list content type ID or a content type name. The list must have it, or it must be in Content Types to Add.
- **Copy View From List** and **View to Copy**: copy a view (title or ID) from another list of the site, the same as **Copy View**. If the list has a view with the same title, the view is not copied, and the output has a warning.
- **Permissions**: give the list its own permissions.
  - **Break Inheritance** (default on), with **Copy Role Assignments** (default off) and **Clear Subscopes** (default off). The node breaks inheritance only when the list inherits.
  - **Role Assignments (JSON)**: the same JSON as **Permission > Set** in JSON mode. The node resolves all principals and roles before it creates the list, so an unknown principal changes nothing. Role assignments always break inheritance first.
  - **Role Assignment Mode**: **Replace** (default) removes the role assignments that are not in the JSON, like Permission > Set. **Add** keeps them.

The node applies the options in this order: list settings, content types, default content type, view, permissions.

The output has:

| Field               | Description                                                       |
| ------------------- | ----------------------------------------------------------------- |
| `id`                | The list ID                                                       |
| `title`             | The list title                                                    |
| `serverRelativeUrl` | The server-relative URL of the list root folder                   |
| `created`           | `true` when the node created the list                             |
| `changes`           | One text for each change, for example `Set OnQuickLaunch to true` |
| `warnings`          | Only when there is a warning                                      |

When the list exists and Apply Settings to Existing List is off, `changes` is empty.

The integration tests do not create lists or change list settings. Only the read test runs: it finds the existing test library. The tests that create a list or apply settings use `it.skip`. Run them only on a test site that you can change and reset.

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

### Get Many field filters

**Field Filters** add conditions on the fields of the list, for example `Status` Equals `Draft`. The **Field** picker lists the fields of the selected list. See [Field Filters](common-concepts.md#field-filters) for the operators and the values. The option **Combine Field Filters** sets how the conditions are combined: **All (AND)** (default) or **Any (OR)**.

Field Filters work in both query modes:

- **Filter** mode: the node makes an OData filter from the conditions. If you also give a **Filter**, the node combines the two with `and`.
- **CAML Query** mode: the node makes a CAML condition. If the query has a `Where` element, the node combines the two with `And`. If not, the node adds a `Where` element. This works for the content of a `Query` element, a `Query` element and a full `View` element.

On a list with more than 5000 items, filter mode can fail when a field in the filter has no index, and with **Recursive**, filter mode uses the `FileDirRef` field. Use CAML mode with the **Folder** option to query one folder.

### Bulk operations

These operations process all input items of one node run together. The node sends the requests in SharePoint `$batch` requests, not one request for each input item.

| Operation        | Description                                                | PnPjs reference                                                                         |
| ---------------- | ---------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| Create Many      | Create one list item for each input item                   | `sp.batched`, `list.addValidateUpdateItemUsingPath`                                     |
| Update Many      | Update one list item for each input item                   | `sp.batched`, `item.validateUpdateListItem`                                             |
| Delete Many      | Delete one list item for each input item, or recycle it    | `sp.batched`, `item.recycle`, `item.delete`                                             |
| Create or Update | Find an item by a key field value. Update it, or create it | `sp.batched`, `items.filter`, `list.getItemsByCAMLQuery`, `item.validateUpdateListItem` |

Inputs for each input item:

- **List**: the list. Input items can use different lists and sites. The node makes one group for each site and list.
- **Item ID** (Update Many, Delete Many): the numeric ID of the item. Use an expression such as `{{ $json.ID }}`.
- **Fields** (Create Many, Update Many, Create or Update): the same field input as Create and Update. Use the field mapper or JSON, with expressions for each input item.
- **Key Field** and **Key Value** (Create or Update): the field and the value that identify the item.

Example of the JSON field input. The keys are internal field names. Use an expression, such as `{{ JSON.stringify({ Title: $json.title }) }}`, to read values from each input item.

```json
{
	"Title": "Example",
	"Status": "Open",
	"Amount": 12.5,
	"Department": 4,
	"Tags": ["A", "B"]
}
```

Options:

| Option                   | Operations                                 | Description                                                                                                                          |
| ------------------------ | ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------ |
| Batch Size               | All                                        | The number of requests in one `$batch` request. The default and the maximum is 100.                                                  |
| Continue on Item Error   | All                                        | Output a failed item with `success: false` and `error`, and continue. If off, a failed item stops the node.                          |
| Create Missing Terms     | Create Many, Update Many, Create or Update | Create unknown managed metadata terms. See [Managed metadata values](#managed-metadata-tag-values).                                  |
| Find in Folder           | Create or Update                           | Look for the item only directly in this folder, with a CAML query. New items go to this folder unless Folder Path is set.            |
| Folder Path              | Create Many, Create or Update              | Create new items in this folder. Empty creates them in the list root.                                                                |
| System Update            | Update Many, Create or Update              | Update without a change to Modified, Modified By or the version number. See [Limits and known behavior](#limits-and-known-behavior). |
| Delete Permanently       | Delete Many                                | Delete the items for good. If off, the items go to the recycle bin.                                                                  |
| If Several Match         | Create or Update                           | **Update First** (the item with the lowest ID, default), **Update All** (up to 5000 items), or **Fail**.                             |
| Skip Update If Unchanged | Create or Update                           | Read the current values of the given fields. If all are equal, do not write the item.                                                |

Output: one item for each input item, in input order, with `pairedItem` set to the input item.

| Property   | Description                                                                                |
| ---------- | ------------------------------------------------------------------------------------------ |
| `id`       | The ID of the item. For Create or Update with Update All, the first ID. `ids` has all IDs. |
| `success`  | `true` when the item was written, `false` when it failed (with Continue on Item Error).    |
| `error`    | The error message, when the item failed.                                                   |
| `action`   | Create or Update only: `created`, `updated` or `unchanged`.                                |
| `recycled` | Delete Many only: `true` when the item went to the recycle bin.                            |

Errors:

- An input error, such as an invalid item ID or invalid JSON, fails only that input item. The node sends no request for it.
- A failed request in a `$batch` request fails only its input item. The other requests in the batch still run.
- A list that cannot be read, for example because it does not exist, fails all input items of that list.
- If Continue on Item Error is off, the node stops with the error of the first failed input item. This occurs after the batch ran, so other items can already be changed. If the node setting **On Error** is **Continue**, the node outputs `{ "error": "..." }` for the failed input item instead.

#### Create or Update key fields

The node finds the item with an OData filter on the key field, in the whole list. With **Find in Folder**, it uses a CAML query that looks only directly in that folder (not in subfolders). These field types can be a key:

| Field type                                             | Key value                              | OData filter          | CAML value                                          |
| ------------------------------------------------------ | -------------------------------------- | --------------------- | --------------------------------------------------- |
| Single line of text, Choice, file name (`FileLeafRef`) | Text. Quotes are escaped.              | `Title eq 'O''Brien'` | `<Value Type="Text">O'Brien</Value>`                |
| Number, Currency, Integer, ID (`Counter`)              | A number                               | `Amount eq 12.5`      | `<Value Type="Number">12.5</Value>`                 |
| Yes/No (`Boolean`)                                     | `true`, `false`, `yes`, `no`, `1`, `0` | `Done eq 1`           | `<Value Type="Boolean">1</Value>`                   |
| Lookup, Person (`User`)                                | The numeric lookup item ID or user ID  | `DepartmentId eq 4`   | `LookupId="TRUE"`, `<Value Type="Lookup">4</Value>` |

Other types, for example multiple lines of text, multi-value fields, dates and managed metadata, cannot be a key. The text comparison in SharePoint ignores case.

On a list above the list view threshold (5000 items), the OData filter fails when the key field has no index. Add an index to the key field, or use Find in Folder with a folder that has fewer than 5000 items.

When no item matches, the node creates an item. It adds the key value to the fields when the fields do not set the key field, for text, choice, number, currency, yes/no and lookup keys. For a person key, give the person in the fields. A document library cannot create files here, so the create path only works in lists.

**Skip Update If Unchanged** can compare text, multiple lines of text, choice, number, yes/no, date, lookup, multi-choice and multi-lookup fields. Other field types, for example person and managed metadata, always count as changed.

## Attachment

PnPjs: [Attachments](https://pnp.github.io/pnpjs/sp/attachments/). Document libraries do not support attachments.

| Operation      | Description                                                  | PnPjs reference                           |
| -------------- | ------------------------------------------------------------ | ----------------------------------------- |
| Add            | Add an attachment to a list item                             | `item.attachmentFiles.add`                |
| Get Many       | Get the attachments of a list item                           | `item.attachmentFiles`                    |
| Download       | Download an attachment as binary data                        | `attachment.getBuffer`                    |
| Update Content | Replace the content of an attachment                         | `attachment.setContent`                   |
| Delete         | Move an attachment to the recycle bin, or delete it for good | `attachment.recycle`, `attachment.delete` |

## Pickers

Many parameters have a **From List** mode, so you can choose a value instead of typing it. Each parameter also accepts a typed value or an expression, and workflows that were saved with plain values keep working.

- **List Item** and **Attachment**: **Item** shows the items of the selected list (newest first). Type text to search by title or file name, or a number to find an ID. **By ID** takes the number.
- **Attachment**: **Attachment Name** shows the attachments of the selected item. **By Name** takes a file name.
- **List Item Get Many**, **Get** (Options), **Create** (Options) and the default column values operations: **Folder** shows the folders of the selected list. Type a path below the list root, for example `Clients/Ac`, to see the matching subfolders.
- **Select Fields** (List Item Get and Get Many) lists the fields of the list. **Other Select Fields** is a text field for names that the list does not show, such as `Author/Title`. The two are added together.
- **Set Default Column Values**: the **Field Name or ID** of each row lists the fields of the list.
- **Field Filters** (List Item Get Many): the **Field Name or ID** of each condition lists the fields of the list.
- **Create Field** with type Lookup: **Lookup Field Name or ID** lists the fields of the lookup list.
- **Group** (Create Field, Update Field) shows the column groups of the site. Use **By Name** to type a new group.

## JSON examples

The **Fields (JSON)** parameter of List Item Create and Update takes an object with internal field names as keys:

```json
{
	"Title": "Example item",
	"Quantity": 5,
	"Status": "Open",
	"Tags": ["Red", "Blue"],
	"DueDate": "2026-01-31T00:00:00Z",
	"DepartmentLookup": 12,
	"Owner": "user@contoso.com",
	"Category": "Example Term"
}
```

- Text, number and choice fields take a plain value. A multi-choice field takes an array.
- A date takes ISO text.
- A lookup field takes the ID of the target item.
- A person field takes a sign-in name or an email address.
- A managed metadata field takes a label, as described below.

The **Other Properties (JSON)** option of List Update Field takes SharePoint field properties as an object. These replace the settings above it:

```json
{
	"MaxLength": 100,
	"Indexed": true
}
```

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
- **Bulk operations** send one `$batch` request for each chunk of up to 100 requests. Update Many with System Update sends two more `$batch` requests for each chunk: one reads the current Modified and Editor values, one converts Modified to the time zone of the site. Create or Update sends one `$batch` request for the lookups and one for the writes, for each chunk.
- **Create or Update** uses an OData filter. On lists above the list view threshold (5000 items), the key field must be indexed, or the lookup fails. Find in Folder uses a CAML query in one folder instead.
- **Item picker** on lists above the list view threshold (5000 items): SharePoint rejects the text search. The picker then scans the items by ID (the creation order), newest first. Each page reads 300 items, shows the items that match the text, sorted by name, and loads the next 300 older items when you scroll. To find an older item faster, type its ID.
- **Default column values** do not support taxonomy fields.
