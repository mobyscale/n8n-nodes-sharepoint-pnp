# Files and folders

The **File**, **Folder** and **File/Folder** resources of the **SharePoint PnP** node. For path formats, see [Paths](common-concepts.md#paths).

## File

PnPjs: [Files](https://pnp.github.io/pnpjs/sp/files/). Copy and Move use the [copy job](#copy-job-versus-direct) by default.

| Operation         | Description                                                       | PnPjs reference                          |
| ----------------- | ----------------------------------------------------------------- | ---------------------------------------- |
| Upload            | Upload a file from binary data. Files above 10 MB go up in chunks | `files.addUsingPath`, `files.addChunked` |
| Download          | Download a file as binary data                                    | `file.getBuffer`                         |
| Get               | Get the properties of a file                                      | `file`                                   |
| Get Versions      | Get the versions of a file                                        | `file.versions`                          |
| Copy              | Copy a file                                                       | `file.copyByPath`, or a copy job         |
| Move              | Move a file                                                       | `file.moveByPath`, or a copy job         |
| Rename            | Rename a file                                                     | `file.moveByPath`                        |
| Delete            | Delete a file, or move it to the recycle bin                      | `file.recycle`, `file.deleteWithParams`  |
| Check Out         | Check out a file                                                  | `file.checkout`                          |
| Check In          | Check in a file                                                   | `file.checkin`                           |
| Discard Check Out | Discard the check out of a file                                   | `file.undoCheckout`                      |
| Publish           | Publish the latest version of a file as a major version           | `file.publish`                           |
| Unpublish         | Unpublish the current major version of a file                     | `file.unpublish`                         |
| Approve           | Approve a file that waits for content approval                    | `file.approve`                           |
| Deny              | Deny a file that waits for content approval                       | `file.deny`                              |
| Restore Version   | Restore a version of a file                                       | `file.versions.restoreByLabel`           |

### Upload options

| Option                 | Description                                                                                                                       |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| Overwrite              | Replace a file with the same name. On by default                                                                                  |
| Create Missing Folders | Create the destination folder and its missing parent folders before the upload. The library must exist                            |
| Fields                 | List item fields to set after the upload, as field name and value pairs. An entry replaces the same field in Fields (JSON)        |
| Fields (JSON)          | An object with internal field names as keys, for example `{"Title": "Report", "Status": "Open"}`. Values use the List Item format |
| Check In               | Check in the file after the upload, with Check-In Comment and Check-In Type. The node does this only when the file is checked out |
| Publish                | Publish the file as a major version after the upload, with Publish Comment. The node does this only when the file is a draft      |

The node sets all fields in one `validateUpdateListItem` request. The request does not make a new version: it changes the version that the upload made. The output has `fields` with the values that SharePoint stored. If SharePoint rejects a value, the node fails, but the file stays uploaded.

Check In and Publish do not fail in a library that does not require check out or has no minor versions. The output has `checkedIn` and `published` (`true` or `false`) when you use these options. The order is: upload, fields, check in, publish.

### Include List Item Fields

File Get and Folder Get have the option **Include List Item Fields**. It adds `ListItemAllFields` to the output, with all list item fields of the file or folder. The default properties do not change. It is the same as `ListItemAllFields` in **Expand**.

### Pickers

- **File** has a **From List** mode. Type a site-relative path, for example `Shared Documents/Reports/`. The list shows the folders (with a `/` at the end) and the files of that folder. Type the folder path to go into a folder. Select a file, not a folder.
- **Restore Version** has a **From List** mode that shows the earlier versions of the selected file, newest first. Use **By Label** to type a label such as `1.0`.

## Folder

PnPjs: [Folders](https://pnp.github.io/pnpjs/sp/folders/). Copy and Move use the [copy job](#copy-job-versus-direct) by default.

| Operation | Description                                    | PnPjs reference                             |
| --------- | ---------------------------------------------- | ------------------------------------------- |
| Create    | Create a folder and all missing parent folders | `folders.addUsingPath`                      |
| Get       | Get the properties of a folder                 | `folder`                                    |
| Copy      | Copy a folder                                  | `folder.copyByPath`, or a copy job          |
| Move      | Move a folder                                  | `folder.moveByPath`, or a copy job          |
| Rename    | Rename a folder                                | `folder.moveByPath`                         |
| Delete    | Delete a folder, or move it to the recycle bin | `folder.recycle`, `folder.deleteWithParams` |

The Folder resource also has five [provisioning operations](#provisioning-operations): Ensure Path, Ensure Folders, Copy Children If Missing, Set Metadata Recursive and Inherit Metadata From Parent.

## File/Folder

PnPjs: [Files](https://pnp.github.io/pnpjs/sp/files/) and [Folders](https://pnp.github.io/pnpjs/sp/folders/). Get Copy Job Status uses the SharePoint migration API. See [Copy job versus Direct](#copy-job-versus-direct).

| Operation           | Description                                                    | PnPjs reference               |
| ------------------- | -------------------------------------------------------------- | ----------------------------- |
| Get Many            | Get many files and folders in a folder                         | `list.renderListDataAsStream` |
| Check Exists        | Check if files or folders exist                                | `file`, `folder`              |
| Compare Structure   | Compare the folders and files of two folders                   | `list.renderListDataAsStream` |
| Get Copy Job Status | Get the progress of a copy or move job that runs in SharePoint | `GetCopyJobProgress` (REST)   |

### Get Many depth

With **Recursive** on, Get Many also reads the subfolders. The direct children of the **Folder** are depth 1, their children are depth 2, and so on. Each entry then has a `depth` property. Without Recursive, the output does not have `depth`.

Two options set the levels to return. They show only when Recursive is on:

- **Min Depth** (default 1): the first level to return. The node still reads the levels above it to find the deeper entries, but does not return them.
- **Max Depth** (default 0, no limit): the last level to return. The node does not read the folders below this level.

For example, a folder `Reports` has `Reports/2024/Q1/summary.pdf`. With Min Depth 2 and Max Depth 2, Get Many on `Reports` returns only `Q1`. With Min Depth 3, it returns only `summary.pdf`.

Without Recursive, Min Depth and Max Depth have no effect.

### Get Many field filters

**Field Filters** add conditions on the list item fields of the files and folders, for example `Title` Equals `Final`. The **Field** picker lists the fields of the library that contains the **Folder**. See [Field Filters](common-concepts.md#field-filters) for the operators and the values.

The option **Combine Field Filters** sets how the conditions are combined: **All (AND)** (default) or **Any (OR)**. The other filters (File Extensions, Modified After, Modified Before, Name Contains) and the **Include** choice always apply with AND.

The node queries each folder on its own with a CAML query. The filter is part of that query. The query stays below the list view threshold when each folder has 5000 items or less. With Recursive, the node also reads the subfolders of each level without the filter, so it finds matching entries in folders that do not match.

### Check Exists outputs

Check Exists has two outputs, **True** and **False**, like the n8n If node. The node makes one output item for each input item and puts it on one of the outputs.

The **Combine** option sets when the item goes to True:

- **All** (default): all paths must exist (AND).
- **Any**: at least one path must exist (OR).

The output item has:

| Field       | Description                                                            |
| ----------- | ---------------------------------------------------------------------- |
| `exists`    | The result for the Combine option. `true` on the True output           |
| `allExist`  | `true` when all paths exist                                            |
| `anyExists` | `true` when at least one path exists                                   |
| `combine`   | `all` or `any`                                                         |
| `results`   | One entry for each path: `path`, `serverRelativeUrl`, `exists`, `type` |

`type` is `file`, `folder` or `null`. The output item does not have the fields of the input item. To use them, refer to the input node, for example `$('Previous Node').item.json`.

When **On Error** is set to continue, an item with an error goes to the **False** output with an `error` field. The n8n If node does the same.

### Compare Structure

Compares a **Source Folder** (for example a template folder) with a **Target Folder** (for example a copy of it). Use it to find the folders that the copy does not have yet. The two folders can be in different libraries of the site.

Options:

- **Include**: Folders (default), Files, or Both.
- **Max Depth**: how many folder levels to compare. 1 (default) compares only the direct children. 0 compares all levels.
- **Case Sensitive** (default off): if on, `Folder` and `folder` are different names. SharePoint names are not case-sensitive.
- **Exclude Names**: names separated by commas or line breaks. `*` matches any text and `?` matches one character. A value with `/` matches the path below the compared folder, for example `Drafts/Old`. The node does not read the content of an excluded folder. Exclude Names applies to both folders.

The output has:

| Field          | Description                                                         |
| -------------- | ------------------------------------------------------------------- |
| `isComplete`   | `true` when the target has all entries of the source                |
| `missing`      | The relative paths that are in the source but not in the target     |
| `extra`        | The relative paths that are in the target but not in the source     |
| `matched`      | The number of source entries that the target has                    |
| `source`       | The server-relative path of the source folder                       |
| `target`       | The server-relative path of the target folder                       |
| `targetExists` | `false` when the target folder does not exist. Then all are missing |

Paths use `/`, for example `A/B`. A folder and a file with the same path do not match. Extra entries do not change `isComplete`.

The node reads each folder level with `list.renderListDataAsStream`, 500 rows for each page, so large folders work. The `Forms` folder of a library root is not in the comparison. If the source folder does not exist, the operation fails.

### Get Copy Job Status job info

**Job Info** is the JSON of a copy or move that returned before the job ended. By default, it is the input item. The job info has this shape:

```json
{
	"jobId": "00000000-0000-0000-0000-000000000000",
	"jobQueueUri": "https://example.queue.core.windows.net/example-queue?sv=...",
	"encryptionKey": "BASE64-KEY"
}
```

The output can also have `sourcePath` and `destinationPath` from the job info.

## Copy job versus Direct

File and Folder Copy and Move have a **Method** option.

- **Copy Job** (default) uses the SharePoint migration API. It scales to large content. A move removes the source in SharePoint.
- **Direct** runs the operation in one request. Use it for small content. The Direct method cannot replace a folder. Use the copy job for that.

A copy job uses the SharePoint migration API `CreateCopyJobs`. See [Microsoft 365 Copy and Move API (CreateCopyJobs)](https://learn.microsoft.com/sharepoint/dev/apis/spod-copy-move-api). The node also uses `GetCopyJobProgress` to read the job status. PnPjs does not wrap these methods, so the node calls them as REST requests through PnPjs.

Copy job behavior:

- **Name Conflict Behavior** is Fail, Keep Both, or Replace.
- **Wait for Completion** is on by default. The node polls the job until it ends or until **Max Wait (Seconds)** passes. The default maximum wait is 300 seconds.
- If the wait time ends, the node does not fail. It returns the job info with the status `running`. The copy continues in SharePoint.
- If you turn **Wait for Completion** off, the node returns at once. The output has `jobId`, `jobQueueUri` and `encryptionKey`.
- Use **File/Folder > Get Copy Job Status** to check a job later. It reads the job info from the input item by default. SharePoint returns each log entry once, so the summary and the messages show only the entries that are new since the last check.
- If the job reports errors, the node fails. Turn on **Continue on Job Errors** to get the job result instead. The errors are in the output in both cases.
- **Reset Role Inheritance After Move** and **Delete Empty Source Folder** run only after a job that finished without errors, and only when the node waits for the job.

The output has `status` (`running`, `completed` or `failed`), `finished`, `summary`, `errors` and `warnings`.

## Provisioning operations

These Folder operations replace groups of nodes. For example, a workflow makes a folder tree such as `Reports/2024/Q1`, tags the new folders, copies the folders of a template into the last folder, and pushes field values down to all content. Use Ensure Path for one branch of folders, and Ensure Folders for a tree with more than one branch. Each operation is safe to run again: a second run with the same input changes nothing.

Field names are internal names, for example `Title` or `Category`. The field values use the same format as **List Item > Update**: commas or an array for multi-choice values, ISO text for dates, IDs for lookup and person fields, and `Label|TermId` for managed metadata fields.

### Ensure Path

Walks a list of segments below the **Folder** (the base folder), creates each missing folder, and sets its fields. The base folder must exist. It can be a library root folder.

Each segment has:

- **Name**: the folder name. It must not have a `/`. Use one segment for each level.
- **Fields When Created (JSON)**: values to set only when this run creates the folder.
- **Fields Always (JSON)**: values to set on each run. The node reads the current values first and does not write a value that is already equal.
- **Append To Multi-Choice Field** and **Append Values**: choices to add to a multi-choice field. The current choices stay. The node writes the field only when a choice is missing.

Set **Segments Input** to **JSON** to give the segments as an array. An entry is a folder name or an object:

```json
[
	"Reports",
	{ "name": "2024", "fieldsWhenCreated": { "Title": "2024" } },
	{
		"name": "Q1",
		"fieldsWhenCreated": { "Title": "Q1" },
		"fieldsAlways": { "Status": "Active" },
		"appendToMultiChoice": { "field": "Tags", "values": ["Final"] }
	}
]
```

The output has `folderPath` (the last folder), `listItemId`, `created` (true when the run made at least one folder), `changed` and `segments`. Each segment entry has `name`, `serverRelativeUrl`, `listItemId`, `created` and `fieldsUpdated` (the names of the fields that the run wrote).

PnPjs: `folders.addUsingPath`, `item.validateUpdateListItem`.

### Ensure Folders

Makes a small tree of folders below the **Folder** (the base folder) in one step. For example, the folders `Reports/2024/Q1` and `Archive/2024`. The base folder must exist. It can be a library root folder.

Each folder row has:

- **Key**: a unique name for the folder, for example `year` or `archive`. The output has the folder under this key.
- **Path**: the path below the base folder. It can have more than one level, for example `Reports/2024/Q1`. Use `/` between levels.
- **Fields When Created (JSON)**, **Fields Always (JSON)**, **Append To Multi-Choice Field** and **Append Values**: the same as for [Ensure Path](#ensure-path). The node sets these fields only on the folder of the row, not on its parents.

The node makes parents before children. A parent that more than one row shares is made once, also when it has no row of its own. A folder that exists is not made again. Names are not case-sensitive, so `Reports` and `reports` are the same folder. Two rows must not have the same path.

Set **Folders Input** to **JSON** to give the folders as an array:

```json
[
	{ "key": "year", "path": "Reports/2024", "fieldsWhenCreated": { "Title": "2024" } },
	{
		"key": "quarter",
		"path": "Reports/2024/Q1",
		"fieldsAlways": { "Status": "Active" },
		"appendToMultiChoice": { "field": "Tags", "values": ["Final"] }
	},
	{ "key": "archive", "path": "Archive/2024" }
]
```

**Output** sets the shape of the output:

- **One Item With All Folders** (default): one item with `folders` (an object with one entry for each key), `created` (true when the run made at least one folder), `changed` and `baseFolder`.
- **One Item Per Folder**: one item for each row, with `key`, the folder fields below, and `baseFolder`.

Each folder entry has:

| Field               | Description                                                |
| ------------------- | ---------------------------------------------------------- |
| `path`              | The path below the base folder, as the row gave it         |
| `serverRelativeUrl` | For example `/sites/Example/Shared Documents/Reports/2024` |
| `siteRelativeUrl`   | For example `Shared Documents/Reports/2024`                |
| `absoluteUrl`       | The full URL, with the folder names URL-encoded            |
| `listItemId`        | The ID of the list item of the folder                      |
| `uniqueId`          | The unique ID of the folder                                |
| `created`           | True when this run made the folder                         |
| `fieldsUpdated`     | The names of the fields that the run wrote                 |

With the default output, use the key to refer to a folder in the next nodes:

```text
{{ $json.folders.quarter.serverRelativeUrl }}
{{ $json.folders.archive.absoluteUrl }}
{{ $json.folders.year.listItemId }}
{{ $('Ensure Folders').item.json.folders.quarter.serverRelativeUrl }}
```

With **One Item Per Folder**, each item has the fields at the top level, for example `{{ $json.key }}` and `{{ $json.serverRelativeUrl }}`.

PnPjs: `folders.addUsingPath`, `item.validateUpdateListItem`.

### Copy Children If Missing

Copies the files and subfolders of the **Folder** (for example a template folder) into the **Destination Folder**, but only the children that are not in the destination. Names are not case-sensitive. A subfolder is copied with all its content.

Options:

- **Include**: Folders (default), Files, or Both.
- **Include Names** and **Exclude Names**: names separated by commas or line breaks. `*` matches any text and `?` matches one character. For example `01 *, Images`. Exclude Names wins over Include Names.
- **Name Conflict Behavior**: what to do with a child that exists in the destination. Skip (default), Fail (the item fails before anything is copied), Replace, or Keep Both. The Direct method cannot replace a folder.
- **Copy Method**: Copy Job (default) or Direct. The copy job starts one job for each child, so the jobs run side by side, and waits for all of them up to **Max Wait (Seconds)** (default 300). A job that runs longer continues in SharePoint and has the status `running`. With Skip, the job uses the name conflict behavior Fail, so a child that appears in the destination during the run is not overwritten.
- **Reset Author and Created**: only for the Direct method. The copy job API (`CopyMigrationOptions`) has no such option, so a copy job keeps the author and the created date of the source.
- **Only If Destination Exists**: if on, a missing destination returns `changed: false` and `destinationExists: false`. If off, a missing destination fails.
- **Ignore Version History** and **Continue on Job Errors**: the same as for **Folder > Copy**.

The output has `copied` (with `destinationPath`, `status` and, for a copy job, `jobId`, `summary` and `errors`), `skipped` (with a `reason`: `exists`, `excluded`, `notIncluded` or `type`), and `changed` (true when the run copied at least one child).

The node does not list the `Forms` folder of a library root, or other folders without a list item.

PnPjs: `folder.folders`, `folder.files`, `CreateCopyJobs` (REST), `folder.copyByPath`, `file.copyByPath`.

### Set Metadata Recursive

Sets field values on all files and folders below the **Folder**. Give the fields as key-value pairs, or set **Fields Input** to **JSON**:

```json
{ "Category": "Report", "Department": "Team A", "Tags": ["Final", "Active"] }
```

The node reads the tree one folder at a time with `list.renderListDataAsStream`, 500 rows for each page. Each request is scoped to one folder, so a large library does not hit the list view threshold. Then it sends the updates (`item.validateUpdateListItem`) in SharePoint batches.

Options:

- **Include**: Both (default), Files, or Folders.
- **Include The Folder Itself**: also update the folder. A library root folder has no fields, so the node skips it.
- **Max Depth**: how many folder levels to go down. 1 updates only the direct children. 0 (default) has no limit.
- **Only Fields In Content Type**: set on each item only the fields that its content type has. The node reads the content types of the list and their fields once. An item without any of the fields is skipped. A field that the library does not have always fails the run before any update.
- **Ignore Empty Values**: leave out empty values, so that they do not clear current values.
- **Skip Unchanged** (on by default): skip a single line of text or choice value that is already equal. The node writes values of other field types each time.
- **System Update**: update without a new version, and keep Modified and Modified By. The node sends the current values back, the same as **List Item > Update** does. This needs one more batch request for each batch, to convert the dates to the time zone of the site. Modified By stays only for users. If an app last changed the item, Modified By can change.
- **Batch Size**: 1 to 100 updates in one batch request. The default is 100.
- **Continue On Item Error**: continue when an update fails. If off, the run stops after the batch with the first error. Items that were updated before stay updated.

The output has `matched` (the items found), `updated`, `skipped`, `failed` and `errors` (`listItemId`, `serverRelativeUrl` and `message` for each failed item).

### Inherit Metadata From Parent

Copies field values from the parent folder of the **Folder** to the folder itself and to all files and folders below it. Use it to push the values of a parent folder down to a new subfolder and its content.

Choose the fields in **Fields To Inherit**. The picker lists the fields of the library that contains the **Folder**. **Other Fields To Inherit** is a text field for more internal names, separated by commas or line breaks, for example `Category, Status`. The node uses the fields of the two parameters together. Workflows that were saved with only the text field keep working.

The node reads the fields of the parent folder. With the option **Search Up To Nearest Folder With Values**, it goes up one folder at a time until a folder has a value for each field. It reads each field from the nearest folder that has a value. The search stops below the library root, because a library root folder has no fields. The folder must not be a library root folder or a direct child of it.

A field that has no value in the source is empty, and it clears the current values. To keep the current values, turn on **Ignore Empty Values**. When no field has a value to write, the node changes nothing.

The options are the same as for [Set Metadata Recursive](#set-metadata-recursive), with two differences:

- **Include The Folder Itself** is on by default. The node updates the selected folder too. Turn it off to update only the items below the folder.
- **Search Up To Nearest Folder With Values** is only for this operation.

The node reads lookup fields as IDs, person fields as login names and link fields as `URL, description`, and writes them in the same format as **List Item > Update**. Managed metadata, calculated and computed fields are not supported. Dates are copied as UTC ISO text.

The output has the fields of Set Metadata Recursive (`matched`, `updated`, `skipped`, `failed`, `errors`) and:

| Field           | Description                                                                     |
| --------------- | ------------------------------------------------------------------------------- |
| `parentFolder`  | The parent folder, where the search starts                                      |
| `sourceFolder`  | The folder that gave the values. With values from more folders, the nearest one |
| `values`        | The values that the node used, keyed by field name                              |
| `valueSources`  | For each field, the folder that gave the value, or `null`                       |
| `missingFields` | The fields that had no value in the source                                      |

This is a separate operation, not an option of Set Metadata Recursive. The two operations have different inputs (field names instead of field values) and a different default for **Include The Folder Itself**. As separate operations, each one shows only the parameters that it uses.

PnPjs: `folder.listItemAllFields`, `list.renderListDataAsStream`, `item.validateUpdateListItem`.

## Limits and known behavior

- **File Upload** replaces a file with the same name by default. Turn off **Overwrite** to fail instead.
- **File Upload** fails when the destination folder does not exist. Turn on **Create Missing Folders** to create it.
- **Check Exists** has two outputs. If you change the operation of a node that has connections, check the connections again.
- **Publish, Unpublish, Approve and Deny** comments have a limit of 1023 characters. Publish and Unpublish need a library with minor versions or content approval. Approve and Deny need content approval.
- **Copy job**: **Reset Role Inheritance After Move** and **Delete Empty Source Folder** do not run when the node does not wait for the job.
