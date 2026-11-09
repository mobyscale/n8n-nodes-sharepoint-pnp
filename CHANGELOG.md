# Changelog

## Unreleased

### Package

- OAuth2 (delegated) and certificate (app-only) credentials. The OAuth2 credential has a Tenant ID option for single-tenant apps.
- PnPjs transport that uses the n8n authentication.
- Property selection contract for all Get and Get Many operations: documented default properties, Additional Properties, Expand and Return All Properties. Include List Item Fields still works for older workflows.

### SharePoint PnP

#### Site

- Get, Get Subsites, Get Property Bag, Set Property Bag Values.
- Create Site, Create Subsite, Get Site Creation Status.
- Get Hub Sites, Join Hub Site, Leave Hub Site.

#### File

- Upload, Download, Get, Get Versions, Rename, Delete, Check Out, Check In, Discard Check Out, Restore Version.
- Copy and Move use the Copy Job method by default. The Direct method stays as a choice.
- Publish, Unpublish, Approve and Deny, each with a comment.
- Upload options: Create Missing Folders, Fields and Fields (JSON) (set in one request, without a new version), Check In and Publish, each with a comment. Check In and Publish run only when they apply.
- Get has the option Include List Item Fields.
- The File picker has a From List mode that lists the folders and files of a typed path.
- Restore Version has a picker with the earlier versions of the file.

#### Folder

- Create, Get, Rename, Delete.
- Get has the option Include List Item Fields.
- Copy and Move use the Copy Job method by default. The Direct method stays as a choice.
- Copy Job options: all options of the SharePoint copy job API, waiting with a maximum wait time, a job summary, Reset Role Inheritance After Move, and Delete Empty Source Folder.
- Ensure Path: create the missing folders of a path, level by level. Each level can set fields only when it is created, set fields on each run (no write when the value is equal), and add choices to a multi-choice field.
- Copy Children If Missing: copy the children of a template folder into a destination when they are not there. Include and exclude names with `*` wildcards, a filter for files or folders, the copy job (default) or the Direct method, and a skip when the destination is missing.
- Set Metadata Recursive: set field values on all files and folders below a folder in batch requests. Options for depth, content type fields, empty values, unchanged values, system update, batch size and item errors.
- Ensure Folders: make a small folder tree (for example team, year and quarter folders with their "Drafts" folders) in one step. Each folder has a key, a path that can have more than one level, and the same field options as Ensure Path. Shared parents are made once. The output has one item with all folders by key (`$json.folders.quarter.serverRelativeUrl`), or one item for each folder.
- Inherit Metadata From Parent: copy field values from the parent folder, or from the nearest folder above it with values, to the folder and all files and folders below it. It uses the options of Set Metadata Recursive.
- Inherit Metadata From Parent: Fields To Inherit is a picker with the fields of the library that contains the folder. Other Fields To Inherit keeps the comma-separated text of older workflows.

#### File/Folder

- Get Many, Check Exists.
- Check Exists has a True and a False output, like the If node. The Combine option (All or Any) sets when the input item goes to True. The output item has `exists`, `allExist`, `anyExists` and one entry in `results` for each path. With Continue On Fail, error items go to False.
- Get Copy Job Status: check a copy or move job that runs in SharePoint. The documentation shows the job info shape.
- Compare Structure: compare a source folder (for example a template) with a target folder. Options: Include (Folders, Files or Both), Max Depth, Case Sensitive and Exclude Names with wildcards. The output has `isComplete`, `missing`, `extra` and `matched`. Folders are read in pages, so large folders work.
- Get Many: the options Min Depth and Max Depth (with Recursive) set the folder levels to return. Max Depth stops the walk. With Recursive, each entry has `depth` (1 for the direct children).
- Get Many: Field Filters (Equals, Not Equals, Contains, Begins With, Greater Than, Less Than, Is Empty, Is Not Empty) on list item fields, combined with AND or OR. The field picker lists the fields of the library that contains the folder. Each folder is queried on its own.

#### List

- Get, Get Many, Create, Update (list settings), Delete.
- Create If Missing (`ensure`): get a list or library by URL or title, and create it if it is missing. Options: Quick Launch, versioning, content types, the default content type, a view copied from another list, and permissions (break inheritance, and role assignments as JSON in Replace or Add mode). Apply Settings to Existing List (default off) also applies the options to a list that exists and writes only changed values. The output has `created` and a `changes` list.
- Get Fields, Create Field, Update Field, Delete Field.
- Get Views, Create View, Copy View. Copy View also accepts a view title.
- Get Content Types, Add Content Type to List, Remove Content Type From List, Set Default Content Type.
- Get Default Column Values, Set Default Column Values.
- Pickers for default column value fields and folders, the lookup field of a Lookup column, and the column group.

#### List Item

- Get, Get Many, Create, Update, Delete.
- Pickers: Item has a From List mode, Select Fields lists the fields of the list, and the Folder options list the folders of the list.
- Get Many: Field Filters with a field picker, combined with AND or OR. In filter mode they are added to the OData filter. In CAML mode they are added to the Where element of the query.
- Create and Update accept term labels, term IDs or "Label|ID" for managed metadata fields. The option Create Missing Terms creates unknown terms.
- Create Many, Update Many and Delete Many: process all input items of a node run in SharePoint `$batch` requests of up to 100 requests. Options: Batch Size, Continue on Item Error, System Update (Update Many), Delete Permanently (Delete Many).
- Create or Update (upsert): find an item by a key field value, then update it or create it. Options: If Several Match, Skip Update If Unchanged, System Update, Find in Folder (a CAML lookup in one folder, for lists above the list view threshold). The output has `action` set to `created`, `updated` or `unchanged`.
- The node can run an operation for all input items in one call (bulk operations). The output keeps one item for each input item, with `pairedItem` and On Error (Continue) for each input item.

#### Attachment

- Add, Get Many, Download, Update Content, Delete.
- Item and Attachment Name have a From List mode.

#### Permission

- Get Many, Get Effective Permissions, Add, Remove, Set, Break Inheritance, Reset Inheritance.
- Pickers for users, principals (groups and users), roles and list items. Group, User and Site owner fields use the same pickers.
- Set and Add options: Clear Subscopes, Always Break Inheritance (Set) and Skip Unresolved Principals, with a `warnings` array in the output.
- Add, Remove and Set have an Input Mode (Fields or JSON). In JSON mode, Role Assignments (JSON) takes many principals and roles in one array.
- Apply Mapping: apply a map of paths, principals and roles to the folders below a base folder. `*` path segments match existing child folders without case. All principals and roles resolve one time before any change. Modes Add and Set (default). Options: Break Inheritance (Always or Only If Inheriting), Copy Current Assignments First, Clear Subscopes, Create Missing Folders, If Folder Is Missing (Skip or Fail), Skip Unresolved Principals, and One Item Per Folder output. The output has `folders`, `warnings` and `missing`.
- Set sends more than 100 role assignment changes in more than one `$batch` request.

#### Sharing

- Create Link, Delete Link, Get Sharing Information, Share With Users, Check Permissions, Unshare.

#### Group and User

- Group: Get, Get Many, Create, Update, Delete, Get Members, Add Member, Remove Member.
- User: Get, Get Many, Ensure.

#### Search

- Query (KQL) and Suggest.

#### Taxonomy

- Managed metadata through the SharePoint CSOM endpoint.
- Get Term Groups, Get Term Sets, Get Term Set, Get Terms, Search Terms, Get Term, Create Term, Update Term, Delete Term.

#### Change

- Get Current Token and Get Many.
- The option On Expired Token sets what to do when a change token is older than the change log (60 days in SharePoint Online). Start from the oldest available change (default), start from now, or fail. Changes after an expired token have `tokenExpired: true`.
- Include Item Details reads the details in bulk: one CAML query for each list and each 500 changed items (it works above the list view threshold), and `$batch` requests of 100 for files, folders and items with expanded fields. Before, it sent one request for each changed item. The output is the same.
- The recycle bin lookup for deleted items asks for 20 items in one request, or reads the recycle bin one time for the run.
- Include Item Details adds `editorEmail` when it is known: the editor email hint of the change, or the only site user with the editor display name.
- New options: Item Kind (Files, Folders, List Items), Max Depth, Exclude Editors and Field Filters (Equals, Not Equals, Contains, Is Empty, Is Not Empty). Field Filters need the item details. The triggers have all of them except Item Kind.

#### Recycle Bin

- Get Many, Restore, Move to Second Stage, Delete Permanently.

#### Pickers (Site, Taxonomy, Recycle Bin, Change)

- Site: lists for Hub Site, Site Design, Language, Property Keys and the property bag Key. Site to Check has From List and By URL.
- Taxonomy: Term and Parent Term pickers, and a language list for labels.
- Recycle Bin: a multi-select for the item IDs.
- Change and triggers: a multi-select for Item Fields and a folder picker for Folder Path Filter.
- Older workflows with plain text values keep working.
- From List pickers for sites, users, groups, principals, group members, folders, files, file versions, terms and items load more results when you scroll. The item search above the list view threshold scans 300 older items for each page.
- Dropdown and multi-select pickers load at most 500 entries, sorted A-Z by name. The recycle bin picker sorts A-Z instead of newest first.

### SharePoint PnP Trigger

- Polls the change log of a site collection, a site or a list.
- Has the On Expired Token option.
- Has the Max Depth, Exclude Editors and Field Filters options. It reads the item details in bulk.

### SharePoint PnP Webhook Trigger

- Starts the workflow from SharePoint list webhook notifications.
- Creates, renews and deletes the list subscription, and checks the client state.
- Reads the changes with the same events, folder path filter and options as the polling trigger.

### Documentation

- Full README with credentials, nodes, common concepts, limits and development notes.
- Manual test checklist in `docs/manual-test-checklist.md`.
