# Changelog

## Unreleased

### Package

- OAuth2 (delegated) and certificate (app-only) credentials.
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

#### Folder

- Create, Get, Rename, Delete.
- Copy and Move use the Copy Job method by default. The Direct method stays as a choice.
- Copy Job options: all options of the SharePoint copy job API, waiting with a maximum wait time, a job summary, Reset Role Inheritance After Move, and Delete Empty Source Folder.

#### File/Folder

- Get Many, Check Exists.
- Get Copy Job Status: check a copy or move job that runs in SharePoint.

#### List

- Get, Get Many, Create, Update (list settings), Delete.
- Get Fields, Create Field, Update Field, Delete Field.
- Get Views, Create View, Copy View.
- Get Content Types, Add Content Type to List, Remove Content Type From List, Set Default Content Type.
- Get Default Column Values, Set Default Column Values.

#### List Item

- Get, Get Many, Create, Update, Delete.
- Create and Update accept term labels, term IDs or "Label|ID" for managed metadata fields. The option Create Missing Terms creates unknown terms.

#### Attachment

- Add, Get Many, Download, Update Content, Delete.

#### Permission

- Get Many, Get Effective Permissions, Add, Remove, Set, Break Inheritance, Reset Inheritance.
- Add, Remove and Set have an Input Mode (Fields or JSON). In JSON mode, Role Assignments (JSON) takes many principals and roles in one array.

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

#### Recycle Bin

- Get Many, Restore, Move to Second Stage, Delete Permanently.

### SharePoint PnP Trigger

- Polls the change log of a site collection, a site or a list.
- Has the On Expired Token option.

### SharePoint PnP Webhook Trigger

- Starts the workflow from SharePoint list webhook notifications.
- Creates, renews and deletes the list subscription, and checks the client state.
- Reads the changes with the same events, folder path filter and options as the polling trigger.

### Documentation

- Full README with credentials, nodes, common concepts, limits and development notes.
- Manual test checklist in `docs/manual-test-checklist.md`.
