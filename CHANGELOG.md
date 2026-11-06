# Changelog

## Unreleased

- Project foundation: OAuth2 (delegated) and certificate (app-only) credentials, PnPjs transport through n8n authentication, and the List resource.
- Site resource: get, subsites, property bag, create site and subsite, site creation status, hub site join and leave, hub sites.
- List setup: update list settings, create, update and delete fields, and manage content types of a list.
- Change resource Get Many and the SharePoint PnP Trigger: the "On Expired Token" option sets what to do when a change token is older than the change log (60 days in SharePoint Online). Start from the oldest available change (default), start from now, or fail. Changes after an expired token have `tokenExpired: true`.
- SharePoint PnP Webhook Trigger: starts the workflow from SharePoint list webhook notifications. It creates, renews and deletes the list subscription, checks the client state, and reads the changes with the same events, folder path filter and options as the polling trigger.
- Folder and file copy and move use the Copy Job method by default. The Direct method stays as a choice.
- Copy Job: all options of the SharePoint copy job API, waiting with a maximum wait time, a job summary, and the options Reset Role Inheritance After Move and Delete Empty Source Folder.
- New operation File/Folder > Get Copy Job Status to check a job that runs in SharePoint.
- Property selection for older resources: File Get and Get Versions, Folder Get, File/Folder Get Many, List Get, Get Many, Get Fields and Get Views, Permission Get Many, Group Get, Get Many and Get Members, User Get and Get Many, and Recycle Bin Get Many now have Additional Properties, Return All Properties and (where it applies) Expand. The default properties are the ones the output had before. Include List Item Fields still works.
- Permission Add, Remove and Set have an Input Mode (Fields or JSON). In JSON mode, Role Assignments (JSON) takes many principals and roles in one array.
- New File operations: Publish, Unpublish, Approve and Deny, each with a comment.
- Taxonomy resource (managed metadata) through the SharePoint CSOM endpoint: term groups, term sets, terms, search by label, and create, update and delete term. List Item Create and Update accept term labels, term IDs or "Label|ID" for managed metadata fields, with the option "Create Missing Terms".
