# Changelog

## Unreleased

- Project foundation: OAuth2 (delegated) and certificate (app-only) credentials, PnPjs transport through n8n authentication, and the List resource.
- Change resource Get Many and the SharePoint PnP Trigger: the "On Expired Token" option sets what to do when a change token is older than the change log (60 days in SharePoint Online). Start from the oldest available change (default), start from now, or fail. Changes after an expired token have `tokenExpired: true`.
- SharePoint PnP Webhook Trigger: starts the workflow from SharePoint list webhook notifications. It creates, renews and deletes the list subscription, checks the client state, and reads the changes with the same events, folder path filter and options as the polling trigger.
- Folder and file copy and move use the Copy Job method by default. The Direct method stays as a choice.
- Copy Job: all options of the SharePoint copy job API, waiting with a maximum wait time, a job summary, and the options Reset Role Inheritance After Move and Delete Empty Source Folder.
- New operation File/Folder > Get Copy Job Status to check a job that runs in SharePoint.
