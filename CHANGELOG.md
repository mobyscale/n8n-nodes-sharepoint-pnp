# Changelog

## Unreleased

- Project foundation: OAuth2 (delegated) and certificate (app-only) credentials, PnPjs transport through n8n authentication, and the List resource.
- Change resource Get Many and the SharePoint PnP Trigger: the "On Expired Token" option sets what to do when a change token is older than the change log (60 days in SharePoint Online). Start from the oldest available change (default), start from now, or fail. Changes after an expired token have `tokenExpired: true`.
