# n8n-nodes-sharepoint-pnp

n8n community nodes for SharePoint Online. The nodes wrap [PnPjs](https://pnp.github.io/pnpjs/) (`@pnp/sp`). PnPjs is the community library for the SharePoint REST API. See the [PnPjs repository](https://github.com/pnp/pnpjs) for the source.

Use these nodes for operations that the built-in Microsoft SharePoint node does not have. Examples are permissions, sharing links, the change log, search, the recycle bin, managed metadata, and folder copy and move with the SharePoint copy job.

These nodes use the SharePoint REST API. They do not use Microsoft Graph.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

## Contents

- [Installation](#installation)
- [Credentials](#credentials)
- [Nodes](#nodes)
- [Common concepts](#common-concepts)
- [Limits and known behavior](#limits-and-known-behavior)
- [Compatibility](#compatibility)
- [Development](#development)
- [Resources and license](#resources-and-license)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation. The package name is `n8n-nodes-sharepoint-pnp`.

PnPjs is bundled into the nodes when the package is built. The package has no runtime dependencies.

## Credentials

The package has two credentials. Each node has an **Authentication** option to choose one of them.

### OAuth2 (delegated)

The **SharePoint PnP OAuth2 API** credential signs in as a user. The node can do what that user can do.

1. Register an app in Microsoft Entra ID.
2. Add a **Web** redirect URI. Use the OAuth redirect URL that n8n shows in the credential.
3. Add delegated permissions for SharePoint, for example `AllSites.FullControl`. Give admin consent if your tenant needs it.
4. Create a client secret.
5. In n8n, create the credential. Enter the client ID and the client secret.
6. Enter the **Subdomain**. For `https://contoso.sharepoint.com`, enter `contoso`.
7. Select **Connect** and sign in.

The credential asks for the scope `https://<subdomain>.sharepoint.com/.default`. The permissions of the app registration decide what the token can do.

### Certificate (app-only)

The **SharePoint PnP Certificate API** credential signs in as the app. No user is involved.

SharePoint REST does not accept app-only tokens from a client secret. A certificate is necessary. The node signs a client assertion with the private key and sends it to Entra to get a token.

1. Register an app in Microsoft Entra ID.
2. Add the SharePoint **application** permission `Sites.FullControl.All` or `Sites.Selected`. Give admin consent. With `Sites.Selected`, grant the app access to each site that it must use.
3. Upload the public certificate to the app on the **Certificates & secrets** page.
4. In n8n, create the credential. Enter the tenant ID, the client ID, and the subdomain.
5. Paste the PEM private key. If the key is encrypted, enter the passphrase.
6. Give the certificate in one of these ways:
   - Paste the certificate PEM in the **Certificate** field. You can also paste it after the key in the **Private Key** field.
   - Enter the SHA-1 thumbprint of the certificate in **Certificate Thumbprint**.

To get PEM files from a `.pfx` file, use these `openssl` commands:

```bash
openssl pkcs12 -in certificate.pfx -nocerts -nodes -out private-key.pem
openssl pkcs12 -in certificate.pfx -clcerts -nokeys -out certificate.pem
```

The first command writes an unencrypted private key. Keep that file secret. Delete it when you no longer need it.

### Which feature needs which rights

| Feature                     | What it needs                                                                                                                                                          |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Taxonomy (term store)       | The node calls SharePoint CSOM. The REST term store API does not accept the app-only SharePoint token, but CSOM does. The identity must have access to the term store. |
| Create Site                 | With app-only authentication, SharePoint needs an owner. Set the **Owner** option to the user principal name of the owner, for example `user@contoso.com`.             |
| Join Hub Site               | The identity needs site collection admin rights on the site.                                                                                                           |
| Set Property Bag Values     | SharePoint can need "custom scripts" to be allowed on the site. If it is not, SharePoint returns an access denied error.                                               |
| Change, deleted item lookup | The node reads the recycle bin of the site collection. If the identity cannot read it, the node reads the recycle bin of the site.                                     |
| Sharing links               | A tenant or site sharing policy can block a link. For example, the policy can block Anyone links, or can require an expiration date.                                   |
| Webhook Trigger             | SharePoint must reach n8n from the internet over HTTPS. See [Webhook requirements](#webhook-requirements).                                                             |

## Nodes

The package has three nodes.

### SharePoint PnP

The main node has a **Resource** and an **Operation** option. Use the **Site** option to choose the site. The site picker has a search list and a **By URL** mode. Most operations use PnPjs directly. The sections below say when a resource uses another SharePoint API.

The node can be used as a tool by an AI agent (`usableAsTool`). It supports **Continue On Fail**.

#### Site

PnPjs: [Webs](https://pnp.github.io/pnpjs/sp/webs/), [Sites](https://pnp.github.io/pnpjs/sp/sites/), [Hub sites](https://pnp.github.io/pnpjs/sp/hubsites/).

Create Site and Get Site Creation Status use the SharePoint `SPSiteManager` REST API. See [Manage modern SharePoint sites using REST](https://learn.microsoft.com/sharepoint/dev/apis/site-creation-rest).

| Operation                | Description                                                                   | PnPjs reference                     |
| ------------------------ | ----------------------------------------------------------------------------- | ----------------------------------- |
| Get                      | Get the selected site, with its site collection information                   | `web`                               |
| Get Subsites             | Get the subsites of the selected site                                         | `web.webs`                          |
| Get Property Bag         | Get all property bag values of the site, or only the given keys               | `web.allProperties`                 |
| Set Property Bag Values  | Set property bag values of the site                                           | `web.allProperties`                 |
| Create Site              | Create a modern Communication site or Team site without a Microsoft 365 group | `SPSiteManager/create` (REST)       |
| Create Subsite           | Create a classic subsite in the selected site                                 | `web.webs.add`                      |
| Get Site Creation Status | Get the creation status of a site                                             | `SPSiteManager/status` (REST)       |
| Get Hub Sites            | Get the hub sites of the tenant                                               | `hubSites`                          |
| Join Hub Site            | Associate the selected site with a hub site                                   | `site.joinHubSite`                  |
| Leave Hub Site           | Remove the association of the selected site with its hub site                 | `site.joinHubSite` with an empty ID |

#### File

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

#### Folder

PnPjs: [Folders](https://pnp.github.io/pnpjs/sp/folders/). Copy and Move use the [copy job](#copy-job-versus-direct) by default.

| Operation | Description                                    | PnPjs reference                             |
| --------- | ---------------------------------------------- | ------------------------------------------- |
| Create    | Create a folder and all missing parent folders | `folders.addUsingPath`                      |
| Get       | Get the properties of a folder                 | `folder`                                    |
| Copy      | Copy a folder                                  | `folder.copyByPath`, or a copy job          |
| Move      | Move a folder                                  | `folder.moveByPath`, or a copy job          |
| Rename    | Rename a folder                                | `folder.moveByPath`                         |
| Delete    | Delete a folder, or move it to the recycle bin | `folder.recycle`, `folder.deleteWithParams` |

#### File/Folder

PnPjs: [Files](https://pnp.github.io/pnpjs/sp/files/) and [Folders](https://pnp.github.io/pnpjs/sp/folders/). Get Copy Job Status uses the SharePoint migration API. See [Copy job](#copy-job-versus-direct).

| Operation           | Description                                                    | PnPjs reference               |
| ------------------- | -------------------------------------------------------------- | ----------------------------- |
| Get Many            | Get many files and folders in a folder                         | `list.renderListDataAsStream` |
| Check Exists        | Check if files or folders exist                                | `file`, `folder`              |
| Get Copy Job Status | Get the progress of a copy or move job that runs in SharePoint | `GetCopyJobProgress` (REST)   |

#### List

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

#### List Item

PnPjs: [Items](https://pnp.github.io/pnpjs/sp/items/).

| Operation | Description                                              | PnPjs reference                                     |
| --------- | -------------------------------------------------------- | --------------------------------------------------- |
| Get       | Get a list item                                          | `items.getById`                                     |
| Get Many  | Get many list items. Use an OData filter or a CAML query | `items`, `list.renderListDataAsStream`              |
| Create    | Create a list item or a folder                           | `items.add`, `items.addValidateUpdateItemUsingPath` |
| Update    | Update the fields of a list item. Can do a system update | `item.update`, `item.validateUpdateListItem`        |
| Delete    | Delete a list item, or move it to the recycle bin        | `item.recycle`, `item.delete`                       |

The fields of Create and Update use a resource mapper. It loads the fields of the list. You can also give the fields as JSON. Managed metadata (tag) fields accept term labels. See [Managed metadata values](#managed-metadata-tag-values).

#### Attachment

PnPjs: [Attachments](https://pnp.github.io/pnpjs/sp/attachments/). Document libraries do not support attachments.

| Operation      | Description                                                  | PnPjs reference                           |
| -------------- | ------------------------------------------------------------ | ----------------------------------------- |
| Add            | Add an attachment to a list item                             | `item.attachmentFiles.add`                |
| Get Many       | Get the attachments of a list item                           | `item.attachmentFiles`                    |
| Download       | Download an attachment as binary data                        | `attachment.getBuffer`                    |
| Update Content | Replace the content of an attachment                         | `attachment.setContent`                   |
| Delete         | Move an attachment to the recycle bin, or delete it for good | `attachment.recycle`, `attachment.delete` |

#### Permission

PnPjs: [Security](https://pnp.github.io/pnpjs/sp/security/). The target can be a site, a list, a folder, a file or a list item.

| Operation                 | Description                                                 | PnPjs reference                                              |
| ------------------------- | ----------------------------------------------------------- | ------------------------------------------------------------ |
| Get Many                  | Get the role assignments of a target                        | `roleAssignments`                                            |
| Get Effective Permissions | Get the permissions that a user has on a target             | `getUserEffectivePermissions`                                |
| Add                       | Give a group or user a role on a target                     | `roleAssignments.add`                                        |
| Remove                    | Remove a role or all roles of a group or user from a target | `roleAssignments.remove`, `roleAssignments.getById().delete` |
| Set                       | Replace all role assignments of a target with a given list  | `roleAssignments`                                            |
| Break Inheritance         | Stop a target from inheriting permissions from its parent   | `breakRoleInheritance`                                       |
| Reset Inheritance         | Make a target inherit permissions from its parent again     | `resetRoleInheritance`                                       |

#### Sharing

PnPjs: [Sharing](https://pnp.github.io/pnpjs/sp/sharing/).

| Operation               | Description                                             | PnPjs reference                          |
| ----------------------- | ------------------------------------------------------- | ---------------------------------------- |
| Create Link             | Create or get a sharing link                            | `getShareLink`                           |
| Delete Link             | Delete a sharing link by its kind or share ID           | `deleteSharingLinkByKind`, `unshareLink` |
| Get Sharing Information | Get the sharing links and the people with access        | `getSharingInformation`                  |
| Share With Users        | Give users access, with an optional email               | `SP.Web.ShareObject` (REST)              |
| Check Permissions       | Check which role users have on a target through sharing | `checkSharingPermissions`                |
| Unshare                 | Remove all sharing from a target                        | `unshare`                                |

#### Group

PnPjs: [Site groups](https://pnp.github.io/pnpjs/sp/site-groups/).

| Operation     | Description                            | PnPjs reference               |
| ------------- | -------------------------------------- | ----------------------------- |
| Get           | Get a SharePoint group                 | `siteGroups.getById`          |
| Get Many      | Get many SharePoint groups of the site | `siteGroups`                  |
| Create        | Create a SharePoint group              | `siteGroups.add`              |
| Update        | Update the properties of a group       | `siteGroups.getById().update` |
| Delete        | Delete a SharePoint group              | `siteGroups.removeById`       |
| Get Members   | Get the members of a group             | `group.users`                 |
| Add Member    | Add a user to a group                  | `group.users.add`             |
| Remove Member | Remove a user from a group             | `group.users.removeById`      |

#### User

PnPjs: [Site users](https://pnp.github.io/pnpjs/sp/site-users/).

| Operation | Description                                                 | PnPjs reference                                     |
| --------- | ----------------------------------------------------------- | --------------------------------------------------- |
| Get       | Get a user of the site by ID, email address or login name   | `siteUsers.getById`, `getByEmail`, `getByLoginName` |
| Get Many  | Get many users of the site                                  | `siteUsers`                                         |
| Ensure    | Get a user and add the user to the site user list if needed | `web.ensureUser`                                    |

#### Search

PnPjs: [Search](https://pnp.github.io/pnpjs/sp/search/).

| Operation | Description                      | PnPjs reference    |
| --------- | -------------------------------- | ------------------ |
| Query     | Run a KQL search query           | `sp.search`        |
| Suggest   | Get query suggestions for a text | `sp.searchSuggest` |

#### Taxonomy

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

#### Change

The change log uses the SharePoint `GetChanges` method with a `ChangeQuery`. PnPjs calls it for sites, site collections and lists. For a folder filter, the node calls `GetChanges` on the folder. See [Query SharePoint change log with ChangeQuery and ChangeToken](https://learn.microsoft.com/sharepoint/dev/solution-guidance/query-sharepoint-change-log-with-changequery-and-changetoken).

| Operation         | Description                                                     | API                                |
| ----------------- | --------------------------------------------------------------- | ---------------------------------- |
| Get Current Token | Get the current change token of a site collection, site or list | `GetChanges`, `CurrentChangeToken` |
| Get Many          | Get many changes from the change log                            | `GetChanges`                       |

#### Recycle Bin

PnPjs: [Recycle bin](https://pnp.github.io/pnpjs/sp/recycle-bin/).

| Operation            | Description                                            | PnPjs reference                                                  |
| -------------------- | ------------------------------------------------------ | ---------------------------------------------------------------- |
| Get Many             | Get many recycle bin items                             | `site.recycleBin`, `web.recycleBin`, `getrecyclebinitems` (REST) |
| Restore              | Restore recycle bin items to their original location   | `recycleBin.getById().restore`                                   |
| Move to Second Stage | Move recycle bin items to the second stage recycle bin | `recycleBin.getById().moveToSecondStage`                         |
| Delete Permanently   | Delete recycle bin items for good                      | `recycleBin.getById().delete`                                    |

#### Copy jobs

The Folder and File resources can run a copy or move as a copy job. A copy job uses the SharePoint migration API `CreateCopyJobs`. See [Microsoft 365 Copy and Move API (CreateCopyJobs)](https://learn.microsoft.com/sharepoint/dev/apis/spod-copy-move-api). The node also uses `GetCopyJobProgress` to read the job status. PnPjs does not wrap these methods, so the node calls them as REST requests through PnPjs.

#### Subscriptions

List webhook subscriptions are used by the Webhook Trigger. The node uses the PnPjs [subscriptions](https://pnp.github.io/pnpjs/sp/subscriptions/) module. The main node has no subscription operations.

### SharePoint PnP Trigger (polling)

The trigger polls the change log of a site collection, a site or a list. It works without a public URL.

Options:

- **Scope**: Site Collection, Site, or List.
- **Events**: File Created, File Updated, File Deleted, Folder Created, Folder Deleted, Item Created, Item Updated, Item Deleted, Moved, Permissions Changed, Renamed, Restored.
- **Folder Path Filter**: only trigger for content in this folder and its subfolders.
- **Options**: Start From Date, Fetch Limit, Include Item Details, Item Fields, On Expired Token, Resolve Deleted Items From Recycle Bin.

How it works:

- When you activate the workflow, the first poll saves the current change token. It returns no items. Only changes after that poll start the workflow. If you set **Start From Date**, the first poll returns the changes after that date.
- Each next poll reads the changes after the saved token. Then it saves the new token.
- In manual mode ("Fetch Test Event"), the node shows the latest matching changes of the last 24 hours. It shows at most 10 changes. It does not save a token.
- If you change the site, scope or list, the node starts again with a new token.

See [The change token contract](#the-change-token-contract) for expired tokens.

### SharePoint PnP Webhook Trigger

The trigger uses SharePoint list webhooks. SharePoint sends a notification when a list or library changes. The node then reads the changes since the last change token. It applies the events and the folder path filter. Then it starts the workflow for the matching changes.

It has the same **Events**, **Folder Path Filter** and **Options** as the polling trigger, except **Start From Date**. The node saves a token at activation, so it does not need a start date.

See [Webhook requirements](#webhook-requirements) and [Polling or webhook](#polling-or-webhook).

## Common concepts

### Paths

Every path option accepts three forms:

| Form            | Example                                                                |
| --------------- | ---------------------------------------------------------------------- |
| Site-relative   | `Shared Documents/Folder`                                              |
| Server-relative | `/sites/Example/Shared Documents/Folder`                               |
| Absolute URL    | `https://contoso.sharepoint.com/sites/Example/Shared Documents/Folder` |

A site-relative path starts at the selected site. The picker in "From List" mode works as a path browser. Type `Documents/Clients/Ac` to list the children of `Documents/Clients` that start with `Ac`.

The site option accepts an absolute site URL, or a server-relative path such as `/sites/Example`.

### Property selection

Get and Get Many operations use one contract:

- Each operation loads a documented set of default properties. The option description lists them.
- **Additional Properties** adds more properties to the default set. Give a comma-separated list. Use `Navigation/Property` for a related object, for example `Author/Title`. The node adds `Navigation` to the expand list for you.
- **Expand** loads related objects.
- **Return All Properties** sends no `$select`. SharePoint then returns its default entity.

The default properties are stable. They are the properties that the output had before the option was added.

### Copy job versus Direct

File and Folder Copy and Move have a **Method** option.

- **Copy Job** (default) uses the SharePoint migration API. It scales to large content. A move removes the source in SharePoint.
- **Direct** runs the operation in one request. Use it for small content. The Direct method cannot replace a folder. Use the copy job for that.

Copy job behavior:

- **Name Conflict Behavior** is Fail, Keep Both, or Replace.
- **Wait for Completion** is on by default. The node polls the job until it ends or until **Max Wait (Seconds)** passes. The default maximum wait is 300 seconds.
- If the wait time ends, the node does not fail. It returns the job info with the status `running`. The copy continues in SharePoint.
- If you turn **Wait for Completion** off, the node returns at once. The output has `jobId`, `jobQueueUri` and `encryptionKey`.
- Use **File/Folder > Get Copy Job Status** to check a job later. It reads the job info from the input item by default. SharePoint returns each log entry once, so the summary and the messages show only the entries that are new since the last check.
- If the job reports errors, the node fails. Turn on **Continue on Job Errors** to get the job result instead. The errors are in the output in both cases.
- **Reset Role Inheritance After Move** and **Delete Empty Source Folder** run only after a job that finished without errors, and only when the node waits for the job.

The output has `status` (`running`, `completed` or `failed`), `finished`, `summary`, `errors` and `warnings`.

### Permission JSON input

Permission Add, Remove and Set have an **Input Mode**: **Fields** or **JSON**. JSON mode changes many principals and roles in one call. **Role Assignments (JSON)** is an array of objects.

| Property        | Description                                                                              |
| --------------- | ---------------------------------------------------------------------------------------- |
| `principal`     | Required. A group name, a group or user ID, or the email address or login name of a user |
| `principalType` | Optional. `user` or `group`. Use it when the name is not clear                           |
| `role`          | One role name or ID                                                                      |
| `roles`         | A list of role names or IDs. Do not use it together with `role`                          |

Example:

```json
[
	{ "principal": "Example Members", "principalType": "group", "role": "Contribute" },
	{ "principal": "user@contoso.com", "roles": ["Read", "Edit"] }
]
```

Rules:

- The node resolves all names to IDs first. It changes nothing in SharePoint if one entry cannot be resolved.
- For Remove, an entry without a role removes all roles of the principal.
- For Set, an empty array removes all role assignments.
- Entries for the same principal are merged. An entry without roles wins.
- An unknown property is an error.

### Managed metadata (tag) values

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

- The node finds the term set from the field. It looks up labels through the term store.
- If a label does not exist, the operation fails. Turn on **Create Missing Terms** to create it. The term set must be open for new terms.
- **Append to Multi-Value Fields** (Update) adds the new values to the current values.
- Default column values do not support taxonomy fields.

### The change token contract

A change token marks a point in the SharePoint change log. The Change resource and the triggers use it.

- **Get Current Token** returns the token for the scope.
- **Get Many** returns `nextChangeToken` when you set **Output** to **Single Item**. Save it and use it as the start of the next run.
- SharePoint Online keeps the change log for about 60 days. A token that is older than the change log has expired. The changes between the token and the start of the change log are lost.
- **On Expired Token** sets what happens:
  - **Start From Oldest Available Change** (default). The node returns all changes that are still in the change log. All of them are newer than the token, so there are no duplicates. Each change has `tokenExpired: true`.
  - **Start From Now**. The node saves the current token and returns no changes.
  - **Fail**. The node stops with an error.
- SharePoint Online does not return an error for an old token. It returns the changes that are still in the log. So the node also checks the age of the token.

### Webhook requirements

- SharePoint must reach n8n from the internet over HTTPS. Set `WEBHOOK_URL` to the public URL of n8n.
- SharePoint validates the URL when it creates the subscription. It sends a `validationtoken` and needs the same text back within 5 seconds.
- When you activate the workflow, the node creates a list subscription. It sets an expiration date of 179 days. The SharePoint maximum is 180 days. The node uses a random client state.
- The node saves the subscription ID, the client state and the current change token in the static data of the workflow.
- Each notification must have the saved client state. The node answers at once, and then reads the changes.
- The node renews the subscription on activation and on a notification when less than 30 days remain. A list that has no changes for more than 150 days can reach the expiration date. To create a new subscription, save the workflow again, or deactivate and activate it.
- "Listen for test event" creates a subscription for the test URL. It expires after one day. n8n deletes it when the test ends.
- When you deactivate the workflow, the node deletes the subscriptions of the list that send to the webhook URL of the node.

### Polling or webhook

|                      | Polling Trigger                | Webhook Trigger      |
| -------------------- | ------------------------------ | -------------------- |
| Public HTTPS URL     | Not needed                     | Needed               |
| Scope                | Site collection, site, or list | One list or library  |
| Delay                | Up to one poll interval        | Seconds              |
| Start date option    | Yes                            | No                   |
| Subscription renewal | Not needed                     | Automatic, see above |

n8n starts a poller for each node type that has a `poll` method. It registers webhooks for each node type that has `webhooks`. A parameter cannot turn one of them off. So the webhook mode is a separate node.

## Limits and known behavior

- **Attachments** are not available in document libraries. Use File operations for libraries.
- **System Update** (List Item Update) uses `validateUpdateListItem` with `bNewDocumentUpdate`. It does not change Modified, Modified By or the version number. It does not start "item updated" workflows or alerts. SharePoint reads the date in the time zone of the site. The editor can only be sent by sign-in name. An app principal cannot be resolved, so the editor can change with app-only authentication.
- **List Item Get Many** with **Recursive** in filter mode uses the `FileDirRef` field. It can fail on lists above the list view threshold. Use CAML mode for those lists.
- **List Item Create** can create folders in a document library. It cannot create files. Use File Upload.
- **Search** reads results in pages of up to 500 rows. SharePoint can repeat a row across pages when the index changes. **Include Raw Result** contains the last page only.
- **Recycle Bin, Site scope** includes items of subsites. Get Many reads the recycle bin of the site collection and filters by path.
- **Sharing link passwords** are not supported. The REST API does not support them.
- **Sharing, Anyone links** can need an expiration date. A tenant policy can block them.
- **Default column values** do not support taxonomy fields.
- **Change log**:
  - The retention is about 60 days.
  - With the List scope and a folder path filter, SharePoint reads the change log of the folder. With other scopes, the node reads the item details of each change to find its path.
  - Delete changes have no name or path. **Resolve Deleted Items From Recycle Bin** looks them up. A very large recycle bin can make this slow.
  - **Latest First** sends only one request.
- **Copy job**: **Reset Role Inheritance After Move** and **Delete Empty Source Folder** do not run when the node does not wait for the job.
- **Site creation** with app-only authentication needs an owner. **Wait Until Ready** stops at **Wait Timeout**. The site can still finish later.
- **Taxonomy** works through CSOM. Delete Term deletes the child terms too.
- **File Upload** replaces a file with the same name by default. Turn off **Overwrite** to fail instead.
- **Publish, Unpublish, Approve and Deny** comments have a limit of 1023 characters. Publish and Unpublish need a library with minor versions or content approval. Approve and Deny need content approval.
- **Property bag writes** can need "custom scripts" on the site.
- **Hub site join** needs site collection admin rights.

## Compatibility

- Package version: 0.1.0 (see `package.json`).
- n8n nodes API version: 1.
- `n8n-workflow` is a peer dependency with any version (`*`). The package is built and tested against `n8n-workflow` 2.46.5 or later in the 2.x line.
- Built with `@n8n/node-cli` 0.55.5. The bundle targets Node.js 20.
- PnPjs `@pnp/sp` 4.22 or later is bundled.
- The trigger reads the poll budget of n8n when it exists (`getPollBudgetMs`). Older n8n versions do not have it, and the trigger then has no time limit.

## Development

```bash
npm install
npm run dev              # starts n8n with the nodes, for manual tests
npm run build
npm test
npm run test:integration # needs a .env file
npm run lint
```

### Build

`npm run build` runs `n8n-node build` and then `scripts/bundle.mjs`. The script uses esbuild to bundle each node entry point and its PnPjs imports into one CommonJS file. PnPjs is ESM-only, and n8n Cloud does not allow runtime dependencies. So PnPjs must be inside the node. Do not add `dependencies` to `package.json`.

### Tests

- `npm test` runs the unit tests in `test/unit`. They do not call SharePoint.
- `npm run test:integration` runs the tests in `test/integration` against a real tenant. Copy `.env.example` to `.env` and fill it in. Never commit `.env`. The tests use the certificate credential, so the Entra app needs `Sites.FullControl.All` or `Sites.Selected` and a certificate.
- Webhook subscription tests need `SP_WEBHOOK_NOTIFICATION_URL`. Without it, they are skipped.

Test safety rules:

- Tests only touch content that they create. Each run creates a folder named `__pnp-test-<timestamp>` in the test library and deletes it at the end.
- Tests never read, copy, move, change or delete content that they did not create.
- Tests never change existing site content, site groups or site-level permissions.
- Tests that create lists or libraries, and tests for site settings, property bag, site creation, hub sites, fields, content types, list settings and term store writes, are skipped by design. Only read tests run for these areas.

For manual tests, see the [manual test checklist](docs/manual-test-checklist.md).

### Contributing

Read `AGENTS.md` before you change code. The main rules:

- Each resource is in `nodes/SharePointPnp/resources/<resource>/index.ts`. Use `resources/list/index.ts` as the model.
- All SharePoint calls go through `createSp()` in `transport/index.ts`. Do not call `fetch` directly.
- Use the property selection contract (`shared/selection.ts`) in every Get and Get Many operation.
- Run `npm run lint`, `npm run build` and `npm test` before each commit.
- This is a public package. Never write real tenant, site, library, folder, file, user, group or term names in code, tests, commits or pull requests. Use placeholders such as `contoso` and `/sites/Example`.

## Resources and license

- [PnPjs documentation](https://pnp.github.io/pnpjs/) and [repository](https://github.com/pnp/pnpjs)
- [SharePoint REST API](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [Change log](CHANGELOG.md)

License: [MIT](LICENSE).
