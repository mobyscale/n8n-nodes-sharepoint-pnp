# Change log and triggers

The **Change** resource of the **SharePoint PnP** node, the **SharePoint PnP Trigger** (polling) and the **SharePoint PnP Webhook Trigger**.

## Change

The change log uses the SharePoint `GetChanges` method with a `ChangeQuery`. PnPjs calls it for sites, site collections and lists. For a folder filter, the node calls `GetChanges` on the folder. See [Query SharePoint change log with ChangeQuery and ChangeToken](https://learn.microsoft.com/sharepoint/dev/solution-guidance/query-sharepoint-change-log-with-changequery-and-changetoken).

| Operation         | Description                                                     | API                                |
| ----------------- | --------------------------------------------------------------- | ---------------------------------- |
| Get Current Token | Get the current change token of a site collection, site or list | `GetChanges`, `CurrentChangeToken` |
| Get Many          | Get many changes from the change log                            | `GetChanges`                       |

## SharePoint PnP Trigger (polling)

The trigger polls the change log of a site collection, a site or a list. It works without a public URL.

Options:

- **Scope**: Site Collection, Site, or List.
- **Events**: File Created, File Updated, File Deleted, Folder Created, Folder Deleted, Item Created, Item Updated, Item Deleted, Moved, Permissions Changed, Renamed, Restored.
- **Folder Path Filter**: only trigger for content in this folder and its subfolders.
- **Options**: Start From Date, Fetch Limit, Include Item Details, Item Fields, On Expired Token, Resolve Deleted Items From Recycle Bin, Max Depth, Exclude Editors, Field Filters. See [Filters](#filters).

How it works:

- When you activate the workflow, the first poll saves the current change token. It returns no items. Only changes after that poll start the workflow. If you set **Start From Date**, the first poll returns the changes after that date.
- Each next poll reads the changes after the saved token. Then it saves the new token.
- In manual mode ("Fetch Test Event"), the node shows the latest matching changes of the last 24 hours. It shows at most 10 changes. It does not save a token.
- If you change the site, scope or list, the node starts again with a new token.

See [The change token contract](#the-change-token-contract) for expired tokens.

## SharePoint PnP Webhook Trigger

The trigger uses SharePoint list webhooks. SharePoint sends a notification when a list or library changes. The node then reads the changes since the last change token. It applies the events and the folder path filter. Then it starts the workflow for the matching changes.

It has the same **Events**, **Folder Path Filter** and **Options** as the polling trigger, except **Start From Date**. The node saves a token at activation, so it does not need a start date.

The node manages the list webhook subscriptions with the PnPjs [subscriptions](https://pnp.github.io/pnpjs/sp/subscriptions/) module. The main node has no subscription operations.

## Item details in bulk

Some options need the item details of a change: Include Item Details, Folder Path Filter (except the folder change log), the trigger events, Field Filters, and sometimes Item Kind and Max Depth. A change has no name or path, so the node reads them.

The node reads the details for each page of changes in bulk:

- List items: one CAML query for each list and for each 500 changed items. The query uses an `In` filter on the `ID` field with `Scope="RecursiveAll"`. `ID` is always indexed, so the query also works in lists and libraries above the list view threshold.
- Item Fields with a slash (lookup and person fields, for example `Editor/Title`): the CAML query cannot expand these fields. The node then reads the items by ID in `$batch` requests of 100.
- Files and folders without a list (File and Folder changes): `$batch` requests of 100.
- Items that the node cannot find go to the recycle bin lookup (see below).
- Each item is read one time in a run, also when it changed more than once.
- The node also reads the web ID one time and the base type of each list one time.

Example: 250 changes in one library need 3 requests (web ID, list type, one CAML query). Before, the node sent one request for each item.

When no filter needs the details, the node only reads the details of the changes up to the limit.

**Include Item Details** also adds `editorEmail`. The node uses the `EditorEmailHint` of the change. When the change has no hint, the node finds the site user with the same display name. It sends one request for each display name in a run. It adds `editorEmail` only when exactly one site user has that name, because display names are not unique. With app-only authentication, SharePoint often leaves `Editor` empty. Then there is no `editorEmail`.

## Filters

The Change resource and the triggers have these options:

| Option          | Get Many | Triggers | How it works                                                                                                                                                                                     |
| --------------- | -------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Item Kind       | Yes      | No       | Files, Folders and List Items. Changes to lists, sites and other objects are removed. The triggers use the events for this.                                                                      |
| Max Depth       | Yes      | Yes      | Folder levels below the Folder Path Filter, or below the scope root: the list or library (List scope) or the site. Level 1 is the direct children. Changes without a path are kept.              |
| Exclude Editors | Yes      | Yes      | A comma-separated list. An entry matches the display name of the editor, the `EditorEmailHint` of the change, or the email or login of the site user with that display name. Not case-sensitive. |
| Field Filters   | Yes      | Yes      | Field, operator (Equals, Not Equals, Contains, Is Empty, Is Not Empty) and value. All filters must match. The node compares the values after it reads the item details.                          |

Notes:

- **Field Filters need item details.** The node reads the details of every change, also when Include Item Details is off. The filter fields are not in the output unless they are also in Item Fields.
- Field Filters compare text without case. A multi-value field matches Equals and Contains when one of its values matches. Numbers and dates compare as values.
- Deleted items, File and Folder changes, and other changes without item fields do not match Field Filters.
- The field picker shows the fields of the selected list (List scope). With other scopes, use an expression. For a lookup or person field, use a name such as `Editor/Title`.
- Item Kind uses the item details. Max Depth uses the path of the change. When the change has no path (SharePoint Online often sends an empty `ServerRelativeUrl`), the node reads the item details.
- Exclude Editors reads a site user only when an entry has `@` or `|` and the display name or hint does not match. The node reads each display name one time in a run.
- The filters do not change the next change token. The token moves past the changes that a filter removed.

## The change token contract

A change token marks a point in the SharePoint change log. The Change resource and the triggers use it.

- **Get Current Token** returns the token for the scope.
- **Get Many** returns `nextChangeToken` when you set **Output** to **Single Item**. Save it and use it as the start of the next run.
- SharePoint Online keeps the change log for about 60 days. A token that is older than the change log has expired. The changes between the token and the start of the change log are lost.
- **On Expired Token** sets what happens:
  - **Start From Oldest Available Change** (default). The node returns all changes that are still in the change log. All of them are newer than the token, so there are no duplicates. Each change has `tokenExpired: true`.
  - **Start From Now**. The node saves the current token and returns no changes.
  - **Fail**. The node stops with an error.
- SharePoint Online does not return an error for an old token. It returns the changes that are still in the log. So the node also checks the age of the token.

## Webhook requirements

- SharePoint must reach n8n from the internet over HTTPS. Set `WEBHOOK_URL` to the public URL of n8n.
- SharePoint validates the URL when it creates the subscription. It sends a `validationtoken` and needs the same text back within 5 seconds.
- When you activate the workflow, the node creates a list subscription. It sets an expiration date of 179 days. The SharePoint maximum is 180 days. The node uses a random client state.
- The node saves the subscription ID, the client state and the current change token in the static data of the workflow.
- Each notification must have the saved client state. The node answers at once, and then reads the changes.
- The node renews the subscription on activation and on a notification when less than 30 days remain. A list that has no changes for more than 150 days can reach the expiration date. To create a new subscription, save the workflow again, or deactivate and activate it.
- "Listen for test event" creates a subscription for the test URL. It expires after one day. n8n deletes it when the test ends.
- When you deactivate the workflow, the node deletes the subscriptions of the list that send to the webhook URL of the node.

## Polling or webhook

|                      | Polling Trigger                | Webhook Trigger      |
| -------------------- | ------------------------------ | -------------------- |
| Public HTTPS URL     | Not needed                     | Needed               |
| Scope                | Site collection, site, or list | One list or library  |
| Delay                | Up to one poll interval        | Seconds              |
| Start date option    | Yes                            | No                   |
| Subscription renewal | Not needed                     | Automatic, see above |

n8n starts a poller for each node type that has a `poll` method. It registers webhooks for each node type that has `webhooks`. A parameter cannot turn one of them off. So the webhook mode is a separate node.

## Limits and known behavior

- The retention is about 60 days.
- With the List scope and a folder path filter, SharePoint reads the change log of the folder. With other scopes, the node reads the item details of the changes in bulk to find their paths.
- Delete changes have no name or path. **Resolve Deleted Items From Recycle Bin** looks them up. The node asks for the UniqueIds of 20 deleted items in one request. A large recycle bin rejects this query. Then the node reads the recycle bin newest first, one time for the run, until it passes the time of the oldest deleted change (at most 25 pages of 200). If the identity cannot read the recycle bin of the site collection, the node reads the recycle bin of the site.
- **Latest First** sends only one request.
- The polling trigger reads the poll budget of n8n when it exists (`getPollBudgetMs`). Older n8n versions do not have it, and the trigger then has no time limit.

## Pickers

- **Folder Path Filter** has **From List** and **By Path**, like the other folder fields. A plain path from an older workflow keeps working.
- **Item Fields** is a multi-select of the fields of the selected list. It needs the List scope, because the other scopes have no single list. An expression can give an array or a comma-separated list, for example `Title,Editor/Title`.
