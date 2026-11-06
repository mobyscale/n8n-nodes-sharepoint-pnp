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
- **Options**: Start From Date, Fetch Limit, Include Item Details, Item Fields, On Expired Token, Resolve Deleted Items From Recycle Bin.

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
- With the List scope and a folder path filter, SharePoint reads the change log of the folder. With other scopes, the node reads the item details of each change to find its path.
- Delete changes have no name or path. **Resolve Deleted Items From Recycle Bin** looks them up. A very large recycle bin can make this slow. If the identity cannot read the recycle bin of the site collection, the node reads the recycle bin of the site.
- **Latest First** sends only one request.
- The polling trigger reads the poll budget of n8n when it exists (`getPollBudgetMs`). Older n8n versions do not have it, and the trigger then has no time limit.
