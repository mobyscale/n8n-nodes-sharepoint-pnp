# n8n-nodes-sharepoint-pnp

n8n community nodes for advanced SharePoint Online operations. The nodes use the SharePoint REST API through [PnPjs](https://pnp.github.io/pnpjs/) ([`@pnp/sp`](https://github.com/pnp/pnpjs)). Use them for operations that the built-in Microsoft SharePoint node does not have, such as permissions, the change log, search, the recycle bin, and folder copy and move.

> This package is in development. The full documentation will be added before the first release.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

## Credentials

### OAuth2 (delegated)

The **SharePoint PnP OAuth2 API** credential signs in as a user. Register an Entra ID app with a web redirect URI from n8n, and give it delegated SharePoint permissions (for example `AllSites.FullControl`). Enter the tenant subdomain, for example `contoso` for `https://contoso.sharepoint.com`.

### Certificate (app-only)

The **SharePoint PnP Certificate API** credential signs in as the app. SharePoint REST does not accept app-only tokens from a client secret, so a certificate is necessary.

1. Register an Entra ID app and give it the SharePoint application permission `Sites.FullControl.All` or `Sites.Selected`.
2. Upload the public certificate to the app.
3. In n8n, enter the tenant ID, client ID, subdomain, and the PEM private key. Paste the certificate PEM, or enter its thumbprint.

To get PEM files from a `.pfx` file:

```bash
openssl pkcs12 -in certificate.pfx -nocerts -nodes -out private-key.pem
openssl pkcs12 -in certificate.pfx -clcerts -nokeys -out certificate.pem
```

## Triggers

The package has two trigger nodes. Both read the SharePoint change log with change tokens, and both have the same events, folder path filter and options.

- **SharePoint PnP Trigger** polls the change log of a site collection, a site or a list. It works without a public URL.
- **SharePoint PnP Webhook Trigger** uses SharePoint list webhooks. SharePoint sends a notification when a list or library changes. The node then reads the changes since the last change token, applies the events and the folder path filter, and starts the workflow for the matching changes.

n8n starts a poller for each node type that has a `poll` method, and registers webhooks for each node type that has `webhooks`. A parameter cannot turn one of them off, so the webhook mode is a separate node.

### Webhook trigger

- SharePoint must reach n8n from the internet over HTTPS. Set `WEBHOOK_URL` to the public URL of n8n. SharePoint validates the URL when it creates the subscription, and needs a response within 5 seconds.
- "Listen for test event" creates a subscription for the test URL. It expires after one day, and n8n deletes it when the test ends.
- When you activate the workflow, the node creates a list subscription with an expiration date of 179 days (the maximum is 180 days) and a random client state. It saves the subscription ID, the client state and the current change token in the static data of the workflow.
- Each notification must have the saved client state. The node responds at once, and then reads the changes.
- The node renews the subscription on activation and on a notification, when less than 30 days remain. A list that has no changes for more than 150 days can get to the expiration date. To create a new subscription, save the workflow again or deactivate and activate it.
- When you deactivate the workflow, the node deletes the subscriptions of the list that send to the webhook URL of the node.

### Expired change tokens

SharePoint Online keeps the change log for 60 days. The **On Expired Token** option sets what happens when a change token is older:

- **Start From Oldest Available Change** (default): return all changes that are still in the change log. All of them are newer than the token, so there are no duplicates. The changes before the start of the change log are lost. Each change has `tokenExpired: true`.
- **Start From Now**: save the current token and return no changes.
- **Fail**: stop with an error.

## Development

```bash
npm install
npm run build
npm test                 # unit tests
npm run test:integration # needs a .env file, see .env.example
```

PnPjs is ESM-only, and n8n Cloud does not allow runtime dependencies, so `npm run build` bundles PnPjs into the node with esbuild.

## Resources

- [PnPjs documentation](https://pnp.github.io/pnpjs/) and [repository](https://github.com/pnp/pnpjs)
- [SharePoint REST API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
