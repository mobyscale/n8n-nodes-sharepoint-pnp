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
