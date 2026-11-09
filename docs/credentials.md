# Credentials

The package has two credentials. Each node has an **Authentication** option to choose one of them.

## OAuth2 (delegated)

The **SharePoint PnP OAuth2 API** credential signs in as a user. The node can do what that user can do.

1. Register an app in Microsoft Entra ID.
2. Add a **Web** redirect URI. Use the OAuth redirect URL that n8n shows in the credential.
3. Add delegated permissions for SharePoint, for example `AllSites.FullControl`. Give admin consent if your tenant needs it.
4. Create a client secret.
5. In n8n, create the credential. Enter the client ID and the client secret.
6. Enter the **Tenant ID** (the directory ID of the app registration). Keep `common` only for a multi-tenant app. A single-tenant app with `common` fails with the Entra error AADSTS50194.
7. Enter the **Subdomain**. For `https://contoso.sharepoint.com`, enter `contoso`.
8. Select **Connect** and sign in.

The credential asks for the scope `https://<subdomain>.sharepoint.com/.default`. The permissions of the app registration decide what the token can do.

## Certificate (app-only)

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

## Which feature needs which rights

| Feature                     | What it needs                                                                                                                                                          |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Taxonomy (term store)       | The node calls SharePoint CSOM. The REST term store API does not accept the app-only SharePoint token, but CSOM does. The identity must have access to the term store. |
| Create Site                 | With app-only authentication, SharePoint needs an owner. Set the **Owner** option to the user principal name of the owner, for example `user@contoso.com`.             |
| Join Hub Site               | The identity needs site collection admin rights on the site.                                                                                                           |
| Set Property Bag Values     | SharePoint can need "custom scripts" to be allowed on the site. If it is not, SharePoint returns an access denied error.                                               |
| Change, deleted item lookup | The node reads the recycle bin of the site collection. If the identity cannot read it, the node reads the recycle bin of the site.                                     |
| Sharing links               | A tenant or site sharing policy can block a link. For example, the policy can block Anyone links, or can require an expiration date.                                   |
| Webhook Trigger             | SharePoint must reach n8n from the internet over HTTPS. See [Webhook requirements](changes-and-triggers.md#webhook-requirements).                                      |
