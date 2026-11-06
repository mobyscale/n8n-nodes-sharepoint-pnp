# Development

```bash
npm install
npm run dev              # starts n8n with the nodes, for manual tests
npm run build
npm test
npm run test:integration # needs a .env file
npm run lint
```

## Build

`npm run build` runs `n8n-node build` and then `scripts/bundle.mjs`. The script uses esbuild to bundle each node entry point and its PnPjs imports into one CommonJS file. PnPjs is ESM-only, and n8n Cloud does not allow runtime dependencies. So PnPjs must be inside the node. Do not add `dependencies` to `package.json`.

## Tests

- `npm test` runs the unit tests in `test/unit`. They do not call SharePoint.
- `npm run test:integration` runs the tests in `test/integration` against a real tenant. Copy `.env.example` to `.env` and fill it in. Never commit `.env`. The tests use the certificate credential, so the Entra app needs `Sites.FullControl.All` or `Sites.Selected` and a certificate.
- Webhook subscription tests need `SP_WEBHOOK_NOTIFICATION_URL`. Without it, they are skipped.

Test safety rules:

- Tests only touch content that they create. Each run creates a folder named `__pnp-test-<timestamp>` in the test library and deletes it at the end.
- Tests never read, copy, move, change or delete content that they did not create.
- Tests never change existing site content, site groups or site-level permissions.
- Tests that create lists or libraries, and tests for site settings, property bag, site creation, hub sites, fields, content types, list settings and term store writes, are skipped by design. Only read tests run for these areas.

For manual tests, see the [manual test checklist](manual-test-checklist.md).

## Contributing

Read `CLAUDE.md` before you change code. The main rules:

- Each resource is in `nodes/SharePointPnp/resources/<resource>/index.ts`. Use `resources/list/index.ts` as the model.
- All SharePoint calls go through `createSp()` in `transport/index.ts`. Do not call `fetch` directly.
- Use the property selection contract (`shared/selection.ts`) in every Get and Get Many operation.
- Run `npm run lint`, `npm run build` and `npm test` before each commit.
- This is a public package. Never write real tenant, site, library, folder, file, user, group or term names in code, tests, commits or pull requests. Use placeholders such as `contoso` and `/sites/Example`.
