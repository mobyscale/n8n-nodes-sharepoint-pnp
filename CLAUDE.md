# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

`@mobyscale/n8n-nodes-sharepoint-pnp` is an n8n community node package for SharePoint Online. It wraps PnPjs (`@pnp/sp`, the
SharePoint REST API, not Microsoft Graph). It is a public npm package that is meant for n8n Cloud verification, so
it follows the strict n8n community node rules (`"strict": true` in `package.json`).

User documentation is in `README.md` and `docs/` (one page for each area). `AGENTS.md` and `.agents/*.md` hold the
general n8n node guidance from the `@n8n/node-cli` template. Read the relevant `.agents` file before you change a
node (`nodes.md`, `properties.md`, `nodes-programmatic.md`) or a credential (`credentials.md`).

## Commands

- Install: `npm ci`. `.npmrc` sets `ignore-scripts=true`, because `isolated-vm` (a dependency of n8n-workflow)
  does not build on current Node versions and is not necessary here.
- `npm run build` runs `n8n-node build`, then `scripts/bundle.mjs` (esbuild bundles each node in `package.json`
  `n8n.nodes` into one CJS file, with `n8n-workflow` as the only external).
- `npm run dev` starts n8n at http://localhost:5678 with the nodes loaded. `docs/manual-test-checklist.md` is the
  manual test plan.
- `npm run lint` (`n8n-node lint`), `npm test` (unit), `npm run test:integration`.
- One test file or test: `npx vitest run test/unit/selection.test.ts -t "expands"`. For integration tests, add
  `--config vitest.integration.config.mts`.
- Integration tests need an uncommitted `.env` (see `.env.example`) and a certificate key pair. The config file
  reads them and passes them to the tests with vitest `provide`, because the lint rules do not allow `process`
  or `fs` in test files. Without `.env`, every integration suite is skipped.
- `n8n-node lint` also scans `.claude/worktrees` if agent worktrees exist. Use
  `npx eslint . --ignore-pattern ".claude/**"` in that case.
- Run `npm run lint`, `npm run build` and `npm test` before each commit. CI (`.github/workflows/ci.yml`) runs the
  same steps. `.github/workflows/publish.yml` publishes to npm with provenance.

## Architecture

### Nodes
- `nodes/SharePointPnp/SharePointPnp.node.ts`: one programmatic node with an Authentication, a Resource and an
  Operation option. The resource registry is `resources/index.ts`.
- `nodes/SharePointPnpTrigger`: polls the change log with change tokens.
- `nodes/SharePointPnpWebhookTrigger`: uses SharePoint list webhooks. It is a separate node because one n8n node
  type cannot both poll and receive webhooks.
- Both triggers use `shared/changes.ts` and `shared/changeTrigger.ts`. The webhook trigger also uses
  `shared/subscriptions.ts`.
- If you add, remove or rename a node or a credential, update `n8n.nodes` and `n8n.credentials` in `package.json`.
  `scripts/bundle.mjs` reads `n8n.nodes` to find its entry points.

### Resources
- Each resource lives in `nodes/SharePointPnp/resources/<resource>/index.ts` and exports a `ResourceModule`
  (`resources/types.ts`). The first entry of `description` must be the `operation` property. Operation handlers
  get `{ sp, siteUrl, itemIndex }` and return JSON data or `INodeExecutionData[]`. The node adds `pairedItem` and
  handles Continue On Fail.
- `resources/list/index.ts` is the reference implementation. Copy its style.
- All operation properties are merged into one node description. Two properties with the same `name` in
  different resources must have the same `type`. If the types are different, give them different names.
- Every Get and Get Many operation uses `shared/selection.ts`: a documented default property set, plus the
  options Additional Properties, Expand and Return All Properties (`selectionOptions`, `applySelection`).
- Use the shared pickers and helpers in `shared/descriptions.ts` (site, list, library, path) and `shared/utils.ts`
  (path handling, `isNotFoundError`, `toNodeError`). Paths accept site-relative, server-relative or absolute URLs.
- Pickers with search (`listSearch`) and the List Item field mapper (`resourceMapping`) are in
  `nodes/SharePointPnp/methods/`.

### Transport and credentials
- All SharePoint calls use PnPjs through `createSp()` in `transport/index.ts`. Do not call `fetch` directly.
  `createSp()` replaces the PnPjs send step with `N8nSend`. It sends each request through
  `httpRequestWithAuthentication`, so n8n handles the token, and returns a fetch `Response`. It retries 429, 503
  and 504 with `Retry-After`. It sends a 401 again so that n8n can refresh the token.
- `SharePointPnpOAuth2Api`: delegated. It extends `microsoftOAuth2Api`, with the scope
  `https://<subdomain>.sharepoint.com/.default`.
- `SharePointPnpCertificateApi`: app-only. SharePoint REST does not accept app-only tokens from a client secret.
  The credential signs a JWT client assertion with `node:crypto` (`credentials/helpers/clientAssertion.ts`) in
  `preAuthentication` and keeps the token in a hidden, expirable field.
- The node gets the credential type from its Authentication parameter (`getCredentialType`).

### Features that do not use PnPjs methods
- Taxonomy uses the CSOM endpoint `/_vti_bin/client.svc/ProcessQuery` (`resources/taxonomy/csom.ts`), because the
  REST v2.1 term store returns 403 for a SharePoint app-only token.
- Copy and move with Copy Job use `CreateCopyJobs` and `GetCopyJobProgress` (`shared/copyJob.ts`). A move is
  complete only when the job state is 0 (after the MoveCleanup phase), not at the first JobEnd.
- Site creation uses `SPSiteManager`.
- Change token contract: SharePoint Online does not return an error for a token older than its change log
  retention (about 60 days). It returns changes from the start of the log. The code finds expired tokens by their
  age, and by default starts from the oldest available change.

### Constraints from the strict n8n lint rules
- No runtime `dependencies`. PnPjs is ESM-only. It is a devDependency that `scripts/bundle.mjs` bundles into
  each node.
- Imports are limited to devDependencies, `n8n-workflow` and `node:crypto`. Do not use `process`, `setTimeout` or
  other Node modules. Use `sleep` from `n8n-workflow`.
- Errors must be `NodeApiError` or `NodeOperationError` (`toNodeError` in `shared/utils.ts`).
- Icons must be SVG.
- Fix every lint and type error. Do not disable a rule unless there is a very specific reason.
- If you change the package version, update `CHANGELOG.md`.

## Tests

- Unit tests: `test/unit/*.test.ts`. They mock PnPjs and do not call SharePoint.
- Integration tests: `test/integration/*.test.ts`. They run against a real production site, so these rules are
  strict:
  - Use `test/integration/harness.ts`. Create and change content only inside the folder from
    `createTestFolder()` (`__pnp-test-<timestamp>` in the test library), and remove it in `afterAll`.
  - Tests may only act on files, folders, list items and groups that the test created in its own setup, and must
    clean them up. Never read, copy, move, change or delete an existing file or folder that the test did not
    create. The test library is only the container for the test folder.
  - Never change existing site content, site groups, or site-level permissions.
  - Tests that create lists or document libraries, and site-level and schema changes (site settings, property bag,
    site creation, hub sites, fields, content types, list settings, term store writes) use `itWrite` from the
    harness. They run only with `SP_RUN_WRITE_TESTS=true`, which is for a site that can be changed and reset, never
    the production site. Without it they are skipped, so only read tests run for these areas.

## Rules for public content

- Never put real tenant, site, library, folder, file, user, group or term names, or tenant and app IDs, in code,
  tests, commit messages or PR text. Use placeholders such as `contoso`, `/sites/Example` and
  `00000000-0000-0000-0000-000000000000`.
- Never commit `.env`, `certs/`, or key and certificate files. `.gitignore` excludes them.

## Git workflow

- One PR for each piece of work, with atomic commits. Merge with a merge commit after CI passes.
