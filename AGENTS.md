# n8n community node

## Overview
This is a project containing code for an n8n community node. n8n is a workflow
automation platform where users build workflows with nodes, which are the
building block of a workflow. Nodes can perform a range of actions, such as
starting a workflow (called a "trigger node"), fetching and sending data, or
processing and manipulating it. Besides that there are credentials - entities
that store sensitive information on how to connect to external services and
APIs. A node can require some credentials to be used. Community nodes are a way
for anyone to create such nodes and add them to be used in n8n. All community
nodes are named in a format: `n8n-nodes-<n>` or `@org/n8n-nodes-<n>`.
Community nodes can also be submitted for approval to be used on n8n Cloud
version. In that case there are rules that the node needs to follow in order to
be approved

## Important notes
- Follow the **rules and guidelines in this document and the linked docs
  below** over any code examples.
- All code blocks in these docs are **illustrative and incomplete**.
  They **MUST NOT** be copied verbatim or assumed to be the final desired code.
- Replace example names like `Example`, `Wordpress`, `wordpressApi`, etc.
  with names that match the **actual service / node** you are building.
- When in doubt, **generalize from the patterns**, don't replicate the exact
  structure, fields, or values from the examples.
- Produce the **full implementation** needed for the current project
  (nodes, credentials, tests, etc.), not just fragments similar to examples.
- If an example omits parts (e.g. types, operations, properties), **infer and
  implement the missing parts** based on the real requirements / API docs.
- Never output `Wordpress`-specific code unless the project is actually about
  WordPress.

## Project structure
There are two main folders in this project:
- `nodes` contains all of the nodes in a package (there can be more than 1).
  The code for each node usually lives in its own folder
- `credentials` contains all of the credentials in a package. Usually it's just
  a single file for every credential
So it looks something like this:
.
├── nodes/
│   └── Example/
│       ├── Example.node.ts
│       └── ...
├── credentials/
│   └── Example.credentials.ts
├── package.json
└── ...
It's important to note that `package.json` has a special field `n8n` that have
information about nodes and credentials in a package:
```json
{
  "name": "n8n-nodes-example",
  "version": "1.0.0",
  "n8n": {
    "n8nNodesApiVersion": 1,
    "strict": true,
    "credentials": [
        "dist/credentials/Example.credentials.js"
    ],
    "nodes": [
      "dist/nodes/Example/Example.node.js"
    ]
  }
}
```
`nodes` and `credentials` keys contain paths to transpiled JS files in a `dist`
folder for the nodes and credentials respectively. If you add/remove/rename
nodes and/or credentials, you need to make sure to update `n8n.nodes` and
`n8n.credentials` keys in `package.json` accordingly. Initial files in the
project _may_ contain example nodes and/or credentials that need to be
**removed or renamed** once you start making an actual node.

## Key guidelines
- Use the `n8n-node` CLI tool **whenever possible** for building, dev mode,
  linting, etc.
- **Always** address any lint/typecheck errors/warnings, unless there is a
  **very specific reason** to ignore/disable it
- Make sure to use **proper types whenever possible**
- If you are updating the npm package version, make sure to **update
  CHANGELOG.md** in the root of the repository
- Read `.agents/workflow.md` for more info

## Context-specific docs
Load these before working on the relevant area:

| Working on...                        | Read first                                                          |
|--------------------------------------|---------------------------------------------------------------------|
| Any node file in `nodes/`            | `.agents/nodes.md` and `.agents/properties.md`                      |
| A declarative-style node             | above + `.agents/nodes-declarative.md`                              |
| A programmatic-style node            | above + `.agents/nodes-programmatic.md`                             |
| Files in `credentials/`              | `.agents/credentials.md`                                            |
| Adding a new version to a node       | `.agents/versioning.md`                                             |
| Starting a new task or planning      | `.agents/workflow.md`                                               |

## Additional resources
If you need any extra information, here are links to n8n's official docs
regarding building community nodes:
- https://docs.n8n.io/integrations/community-nodes/build-community-nodes/
- https://docs.n8n.io/integrations/creating-nodes/overview/
- https://docs.n8n.io/integrations/creating-nodes/build/reference/
- https://docs.n8n.io/integrations/creating-nodes/build/reference/ux-guidelines/

## This project: n8n-nodes-sharepoint-pnp

### Architecture
- One programmatic node, `nodes/SharePointPnp/SharePointPnp.node.ts`, with a Resource and an Operation dropdown.
- Each resource lives in `nodes/SharePointPnp/resources/<resource>/index.ts` and exports a `ResourceModule`
  (`resources/types.ts`). The first entry of `description` must be the `operation` property. Operation handlers
  get `{ sp, siteUrl, itemIndex }` and return JSON data or `INodeExecutionData[]`. The node adds `pairedItem`.
- `resources/list/index.ts` is the reference implementation. Copy its style.
- All SharePoint calls use PnPjs (`@pnp/sp`) through `createSp()` in `transport/index.ts`. Requests go through
  `httpRequestWithAuthentication`, so n8n handles the token. Do not call `fetch` directly.
- Every Get and Get Many operation uses `shared/selection.ts`: a documented default property set, plus the
  options Additional Properties, Expand and Return All Properties (`selectionOptions`, `applySelection`).
- Use the shared pickers and helpers in `shared/descriptions.ts` (site, list, library, path) and `shared/utils.ts`
  (path handling, `isNotFoundError`, `toNodeError`). Paths accept site-relative, server-relative or absolute URLs.
- PnPjs is ESM-only. It is a devDependency that `scripts/bundle.mjs` bundles into the node. Never add `dependencies`.
- Do not use `setTimeout`. Use `sleep` from `n8n-workflow`. Do not use `process` or Node modules except `node:crypto`.

### Tests
- Unit tests: `test/unit/*.test.ts` (`npm test`). Do not call SharePoint in unit tests.
- Integration tests: `test/integration/*.test.ts` (`npm run test:integration`). Use `test/integration/harness.ts`.
  Create and change content only inside the folder from `createTestFolder()`, and remove it in `afterAll`.
- Tests that create lists or document libraries must use `it.skip`. They are hard to delete and need manual cleanup.
- Tests may only act on files, folders, list items and groups that the test created in its own setup, and must
  clean them up afterward. Never read, copy, move, change or delete an existing file or folder that the test did not
  create. The test library is only the container for the test folder.
- Never change existing site content, site groups, or site-level permissions in tests.
- Site-level and schema changes (site settings, property bag, site creation, hub sites, fields, content types,
  list settings, term store writes) never run in integration tests. Only read tests run. Write tests use `it.skip`.

### Rules for public content
- This is a public npm package. Never put real tenant, site, library, folder, file, user or group names in code,
  tests, commit messages or PR text. Use placeholders such as `contoso` and `/sites/Example`.
- Run `npm run lint`, `npm run build` and `npm test` before each commit.
