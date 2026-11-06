# Common concepts

These rules apply to all resources of the **SharePoint PnP** node.

## Site

Use the **Site** option to choose the site. The site picker has a search list and a **By URL** mode. It accepts an absolute site URL, or a server-relative path such as `/sites/Example`.

## Paths

Every path option accepts three forms:

| Form            | Example                                                                |
| --------------- | ---------------------------------------------------------------------- |
| Site-relative   | `Shared Documents/Folder`                                              |
| Server-relative | `/sites/Example/Shared Documents/Folder`                               |
| Absolute URL    | `https://contoso.sharepoint.com/sites/Example/Shared Documents/Folder` |

A site-relative path starts at the selected site. The picker in "From List" mode works as a path browser. Type `Documents/Clients/Ac` to list the children of `Documents/Clients` that start with `Ac`.

## Property selection

Get and Get Many operations use one contract:

- Each operation loads a documented set of default properties. The option description lists them.
- **Additional Properties** adds more properties to the default set. Give a comma-separated list. Use `Navigation/Property` for a related object, for example `Author/Title`. The node adds `Navigation` to the expand list for you.
- **Expand** loads related objects.
- **Return All Properties** sends no `$select`. SharePoint then returns its default entity.

The default properties are stable. They are the properties that the output had before the option was added.

## AI agents and errors

The node can be used as a tool by an AI agent (`usableAsTool`). It supports **Continue On Fail**.
