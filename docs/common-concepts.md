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

## Pickers

- **From List** pickers load more results when you scroll. Most of them sort the results A-Z. Sites keep the relevance order of the search. File versions show the newest version first.
- Dropdown and multi-select pickers cannot load more results. They load at most 500 entries, sorted A-Z by name. For a value that is not in the list, use an expression, or the **By ID** or **By Name** mode when the field has one.

## Field Filters

**List Item > Get Many** and **File/Folder > Get Many** have **Field Filters**. Each condition has a **Field**, an **Operator** and a **Value**.

| Operator     | Use                                                  |
| ------------ | ---------------------------------------------------- |
| Equals       | The value is equal                                   |
| Not Equals   | The value is not equal                               |
| Contains     | The value contains the text. Only for text fields    |
| Begins With  | The value starts with the text. Only for text fields |
| Greater Than | The value is greater (numbers, dates, text)          |
| Less Than    | The value is less (numbers, dates, text)             |
| Is Empty     | The field has no value. There is no Value parameter  |
| Is Not Empty | The field has a value. There is no Value parameter   |

The node reads the type of each field and writes the value in the format of that type:

- Text, multiple lines of text, choice and multi-choice fields: the text. For a multi-choice field, Equals matches when one of the choices is equal.
- Number, currency, integer and counter (ID) fields: a number.
- Date fields: an ISO date. A date without a time, such as `2024-01-31`, compares only the date. A date with a time compares the full value.
- Yes/no fields: `1` or `0`. `true`, `false`, `yes` and `no` also work.
- Lookup and person fields: the ID of the lookup item or of the user, not the display text.
- Other field types: the text.

The option **Combine Field Filters** sets how the conditions are combined: **All (AND)** (default) or **Any (OR)**.

**Limit**: SharePoint has a list view threshold of 5000 items. A filter on a field that has no index can fail when one folder or list scope has more than 5000 items. File/Folder Get Many queries each folder on its own. List Item Get Many in CAML mode with the **Folder** option queries one folder. To filter a larger scope, add an index to the field in the list settings.

## AI agents and errors

The node can be used as a tool by an AI agent (`usableAsTool`). It supports **Continue On Fail**.
