# Permissions, sharing, groups and users

The **Permission**, **Sharing**, **Group** and **User** resources of the **SharePoint PnP** node.

## Permission

PnPjs: [Security](https://pnp.github.io/pnpjs/sp/security/). The target can be a site, a list, a folder, a file or a list item.

| Operation                 | Description                                                 | PnPjs reference                                              |
| ------------------------- | ----------------------------------------------------------- | ------------------------------------------------------------ |
| Get Many                  | Get the role assignments of a target                        | `roleAssignments`                                            |
| Get Effective Permissions | Get the permissions that a user has on a target             | `getUserEffectivePermissions`                                |
| Add                       | Give a group or user a role on a target                     | `roleAssignments.add`                                        |
| Remove                    | Remove a role or all roles of a group or user from a target | `roleAssignments.remove`, `roleAssignments.getById().delete` |
| Set                       | Replace all role assignments of a target with a given list  | `roleAssignments`                                            |
| Break Inheritance         | Stop a target from inheriting permissions from its parent   | `breakRoleInheritance`                                       |
| Reset Inheritance         | Make a target inherit permissions from its parent again     | `resetRoleInheritance`                                       |
| Apply Mapping             | Apply a map of paths, principals and roles to many folders  | `roleAssignments`, `breakRoleInheritance`                    |

### Permission JSON input

Permission Add, Remove and Set have an **Input Mode**: **Fields** or **JSON**. JSON mode changes many principals and roles in one call. **Role Assignments (JSON)** is an array of objects.

| Property        | Description                                                                              |
| --------------- | ---------------------------------------------------------------------------------------- |
| `principal`     | Required. A group name, a group or user ID, or the email address or login name of a user |
| `principalType` | Optional. `user` or `group`. Use it when the name is not clear                           |
| `role`          | One role name or ID                                                                      |
| `roles`         | A list of role names or IDs. Do not use it together with `role`                          |

Example:

```json
[
	{ "principal": "Example Members", "principalType": "group", "role": "Contribute" },
	{ "principal": "user@contoso.com", "roles": ["Read", "Edit"] },
	{ "principal": "Example Owners", "role": "Edit" }
]
```

Rules:

- The node resolves all names to IDs first. It changes nothing in SharePoint if one entry cannot be resolved.
- For Remove, an entry without a role removes all roles of the principal.
- For Set, an empty array removes all role assignments.
- Entries for the same principal are merged. An entry without roles wins.
- An unknown property is an error.

### Pickers

The principal, user, role, group and item fields are resource locators. Use **From List** to search, or a free-text mode:

| Field                                               | From List                                                  | Other modes                             |
| --------------------------------------------------- | ---------------------------------------------------------- | --------------------------------------- |
| User (Permission, Group, User, Site owner)          | Site users. For a typed text, also people of the directory | By Email or Login                       |
| Principal (Set and Apply Mapping rows, group owner) | Groups and users, shown as `Group: name` or `User: name`   | By Email or Login, By Group Name, By ID |
| Role                                                | Visible permission levels of the site                      | By Name, By ID                          |
| Item (List Item target)                             | Items of the selected list                                 | By ID                                   |
| Remove Member user                                  | Members of the selected group                              | By Email or Login                       |

A workflow can still give a plain string, for example from an expression.

### Set and Add options

| Option                     | Operations | Description                                                                                                                                                            |
| -------------------------- | ---------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Clear Subscopes            | Add, Set   | Passed to `breakRoleInheritance` when the operation breaks inheritance. It removes unique permissions of child objects                                                 |
| Always Break Inheritance   | Set        | Call `breakRoleInheritance` even when the target has unique permissions. SharePoint can keep the existing assignments, so Set still adds and removes to match the list |
| Skip Unresolved Principals | Add, Set   | Skip an entry whose principal or role cannot be found. The output has a `warnings` array. Off by default: the operation fails and changes nothing                      |

### Apply Mapping

Apply Mapping gives many folders their permissions in one step, for example the folders of a tree such as `Reports/2024`. Each mapping has a **Path**, a **Principal** and a **Role**. The paths are relative to the **Base Folder**.

Use the **Mappings** rows, or set **Input Mode** to **JSON** and use **Mappings (JSON)**. The JSON is an array of objects with the properties of [Permission JSON input](#permission-json-input) and a `path`:

```json
[
	{ "path": "Contracts", "principal": "Example Members", "principalType": "group", "role": "Read" },
	{ "path": "Reports/*/Drafts", "principal": "Example Visitors", "role": "Read" },
	{ "path": "Reports/*/Contracts", "principal": "user@contoso.com", "roles": ["Read", "Edit"] },
	{ "path": "Archive/**/Final", "principal": "Example Visitors", "role": "Read" },
	{ "path": "Templates/**", "principal": "Example Members", "role": "Read" },
	{ "path": ".", "principal": "Example Owners", "role": "Full Control" }
]
```

Path rules:

- A path is relative to the base folder. `.` is the base folder.
- `*` matches all child folders that exist, one level. It must be a whole segment: `Reports/*/Drafts` is correct, `Rep*/Drafts` is an error.
- `**` matches folders at any depth. `Archive/**/Final` matches each `Final` folder below `Archive`, for example `Archive/Final` and `Archive/2024/Q1/Final`. At the end of a path, `**` matches all folders below: `Templates/**` matches each folder below `Templates`, but not `Templates` itself. It must be a whole segment too.
- The node lists each folder that a `**` reaches, one request for each folder. On a large tree, start the path at the deepest folder that you know.
- Names match without case. `reports/*/drafts` matches `Reports/2024/Drafts`.
- A path without `*` or `**` is one folder. When it does not exist, **If Folder Is Missing** sets what to do: **Skip** (default) adds it to `missing`, **Fail** stops before any permission change. With **Create Missing Folders**, the node creates the folder and its missing parents.
- A path with `*` or `**` only matches folders that exist. When it matches no folders, the output has a warning.
- `..` is not allowed.

How the node works:

1. It resolves all principals and roles one time for the whole mapping. If a name cannot be resolved, it fails and changes nothing. With **Skip Unresolved Principals**, it skips that mapping and adds a warning.
2. It expands the paths and groups the mapped role assignments by folder. A folder that more than one mapping matches gets the role assignments of all these mappings.
3. It applies the role assignments to each folder. It sends the changes of a folder in `$batch` requests.

**Mode**:

| Mode          | Result                                                                                                             |
| ------------- | ------------------------------------------------------------------------------------------------------------------ |
| Set (default) | Each folder has exactly the mapped role assignments. The node removes other role assignments, like Set             |
| Add           | The node adds the mapped role assignments. Other role assignments stay. When it breaks inheritance, it copies them |

Options:

| Option                         | Description                                                                                                                        |
| ------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------- |
| Break Inheritance              | **Only If Inheriting** (default) breaks inheritance on folders that inherit. **Always** breaks it on each folder                   |
| Copy Current Assignments First | Set mode only. Copy the current role assignments when the node breaks inheritance. Add mode always copies them                     |
| Clear Subscopes                | Passed to `breakRoleInheritance`. It removes unique permissions of child objects                                                   |
| Create Missing Folders         | Create a folder of a path without `*` or `**` when it does not exist                                                               |
| If Folder Is Missing           | **Skip** or **Fail** when a folder of a path without `*` or `**` does not exist and Create Missing Folders is off                  |
| Skip Unresolved Principals     | Skip a mapping whose principal or role cannot be found, and add a warning. Off by default: the operation fails and changes nothing |
| Output                         | **One Item** (default) or **One Item Per Folder**                                                                                  |

Output with **One Item**:

```json
{
	"baseFolder": "/sites/Example/Shared Documents/Archive",
	"mode": "set",
	"folders": [
		{
			"path": "Reports/2024/Drafts",
			"serverRelativeUrl": "/sites/Example/Shared Documents/Archive/Reports/2024/Drafts",
			"matchedFrom": ["Reports/*/Drafts"],
			"inheritanceBroken": true,
			"added": 1,
			"removed": 0,
			"unchanged": 0,
			"created": false
		}
	],
	"warnings": [],
	"missing": [
		{
			"path": "Contracts",
			"serverRelativeUrl": "/sites/Example/Shared Documents/Archive/Contracts",
			"matchedFrom": ["Contracts"]
		}
	]
}
```

- `matchedFrom` lists the mapping paths that matched the folder.
- `added` and `removed` count role assignments (one principal and one role). `unchanged` counts mapped role assignments that the folder had already. A second run with the same mapping has `added: 0` and `removed: 0`.
- With **One Item Per Folder**, each folder and each missing folder is one item with `baseFolder`, `mode`, the folder properties, `missing` (`true` or `false`) and `warnings`. If there are no folders, the output is one item like **One Item**.

## Sharing

PnPjs: [Sharing](https://pnp.github.io/pnpjs/sp/sharing/).

| Operation               | Description                                             | PnPjs reference                          |
| ----------------------- | ------------------------------------------------------- | ---------------------------------------- |
| Create Link             | Create or get a sharing link                            | `getShareLink`                           |
| Delete Link             | Delete a sharing link by its kind or share ID           | `deleteSharingLinkByKind`, `unshareLink` |
| Get Sharing Information | Get the sharing links and the people with access        | `getSharingInformation`                  |
| Share With Users        | Give users access, with an optional email               | `SP.Web.ShareObject` (REST)              |
| Check Permissions       | Check which role users have on a target through sharing | `checkSharingPermissions`                |
| Unshare                 | Remove all sharing from a target                        | `unshare`                                |

## Group

PnPjs: [Site groups](https://pnp.github.io/pnpjs/sp/site-groups/).

| Operation     | Description                            | PnPjs reference               |
| ------------- | -------------------------------------- | ----------------------------- |
| Get           | Get a SharePoint group                 | `siteGroups.getById`          |
| Get Many      | Get many SharePoint groups of the site | `siteGroups`                  |
| Create        | Create a SharePoint group              | `siteGroups.add`              |
| Update        | Update the properties of a group       | `siteGroups.getById().update` |
| Delete        | Delete a SharePoint group              | `siteGroups.removeById`       |
| Get Members   | Get the members of a group             | `group.users`                 |
| Add Member    | Add a user to a group                  | `group.users.add`             |
| Remove Member | Remove a user from a group             | `group.users.removeById`      |

## User

PnPjs: [Site users](https://pnp.github.io/pnpjs/sp/site-users/).

| Operation | Description                                                 | PnPjs reference                                     |
| --------- | ----------------------------------------------------------- | --------------------------------------------------- |
| Get       | Get a user of the site by ID, email address or login name   | `siteUsers.getById`, `getByEmail`, `getByLoginName` |
| Get Many  | Get many users of the site                                  | `siteUsers`                                         |
| Ensure    | Get a user and add the user to the site user list if needed | `web.ensureUser`                                    |

## Limits and known behavior

- **Sharing link passwords** are not supported. The REST API does not support them.
- **Sharing, Anyone links** can need an expiration date. A tenant policy can block them.
