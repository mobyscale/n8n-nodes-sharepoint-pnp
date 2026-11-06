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
	{ "principal": "user@contoso.com", "roles": ["Read", "Edit"] }
]
```

Rules:

- The node resolves all names to IDs first. It changes nothing in SharePoint if one entry cannot be resolved.
- For Remove, an entry without a role removes all roles of the principal.
- For Set, an empty array removes all role assignments.
- Entries for the same principal are merged. An entry without roles wins.
- An unknown property is an error.

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
