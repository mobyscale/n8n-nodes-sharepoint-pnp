# Manual test checklist

Use this list to test the nodes by hand in n8n. Run `npm run dev` to start n8n with the nodes.

Items that change content have the mark **[WRITE]**. All other items only read.

## Before you start

- [ ] Use a test site, or a test folder in a test library. Do not test on production content.
- [ ] Create one test folder, for example `Shared Documents/__manual-test`. Do all **[WRITE]** items inside it.
- [ ] Use placeholder names in screenshots and notes. Do not write real tenant, site, user or group names in tickets or pull requests.
- [ ] At the end, delete the test folder, the test items, the test groups and the test terms that you created. Empty the recycle bin items that you made.
- [ ] Do not run the **[WRITE]** items for Site, List schema, Taxonomy and Permission on a site that you do not own.

## Credentials

### OAuth2 (delegated)

- [ ] Create a **SharePoint PnP OAuth2 API** credential with the client ID, the client secret and the subdomain.
- [ ] Select **Connect** and sign in. The credential shows as connected.
- [ ] Run the credential test. It passes.
- [ ] Enter a wrong subdomain. The test fails with a clear message.

### Certificate (app-only)

- [ ] Create a **SharePoint PnP Certificate API** credential with the tenant ID, client ID, subdomain, private key and certificate PEM.
- [ ] Run the credential test. It passes.
- [ ] Remove the certificate PEM and enter the thumbprint instead. The test passes.
- [ ] Enter a wrong private key. The test fails with a clear message.
- [ ] Use the **Authentication** option in a node to switch between the two credentials.

## Pickers

Use the SharePoint PnP node. Open each picker and check that it loads.

- [ ] **Site**, "From List": type part of a site name. The list shows matching sites.
- [ ] **Site**, "By URL": enter an absolute URL. Then enter a server-relative path such as `/sites/Example`. Both work.
- [ ] **List** (List > Get): the list shows lists and libraries. "By Title" and "By ID" also work. A bad ID shows a validation error.
- [ ] **Folder path browser** (File > Upload > Destination Folder): type `Shared Documents/` in "From List" mode. The list shows the subfolders. Type more letters. The list narrows.
- [ ] **Path**, "By Path": try a site-relative path, a server-relative path and an absolute URL. All three give the same result.
- [ ] **Fields** (List > Create View, "Fields to show"): the list shows the fields of the list.
- [ ] **Target field** (List > Create Field, Target = Site): the options load.
- [ ] **Groups** (Group > Get): the list shows the SharePoint groups of the site.
- [ ] **Role definitions** (Permission > Add > Role): the list shows the permission levels of the site.
- [ ] **Term groups** (Taxonomy > Get Term Sets): the list shows the term groups.
- [ ] **Term sets** (Taxonomy > Get Terms): the list shows the term sets.
- [ ] **Content types** (List > Add Content Type to List): the list shows the site content types. For Remove and Set Default, the list shows the content types of the list.
- [ ] **Hub sites** (Site > Join Hub Site): the list shows the hub sites.

## List Item resource mapper

- [ ] List Item > Create: choose a list. The **Fields** mapper loads the fields of the list.
- [ ] Check that read-only fields are not in the mapper.
- [ ] **[WRITE]** Create an item with a text field, a number, a date, a choice, and a multi-choice field. Check the item in SharePoint.
- [ ] **[WRITE]** Create an item with a lookup field and a person field. Check the values.
- [ ] Switch **Fields Input** to JSON. **[WRITE]** Create an item. Check it.
- [ ] **[WRITE]** List Item > Update: change one field. Other fields stay the same.
- [ ] **[WRITE]** Update with **Append to Multi-Value Fields** on a multi-choice field. The old values stay.
- [ ] **[WRITE]** Update with **System Update**. The Modified date does not change.
- [ ] List Item > Get and Get Many: test a filter, Order By, and a CAML query. Test **Return All** and **Limit**.
- [ ] Test **Additional Properties**, **Expand** and **Return All Properties** on Get.
- [ ] **[WRITE]** List Item > Delete: the item goes to the recycle bin. Then test **Delete Permanently** on another item.

## Binary upload and download

- [ ] Add a node that gives binary data, for example Read/Write Files from Disk.
- [ ] **[WRITE]** File > Upload with a small file. Check the file in SharePoint.
- [ ] **[WRITE]** Upload the same file again with **Overwrite** on. It replaces the file.
- [ ] **[WRITE]** Upload the same file with **Overwrite** off. It fails with a clear message.
- [ ] **[WRITE]** Upload a file larger than 10 MB. Check that it opens.
- [ ] File > Download. The output has the binary field. Save it and compare it with the original.
- [ ] **[WRITE]** Attachment > Add on a list item (not a library). Then Get Many and Download. The content is the same.
- [ ] **[WRITE]** Attachment > Update Content and Delete.
- [ ] Attachment on a document library gives a clear error.

## File and Folder copy and move

Make two folders in the test folder: `source` and `target`. Put a few files in `source`.

### Copy job, wait

- [ ] **[WRITE]** Folder > Copy with Method = Copy Job and **Wait for Completion** on. The output has `status: completed` and a summary.
- [ ] **[WRITE]** File > Copy with the same settings. The output has the new path.
- [ ] **[WRITE]** Folder > Move. The source folder is gone after the job.
- [ ] **[WRITE]** File > Move. The file is in the new place.

### Copy job, no wait

- [ ] **[WRITE]** Folder > Copy with **Wait for Completion** off. The node returns at once with `jobId`, `jobQueueUri` and `encryptionKey`.
- [ ] Connect File/Folder > Get Copy Job Status. It reads the job info from the input item. The status changes to `completed`.
- [ ] Run Get Copy Job Status again. The messages show only new log entries.
- [ ] **[WRITE]** Set **Max Wait (Seconds)** to 1 on a big copy. The node returns `status: running` and does not fail.

### Conflicts

- [ ] **[WRITE]** Copy a file to a folder that has a file with the same name. With **Fail**, the job reports an error and the node fails.
- [ ] **[WRITE]** Same test with **Continue on Job Errors** on. The node returns the result with the errors.
- [ ] **[WRITE]** Same test with **Keep Both**. The new file has a number in its name.
- [ ] **[WRITE]** Same test with **Replace**. The file is replaced.
- [ ] **[WRITE]** Method = Direct. Test **If Exists** options on a file copy.
- [ ] **[WRITE]** Folder move with **Reset Role Inheritance After Move**. The folder uses the permissions of the new parent.

## Permissions

Use a folder or an item in the test folder.

### Fields mode

- [ ] Permission > Get Many on the test folder. The output lists role assignments.
- [ ] **[WRITE]** Break Inheritance with **Copy Role Assignments** on.
- [ ] **[WRITE]** Add a test group with the role Read. Check in SharePoint.
- [ ] **[WRITE]** Add a user by email address (`user@contoso.com` style) with the role Contribute.
- [ ] Get Effective Permissions for that user.
- [ ] **[WRITE]** Remove the role of the user. Then Remove with no role. All roles of the principal go.
- [ ] **[WRITE]** Set with two entries. Other assignments go away.
- [ ] **[WRITE]** Reset Inheritance. The folder inherits again.

### JSON mode

- [ ] **[WRITE]** Add with **Input Mode** = JSON:

  ```json
  [
  	{ "principal": "Example Members", "principalType": "group", "role": "Read" },
  	{ "principal": "user@contoso.com", "roles": ["Read", "Edit"] }
  ]
  ```

- [ ] Use a wrong property name. The node fails with a message that lists the allowed properties.
- [ ] Use a principal that does not exist. The node fails, and SharePoint has no change.
- [ ] **[WRITE]** Remove with an entry that has no role. All roles of the principal go.
- [ ] **[WRITE]** Set with `[]`. All role assignments go. Do this only on an item with unique permissions in the test folder.

## Sharing

- [ ] **[WRITE]** Sharing > Create Link on a test file. Try View and Edit, with the scopes Organization and Anyone.
- [ ] Try **Expiration Date**. The output has the expiry.
- [ ] If the tenant blocks Anyone links, the error says that a sharing policy probably blocks them.
- [ ] Get Sharing Information. The new link is in the output.
- [ ] **[WRITE]** Share With Users with **Send Email** off.
- [ ] Check Permissions for the same user.
- [ ] **[WRITE]** Delete Link by kind. Then Unshare.

## Taxonomy

- [ ] Taxonomy > Get Term Groups. The output lists the groups.
- [ ] Get Term Sets with and without a term group.
- [ ] Get Term Set, Get Terms (with and without **Include Children**), and Get Term.
- [ ] Search Terms with the Match options. The case is ignored.
- [ ] **[WRITE]** Create Term in a test term set. Then Update Term and Delete Term. Use a test term set only.
- [ ] **[WRITE]** List Item Update on a managed metadata field with a term label. The item shows the term.
- [ ] **[WRITE]** Use a term ID, a `Label|ID` value, and a list that is separated with semicolons.
- [ ] **[WRITE]** Use an unknown label with **Create Missing Terms** off. The node fails.
- [ ] **[WRITE]** Use an unknown label with **Create Missing Terms** on. The term is created.

## Polling trigger

- [ ] Add SharePoint PnP Trigger. Select the test library, the scope List, and the events File Created and File Updated.
- [ ] Use "Fetch Test Event". The node shows changes of the last 24 hours, at most 10. If there are none, **[WRITE]** upload a file in the test folder and fetch again.
- [ ] Activate the workflow. The first poll saves the token and starts no execution.
- [ ] **[WRITE]** Upload a file in the test folder. The next poll starts an execution with the change.
- [ ] Test **Folder Path Filter** with the test folder. A change outside the folder does not start the workflow.
- [ ] Test the scopes Site and Site Collection.
- [ ] Test **Include Item Details**. The output has the name and the path.
- [ ] **[WRITE]** Delete a file. With the event File Deleted and **Resolve Deleted Items From Recycle Bin** on, the output has the name.
- [ ] Deactivate and activate the workflow. No old changes come again.
- [ ] Set **Start From Date** on a new workflow. The first poll returns the changes after that date.

## Webhook trigger

- [ ] Make n8n reachable over public HTTPS. Set `WEBHOOK_URL` to the public URL. Use a tunnel if needed.
- [ ] Add SharePoint PnP Webhook Trigger. Select the test library and the events.
- [ ] Use "Listen for test event". The node creates a subscription.
- [ ] **[WRITE]** Upload a file in the test folder. The event arrives within seconds.
- [ ] Activate the workflow. **[WRITE]** Upload a file again. An execution starts.
- [ ] Check that a change outside the **Folder Path Filter** does not start an execution.
- [ ] Deactivate the workflow. The subscription of the node is gone.
- [ ] Set `WEBHOOK_URL` to a URL that SharePoint cannot reach. Activation fails with the message about the webhook URL.

## Other resources

Do these only if you have time.

- [ ] Search > Query with a KQL text, **Return All** and **Row Limit per Page** = 2.
- [ ] Recycle Bin > Get Many with the filters. **[WRITE]** Restore one item that you deleted yourself.
- [ ] Change > Get Current Token, then Get Many with the token and **Output** = Single Item. The output has `nextChangeToken`.
- [ ] Change > Get Many with an old token (older than 60 days). The output has `tokenExpired: true`. Test all three **On Expired Token** values.
- [ ] Group and User operations on a test group that you create. **[WRITE]** Delete the group at the end.

## Error messages

- [ ] File > Get with a path that does not exist. The error message is clear and has the SharePoint text.
- [ ] Folder > Get with a bad path.
- [ ] Use a site URL that does not exist. The error is clear.
- [ ] Use an identity without access to a site, or a list with no rights. The node shows an access denied message. It does not show a stack trace.
- [ ] Use a certificate that Entra does not know. The error comes from the token request and is clear.

## Continue On Fail

- [ ] Send three input items to File > Get. Make the second path wrong. Turn on **Continue On Fail**.
- [ ] The output has three items. The second has an `error` field. The first and third are correct.
- [ ] Turn **Continue On Fail** off. The workflow stops at the second item.
- [ ] Check that the output items are paired with the input items.

## AI tool use

The main node has `usableAsTool` set.

- [ ] Add an AI Agent node. Connect SharePoint PnP as a tool.
- [ ] Give a read-only task, for example "list the files in the test folder". The agent calls File/Folder > Get Many and answers.
- [ ] Check that the tool shows its resource and operation in the agent log.
- [ ] Do not give the agent **[WRITE]** tools on production content.
