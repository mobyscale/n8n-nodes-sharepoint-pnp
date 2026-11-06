# Files and folders

The **File**, **Folder** and **File/Folder** resources of the **SharePoint PnP** node. For path formats, see [Paths](common-concepts.md#paths).

## File

PnPjs: [Files](https://pnp.github.io/pnpjs/sp/files/). Copy and Move use the [copy job](#copy-job-versus-direct) by default.

| Operation         | Description                                                       | PnPjs reference                          |
| ----------------- | ----------------------------------------------------------------- | ---------------------------------------- |
| Upload            | Upload a file from binary data. Files above 10 MB go up in chunks | `files.addUsingPath`, `files.addChunked` |
| Download          | Download a file as binary data                                    | `file.getBuffer`                         |
| Get               | Get the properties of a file                                      | `file`                                   |
| Get Versions      | Get the versions of a file                                        | `file.versions`                          |
| Copy              | Copy a file                                                       | `file.copyByPath`, or a copy job         |
| Move              | Move a file                                                       | `file.moveByPath`, or a copy job         |
| Rename            | Rename a file                                                     | `file.moveByPath`                        |
| Delete            | Delete a file, or move it to the recycle bin                      | `file.recycle`, `file.deleteWithParams`  |
| Check Out         | Check out a file                                                  | `file.checkout`                          |
| Check In          | Check in a file                                                   | `file.checkin`                           |
| Discard Check Out | Discard the check out of a file                                   | `file.undoCheckout`                      |
| Publish           | Publish the latest version of a file as a major version           | `file.publish`                           |
| Unpublish         | Unpublish the current major version of a file                     | `file.unpublish`                         |
| Approve           | Approve a file that waits for content approval                    | `file.approve`                           |
| Deny              | Deny a file that waits for content approval                       | `file.deny`                              |
| Restore Version   | Restore a version of a file                                       | `file.versions.restoreByLabel`           |

## Folder

PnPjs: [Folders](https://pnp.github.io/pnpjs/sp/folders/). Copy and Move use the [copy job](#copy-job-versus-direct) by default.

| Operation | Description                                    | PnPjs reference                             |
| --------- | ---------------------------------------------- | ------------------------------------------- |
| Create    | Create a folder and all missing parent folders | `folders.addUsingPath`                      |
| Get       | Get the properties of a folder                 | `folder`                                    |
| Copy      | Copy a folder                                  | `folder.copyByPath`, or a copy job          |
| Move      | Move a folder                                  | `folder.moveByPath`, or a copy job          |
| Rename    | Rename a folder                                | `folder.moveByPath`                         |
| Delete    | Delete a folder, or move it to the recycle bin | `folder.recycle`, `folder.deleteWithParams` |

## File/Folder

PnPjs: [Files](https://pnp.github.io/pnpjs/sp/files/) and [Folders](https://pnp.github.io/pnpjs/sp/folders/). Get Copy Job Status uses the SharePoint migration API. See [Copy job versus Direct](#copy-job-versus-direct).

| Operation           | Description                                                    | PnPjs reference               |
| ------------------- | -------------------------------------------------------------- | ----------------------------- |
| Get Many            | Get many files and folders in a folder                         | `list.renderListDataAsStream` |
| Check Exists        | Check if files or folders exist                                | `file`, `folder`              |
| Get Copy Job Status | Get the progress of a copy or move job that runs in SharePoint | `GetCopyJobProgress` (REST)   |

## Copy job versus Direct

File and Folder Copy and Move have a **Method** option.

- **Copy Job** (default) uses the SharePoint migration API. It scales to large content. A move removes the source in SharePoint.
- **Direct** runs the operation in one request. Use it for small content. The Direct method cannot replace a folder. Use the copy job for that.

A copy job uses the SharePoint migration API `CreateCopyJobs`. See [Microsoft 365 Copy and Move API (CreateCopyJobs)](https://learn.microsoft.com/sharepoint/dev/apis/spod-copy-move-api). The node also uses `GetCopyJobProgress` to read the job status. PnPjs does not wrap these methods, so the node calls them as REST requests through PnPjs.

Copy job behavior:

- **Name Conflict Behavior** is Fail, Keep Both, or Replace.
- **Wait for Completion** is on by default. The node polls the job until it ends or until **Max Wait (Seconds)** passes. The default maximum wait is 300 seconds.
- If the wait time ends, the node does not fail. It returns the job info with the status `running`. The copy continues in SharePoint.
- If you turn **Wait for Completion** off, the node returns at once. The output has `jobId`, `jobQueueUri` and `encryptionKey`.
- Use **File/Folder > Get Copy Job Status** to check a job later. It reads the job info from the input item by default. SharePoint returns each log entry once, so the summary and the messages show only the entries that are new since the last check.
- If the job reports errors, the node fails. Turn on **Continue on Job Errors** to get the job result instead. The errors are in the output in both cases.
- **Reset Role Inheritance After Move** and **Delete Empty Source Folder** run only after a job that finished without errors, and only when the node waits for the job.

The output has `status` (`running`, `completed` or `failed`), `finished`, `summary`, `errors` and `warnings`.

## Limits and known behavior

- **File Upload** replaces a file with the same name by default. Turn off **Overwrite** to fail instead.
- **Publish, Unpublish, Approve and Deny** comments have a limit of 1023 characters. Publish and Unpublish need a library with minor versions or content approval. Approve and Deny need content approval.
- **Copy job**: **Reset Role Inheritance After Move** and **Delete Empty Source Folder** do not run when the node does not wait for the job.
