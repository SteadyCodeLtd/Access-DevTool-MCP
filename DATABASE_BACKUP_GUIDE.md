# Database Backup Guide

This guide explains how to back up Microsoft Access databases using the `export_database_objects` tool in Access-DevTool-MCP.

## Simple Claude Prompt

Use this prompt when you want Claude to perform a backup:

```text
Please back up my Access database to a structured folder.

Database: C:\path\to\MyDatabase.accdb
Output: C:\Backups\MyDatabase_Backup

Export all forms, reports, queries, and modules into subfolders.
```

## Recommended Backup Approach

The most token-efficient approach is:

1. Claude connects to the database.
2. Claude calls `export_database_objects` once.
3. The server returns JSON with the exported object code.
4. Claude writes the files locally into a folder structure.

This keeps the MCP server focused on Access extraction and avoids many separate tool calls.

## Folder Structure

A typical backup folder looks like this:

```text
Database_Backup/
├── forms/
├── reports/
├── queries/
├── modules/
└── BACKUP_SUMMARY.txt
```

### File Extensions

- **Forms**: `.txt`
- **Reports**: `.txt`
- **Queries**: `.sql`
- **Modules**: `.bas`

## Full Backup Workflow

### 1. Connect to the database

```text
connect_access(database_path)
```

### 2. Export all object types

Use this payload:

```json
{
  "object_types": {
	"forms": [],
	"reports": [],
	"queries": [],
	"modules": []
  }
}
```

An empty array means export all objects of that type.

### 3. Write files locally

Claude should create the folder structure and write each exported object to the correct subfolder.

### 4. Create a summary file

Write `BACKUP_SUMMARY.txt` with:

- database path
- timestamp
- object counts
- status message
- any errors returned by the tool

### 5. Disconnect

```text
disconnect_access()
```

## Selective Backup Workflow

You can export only specific objects.

### Example Prompt

```text
Export only these objects from my Access database:
- Forms: frmMain, frmCustomer
- Modules: basUtility

Save them to C:\Backups\Selective
```

### Example Payload

```json
{
  "object_types": {
	"forms": ["frmMain", "frmCustomer"],
	"modules": ["basUtility"]
  }
}
```

## Best Practices

- Always disconnect when finished.
- Use a short output path to avoid long Windows paths.
- Check the `errors` array in the response.
- Verify the file counts in `BACKUP_SUMMARY.txt`.
- Keep backups under version control if that fits your workflow.

## Troubleshooting

### No objects exported
- Confirm the database path is correct.
- Confirm the object names exist.
- Check whether Access is open elsewhere.

### Export errors
- Review the `errors` array returned by the tool.
- Some objects may be locked or corrupted.

### Path issues
- Use a shorter output folder if the path is too long.

## Example Backup Session

```text
User: Back up C:\Projects\Invoicing.accdb to C:\Backups\Invoicing_Backup

Claude:
1. connect_access("C:\Projects\Invoicing.accdb")
2. export_database_objects({ forms: [], reports: [], queries: [], modules: [] })
3. Create forms, reports, queries, modules subfolders
4. Write exported files to disk
5. Create BACKUP_SUMMARY.txt
6. disconnect_access()
```

## Related Files

- `.claude/skills/access-ace-helper.md` - Claude skill guidance
- `README.md` - Project overview

---

If you want, I can also add a shorter version of this guide to the README or create a restore guide next.
