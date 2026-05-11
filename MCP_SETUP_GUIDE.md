# Access-ACE-MCP: Setup and Usage Guide

## Overview

**Access-ACE-MCP** is a Model Context Protocol (MCP) server that enables Claude to interact with Microsoft Access databases (.accdb and .mdb files) through automated COM automation. This allows Claude to read, modify, and manage Access database objects including tables, queries, forms, reports, modules, and VBA code.

The server provides comprehensive access to Access's object model, allowing you to:
- Connect to and manage Access databases
- Read and modify forms, reports, and modules
- Execute VBA procedures and Access commands
- Export and import database objects
- Modify form controls and properties
- Work with Access application settings and expressions

## ⚠️ CRITICAL: Backup Your Database First

**Before connecting to any Access database with this tool, you MUST create a backup copy.** This MCP server allows full programmatic modification of your database, including forms, reports, VBA code, and application settings.

### How to Backup Your Database

1. **Locate your .accdb or .mdb file** on disk
2. **Right-click the file** and select "Copy"
3. **Paste it in the same folder or a different safe location**
4. **Rename the copy** with a date or "backup" suffix, e.g., `MyDatabase_backup_2026-05-09.accdb`
5. **Store the backup safely** where it won't be accidentally modified

If anything goes wrong during Claude's modifications, you'll have your original database to restore from.

---

## System Requirements

### Prerequisites
- **Windows operating system** (this tool is Windows-only)
- **32-bit Microsoft Access** (Office 2010, 2013, 2016, 2019, 2021, or Microsoft 365)
  - **IMPORTANT:** 64-bit Access is **NOT supported**. If you have 64-bit Office installed, you must install the 32-bit version side-by-side
- **.NET Runtime** (the published executable handles this automatically with the runtime config)

### Checking Your Access Installation

To verify which version of Access you have:

1. Open Microsoft Access
2. Go to **File > Account** (or Help > About Microsoft Access in older versions)
3. Look for "64-bit" or "32-bit" in the version information

**If you see "64-bit":** You need to install the 32-bit version of Office/Access alongside your current installation.

---

## Installation

### Step 1: Build the Project

If you haven't already built the project:

1. Open the solution in **Visual Studio 2022** or later:
   ```
   C:\GitHub\Access-ACE-MCP\Access-ACE-MCP.slnx
   ```

2. Build the solution:
   - **Solution > Build Solution** (Ctrl+Shift+B)
   - Or: **Release > Build**

3. After building, publish the main project:
   - Right-click **Access-ACE-MCP** project
   - Select **Publish**
   - Follow the publish wizard to create a release build

The published executable will be located in the **Published** folder at the root of the solution.

### Step 2: Locate the Published Executable

After publishing, the MCP server executable and all dependencies will be in:
```
C:\GitHub\Access-ACE-MCP\Published\Access-ACE-MCP.exe
```

This folder contains:
- `Access-ACE-MCP.exe` - The main MCP server executable
- `Access-ACE-Agent.exe` - The .NET 4.8 COM worker (must be present)
- All required .NET runtime dependencies and DLLs

---

## Configuration for Claude Code

Access-ACE-MCP is configured as an MCP server that Claude Code can use. You can set it up in two ways:

### Option A: Claude Code CLI

1. **Locate your settings.json file:**
   - On Windows: `%APPDATA%\.claude\settings.json` or `~/.claude/settings.json`

2. **Add the MCP server configuration:**

   ```json
   "mcpServers": {
     "access-ace": {
       "command": "C:\\GitHub\\Access-ACE-MCP\\Published\\Access-ACE-MCP.exe",
       "args": []
     }
   }
   ```

3. **To connect to a specific database by default:**

   Pass the path to the .accdb file as an argument:

   ```json
   "mcpServers": {
     "access-ace": {
       "command": "C:\\GitHub\\Access-ACE-MCP\\Published\\Access-ACE-MCP.exe",
       "args": ["C:\\path\\to\\your\\database.accdb"]
     }
   }
   ```

4. **Save settings.json** and restart Claude Code CLI

### Option B: Claude Code Desktop

1. **Open Claude Code Desktop**

2. **Go to Settings > MCP Servers**

3. **Click "Add Server"** and enter:
   - **Name:** `access-ace` (or any name you prefer)
   - **Command:** `C:\GitHub\Access-ACE-MCP\Published\Access-ACE-MCP.exe`
   - **Arguments:** (leave empty, or add database path below)

4. **To auto-connect to a specific database:**
   - Add the argument: `C:\path\to\your\database.accdb`

5. **Click "Add"** and the server will connect

### Option C: Manual Connection in Claude Sessions

You don't need to pre-configure the server. In any Claude conversation, you can ask Claude to use the Access tools, and it will:
1. Use the MCP server from your settings if configured
2. Call `connect_access` with your database path when needed
3. Perform the requested operations

---

## Usage Examples

### Example 1: Connect to a Database and List Forms

**Your task:** "Connect to my database at C:\Users\DJ\MyDatabase.accdb and list all forms."

**What Claude will do:**
1. Call `connect_access` with the database path
2. Verify that 32-bit Access is installed
3. Launch the Access-ACE-Agent worker process
4. Call `get_forms` to retrieve the list of forms
5. Return the results

**Important:** Always ensure you've backed up your database before this step.

### Example 2: Modify a Form Control

**Your task:** "In my database at C:\Users\DJ\MyDatabase.accdb, find the 'UserForm' and change the caption of the 'SubmitButton' to 'Save Changes'."

**What Claude will do:**
1. Connect to the database (if not already connected)
2. Call `open_form` to open UserForm in design mode
3. Call `get_control_properties` to verify the button exists
4. Call `set_control_property` to change the Caption property
5. The form is automatically saved

### Example 3: Export a VBA Module for Review

**Your task:** "Export the VBA code from the 'Utils' module in my database so I can review it."

**What Claude will do:**
1. Connect to the database
2. Call `get_vba_code` with module name "Utils"
3. Return the complete VBA source code
4. Display it in a code block for your review

**Note:** This is a read-only operation and makes no changes to your database.

### Example 4: Add a VBA Procedure

**Your task:** "Add a new VBA procedure called 'LogEvent' to my 'Modules' module that logs a message to the console."

**What Claude will do:**
1. Connect to the database
2. Provide Claude with the VBA code for the procedure
3. Call `add_vba_procedure` to append the code to the module
4. Return confirmation of the addition

**Important:** Always review the generated code before asking Claude to add it to your database.

### Example 5: Get Database Metadata

**Your task:** "Show me all the queries in my database and their SQL text."

**What Claude will do:**
1. Connect to the database
2. Call `get_queries` to retrieve all queries with their SQL definitions
3. Return formatted results showing query names and SQL text

---

## Connecting to a Different Database

### Within an Existing Session

If you're already connected to one database and want to switch to another:

1. **Ask Claude:** "Disconnect from the current database and connect to C:\Users\DJ\OtherDatabase.accdb"

2. **Claude will:**
   - Call `disconnect_access` to close the current connection
   - Call `connect_access` with the new database path
   - Establish a fresh connection to the new database

### In a New Session

Simply provide the database path in your first request:

"In C:\Users\DJ\MyDatabase.accdb, list all the reports."

Claude will automatically connect using that path.

---

## Available Tools

The Access-ACE-MCP server exposes the following tools to Claude:

### Connection Management
- **connect_access** - Connect to an Access database (.accdb or .mdb)
- **disconnect_access** - Close the connection and exit Access
- **is_connected** - Check if a database is currently connected

### Application & Settings
- **get_application_info** - Get Access version, database name, and other metadata
- **get_option** / **set_option** - Read and modify Access application settings
- **eval_expression** - Evaluate Access expressions

### Procedures & Commands
- **run_procedure** - Execute a VBA Sub or Function
- **run_command** - Execute Access DoCmd commands by command ID
- **invoke_application_method** - Call any Access Application method
- **invoke_docmd_method** - Call any DoCmd method

### Database Objects (Generic)
- **export_object_to_text** - Export any object (form, report, module, macro, query) as text
- **import_object_from_text** - Import or replace an object from text definition
- **delete_object** - Delete any object from the database

### Enumeration
- **get_forms** - List all forms in the database
- **get_reports** - List all reports
- **get_modules** - List all VBA modules
- **get_macros** - List all macros
- **get_queries** - List all queries with their SQL text
- **get_vba_projects** - List VBA projects and their components

### VBA Code Management
- **get_vba_code** - Get complete source code of a module
- **set_vba_code** - Replace the entire source of a module
- **add_vba_procedure** - Append a procedure to an existing module
- **compile_vba** - Compile all VBA modules

### Form Design
- **open_form** - Open a form in design view
- **close_form** - Close a form without saving
- **get_form_controls** - List all controls on a form
- **get_control_properties** - Get all properties of a specific control
- **set_control_property** - Modify a single control property
- **export_form_to_text** - Export a form definition
- **import_form_from_text** - Import or replace a form definition
- **delete_form** - Delete a form

### Report Design
- **export_report_to_text** - Export a report definition
- **import_report_from_text** - Import or replace a report definition
- **delete_report** - Delete a report

---

## Troubleshooting

### "64-bit Microsoft Access is installed but is not supported"

**Cause:** You have 64-bit Access installed.

**Solution:**
1. Uninstall the 64-bit version of Microsoft Office
2. Download and install the **32-bit version** of Microsoft Office/Access
3. If you need both versions, you can install 32-bit Office/Access alongside your existing 64-bit setup (requires careful installation order)

### "Access-ACE-Agent.exe not found"

**Cause:** The Published folder is missing the agent executable.

**Solution:**
1. Rebuild the entire solution in Visual Studio
2. Re-publish the main Access-ACE-MCP project
3. Ensure both executables are in the Published folder:
   - `Access-ACE-Agent.exe`
   - `Access-ACE-MCP.exe`

### Connection times out (30 seconds)

**Cause:** The Access-ACE-Agent process failed to start or the named pipe connection failed.

**Possible solutions:**
1. Verify 32-bit Access is installed correctly
2. Check that your database file exists and is readable
3. Ensure the database isn't locked or corrupted
4. Try opening the database manually in Access first
5. Check Windows Event Viewer for any error messages

### "The database cannot be found"

**Cause:** The file path is incorrect or the file doesn't exist.

**Solution:**
1. Double-check the exact file path
2. Verify the file exists at that location
3. Use absolute paths (e.g., `C:\Users\DJ\MyDatabase.accdb`)
4. Ensure the file has read/write permissions

### Database becomes corrupted or locked

**Cause:** Access or the agent process terminated unexpectedly while the database was open.

**Solution:**
1. **Immediately restore from backup** (this is why backups are critical)
2. Close all instances of Microsoft Access
3. Delete any `.laccdb` lock files in the same folder as your database
4. Reopen the database to let Access compact/repair it
5. Contact Microsoft Access support if the database is still corrupted

---

## Security & Best Practices

### Backups
- ✅ **Always maintain recent backups** of your database
- ✅ **Test your backups** by restoring them to verify they work
- ✅ **Keep multiple versions** in case you need to roll back further

### Code Review
- ✅ **Review all VBA code** before asking Claude to add it to your database
- ✅ **Start with read-only operations** (like listing forms and queries)
- ✅ **Make incremental changes** rather than large bulk modifications
- ✅ **Test changes on a copy first** before applying to your production database

### Version Control
- ✅ **Export your database objects to text** using the export tools
- ✅ **Store exports in Git** for version control
- ✅ **Use this to track changes** over time

### Permissions
- ✅ **Run Claude with the same user account** that has write access to your database
- ✅ **Ensure the database file has read/write permissions** for your user account
- ✅ **If network drives are involved, ensure consistent access**

---

## Advanced: Batch Operations

You can ask Claude to perform multiple operations in sequence:

**Example:** "In my database, I want to:
1. Export the current 'MainForm' definition
2. Make a copy of 'MainForm' named 'MainForm_Backup'
3. List all controls on the original MainForm
4. Change the BackColor of the form to a light blue (#E6F2FF)
5. Verify the changes by re-exporting the form"

Claude will execute all these steps in order, handling connections and confirmations as needed.

---

## Support & Feedback

If you encounter issues or have feature requests:

1. **Check the Troubleshooting section** above
2. **Verify prerequisites** are installed correctly
3. **Restore from backup** if anything goes wrong
4. **Contact support** with details about:
   - The exact error message
   - What you were trying to do
   - Your Access version and Windows version
   - Whether it's a read or write operation

---

## Summary

| Task | Tool | Purpose |
|------|------|---------|
| Set up the MCP server | Installation steps above | One-time setup |
| Backup your database | Manual file copy | Critical safety step |
| Connect to a database | `connect_access` | Required before all operations |
| List database objects | `get_forms`, `get_queries`, etc. | Read-only exploration |
| Modify forms | `open_form`, `set_control_property` | Design changes |
| Modify VBA code | `get_vba_code`, `set_vba_code`, `add_vba_procedure` | Code changes |
| Export objects | `export_form_to_text`, `export_report_to_text` | Backup & version control |
| Execute VBA | `run_procedure` | Custom database automation |

---

**Last Updated:** May 9, 2026
**Version:** 1.0
