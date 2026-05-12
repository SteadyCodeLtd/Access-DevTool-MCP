# Access-ACE-MCP

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET Version](https://img.shields.io/badge/.NET-10.0-blue)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2032--bit-blue)](https://www.microsoft.com/)

**Access-ACE-MCP** is a Model Context Protocol (MCP) server that enables Claude AI to interact with Microsoft Access databases through COM automation. Automate database management, modify forms and reports, execute VBA code, and manage database objects programmatically.

## Features

- 🗄️ **Database Connection** - Connect to and manage .accdb and .mdb files
- 📋 **Object Management** - Create, read, update, and delete forms, reports, modules, queries, and macros
- 💾 **VBA Automation** - Execute VBA procedures, read/modify VBA source code, compile modules
- 🎨 **Form Design** - Inspect and modify form controls, properties, and layouts
- 📊 **Report Management** - Export, import, and delete reports
- ⚙️ **Application Control** - Execute Access commands, invoke COM methods, evaluate expressions
- 🔍 **Database Inspection** - List and explore all database objects with metadata
- 📤 **Export/Import** - Round-trip export/import of database objects as text (SaveAsText/LoadFromText)

## Quick Links

- 📖 **[Setup Guide](MCP_SETUP_GUIDE.md)** - Complete installation and usage guide
- ⚡ **[Quick Start](QUICK_START.md)** - Get running in 30 seconds
- 🏗️ **[Architecture](ARCHITECTURE.md)** - Technical deep dive and design documentation
- 🤖 **[Claude Code Skill](.claude/skills/access-ace-helper.md)** - Expert guidance for using Access-ACE-MCP

## Requirements

### System
- **Windows Operating System** (Windows 10 or later recommended)
- **32-bit Microsoft Access** (Office 2010, 2013, 2016, 2019, 2021, or Microsoft 365)
  - ⚠️ **64-bit Access is NOT supported**
- **.NET Runtime** (included with published builds)

### Software
- **Microsoft Visual Studio 2022** or later (for building from source)
- **Git** (optional, for cloning the repository)

## Installation

### Option 1: Use Published Build (Recommended)

1. **Download** the latest release from the [Releases](../../releases) page
2. **Extract** to a location like `C:\Tools\Access-ACE-MCP\`
3. **Verify** both executables are present:
   - `Access-ACE-MCP.exe`
   - `Access-ACE-Agent.exe`
4. Skip to [Configuration](#configuration)

### Option 2: Build from Source

```bash
# Clone the repository
git clone https://github.com/YOUR_ORG/Access-ACE-MCP.git
cd Access-ACE-MCP

# Open in Visual Studio
# Build > Build Solution (Ctrl+Shift+B)

# Publish the main project
# Right-click Access-ACE-MCP > Publish
```

Published files will be in the `Published` folder.

## Configuration

### Claude Code CLI

Add to `~/.claude/settings.json`:

```json
{
  "mcpServers": {
    "access-ace": {
      "command": "C:\\Tools\\Access-ACE-MCP\\Access-ACE-MCP.exe",
      "args": []
    }
  }
}
```

### Claude Code Desktop

1. Open **Settings > MCP Servers**
2. Click **"Add Server"**
3. Enter:
   - **Name:** `access-ace`
   - **Command:** `C:\Tools\Access-ACE-MCP\Access-ACE-MCP.exe`
4. Click **"Add"**

### Auto-Connect to Specific Database

Add the database path to the `args` array:

```json
{
  "mcpServers": {
    "access-ace": {
      "command": "C:\\Tools\\Access-ACE-MCP\\Access-ACE-MCP.exe",
      "args": ["C:\\Users\\YourName\\MyDatabase.accdb"]
    }
  }
}
```

## Usage

### ⚠️ Before You Start: Backup Your Database

**Always create a backup of your Access database before using this tool.** This MCP server provides full programmatic access to your database, including the ability to modify code, forms, and data.

```bash
# Simple backup example
copy "C:\Users\YourName\MyDatabase.accdb" "C:\Users\YourName\MyDatabase_backup_2026-05-09.accdb"
```

### 🤖 Using the Claude Code Skill for Guidance

**Access-ACE-MCP includes a built-in Claude Code skill** that provides expert guidance on using all 40+ tools safely and effectively. When working with this project in Claude Code, the skill is automatically available.

**The skill helps you with:**
- **Safe operation patterns** - Recommended workflows for connecting, modifying objects, and disconnecting
- **All 40+ tools** - Detailed reference for each tool, object type codes, and common operations
- **Best practices** - Backup strategies, code review, version control, and gradual testing
- **Troubleshooting** - Solutions for common connection issues, dialog handling, and lockfile behavior
- **Code examples** - Ready-to-use patterns for forms, reports, VBA, and database inspection

Simply ask Claude for help with any task, and the skill will provide context-aware guidance tailored to Access-ACE-MCP. For example:
- "How do I safely modify a form control?"
- "What's the pattern for exporting forms?"
- "Help me understand the lockfile behavior"

### Important Usage Notes

#### Who Should Use This Tool

**Access-ACE-MCP is intended for developers and power users who need to modify Access applications**, not for end users. This tool provides full programmatic access to database structure, code, and design. Use it only if you understand database design and are comfortable modifying VBA code and database objects.

#### Getting Started Checklist

1. **Close Access before connecting** - Make sure Microsoft Access is completely closed before you ask Claude to connect to your database
2. **Ask Claude to disconnect when finished** - Always request Claude to disconnect from the database when your work is complete
3. **Keep Access visible** - Once connected, Access may display dialog boxes during automation. Keep the Access window visible and be ready to click appropriate buttons if needed
4. **Understand the lockfile** - While connected, Access creates a lockfile (`.ldb` or `.laccdb`) that prevents other connections from saving changes. Disconnecting removes this lockfile and allows normal Access usage

#### Connection Behavior

- **One connection at a time** - The MCP server maintains one connection per database
- **Lockfile prevents concurrent saves** - While connected via the MCP server, other Access instances cannot save changes to the same database
- **Dialog boxes require interaction** - Some Access operations may display dialog boxes. You need to interact with these dialogs by clicking buttons in the Access window

#### Disconnection Important

Always disconnect when finished:
```
User: "Please disconnect from the database"
```

Disconnecting:
- Closes the Access connection
- Removes the lockfile
- Allows other Access instances to save changes normally
- Prevents accidental modifications

### Basic Examples

#### 1. List All Forms
```
User: "In my database at C:\Users\<UserName>\MyDB.accdb, list all forms"

Claude will:
- Connect to the database
- Enumerate all forms
- Return the list of form names
```

#### 2. Export a Form for Review
```
User: "Export the 'ContactForm' from my database so I can review it"

Claude will:
- Open the form in design view
- Extract the complete form definition
- Display it as text for your review
```

#### 3. Modify a Form Control
```
User: "Change the label text on 'Label1' in the 'MainForm' to 'New Label Text'"

Claude will:
- Open the form
- Locate the control
- Change the property
- Save the form
```

#### 4. Execute VBA Code
```
User: "Add a procedure called 'LogEvent' to the 'Utilities' module that logs messages"

Claude will:
- Review the code with you first
- Add the procedure to the module
- Compile to verify syntax
```

#### 5. Get Database Metadata
```
User: "Show me all the queries in my database and their SQL"

Claude will:
- Connect to the database
- Extract all query definitions with SQL text
- Display them in a readable format
```

## Available Tools

Access-ACE-MCP exposes 40+ tools for database automation:

| Category | Tools |
|----------|-------|
| **Connection** | connect_access, disconnect_access, is_connected |
| **Enumeration** | get_forms, get_reports, get_modules, get_queries, get_macros, get_vba_projects |
| **Form Design** | open_form, close_form, get_form_controls, get_control_properties, set_control_property |
| **VBA** | get_vba_code, set_vba_code, add_vba_procedure, compile_vba |
| **Objects** | export_object_to_text, import_object_from_text, delete_object |
| **Application** | get_application_info, get_option, set_option, eval_expression |
| **Execution** | run_procedure, run_command, invoke_application_method, invoke_docmd_method |
| **Reports** | export_report_to_text, import_report_from_text, delete_report |

See [MCP_SETUP_GUIDE.md](MCP_SETUP_GUIDE.md#available-tools) for complete tool documentation.

## Troubleshooting

### "64-bit Microsoft Access is installed but is not supported"

**Solution:** Install 32-bit Office/Access. 

```bash
# Check your Access version
# File > Account > About Access button Look for "32-bit" or "64-bit"
```

### "Access-ACE-Agent.exe not found"

**Solution:** Rebuild and republish the solution. The agent must be in the same folder as the MCP server executable.

### Connection Timeout

**Solution:** Verify:
- 32-bit Access is installed
- Database file exists and is readable
- Database file is not corrupted
- You have read/write permissions on the file

### Database Becomes Corrupted

**Solution:** Immediately restore from your backup.

For more troubleshooting, see [MCP_SETUP_GUIDE.md#troubleshooting](MCP_SETUP_GUIDE.md#troubleshooting).

## Architecture

Access-ACE-MCP uses a two-process architecture:

```
Claude → Access-ACE-MCP.exe (.NET 10) 
         ↓ (Named Pipe IPC)
         Access-ACE-Agent.exe (.NET 4.8 COM Worker)
         ↓ (COM Automation)
         Microsoft Access
```

For technical details, see [ARCHITECTURE.md](ARCHITECTURE.md).

## Security & Best Practices

### Backups
- ✅ Always maintain recent backups before making changes
- ✅ Test backup restoration procedures
- ✅ Keep multiple versions for rollback capability

### Code Review
- ✅ Review all VBA code before adding it to your database
- ✅ Start with read-only operations (list, export)
- ✅ Test changes on a copy first

### Version Control
- ✅ Use export tools to backup database objects
- ✅ Store exports in Git for version tracking
- ✅ Track changes over time

## Limitations

- **32-bit only** - 64-bit Access not supported
- **Single connection** - One database at a time per instance
- **Windows only** - Requires Windows and COM automation
- **No datasheet operations** - Table data editing requires other tools
- **No report preview** - Can modify reports but not preview them

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Building from Source

### Prerequisites
- Visual Studio 2022 or later
- .NET 10 SDK
- .NET Framework 4.8 (for the COM worker)

### Build Steps

```bash
# Clone and navigate
git clone https://github.com/YOUR_ORG/Access-ACE-MCP.git
cd Access-ACE-MCP

# Build all projects
dotnet build

# Or use Visual Studio
# Build > Build Solution (Ctrl+Shift+B)

# Publish for distribution
# Right-click Access-ACE-MCP project > Publish
# Output goes to ./Published folder
```

## Testing

```bash
# Run unit tests
dotnet test Access-ACE-MCP-Tests.csproj
dotnet test Access-ACE-MCP-Tests-Net10.csproj
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- 🤖 **Interactive assistance:** Use the [Access-ACE-MCP Helper skill](.claude/skills/access-ace-helper.md) in Claude Code for expert guidance
- 📖 **Full documentation:** See [MCP_SETUP_GUIDE.md](MCP_SETUP_GUIDE.md)
- ⚡ **Quick start:** See [QUICK_START.md](QUICK_START.md)
- 🏗️ **Architecture details:** See [ARCHITECTURE.md](ARCHITECTURE.md)
- 🐛 **Report issues:** Use the [Issues](../../issues) page
- 💬 **Discussions:** Start a [Discussion](../../discussions)

## Roadmap

- [ ] Multi-database session support
- [ ] Transaction batching
- [ ] Datasheet view operations
- [ ] Report preview generation
- [ ] VBA syntax validation
- [ ] Database schema export (DDL scripts)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and updates.

## Credits

- Built with [Model Context Protocol](https://modelcontextprotocol.io/)
- Uses Microsoft Office Interop for COM automation
- Designed to work seamlessly with Claude AI

## Disclaimer

This tool provides full programmatic access to Microsoft Access databases. Use it carefully and maintain backups of your data. The developers are not responsible for data loss or corruption resulting from incorrect usage.

---

**Last Updated:** May 9, 2026  
**Version:** 1.0.0  
**Status:** Active Development