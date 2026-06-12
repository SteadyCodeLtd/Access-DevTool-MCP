# Access-DevTool-MCP: Quick Start Guide

**TL;DR:** Build the project, configure Claude Code, back up your database, then ask Claude to automate your Access database.

## 30-Second Setup

### 1. Build and Publish
```bash
# In Visual Studio
Build > Build Solution
# Then publish the main Access-DevTool-MCP project
```

Published executable location:
```
C:\\GitHub\\Access-DevTool-MCP\\Published\\Access-DevTool-MCP.exe
```

### 2. Configure Claude Code

**Add to `~/.claude/settings.json` (CLI):**
```json
"mcpServers": {
  "access-devtool-mcp": {
    "command": "C:\\GitHub\\Access-DevTool-MCP\\Published\\Access-DevTool-MCP.exe",
    "args": []
  }
}
```

Or **Settings > MCP Servers (Desktop)** - Add server with command above.

### 3. Backup Your Database
```
Copy your .accdb file to a safe location with a date-stamped name
```

### 4. Start Using It
Ask Claude: *"List all the forms in my database at C:\path\to\database.accdb"*

That's it! Claude will handle the connection and return the results.

---

## Common Tasks (One-Liners)

**List database contents:**
> "List all forms, reports, and queries in C:\Users\<UserName>\MyDB.accdb"

**Export a form for review:**
> "Export the 'MyForm' from C:\Users\<UserName>\MyDB.accdb so I can review its definition"

**Modify a form control:**
> "In C:\Users\<UserName>\MyDB.accdb, change the label text of control 'Label1' on form 'MyForm' to 'New Text'"

**Export VBA code:**
> "Export the VBA code from the 'Utilities' module in C:\Users\<UserName>\MyDB.accdb"

**Add a VBA procedure:**
> "Add a procedure called 'HelloWorld' to the 'Utilities' module that prints 'Hello World'"

---

## Checklist

- [ ] Microsoft Access is installed (32-bit or 64-bit)
- [ ] Solution is built in Visual Studio
- [ ] Published folder exists with both executables
- [ ] MCP server is configured in Claude Code settings
- [ ] **Database backup exists**
- [ ] You're ready to automate!

---

## Troubleshooting Quick Fixes

| Problem | Fix |
|---------|-----|
| "Agent.exe not found" | Rebuild solution & republish |
| "Connection timeout" | Verify database path is correct |
| "Database corrupted" | Restore from backup immediately |

---

## Next Steps

→ See **MCP_SETUP_GUIDE.md** for complete documentation  
→ Read **Tool descriptions** in the guide for all available operations  
→ **Always maintain backups** before making changes

---

**Key Rule:** ⚠️ Backup your database before you start.
