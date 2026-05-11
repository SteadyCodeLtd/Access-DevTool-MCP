# Access-ACE-MCP: Architecture Overview

## System Design

Access-ACE-MCP uses a **two-process architecture** to bridge Claude's .NET 10 world with Access's COM requirements:

```
┌─────────────────────────────────────────────────────────────┐
│ Claude Code / MCP Client                                     │
│ (Local or Remote)                                            │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        │ stdio (JSON-RPC)
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ Access-ACE-MCP.exe (.NET 10, x86)                           │
│ ┌────────────────────────────────────────────────────────┐  │
│ │ - MCP Server Handler (ModelContextProtocol)           │  │
│ │ - Tool definitions & descriptions                      │  │
│ │ - Access bitness detection (x86 vs x64)               │  │
│ │ - Worker process management                            │  │
│ └────────────────────────────────────────────────────────┘  │
│                        ↓                                      │
│                  Named Pipe Channel                          │
│            (IPC: access-ace-{ProcessID})                    │
│                        ↓                                      │
│ ┌────────────────────────────────────────────────────────┐  │
│ │ Access-ACE-Agent.exe (.NET 4.8, x86 COM worker)      │  │
│ │ ┌──────────────────────────────────────────────────┐  │  │
│ │ │ - COM Interop (Microsoft.Office.Interop.Access)│  │  │
│ │ │ - Database connection management                 │  │  │
│ │ │ - VBA manipulation via DAO/ADO/COM             │  │  │
│ │ │ - SaveAsText / LoadFromText operations          │  │  │
│ │ │ - Error handling & JSON serialization           │  │  │
│ │ └──────────────────────────────────────────────────┘  │  │
│ │                     ↓                                   │  │
│ │            COM Automation (DCOM)                       │  │
│ │                     ↓                                   │  │
│ │         ┌──────────────────────┐                       │  │
│ │         │ Microsoft Access COM │                       │  │
│ │         │ (.accdb / .mdb)      │                       │  │
│ │         └──────────────────────┘                       │  │
│ └────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

## Component Breakdown

### 1. Access-ACE-MCP.exe (Main Server)
**Language:** C# / .NET 10.0  
**Platform:** Windows, x86 only  
**Role:** MCP server entry point

**Key Classes:**
- `Program.cs` - Bootstraps the server, launches worker process, establishes named pipe
- `AccessTools.cs` - Defines all MCP tools and delegates to worker via named pipe
- `PipeChannel.cs` - Named pipe client for IPC communication

**Key Responsibilities:**
- Listens on stdio for MCP requests from Claude
- Spawns the Access-ACE-Agent worker process with a unique pipe name
- Detects installed Access bitness (32-bit vs 64-bit) and rejects x64
- Routes tool calls to the worker process via named pipes
- Handles application lifecycle (startup, shutdown, error handling)

### 2. Access-ACE-Agent.exe (COM Worker)
**Language:** C# / .NET 4.8  
**Platform:** Windows, x86 only  
**Role:** COM interop layer with Microsoft Access

**Key Responsibilities:**
- Receives JSON-RPC calls over the named pipe
- Creates/maintains COM connections to Microsoft Access
- Executes DAO/ADO operations on the database
- Uses SaveAsText / LoadFromText for object export/import
- Serializes results back to JSON for the parent process

**Why .NET 4.8?**
- Microsoft.Office.Interop.Access (COM Interop) has better stability on .NET Framework 4.8
- .NET 10 can call .NET 4.8 via inter-process communication (pipes)
- Avoids potential issues with direct COM interop from .NET 10

### 3. Named Pipe Channel (IPC)
**Protocol:** JSON-RPC 2.0 over named pipes  
**Pipe Name:** `access-ace-{ProcessID}` (unique per invocation)

**Example Call:**
```json
{
  "jsonrpc": "2.0",
  "method": "connect_access",
  "params": { "database_path": "C:\\Users\\DJ\\MyDB.accdb" },
  "id": 1
}
```

**Example Response:**
```json
{
  "jsonrpc": "2.0",
  "result": {
    "success": true,
    "database": "MyDB.accdb",
    "version": "16.0"
  },
  "id": 1
}
```

---

## Data Flow Example: Connect and List Forms

```
User Request:
"List all forms in C:\Users\DJ\MyDB.accdb"
        ↓
Claude Client calls connect_access("C:\Users\DJ\MyDB.accdb")
        ↓
MCP Server receives JSON-RPC call on stdio
        ↓
AccessTools.ConnectAccess()
  - Checks Access bitness via Windows Registry
  - If x64: returns error immediately
  - If x86: delegates to channel
        ↓
PipeChannel.CallAsync("connect_access", { "database_path": "..." })
  - Sends JSON-RPC to worker process via pipe
  - Waits for response (blocks)
        ↓
Agent.exe receives JSON-RPC on pipe
        ↓
Agent loads COM Interop:
  - Creates Access.Application COM object
  - Opens database file
  - Initializes DAO/ADO objects
        ↓
Agent sends success response via pipe
        ↓
MCP Server returns result to Claude
        ↓
Claude client calls get_forms()
        ↓
AccessTools.GetForms() → PipeChannel.CallAsync("get_forms")
        ↓
Agent.exe queries Access object model:
  - Iterates Application.Modules
  - Collects form names, types, properties
  - Serializes to JSON
        ↓
Agent sends JSON-RPC response
        ↓
MCP Server returns JSON to Claude
        ↓
Claude displays form list to user
```

---

## Tool Categories & Implementation

### Connection Management
- **connect_access** - Creates COM object, opens database
- **disconnect_access** - Closes COM connection, kills Access process
- **is_connected** - Checks if Access object exists

### Database Inspection
- **get_forms**, **get_reports**, **get_modules**, **get_queries** - Enumerate collections
- **get_form_controls**, **get_control_properties** - Inspect form structure

### Object Manipulation
- **export_object_to_text** - Calls `SaveAsText` on Access objects
- **import_object_from_text** - Calls `LoadFromText` to replace objects
- **delete_object** - Calls `Delete()` on containers

### VBA Management
- **get_vba_code** - Reads `Module.CodeModule.Lines` collection
- **set_vba_code** - Replaces entire module source via `AddFromString()`
- **add_vba_procedure** - Appends to existing module
- **compile_vba** - Calls `CodeModule.CommitLines()`

### Property Modification
- **open_form** - Opens form in design view
- **set_control_property** - Sets individual properties via reflection
- **set_option** - Sets Application properties

### VBA Execution
- **run_procedure** - Calls `Application.Run()`
- **run_command** - Calls `DoCmd.RunCommand()`
- **invoke_application_method** - Generic method invocation via reflection
- **invoke_docmd_method** - Generic DoCmd method invocation

---

## Access Bitness Detection

Access-ACE-MCP enforces **32-bit only** because:
1. COM interop is simpler and more stable on 32-bit
2. Both MCP server and agent are x86 (32-bit)
3. Registry keys differ for 32-bit vs 64-bit Office

**Detection Logic (in AccessTools.cs):**
1. Check Registry: `HKLM\SOFTWARE\Microsoft\Office\{version}\Outlook`
2. Read `Bitness` value ("x86" or "x64")
3. Fallback: Check `InstallRoot` locations in native vs WOW6432Node hives
4. Return "x86", "x64", or "Not Found"

---

## Process Lifecycle

### Startup
```
User runs: Access-ACE-MCP.exe [optional_database_path]
  ↓
Program.cs detects Agent.exe location
  ↓
Unique pipe name generated: access-ace-{PID}
  ↓
Agent process spawned with: --pipe {name} [{database_path}]
  ↓
Main process waits up to 30 seconds for agent to start pipe server
  ↓
Connection established
  ↓
MCP server starts listening on stdio
```

### Shutdown
```
Claude closes connection or session ends
  ↓
IHostApplicationLifetime.ApplicationStopping event fires
  ↓
PipeChannel is disposed (closes pipes)
  ↓
Worker process is killed (entire tree)
  ↓
Access.Application is released (database closes)
  ↓
Processes exit cleanly
```

---

## Error Handling

### Registry Check Fails
- Returns "Not Found" for bitness
- connect_access returns error suggesting manual Access installation check

### Database File Not Found
- Agent sends JSON error response
- MCP passes error to Claude
- Claude displays message to user

### COM Interop Fails
- Agent catches COM exceptions
- Serializes exception message to JSON
- MCP receives and returns to Claude

### Pipe Connection Timeout
- 30-second timeout on initial connection
- Suggests agent didn't start (likely Access installation issue)

---

## Building & Publishing

### Project Dependencies
- **Access-ACE-MCP.csproj**
  - ModelContextProtocol (NuGet, prerelease)
  - Microsoft.Extensions.Hosting (v9.*)

- **Access-ACE-Agent.csproj**
  - Microsoft.Office.Interop.Access (COM Interop wrapper)
  - System.ComponentModel
  - No MCP dependencies

### Build Targets (in Access-ACE-MCP.csproj)
1. `BuildAccessAceAgentForPublish` - Pre-publish: builds agent
2. `CopyAccessAceAgentToPublish` - Post-publish: copies agent.exe to output

### Published Output
```
Published/
  ├── Access-ACE-MCP.exe          (Main server executable)
  ├── Access-ACE-MCP.dll          (Runtime assembly)
  ├── Access-ACE-MCP.pdb          (Debug symbols)
  ├── Access-ACE-Agent.exe        (Worker process)
  ├── Access-ACE-Agent.exe.config (COM worker config)
  ├── Access-ACE-Agent.pdb        (Worker symbols)
  ├── *.deps.json                 (Dependency manifest)
  ├── *.runtimeconfig.json        (Runtime configuration)
  └── [Dependencies]              (All .NET framework DLLs)
```

---

## Security Considerations

### Access Database Permissions
- Agent runs with user credentials
- Read/write access depends on file system permissions
- No special elevation required unless database is in protected folder

### Code Execution
- **VBA Execution:** Full code execution in Access context
- **COM Invocation:** Generic method invocation can call any Access COM method
- **Backup is essential:** No sandboxing between Claude and actual database

### Named Pipe Security
- Pipes are user-scoped on Windows
- Only the process that spawned the pipe can connect
- Not suitable for multi-user or network scenarios

---

## Testing

### Unit Tests (Access-ACE-MCP-Tests.csproj)
- .NET Framework 4.8 compatibility tests
- COM interop verification
- SaveAsText / LoadFromText round-trip tests

### Integration Tests (Access-ACE-MCP-Tests-Net10.csproj)
- Full MCP protocol validation
- End-to-end database operations
- Multiple tool invocation sequences

---

## Limitations & Known Issues

1. **32-bit only** - No 64-bit Access support
2. **Single connection** - One database at a time per instance
3. **Windows only** - COM automation is Windows-specific
4. **No transaction support** - Each tool call is independent
5. **Form in design mode** - open_form can lock form temporarily
6. **VBA debugging** - No integrated debugger, code review recommended

---

## Future Enhancements

- [ ] Support for multiple simultaneous database connections (separate instances)
- [ ] Transaction batching (multiple operations in one batch)
- [ ] Datasheet view operations (reading/modifying table data)
- [ ] Report preview generation
- [ ] VBA syntax validation before compilation
- [ ] Database schema export (create DDL scripts)

---

**Architecture Last Updated:** May 9, 2026
