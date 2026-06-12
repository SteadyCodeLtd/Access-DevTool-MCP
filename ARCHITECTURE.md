# Access-DevTool-MCP: Architecture Overview

## System Design

Access-DevTool-MCP uses a **two-process architecture** to bridge Claude's .NET 10 world with Access's COM requirements:

```
┌─────────────────────────────────────────────────────────────┐
│ Claude Code / MCP Client                                     │
│ (Local or Remote)                                            │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        │ stdio (JSON-RPC)
                        ↓
┌─────────────────────────────────────────────────────────────┐
│ Access-DevTool-MCP.exe (.NET 10, x86)                           │
│ ┌────────────────────────────────────────────────────────┐  │
│ │ - MCP Server Handler (ModelContextProtocol)           │  │
│ │ - Tool definitions & descriptions                      │  │
│ │ - Access bitness detection (x86 vs x64)               │  │
│ │ - Worker process management                            │  │
│ └────────────────────────────────────────────────────────┘  │
│                        ↓                                      │
│                  Named Pipe Channel                          │
│            (IPC: Access-DevTool-{ProcessID})                    │
│                        ↓                                      │
│ ┌────────────────────────────────────────────────────────┐  │
│ │ Access-DevTool-Agent.exe (.NET 4.8, x86 COM Interop)     │  │
│ │ ┌──────────────────────────────────────────────────┐  │  │
│ │ │ - COM Interop (Microsoft.Office.Interop.Access)│  │  │
│ │ │ - Database connection management                 │  │  │
│ │ │ - VBA manipulation via DAO/ADO/COM Interop     │  │  │
│ │ │ - SaveAsText / LoadFromText operations          │  │  │
│ │ │ - Error handling & JSON serialization           │  │  │
│ │ └──────────────────────────────────────────────────┘  │  │
│ │                     ↓                                   │  │
│ │         Access COM Interop (DCOM)                      │  │
│ │                     ↓                                   │  │
│ │         ┌──────────────────────┐                       │  │
│ │         │ Microsoft Access     │                       │  │
│ │         │ (.accdb / .mdb)      │                       │  │
│ │         └──────────────────────┘                       │  │
│ └────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

## Component Breakdown

### 1. Access-DevTool-MCP.exe (Main Server)
**Language:** C# / .NET 10.0  
**Platform:** Windows  
**Role:** MCP server entry point

**Key Classes:**
- `Program.cs` - Bootstraps the server, launches worker process, establishes named pipe
- `AccessTools.cs` - Defines all MCP tools and delegates to worker via named pipe
- `PipeChannel.cs` - Named pipe client for IPC communication

**Key Responsibilities:**
- Listens on stdio for MCP requests from Claude
- Spawns the Access-DevTool-Agent worker process with a unique pipe name
- Detects installed Access bitness (32-bit vs 64-bit) and logs the result
- Routes tool calls to the worker process via named pipes
- Handles application lifecycle (startup, shutdown, error handling)

### 2. Access-DevTool-Agent.exe (COM Interop Worker)
**Language:** C# / .NET 4.8  
**Platform:** Windows  
**Role:** COM Interop layer with Microsoft Access

**Key Responsibilities:**
- Receives JSON-RPC calls over the named pipe
- Creates/maintains COM connections to Microsoft Access
- Executes DAO/ADO operations on the database
- Uses SaveAsText / LoadFromText for object export/import
- Serializes results back to JSON for the parent process

**Why .NET 4.8?**
- Microsoft.Office.Interop.Access has well-proven stability on .NET Framework 4.8
- .NET Framework was the original platform designed for COM Interop and Office automation
- Decades of COM Interop experience and patterns built into the runtime
- .NET 10 can reliably call .NET 4.8 via inter-process communication (pipes)

**Important:** While .NET 10 has improved COM support, the two-process architecture is deliberately maintained for its significant architectural advantages (see "Why Two-Process Architecture?" below).

---

## Why Two-Process Architecture?

The separation of the MCP Server (.NET 10) and COM Interop Worker (.NET 4.8) into distinct processes provides significant architectural advantages:

### 1. **Crash Isolation** ⚠️ Critical
**Problem:** Direct COM Interop automation can fail unexpectedly
- Access may encounter unrecoverable errors
- COM objects can become corrupted or unstable
- A crash in Access takes down the entire process

**Solution:** Separate worker process isolates failures
- COM Interop crash in Agent.exe doesn't affect MCP Server
- Claude session remains responsive even if Access crashes
- Users can restart the agent without losing MCP connection context
- Database lockfile is cleaned up automatically when agent exits

**Impact:** Users can recover gracefully from Access failures without losing their Claude conversation context.

### 2. **Threading Isolation** 🔒 Performance
**Problem:** COM requires Single-Threaded Apartment (STA) threading
- Access COM objects MUST be accessed from a single thread
- Creating dedicated STA threads in a .NET 10 async context is complex and limits concurrency
- Mixing STA and thread-pool threads causes synchronization issues

**Solution:** Dedicated worker process with dedicated STA threads
- Agent.exe uses simple, predictable STA thread model
- MCP Server stays completely async and responsive
- No thread-pool starvation or deadlocks
- Claude requests are handled by server while agent works on COM Interop operations
- Clear separation of async (server) vs. synchronous (agent) operations

**Impact:** Better responsiveness; server never blocks on COM Interop operations.

### 3. **Debuggability** 🐛 Development
**Advantage:** Independent process debugging
- Can attach Visual Studio debugger to either process independently
- MCP Server and COM Worker can be debugged simultaneously
- Breakpoints in one process don't freeze the other
- Easier to diagnose protocol vs. COM issues
- Named pipe communication is traceable and loggable

**Impact:** Faster development cycles and easier troubleshooting.

### 4. **Separation of Concerns** 🏗️ Architecture
**MCP Server** (Access-DevTool-MCP.exe)
- Stateless, request-response handler
- Implements MCP protocol specification
- Manages tool definitions and descriptions
- Zero knowledge of COM Interop complexity
- Can be updated without touching COM Interop code

**COM Interop Worker** (Access-DevTool-Agent.exe)
- Stateful, COM Interop object management
- Manages database connections and Access lifecycle
- Zero knowledge of MCP protocol
- Can be updated without affecting server
- Easier to unit test COM Interop operations

**Impact:** Cleaner codebase; easier to maintain and extend each component independently.

### 5. **Graceful Degradation** 📉 Resilience
- Server continues running if agent crashes
- Agent can be respawned for new database connections
- Multiple independent database sessions possible (multiple agent instances)
- Server can detect and report agent failures cleanly

### Trade-offs Accepted
The advantages above justify accepting these trade-offs:
- **IPC Latency:** Named pipe communication adds ~1-5ms per call (negligible for interactive use)
- **Deployment Complexity:** Both .exe files must be present and discoverable
- **Process Overhead:** Two processes instead of one (~20MB additional memory)

---

### 3. Named Pipe Channel (IPC)
**Protocol:** JSON-RPC 2.0 over named pipes  
**Pipe Name:** `Access-DevTool-{ProcessID}` (unique per invocation)

**Example Call:**
```json
{
  "jsonrpc": "2.0",
  "method": "connect_access",
  "params": { "database_path": "C:\\Users\\<UserName>\\MyDB.accdb" },
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
"List all forms in C:\Users\<UserName>\MyDB.accdb"
        ↓
Claude Client calls connect_access("C:\Users\<UserName>\MyDB.accdb")
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
- **connect_access** - Creates COM Interop object, opens database
- **disconnect_access** - Closes COM Interop connection, kills Access process
- **is_connected** - Checks if COM Interop Access object exists

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

Access-DevTool-MCP detects the installed Access bitness to provide appropriate feedback. Both 32-bit and 64-bit Access are supported.

**Detection Logic (in AccessTools.cs):**
1. Check Registry: `HKLM\SOFTWARE\Microsoft\Office\{version}\Outlook`
2. Read `Bitness` value ("x86" or "x64")
3. Fallback: Check `InstallRoot` locations in native vs WOW6432Node hives
4. Return "x86", "x64", or "Not Found"

---

## Process Lifecycle

### Startup
```
User runs: Access-DevTool-MCP.exe [optional_database_path]
  ↓
Program.cs detects Agent.exe location
  ↓
Unique pipe name generated: Access-DevTool-{PID}
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
- **Access-DevTool-MCP.csproj**
  - ModelContextProtocol (NuGet, prerelease)
  - Microsoft.Extensions.Hosting (v9.*)

- **Access-DevTool-Agent.csproj**
  - Microsoft.Office.Interop.Access (COM Interop wrapper)
  - System.ComponentModel
  - No MCP dependencies

### Build Targets (in Access-DevTool-MCP.csproj)
1. `BuildAccessAceAgentForPublish` - Pre-publish: builds agent
2. `CopyAccessAceAgentToPublish` - Post-publish: copies agent.exe to output

### Published Output
```
Published/
  ├── Access-DevTool-MCP.exe          (Main server executable)
  ├── Access-DevTool-MCP.dll          (Runtime assembly)
  ├── Access-DevTool-MCP.pdb          (Debug symbols)
  ├── Access-DevTool-Agent.exe        (COM Interop worker process)
  ├── Access-DevTool-Agent.exe.config (COM Interop worker config)
  ├── Access-DevTool-Agent.pdb        (Worker symbols)
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
- **COM Interop Invocation:** Generic method invocation can call any Access COM method via Interop
- **Backup is essential:** No sandboxing between Claude and actual database

### Named Pipe Security
- Pipes are user-scoped on Windows
- Only the process that spawned the pipe can connect
- Not suitable for multi-user or network scenarios

---

## Testing

### Unit Tests (Access-DevTool-MCP-Tests.csproj)
- .NET Framework 4.8 compatibility tests
- COM Interop verification
- SaveAsText / LoadFromText round-trip tests

### Integration Tests (Access-DevTool-MCP-Tests-Net10.csproj)
- Full MCP protocol validation
- End-to-end database operations
- Multiple tool invocation sequences

---

## Limitations & Known Issues

1. **Single connection** - One database at a time per instance
2. **Windows only** - COM Interop is Windows-specific
3. **No transaction support** - Each tool call is independent
4. **Form in design mode** - open_form can lock form temporarily
5. **VBA debugging** - No integrated debugger, code review recommended

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
