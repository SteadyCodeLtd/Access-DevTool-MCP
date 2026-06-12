# Copilot Instructions for Access-ACE-MCP

## Overview
This repository is a cross-process automation bridge for Microsoft Access, consisting of:
- **Access-ACE-Agent** (`Access_ACE_Agent/`): .NET Framework 4.8.1 x86 EXE. Hosts and automates Access via COM, exposing a pipe-based protocol for higher-level orchestration. 
- **Access-ACE-MCP** (`Access-ACE-MCP/`): .NET 10 host. Launches the agent as a subprocess, communicates via named pipes, and exposes a modern tool-based protocol (MCP) for clients.
- **Tests**: Two test projects:
  - `Access-ACE-MCP-Tests-Net10`: End-to-end integration tests (pipe, protocol, round-trip, etc).
  - `Access-ACE-MCP-Tests`: Legacy .NET 4.8.1 COM tests (direct agent API).

## Architecture & Data Flow
- The .NET 10 host (`Access-ACE-MCP`) launches `Access-ACE-Agent.exe` as a child process, passing a named pipe and (optionally) a database path.
- If the agent process fails to start or the pipe connection times out, check that `Access-ACE-Agent.exe` is present in the same directory as the host, that 32-bit Access is installed, and consult the output pane for the agent's stderr output.
- All Access automation occurs in the agent, on a dedicated STA thread, via COM.
- The host and agent communicate using a JSON-over-pipe protocol. The host exposes a tool-based API (see `AccessTools.cs`).
- The agent disables macros on connect to avoid startup errors from missing providers.

## Build & Publish
- **Build everything:**
  - Use Visual Studio (solution includes all projects) or `dotnet build` for .NET 10 projects and MSBuild for the .NET Framework agent.
- **Publish:**
  - Publishing `Access-ACE-MCP.csproj` will automatically build and copy the latest `Access-ACE-Agent.exe` (and all its output files) into the publish folder.
  - The host (`Access-ACE-MCP`) expects `Access-ACE-Agent.exe` to be present in the same directory at runtime.

## Key Conventions & Patterns
- **Bitness detection:**
  - The agent and host detect the installed Access bitness (x86 or x64) for diagnostic purposes and to provide appropriate feedback. Both 32-bit and 64-bit Access installations are supported.
- **COM/STA threading:**
  - All COM calls in the agent are dispatched on a dedicated STA thread. Tests use STA threads for all COM interactions.
- **Integration tests:**
  - `Access-ACE-MCP-Tests-Net10` launches the full host/agent stack and uses the pipe protocol. Tests are skipped (Inconclusive) if the environment is not set up (e.g., missing Access, missing test DB).
- **Error handling:**
  - Many tests and protocol methods treat known environment errors (e.g., "Provider cannot be found", "exclusive access") as Inconclusive/skipped, not failures.
  - If the agent process fails to start or the pipe connection times out, check that `Access-ACE-Agent.exe` is present in the same directory as the host, that Microsoft Access is installed, and consult the output pane for the agent's stderr output.
- **Tool-based API:**
  - All automation is exposed as "tools" (see `AccessTools.cs`), with strong conventions for naming, argument schemas, and JSON serialization.

## Integration Points
- **Microsoft Access:**
  - Requires Access installed (with ACE OLEDB provider). Both 32-bit and 64-bit Access are supported.
  - Test DB path is hardcoded in tests (`C:\GitHub\NorthwindAccess\AccessMcp.accdb`).
  - To use a different test DB path, update the hardcoded path constant in `Access-ACE-MCP-Tests-Net10/WorkerPipeIntegrationTests.cs`. Consider using an environment variable or a config file to avoid modifying source for environment-specific paths.
- **Named pipes:**
  - Used for all host/agent communication.
- **Registry:**
  - Used to detect Access/ACE bitness and provider presence.

## Examples
- To add a new automation tool: (1) implement a `[McpServerTool]` method in `Access-ACE-MCP/AccessTools.cs`, (2) add the corresponding handler in `Access_ACE_Agent/AccessComService.cs`, and (3) update the JSON-over-pipe protocol to include the new command name and argument schema.
- To run integration tests, ensure Access and the test DB are present, then run tests in `Access-ACE-MCP-Tests-Net10`.

## Key Files
- `Access_ACE_Agent/AccessComService.cs`: All Access COM automation logic.
- `Access-ACE-MCP/AccessTools.cs`: Tool API surface for automation.
- `Access-ACE-MCP/Program.cs`: Host process/pipe/agent orchestration.
- `Access-ACE-MCP-Tests-Net10/WorkerPipeIntegrationTests.cs`: End-to-end integration tests.

---
If any section is unclear or missing important project-specific details, please provide feedback for further refinement.
