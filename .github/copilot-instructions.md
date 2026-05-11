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
- **Bitness enforcement:**
  - The agent and host enforce that only 32-bit Access/ACE is used. If 64-bit Access is detected, or the 32-bit ACE provider is missing, connection is blocked with a clear error.
- **COM/STA threading:**
  - All COM calls in the agent are dispatched on a dedicated STA thread. Tests use STA threads for all COM interactions.
- **Integration tests:**
  - `Access-ACE-MCP-Tests-Net10` launches the full host/agent stack and uses the pipe protocol. Tests are skipped (Inconclusive) if the environment is not set up (e.g., missing Access, missing test DB).
- **Error handling:**
  - Many tests and protocol methods treat known environment errors (e.g., "Provider cannot be found", "exclusive access") as Inconclusive/skipped, not failures.
- **Tool-based API:**
  - All automation is exposed as "tools" (see `AccessTools.cs`), with strong conventions for naming, argument schemas, and JSON serialization.

## Integration Points
- **Microsoft Access:**
  - Requires Access installed (32-bit, with ACE OLEDB provider).
  - Test DB path is hardcoded in tests (`C:\GitHub\NorthwindAccess\AccessMcp.accdb`).
- **Named pipes:**
  - Used for all host/agent communication.
- **Registry:**
  - Used to detect Access/ACE bitness and provider presence.

## Examples
- To add a new automation tool, implement a `[McpServerTool]` method in `AccessTools.cs` and ensure it is surfaced in the agent protocol.
- To run integration tests, ensure Access and the test DB are present, then run tests in `Access-ACE-MCP-Tests-Net10`.

## Key Files
- `Access_ACE_Agent/AccessComService.cs`: All Access COM automation logic.
- `Access-ACE-MCP/AccessTools.cs`: Tool API surface for automation.
- `Access-ACE-MCP/Program.cs`: Host process/pipe/agent orchestration.
- `Access-ACE-MCP-Tests-Net10/WorkerPipeIntegrationTests.cs`: End-to-end integration tests.

---
If any section is unclear or missing important project-specific details, please provide feedback for further refinement.
