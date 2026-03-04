# MCP Server Daemon Unification Specification

## Implementation Status

> **⚠️ SUPERSEDED** - This spec described the shared-daemon architecture. The actual implementation uses a hybrid model:
> - **MCP Server**: Fully in-process PptMcpService with direct method calls (no named pipe)
> - **CLI**: Daemon process with named pipe (`PptMcp-cli-{SID}`) and system tray
>
> See `architecture-patterns.instructions.md` for current architecture.

### Completed Features (Phase 1)

- ✅ **Rename Daemon to PptMcp Service** - All code, pipes, mutex, lock files updated
- ✅ **Session Origin Tracking** - Sessions labeled [CLI] or [MCP] in tray UI
- ✅ **About Dialog** - Version info and helpful links in tray menu
- ✅ **Removed Manual Daemon Commands** - No more `daemon start/stop/status` commands
- ✅ **Service Client Library** - Shared `ServiceClient/` in ComInterop for CLI and MCP
- ✅ **MCP Server Infrastructure** - Service mode detection and forwarding framework
- ✅ **All MCP Tools Forward to Service** - Removed standalone mode, all tools use `ForwardToService` pattern
- ✅ **Removed Standalone Mode** - No more `PptMcp_STANDALONE` or `UseServiceMode` toggles

### In Progress (Phase 2 - Unified Package)

- 🔄 **Bundle CLI with MCP Server Package** - Single NuGet package includes both `PptMcp.exe` and `pptcli.exe`
- 🔄 **Deprecate Separate CLI Package** - `PptMcp.CLI` deprecated, points to unified package
- ⏳ **Update ServiceLauncher** - Find `pptcli.exe` next to current executable
- ⏳ **Deduplicate Update Notifications** - Single notification per process lifetime
- ⏳ **Update Release Workflow** - Single unified release artifact

### Problem Discovered During Testing

MCP Server tests fail because:
1. Service lives in CLI project (`pptcli service run`)
2. Tests only build MCP Server, not CLI
3. `ServiceLauncher` can't find `pptcli.exe`
4. Installing MCP-only doesn't include the service

### Solution: Unified Package (Simpler Than Service Extraction)

Instead of extracting a separate service project, **bundle CLI with MCP Server**:

```
PptMcp.McpServer  → PptMcp.exe + pptcli.exe (both included)
PptMcp.CLI        → DEPRECATED (points to McpServer package)
```

**Benefits:**
- ✅ No version mismatch possible (everything upgrades together)
- ✅ No new project needed (keep service in CLI)
- ✅ Simpler release (one package)
- ✅ MCP always finds service (pptcli.exe next to PptMcp.exe)

**Installation (After):**
```powershell
# One package, both tools
dotnet tool install --global PptMcp.McpServer

# Both commands available
PptMcp    # MCP Server for AI assistants
pptcli    # CLI for coding agents
```

### Architecture

**Service-Only Mode**: MCP Server is now a thin JSON-over-named-pipe layer that forwards ALL requests to the PptMcp Service. This enables CLI and MCP Server to share sessions transparently.

```
MCP Client (VS Code, etc.)
    │
    ▼
┌──────────────────────────┐
│     MCP Server           │
│  ForwardToService()      │  ──────► Named Pipe: PptMcp-{UserSid}
│  (no local Core cmds)    │
└──────────────────────────┘
                                      │
                                      ▼
                           ┌──────────────────────────┐
                           │   PptMcp Service       │
                           │  (runs via pptcli)     │
                           │  ┌────────────────────┐  │
                           │  │  SessionManager    │  │
                           │  │  (shared sessions) │  │
                           │  └────────────────────┘  │
                           └──────────────────────────┘
```

---

## Phase 2: Service Extraction (Current Work)

### Problem: Deployment Mismatch

**User installs ONLY MCP Server:**
```powershell
dotnet tool install --global PptMcp.McpServer
```
- MCP Server tries to start `pptcli.exe service run`
- `pptcli.exe` doesn't exist because CLI isn't installed
- **All operations fail** ❌

### Solution: Separate Service Project

Create `PptMcp.Service` as an independent project that produces `excelservice.exe`:

```
src/
  PptMcp.Service/              ← NEW PROJECT
    PptMcp.Service.csproj      ← net10.0-windows (WinForms for tray)
    Program.cs                   ← Entry point
    PptMcpService.cs           ← Moved from CLI/Service/
    ServiceTray.cs               ← Moved from CLI/Service/
    ...

  PptMcp.CLI/
    PptMcp.CLI.csproj          ← BUNDLES excelservice.exe
    Commands/                     ← CLI commands only

  PptMcp.McpServer/
    PptMcp.McpServer.csproj    ← BUNDLES excelservice.exe
    Tools/                        ← MCP tools only
```

### Deployment Scenarios

**User installs CLI only:**
```
~/.dotnet/tools/
  pptcli.exe              ← CLI tool
  excelservice.exe          ← Bundled service
```
✅ CLI finds service next to itself

**User installs MCP only:**
```
~/.dotnet/tools/
  PptMcp.exe              ← MCP Server
  excelservice.exe          ← Bundled service
```
✅ MCP finds service next to itself

**User installs BOTH:**
```
~/.dotnet/tools/
  pptcli.exe
  PptMcp.exe
  excelservice.exe          ← One copy, shared
```
✅ Either can start it, sessions are shared

### Version Mismatch Handling

**Scenario:** User has CLI v1.5 (with Service v1.5) and updates MCP to v1.6 (with Service v1.6)

**Problem:**
- CLI starts service v1.5
- MCP connects and expects v1.6 protocol
- Potential incompatibility!

**Solution: "Latest Wins" Strategy**

```csharp
// On client startup (both CLI and MCP):
public async Task<bool> EnsureServiceRunningAsync()
{
    var runningVersion = await GetRunningServiceVersionAsync();
    var bundledVersion = GetBundledServiceVersion();
    
    if (runningVersion == null)
    {
        // No service running, start bundled version
        return await StartServiceAsync();
    }
    
    if (bundledVersion > runningVersion)
    {
        // Bundled version is newer - upgrade!
        await RequestServiceShutdownAsync();
        await WaitForServiceExitAsync();
        return await StartServiceAsync();
    }
    
    // Running version is same or newer - use it
    return true;
}
```

**Protocol Additions:**

```json
// Ping response includes version
{
  "success": true,
  "version": "1.6.0",
  "uptime": "00:15:30"
}

// Graceful shutdown command
{
  "command": "service.shutdown",
  "reason": "upgrade"
}
```

**Compatibility Rules:**
- Same major version = compatible (v1.5 client can use v1.6 service)
- Different major version = force upgrade (v2.0 client shuts down v1.x service)
- Service maintains backward compatibility within major version

### Files to Move

**From `CLI/Service/` to new `Service/` project:**
- `PptMcpService.cs` (2282 lines - the main service)
- `ServiceTray.cs` - Windows Forms tray icon
- `DialogService.cs` - About dialog
- `ServiceProtocol.cs` - Command routing
- `ServiceSecurity.cs` (service-side parts) - Lock files, mutex

**Keep in ComInterop (shared client code):**
- `ServiceClient/ExcelServiceClient.cs` - Named pipe client
- `ServiceClient/ServiceLauncher.cs` - Find and start service
- `ServiceClient/ServiceSecurity.cs` (read-only parts) - Check if running

### NuGet Packaging

Both CLI and MCP Server `.csproj` files need to bundle `excelservice.exe`:

```xml
<ItemGroup>
  <!-- Bundle the service executable -->
  <None Include="$(OutputPath)\..\PptMcp.Service\net10.0-windows\excelservice.exe"
        Pack="true"
        PackagePath="tools\net10.0-windows\any\" />
</ItemGroup>
```

### ServiceLauncher Simplification

```csharp
private static ProcessStartInfo? GetServiceStartInfo()
{
    // Primary: Look next to current executable
    var serviceExe = Path.Combine(AppContext.BaseDirectory, "excelservice.exe");
    if (File.Exists(serviceExe))
    {
        return new ProcessStartInfo
        {
            FileName = serviceExe,
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
    }
    
    // Fallback: Global tools location
    var globalTools = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".dotnet", "tools", "excelservice.exe");
    
    if (File.Exists(globalTools))
    {
        return new ProcessStartInfo
        {
            FileName = globalTools,
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
    }
    
    return null;
}
```

---

## Overview

Unify the MCP Server with the existing CLI daemon architecture to provide persistent session management across both interfaces.

## Problem Statement (Phase 1 - Completed)

### Current Architecture

```
┌──────────────────────────┐     ┌──────────────────────────┐
│     MCP Server #1        │     │     MCP Server #2        │
│  ┌────────────────────┐  │     │  ┌────────────────────┐  │
│  │  SessionManager    │  │     │  │  SessionManager    │  │
│  │  (isolated)        │  │     │  │  (isolated)        │  │
│  └────────────────────┘  │     │  └────────────────────┘  │
│           │              │     │           │              │
│     Excel Process A      │     │     Excel Process B      │
└──────────────────────────┘     └──────────────────────────┘
           ↑                                ↑
           │                                │
      File: test.xlsx ─────────────── File: test.xlsx
                        ❌ CONFLICT!
```

**Issues:**
1. Each MCP server process has its own `SessionManager`
2. Each opens separate Excel processes
3. File locking conflicts when multiple MCP servers access the same file
4. Sessions lost when MCP server process restarts
5. No visibility into sessions (no tray UI)

### CLI Daemon Architecture (Already Working)

```
┌─────────────────┐    Named Pipe    ┌──────────────────────────────────┐
│  CLI Command 1  │ ──────────────── │                                  │
└─────────────────┘                  │       ExcelDaemon                │
                                     │                                  │
┌─────────────────┐    Named Pipe    │  • SessionManager (singleton)    │
│  CLI Command 2  │ ──────────────── │  • Tray Icon                     │
└─────────────────┘                  │  • 10-min idle timeout           │
                                     │  • Single instance mutex         │
                                     │                                  │
                                     └──────────────────────────────────┘
```

**Benefits:**
- Single `SessionManager` across all CLI invocations
- Sessions persist between commands
- Tray UI shows active sessions
- Automatic cleanup via idle timeout

## Proposed Architecture

```
                                    ┌──────────────────────────────────┐
┌─────────────────┐                 │                                  │
│  CLI Commands   │──Named Pipe────▶│                                  │
└─────────────────┘                 │       ExcelDaemon                │
                                    │       (Unified)                  │
┌─────────────────┐                 │                                  │
│  MCP Server #1  │──Named Pipe────▶│  • SessionManager (singleton)    │
└─────────────────┘                 │  • Tray Icon (all sessions)      │
                                    │  • 10-min idle timeout           │
┌─────────────────┐                 │  • Single instance mutex         │
│  MCP Server #2  │──Named Pipe────▶│  • Core Commands                 │
└─────────────────┘                 │                                  │
                                    └──────────────────────────────────┘
                                                   │
                                           Excel Processes
                                        (one per open file)
```

**Benefits:**
1. ✅ Single `SessionManager` for CLI and MCP
2. ✅ No file locking conflicts between MCP instances
3. ✅ Sessions survive MCP server restarts
4. ✅ Unified tray UI shows all sessions
5. ✅ MCP Server becomes thin wrapper (less code to maintain)
6. ✅ LLM tests can use multiple turns without race conditions

## Implementation Plan

### Phase 1: Extract Daemon Client Library

**Goal:** Create reusable client library that both CLI and MCP can use.

**New Project:** `PptMcp.Daemon.Client`

```csharp
namespace PptMcp.Daemon.Client;

public class DaemonClient : IDisposable
{
    public static DaemonClient Connect(bool autoStartDaemon = true);
    public Task<string> SendCommandAsync(string toolName, string action, Dictionary<string, object> parameters);
    public bool IsDaemonRunning { get; }
}
```

**Files to create:**
- `src/PptMcp.Daemon.Client/DaemonClient.cs`
- `src/PptMcp.Daemon.Client/DaemonProtocol.cs` (shared message types)
- `src/PptMcp.Daemon.Client/DaemonLauncher.cs` (auto-start logic)

### Phase 2: Refactor CLI to Use Client Library

**Goal:** CLI uses `DaemonClient` instead of direct pipe operations.

**Changes:**
- Extract pipe communication from `ExcelDaemon.cs` into shared protocol
- CLI commands use `DaemonClient.SendCommandAsync()`
- Verify existing CLI tests still pass

### Phase 3: MCP Server Uses Daemon

**Goal:** MCP Server tools forward requests to daemon.

**Before:**
```csharp
public class PptFileTool
{
    private static readonly SessionManager _sessionManager = new();
    
    public static string Open(string path, bool showExcel)
    {
        var batch = _sessionManager.CreateSession(path, showExcel);
        // Complex session management...
    }
}
```

**After:**
```csharp
public class PptFileTool
{
    public static async Task<string> Open(string path, bool showExcel)
    {
        using var client = DaemonClient.Connect();
        return await client.SendCommandAsync("file", "open", new { path, showExcel });
    }
}
```

### Phase 4: Enhanced Tray UI

**Goal:** Tray shows session source (CLI vs MCP).

**Changes:**
- Track session origin in `SessionManager`
- Show in tray tooltip: "2 MCP sessions, 1 CLI session"
- Context menu: "Close all MCP sessions", "Close all CLI sessions"

## Protocol Design

### Request Format (JSON over Named Pipe)

```json
{
  "id": "uuid-v4",
  "tool": "file",
  "action": "open",
  "parameters": {
    "excelPath": "C:\\test.xlsx",
    "showExcel": false
  },
  "source": "mcp-server"
}
```

### Response Format

```json
{
  "id": "uuid-v4",
  "success": true,
  "result": {
    "sessionId": "abc123",
    "filePath": "C:\\test.xlsx"
  }
}
```

### Error Response

```json
{
  "id": "uuid-v4",
  "success": false,
  "error": {
    "message": "File not found",
    "code": "FILE_NOT_FOUND"
  }
}
```

## Migration Strategy

> **Note:** This section describes the original migration plan. As of February 2026, standalone mode has been **removed entirely**. The MCP Server now operates exclusively in service mode using `ForwardToService` pattern.

### ~~Backward Compatibility~~ (Superseded)

~~1. MCP Server can work in **two modes:**~~
   ~~- **Daemon mode** (default): Forward to daemon~~
   ~~- **Standalone mode** (fallback): Use embedded `SessionManager`~~

**Current Implementation:** Service-only mode. All MCP tools use `ForwardToService()` to send commands to the PptMcp Service via named pipe.

### Testing Strategy

1. **Unit tests:** Mock `DaemonClient`, test tool logic
2. **Integration tests:** Start daemon, run MCP tests
3. **LLM tests:** Multi-turn workflows (the original problem!)

## File Structure Changes

```
src/
├── PptMcp.Daemon.Client/           # NEW: Shared client library
│   ├── DaemonClient.cs
│   ├── DaemonProtocol.cs
│   └── DaemonLauncher.cs
├── PptMcp.CLI/
│   ├── Daemon/
│   │   └── ExcelDaemon.cs            # MODIFIED: Use shared protocol
│   └── Commands/                      # MODIFIED: Use DaemonClient
├── PptMcp.McpServer/
│   └── Tools/                         # MODIFIED: Use DaemonClient
└── PptMcp.Core/                     # UNCHANGED
```

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Daemon startup latency | First call slow | Pre-launch daemon on install, lazy connect |
| Daemon crashes | All sessions lost | Robust error handling, reconnect logic |
| Protocol versioning | Breaking changes | Version field in protocol, "latest wins" upgrade |
| Security | Named pipe access | Keep existing security (per-user pipe) |
| Debugging complexity | Two processes | Unified logging, trace correlation |
| Version mismatch CLI/MCP | Incompatible protocols | Service version check, automatic upgrade |
| Duplicate services | Race condition on startup | Mutex + lock file, version-aware handoff |

## Success Criteria

### Phase 1 (Completed)
1. ✅ MCP Server can complete 5-turn workflow without file locking
2. ✅ CLI and MCP sessions visible in same tray UI
3. ✅ Session survives MCP server restart
4. ✅ No performance regression (< 50ms added latency)
5. ✅ Removed standalone mode - service-only architecture

### Phase 2 (In Progress)
6. ⏳ MCP-only install works (no CLI required)
7. ⏳ CLI-only install works (no MCP required)
8. ⏳ Version mismatch auto-upgrades service
9. ⏳ Single update notification per process lifetime
10. ⏳ All MCP Server tests pass

## Timeline Estimate

### Phase 1 (Completed)
- ✅ Extract client library: 2 days
- ✅ Refactor CLI to use client: 1 day
- ✅ MCP integration: 3 days
- ✅ Tray enhancements: 1 day
- ✅ Remove standalone mode: 1 day

### Phase 2 (Current)
- 🔄 Create PptMcp.Service project: 1 day
- ⏳ Move service code from CLI: 1 day
- ⏳ Bundle service in NuGet packages: 1 day
- ⏳ Version check and upgrade logic: 1 day
- ⏳ Fix duplicate update notifications: 0.5 day
- ⏳ Update tests: 1 day
- ⏳ Documentation: 0.5 day

**Phase 2 Total:** ~6 days

## Open Questions (Updated)

1. ~~Should daemon auto-start when MCP server connects?~~
   - **RESOLVED:** Yes, always auto-start

2. ~~Should we support multiple daemons (per-workspace)?~~
   - **RESOLVED:** No, single daemon per user

3. ~~What happens if daemon exits while MCP is running?~~
   - **RESOLVED:** Client automatically reconnects and restarts service

4. **NEW:** What if both CLI and MCP try to upgrade service simultaneously?
   - **Recommendation:** First one wins (mutex), second waits and connects

5. **NEW:** Should we show "upgrade in progress" to user?
   - **Recommendation:** Yes, brief tray notification

6. **NEW:** How long to wait for old service to shut down?
   - **Recommendation:** 5 seconds timeout, then force-kill process
