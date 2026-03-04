# Validation and Rollback Plan: MCP SDK 0.5.0-preview.1 Upgrade

**Created**: 2025-12-14
**Branch**: `001-upgrade-mcp-sdk`
**Target SDK**: `ModelContextProtocol` 0.5.0-preview.1

---

## Validation Checklist

### Pre-Merge Validation

| Step | Command | Expected Result | Gate |
|------|---------|-----------------|------|
| 1. Build | `dotnet build` | 0 warnings, 0 errors | ✅ PASS |
| 2. MCP Server Tests | `dotnet test tests/PptMcp.McpServer.Tests/` | 66/66 passing | ✅ PASS |
| 3. CLI Tests | `dotnet test tests/PptMcp.CLI.Tests/` | 2/2 passing | ✅ PASS |
| 4. Core Feature Tests | `dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"` | 49/49 passing | ✅ PASS |
| 5. Tables Feature Tests | `dotnet test --filter "Feature=Tables&RunType!=OnDemand"` | 20/20 passing | ✅ PASS |

### Post-Merge Validation (CI/CD)

| Step | Workflow | Expected Result |
|------|----------|-----------------|
| 1. PR Build | `build-mcp-server.yml` | Green ✅ |
| 2. PR Build | `build-cli.yml` | Green ✅ |
| 3. Integration Tests | `integration-tests.yml` | Green ✅ |
| 4. CodeQL | `codeql.yml` | No new security issues |

### Manual Smoke Tests

| Test | Procedure | Expected Result |
|------|-----------|-----------------|
| 1. Server Startup | `dotnet run --project src/PptMcp.McpServer` | Starts without error |
| 2. Claude Desktop | Connect via MCP config | Tools discovered |
| 3. file | Open test file | Session created |
| 4. worksheet | List sheets | Sheets returned |

---

## Decision Gates

### Go/No-Go Criteria

| Gate | Criteria | Status |
|------|----------|--------|
| **BUILD** | 0 warnings, 0 errors | ✅ Required for merge |
| **TESTS** | No new test failures | ✅ Required for merge |
| **SECURITY** | No new CodeQL alerts | ✅ Required for merge |
| **REVIEW** | Approved by 1+ reviewer | ✅ Required for merge |

### Acceptable Conditions

| Condition | Decision |
|-----------|----------|
| Pre-existing test failures | ⚠️ Acceptable (documented in impact-report.md) |
| New SDK deprecation warnings | ⚠️ Acceptable if suppressed with justification |
| Preview package stability | ⚠️ Acceptable for development builds |

### Blocking Conditions

| Condition | Decision |
|-----------|----------|
| New test failures | ❌ BLOCK - Fix before merge |
| Build errors | ❌ BLOCK - Fix before merge |
| New security alerts | ❌ BLOCK - Assess severity |
| MCP tools not discoverable | ❌ BLOCK - Fix before merge |

---

## Rollback Procedure

### Trigger Conditions

Rollback is required if ANY of the following occur after merge:

1. **Build Failures**: `main` branch fails to build
2. **Test Regressions**: Previously passing tests fail
3. **Runtime Errors**: MCP server crashes or hangs
4. **Protocol Errors**: Clients cannot connect/discover tools

### Rollback Steps

#### Step 1: Immediate Mitigation

```bash
# Create hotfix branch from main (pre-merge state)
git checkout main
git checkout -b hotfix/revert-sdk-upgrade

# Revert the merge commit
git revert <merge-commit-sha> --no-edit

# Push hotfix
git push origin hotfix/revert-sdk-upgrade
```

#### Step 2: Dependency Revert

Edit `Directory.Packages.props`:

```xml
<!-- Revert from -->
<PackageVersion Include="ModelContextProtocol" Version="0.5.0-preview.1" />

<!-- Revert to -->
<PackageVersion Include="ModelContextProtocol" Version="0.4.1-preview.1" />
```

#### Step 3: Verification

```bash
# Verify build
dotnet build

# Verify tests
dotnet test tests/PptMcp.McpServer.Tests/

# Verify smoke test
dotnet run --project src/PptMcp.McpServer
```

#### Step 4: Communication

1. **GitHub Issue**: Create issue documenting the failure
2. **PR Comment**: Add rollback details to original PR
3. **Team Notification**: Alert maintainers via GitHub mentions

---

## Release Timeline

| Phase | Target | Status |
|-------|--------|--------|
| Feature Branch | 2025-12-14 | ✅ Complete |
| Code Review | 2025-12-14 | 🔄 Pending |
| Merge to Main | After approval | ⏳ Pending |
| Release Tag | Next release | ⏳ Pending |

---

## Sign-Off

| Role | Name | Date | Approval |
|------|------|------|----------|
| Developer | GitHub Copilot | 2025-12-14 | ✅ Implemented |
| Reviewer | | | ⏳ Pending |
| Release Manager | | | ⏳ Pending |

---

## Appendix: Files Changed

| File | Change Type | Description |
|------|-------------|-------------|
| `Directory.Packages.props` | Modified | SDK version bump |
| `tests/.../McpServerIntegrationTests.cs` | Modified | API rename |
| `tests/.../McpServerSmokeTests.cs` | Modified | Test isolation |
| `tests/.../PptFileToolOperationTrackingTests.cs` | Modified | Test isolation |
| `tests/.../ProgramTransportTestCollection.cs` | Created | xUnit collection |
| `src/PptMcp.CLI/Commands/Sheet/SheetCommand.cs` | Modified | JSON output for mutations |
| `src/PptMcp.McpServer/Program.cs` | Modified | Exit code handling (0/1) |
| `src/PptMcp.Core/.../PivotTableCommands.Fields.cs` | Modified | stderr for warnings |
| `src/PptMcp.Core/.../PivotTableCommands.Lifecycle.cs` | Modified | stderr for warnings |
| `src/PptMcp.Core/.../RegularPivotTableFieldStrategy.cs` | Modified | stderr for warnings |
| `src/PptMcp.Core/.../OlapPivotTableFieldStrategy.cs` | Modified | stderr for warnings |
| `specs/001-upgrade-mcp-sdk/impact-report.md` | Created | Impact documentation |
| `specs/001-upgrade-mcp-sdk/validation-plan.md` | Created | This file |
| `specs/001-upgrade-mcp-sdk/tasks.md` | Modified | Task tracking |
