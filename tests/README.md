# PptMcp Tests

> **Integration-first, not integration-only.** PowerPoint COM cannot be meaningfully
> mocked, so anything that touches COM is tested against a real PowerPoint instance.
> A small number of genuinely COM-free helpers (enum mapping, string transforms,
> skill-file generation) are covered by `Category=Unit` tests under each project's
> `Unit/` folder. See [`docs/ADR-001-NO-UNIT-TESTS.md`](../docs/ADR-001-NO-UNIT-TESTS.md)
> for the rationale, and Rule 30 in
> [Critical Rules](../.github/instructions/critical-rules.instructions.md) for when a
> unit test is acceptable.

## Read this before copying a filter

`dotnet test --filter <expr>` prints `No test matches the given testcase filter` and
then **exits 0**. A filter naming a trait that does not exist is indistinguishable
from a run in which everything passed.

Always run filtered tests through the guard, which fails on a zero match:

```powershell
.\scripts\Invoke-GuardedTest.ps1 -Project tests\PptMcp.Core.Tests -Filter "Feature=Slide"
```

`scripts\check-test-filters.ps1` (wired into pre-commit) fails the build if any
document in this repository cites a `Feature=` value that no test carries.

### Traits that are NOT usable as filters

`Category` is applied to only a handful of classes - `Category=Integration` matches
**1 of 322** test methods in `PptMcp.Core.Tests`. Do not filter on it, and do not
reintroduce the historical `Category=Integration&RunType!=OnDemand&...` recipe: it
selected nothing while appearing to pass.

## Quick Start

```powershell
# Everything except the slow COM/session suite
.\scripts\Invoke-GuardedTest.ps1 -Project PptMcp.sln -Filter "RunType!=OnDemand"

# One project
.\scripts\Invoke-GuardedTest.ps1 -Project tests\PptMcp.Core.Tests -Filter "RunType!=OnDemand"

# Session/batch changes (MANDATORY when modifying session/batch code)
.\scripts\Invoke-GuardedTest.ps1 -Project PptMcp.sln -Filter "RunType=OnDemand"
```

## Documentation

**For complete testing guidance, see:**

- **[Testing Strategy](../.github/instructions/testing-strategy.instructions.md)** - Quick reference, templates, common mistakes
- **[Critical Rules](../.github/instructions/critical-rules.instructions.md)** - Mandatory development rules (Rule 14: Save, Rule 16: test scope, Rule 30: integration tests)

## Test Architecture

```
tests/
├── PptMcp.Core.Tests/            # Core business logic (322 tests)
├── PptMcp.McpServer.Tests/       # MCP protocol layer (99 tests)
├── PptMcp.ComInterop.Tests/      # COM utilities, sessions, batching (86 tests)
├── PptMcp.CLI.Tests/             # CLI wrapper (52 tests)
└── PptMcp.SkillGeneration.Tests/ # Generated skill/doc consistency (11 tests)

llm-tests/                        # LLM tool behavior validation (manual)
```

Counts are current as of the commit that introduced this section; they are
illustrative of relative size, not a contract.

## Feature-Specific Tests

Rule 16: run only the feature you changed. These are the `Feature` trait values that
actually exist:

```
ActionEnums    ActionValidation  Batch          Configuration  Design
Diag           Export            File           FileLocking    Master
McpProtocol    ParameterTransforms              PptBatch       PptMcpService
PptSession     ServiceDaemon     ServiceRegistry               ServiceRouting
SessionManager SkillGeneration   Slide          StreamJsonRpc  VersionCheck
```

```powershell
.\scripts\Invoke-GuardedTest.ps1 -Project tests\PptMcp.Core.Tests -Filter "Feature=Slide&RunType!=OnDemand"
.\scripts\Invoke-GuardedTest.ps1 -Project tests\PptMcp.CLI.Tests  -Filter "Feature=Design&RunType!=OnDemand"
```

Regenerate the list rather than trusting this block:

```powershell
Get-ChildItem tests -Recurse -Filter *.cs |
  Select-String -Pattern 'Trait\("Feature",\s*"([^"]+)"' -AllMatches |
  ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value } |
  Sort-Object -Unique
```

There are no `Shape`, `Text`, `Chart`, `Table`, `Animation`, `VBA`, `VBATrust` or
`Screenshot` features. Filters naming them - which appeared throughout this
repository's documentation until issue #127 - match zero tests and pass vacuously.

## When to Run Which Tests

| Scenario | Command |
|----------|---------|
| **Changed one feature** | `.\scripts\Invoke-GuardedTest.ps1 -Project <project> -Filter "Feature=<name>&RunType!=OnDemand"` |
| **Before commit** | `.\scripts\pre-commit.ps1` |
| **Modified session/batch code** | `.\scripts\Invoke-GuardedTest.ps1 -Project PptMcp.sln -Filter "RunType=OnDemand"` (see [Rule 3](../.github/instructions/critical-rules.instructions.md#rule-3-session-cleanup-tests)) |
| **Changed tool descriptions or skills** | `.\scripts\Test-LlmRegressionGate.ps1` |

## LLM Tests

The `llm-tests/` project validates that LLMs correctly use PowerPoint MCP Server and CLI tools using [pytest-aitest](https://github.com/trsdn/pytest-aitest).

### When to Run LLM Tests

- **Manual/on-demand only** - Not part of CI/CD
- After changing tool descriptions or adding new tools
- To validate LLM behavior patterns (e.g., incremental updates vs rebuild)

### Running LLM Tests

```powershell
# From llm-tests/
uv sync
uv run pytest -m aitest -v
```

### Canonical regression gate

Use the repository-level helper when you want the standard manual gate instead of the full suite:

```powershell
.\scripts\Test-LlmRegressionGate.ps1
```

The canonical gate runs three CLI scenarios plus the matching three MCP scenarios and is the recommended check after changing tool descriptions, skill guidance, or CLI help output.

### Prerequisites

- `AZURE_OPENAI_ENDPOINT` environment variable
- Windows desktop with PowerPoint installed
- pytest-aitest dependency (local path via uv)

**See [LLM Tests README](../llm-tests/README.md) for complete documentation.**

## VBA Testing

**There are currently no VBA tests in this repository.** Until issue #127 this file
documented a `tests/PptMcp.Core.Tests/Integration/Commands/Script/` directory with
four test files and a `ScriptAndSetupCommandsTests.cs` in the CLI project. None of
those files exist, and no test carries `Feature=VBA` or `Feature=VBATrust`. <!-- ghost-filter-ok -->

VBA functionality (`ScriptCommands`, VBA trust detection) is therefore **untested**.
If you add coverage, tag it `[Trait("Feature", "VBA")]`, add the trait to the list
above, and remove this notice.

VBA work against a real PowerPoint install needs trust enabled:

```powershell
# Enable VBA trust (development machines only)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "AccessVBOM" -Value 1

# Verify setting
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security" -Name "AccessVBOM"
```

**Security Note:** Only enable VBA trust in development environments. Production systems should keep this disabled.

## Key Principles

- ✅ **File Isolation** - Each test creates unique file (no sharing)
- ✅ **Binary Assertions** - Pass OR fail, never "accept both"
- ✅ **Verify PowerPoint State** - Always verify actual PowerPoint state after operations
- ✅ **Guarded Filters** - A filter that selects nothing must fail, never pass
- ❌ **No Save** - Unless testing persistence (see [Rule 14](../.github/instructions/critical-rules.instructions.md#rule-14-no-save-unless-testing-persistence))

## Getting Help

- **Test failures**: Check test output for detailed error messages
- **PowerPoint issues**: Ensure PowerPoint 2016+ installed and activated
- **Session/batch issues**: Run OnDemand tests to verify cleanup
- **Writing tests**: See [Testing Strategy](../.github/instructions/testing-strategy.instructions.md)

