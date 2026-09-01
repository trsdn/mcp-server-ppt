#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Git pre-commit hook to check for COM object leaks, Core Commands coverage, naming consistency, Success flag violations, CLI workflow, and MCP Server functionality

.DESCRIPTION
    Runs checks before allowing commits:
    0. Process cleanup - kills stale PowerPoint, pptcli, and MCP server processes to prevent file locks
    1. COM leak checker - ensures no PowerPoint COM objects are leaked
    2. Coverage audit - ensures 100% Core Commands are exposed via MCP Server
    3. Naming consistency - ensures enum names match Core method names exactly
    4. Success flag validation - ensures Success=true never paired with ErrorMessage (Rule 0)
    5. CLI workflow smoke test - validates end-to-end CLI functionality
    6. MCP Server smoke test - validates all MCP tools work correctly

    Ensures code quality and prevents regression.

.EXAMPLE
    .\pre-commit.ps1

.NOTES
    This script is called by the Git pre-commit hook.
    To install: Copy .git/hooks/pre-commit (bash) or configure Git to use this PowerShell version.
#>

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

# CRITICAL: Check branch FIRST - never commit directly to main (Rule 6)
Write-Host "Checking current branch..." -ForegroundColor Cyan
$currentBranch = git branch --show-current

if ($currentBranch -eq "main") {
    Write-Host ""
    Write-Host "BLOCKED: Cannot commit directly to 'main' branch!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   Rule 6: All Changes Via Pull Requests" -ForegroundColor Yellow
    Write-Host "   'Never commit to main. Create feature branch -> PR -> CI/CD + review -> merge.'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   To fix:" -ForegroundColor Cyan
    Write-Host "   1. git stash                                    # Save your changes" -ForegroundColor White
    Write-Host "   2. git checkout -b feature/your-feature-name    # Create feature branch" -ForegroundColor White
    Write-Host "   3. git stash pop                                # Restore changes" -ForegroundColor White
    Write-Host "   4. git add <files>                              # Stage changes" -ForegroundColor White
    Write-Host "   5. git commit -m 'your message'                 # Commit to feature branch" -ForegroundColor White
    Write-Host ""
    exit 1
}

Write-Host "Branch check passed - on '$currentBranch' (not main)" -ForegroundColor Green
Write-Host ""

# Kill stale PowerPoint and MCP server processes to avoid file locks on Release binaries
Write-Host "Killing stale PowerPoint and server processes..." -ForegroundColor Cyan

$killedProcesses = @()
foreach ($procName in @("POWERPNT", "pptcli", "PptMcp.McpServer", "PptMcp.Service")) {
    $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($procs) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        $killedProcesses += "$procName ($($procs.Count))"
    }
}

if ($killedProcesses.Count -gt 0) {
    Write-Host "   Killed: $($killedProcesses -join ', ')" -ForegroundColor Yellow
    # Brief pause to let file handles release
    Start-Sleep -Milliseconds 500
}
else {
    Write-Host "   No stale processes found" -ForegroundColor Gray
}

Write-Host "Process cleanup done" -ForegroundColor Green
Write-Host ""

Write-Host "Checking for COM object leaks..." -ForegroundColor Cyan

try {
    $leakCheckScript = Join-Path $rootDir "scripts\check-com-leaks.ps1"
    & $leakCheckScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "COM object leaks detected! Fix them before committing." -ForegroundColor Red
        exit 1
    }

    Write-Host "COM leak check passed" -ForegroundColor Green
}
catch {
    # A check that cannot run has not passed. Swallowing the error here would let
    # a broken gate look like a green one, which is exactly how the removed
    # coverage audit stayed unnoticed for so long.
    Write-Host "Error running COM leak check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking documented tool and operation counts..." -ForegroundColor Cyan

# Replaces the former audit-core-coverage.ps1 and check-mcp-core-implementations.ps1.
# Both were written against a hand-maintained ToolActions.cs/ActionExtensions.cs that
# the source generator architecture removed, so one crashed and the other reported
# "100% coverage" while detecting zero methods. Parity between the CLI and the MCP
# server is now structural - both are generated from the same Core interfaces - so
# what actually needs guarding is that the published counts still match the
# generated tool surface.
try {
    $countsScript = Join-Path $rootDir "scripts\check-documented-counts.ps1"
    & $countsScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Documented counts do not match the generated tool surface!" -ForegroundColor Red
        Write-Host "   Update the affected documents, or regenerate FEATURES.md." -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host ""
    Write-Host "Error running documented count check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking CLI Settings properties are forwarded to the daemon..." -ForegroundColor Cyan

try {
    $settingsScript = Join-Path $rootDir "scripts\check-cli-settings-usage.ps1"
    & $settingsScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "CLI Settings properties are defined but never forwarded!" -ForegroundColor Red
        Write-Host "   User-supplied values would be silently ignored." -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host ""
    Write-Host "Error running CLI Settings usage check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking Success flag violations (Rule 0)..." -ForegroundColor Cyan

try {
    $successFlagScript = Join-Path $rootDir "scripts\check-success-flag.ps1"
    & $successFlagScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Success flag violations detected!" -ForegroundColor Red
        Write-Host "   CRITICAL: Success=true with ErrorMessage confuses LLMs and causes data corruption." -ForegroundColor Red
        Write-Host "   Fix the violations before committing (add Success=false in catch blocks)." -ForegroundColor Red
        exit 1
    }

    Write-Host "Success flag check passed - all flags match reality" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running success flag check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# NOTE: CLI coverage checks removed - commands are now auto-generated by Roslyn source generators
# The CLI generator produces all command classes and registration from Core interfaces
# Validation is handled by:
# - Build-time generator errors if interfaces are malformed
# - CLI workflow smoke test below (end-to-end validation)

Write-Host ""
Write-Host "Auto-staging generated SKILL.md files..." -ForegroundColor Cyan

try {
    # SKILL.md files are generated during Release build from templates + source generators.
    # The Release build already ran (required for CLI smoke test below), so SKILL.md files
    # are up to date on disk. Auto-stage them so developers never have to think about it.
    # SKILL.md + references are generated during Release build.
    # Auto-stage all of them so developers never have to think about it.
    $skillPaths = @(
        "skills/ppt-mcp/SKILL.md",
        "skills/ppt-cli/SKILL.md",
        "skills/ppt-mcp/references/",
        "skills/ppt-cli/references/"
    )
    $skillDiff = git diff --name-only -- @skillPaths 2>&1
    $untrackedSkills = git ls-files --others --exclude-standard -- @skillPaths 2>&1

    $allChanges = @()
    if ($skillDiff) { $allChanges += $skillDiff }
    if ($untrackedSkills) { $allChanges += $untrackedSkills }

    if ($allChanges.Count -gt 0) {
        git add -- @skillPaths
        Write-Host "Skill files were regenerated and auto-staged ($($allChanges.Count) files)" -ForegroundColor Green
        $allChanges | ForEach-Object { Write-Host "   + $_" -ForegroundColor DarkGray }
    } else {
        Write-Host "Skill files are already up to date" -ForegroundColor Green
    }
}
catch {
    # Not swallowed: if regenerated SKILL.md files fail to stage, the commit ships
    # documentation that no longer matches the build output.
    Write-Host "Error auto-staging SKILL.md files: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Running CLI workflow smoke test..." -ForegroundColor Cyan

try {
    $cliWorkflowScript = Join-Path $rootDir "scripts\Test-CliWorkflow.ps1"
    $cliWorkflowOutput = & $cliWorkflowScript 2>&1 | Out-String
    $cliWorkflowExitCode = $LASTEXITCODE

    if ($cliWorkflowExitCode -ne 0) {
        Write-Host ""
        Write-Host "CLI workflow smoke test failed!" -ForegroundColor Red
        Write-Host "   This test validates the end-to-end CLI workflow." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $cliWorkflowOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "CLI workflow smoke test passed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running CLI workflow smoke test: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Running MCP Server smoke test..." -ForegroundColor Cyan

# Stop PptMcp Service before smoke test to prevent DLL locking
& "$PSScriptRoot\Stop-PptMcpProcesses.ps1"

try {
    # Run the smoke tests - validate all MCP tools are both discoverable AND routable.
    # Discovery alone is not a gate: before #124, 11 tools were listed by ListToolsAsync
    # while every invocation returned "Unknown command category". Match the whole
    # SmokeTest_ prefix so a new smoke test is picked up without editing this script.
    $smokeTestFilter = "FullyQualifiedName~McpServerIntegrationTests.SmokeTest_|FullyQualifiedName~ServiceRoutingTests"

    Write-Host "   dotnet test --filter `"$smokeTestFilter`"" -ForegroundColor Gray

    # Force English CLI output. The success check below parses the summary line, and on a
    # localized machine dotnet prints "Bestanden! ... erfolgreich: 1", which the English
    # pattern misses - turning the zero-test guard into a false alarm on a passing test.
    $previousUiLanguage = $env:DOTNET_CLI_UI_LANGUAGE
    $env:DOTNET_CLI_UI_LANGUAGE = "en"
    try {
        # Capture output to verify tests actually ran (dotnet test returns 0 even if no tests match!)
        $testOutput = dotnet test --filter $smokeTestFilter --verbosity minimal 2>&1 | Out-String
        $testExitCode = $LASTEXITCODE
    }
    finally {
        $env:DOTNET_CLI_UI_LANGUAGE = $previousUiLanguage
    }

    # Check that enough tests actually ran (critical - filter typos cause silent failures!)
    # The routing theory runs once per generated category (33 today), so a healthy run is
    # well above 20. A bare "Passed: 1" would satisfy a naive check even if one filter
    # clause stopped matching, which is exactly how a gate goes quiet without failing.
    $passedCount = 0
    if ($testOutput -match "Passed!.*Passed:\s*(\d+)") {
        $passedCount = [int]$Matches[1]
    }

    if ($passedCount -lt 20) {
        Write-Host ""
        Write-Host "CRITICAL: Only $passedCount smoke tests passed (expected 20+)." -ForegroundColor Red
        Write-Host "   Filter: $smokeTestFilter" -ForegroundColor Yellow
        Write-Host "   A filter clause likely matched zero tests." -ForegroundColor Yellow
        Write-Host "   Verify the tests exist: McpServerIntegrationTests.SmokeTest_* and ServiceRoutingTests" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    if ($testExitCode -ne 0) {
        Write-Host ""
        Write-Host "MCP Server smoke test failed! Core functionality is broken." -ForegroundColor Red
        Write-Host "   This test validates all MCP tools work correctly." -ForegroundColor Red
        Write-Host "   Fix the issues before committing." -ForegroundColor Red
        Write-Host ""
        Write-Host $testOutput -ForegroundColor Gray
        exit 1
    }

    Write-Host "MCP Server smoke test passed - all tools functional" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error running smoke test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Ensure PowerPoint is installed and accessible." -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "Checking for undocumented ((dynamic)) casts..." -ForegroundColor Cyan

try {
    $dynamicCastScript = Join-Path $rootDir "scripts\check-dynamic-casts.ps1"
    & $dynamicCastScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Undocumented ((dynamic)) casts detected!" -ForegroundColor Red
        Write-Host "   Add a justification comment (// PIA gap:, // TODO:, or // Reason:) before each cast." -ForegroundColor Red
        Write-Host "   See docs/PIA-COVERAGE.md for guidance." -ForegroundColor Red
        exit 1
    }

    Write-Host "Dynamic cast check passed - no new undocumented casts" -ForegroundColor Green
}
catch {
    # A check that cannot run has not passed.
    Write-Host "Error running dynamic cast check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Checking NuGet flat container package ids..." -ForegroundColor Cyan

try {
    $flatContainerScript = Join-Path $rootDir "scripts\check-nuget-flatcontainer-ids.ps1"
    & $flatContainerScript

    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "Mixed-case NuGet flat container package id detected!" -ForegroundColor Red
        Write-Host "   The flat container API 404s on any casing other than lowercase," -ForegroundColor Red
        Write-Host "   and callers swallow that 404, so the failure is silent at runtime." -ForegroundColor Red
        exit 1
    }

    Write-Host "Flat container id check passed" -ForegroundColor Green
}
catch {
    # A check that cannot run has not passed.
    Write-Host "Error running flat container id check: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "All pre-commit checks passed!" -ForegroundColor Green
exit 0
