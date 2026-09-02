# Runs the PowerPoint integration suite on this machine and records what actually ran.
#
# WHY THIS EXISTS
#
# The powerpoint-integration job in .github/workflows/integration-tests.yml cannot run.
# It needs a self-hosted Windows runner with PowerPoint installed, and that runner will
# not be provisioned (see docs/AZURE_SELFHOSTED_RUNNER_SETUP.md for the decision). So the
# only machine that can execute these tests is a maintainer's desktop.
#
# An unenforced "please run the tests locally" is not a gate. This script makes the local
# run produce EVIDENCE: a manifest bound to an exact commit SHA, listing every suite, the
# number of tests that ran, and the result. scripts/check-integration-evidence.ps1 then
# verifies that evidence during pre-push, so a commit cannot reach the remote claiming
# coverage that nobody produced.
#
# The suite list deliberately mirrors the steps of the unreachable powerpoint-integration
# job, so this is the same coverage that CI would have provided, plus the ComInterop
# assembly - which was too slow to include before issue #148 cut it from 33m to 7m.
#
# EVIDENCE IS BOUND TO A COMMIT, NOT TO A WORKING TREE.
#
# A manifest describes the content of one commit. If the working tree has uncommitted
# changes under src/ or tests/, the tests exercised something that no commit contains,
# so the run cannot be evidence for anything. That case is refused rather than recorded,
# unless -AllowDirty is passed, which marks the manifest binding as "none" - readable by
# a human, and rejected by the pre-push check.
#
# A GATE THAT INSPECTS NOTHING MUST FAIL. If no suite ran, or a suite reported no test
# count, the manifest records a failure. There is no path through this script that writes
# result=pass without per-suite evidence.
#
# ASCII only - see scripts/check-test-filters.ps1 for why.

[CmdletBinding()]
param(
    # Where the manifest is written. Gitignored: evidence is a local artifact, never a
    # committed claim.
    [string]$OutputPath = '.integration-evidence/manifest.json',

    # Produce a manifest even though the working tree is dirty. The manifest is marked
    # binding=none and will NOT satisfy the pre-push check.
    [switch]$AllowDirty,

    # Skip the build step when the tree is already built.
    [switch]$NoBuild
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$started = Get-Date

Write-Host ""
Write-Host "Local PowerPoint Integration Gate" -ForegroundColor Cyan
Write-Host "=================================="
Write-Host ""

# ---------------------------------------------------------------- commit binding
$commit = (git rev-parse HEAD).Trim()
$branch = (git rev-parse --abbrev-ref HEAD).Trim()

# Only src/ and tests/ decide the binding. A dirty README does not change what the tests
# exercised, and refusing to run over it would just train people to use -AllowDirty.
$dirty = @(git status --porcelain -- src tests | Where-Object { $_ -ne '' })
$binding = 'commit'

if ($dirty.Count -gt 0) {
    if (-not $AllowDirty) {
        Write-Host "REFUSED: uncommitted changes under src/ or tests/." -ForegroundColor Red
        Write-Host ""
        Write-Host "Evidence is bound to a commit. Running now would test content that no"
        Write-Host "commit holds, so the manifest could not honestly name a SHA."
        Write-Host ""
        Write-Host "Changed:"
        $dirty | Select-Object -First 20 | ForEach-Object { Write-Host "  $_" }
        Write-Host ""
        Write-Host "Commit first, then re-run. To produce an advisory-only manifest that"
        Write-Host "will NOT satisfy pre-push, pass -AllowDirty."
        exit 1
    }

    $binding = 'none'
    Write-Host "WARNING: working tree is dirty, so this run cannot be evidence for a commit." -ForegroundColor Yellow
    Write-Host "         The manifest will be marked binding=none and pre-push will reject it."
    Write-Host ""
}

Write-Host "Commit: $commit ($branch)"

# ---------------------------------------------------------------- environment
function Get-PowerPointVersion {
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE'
    try {
        $exe = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
        if ($exe -and (Test-Path $exe)) {
            return (Get-Item $exe).VersionInfo.ProductVersion
        }
    }
    catch { }
    return $null
}

$powerPointVersion = Get-PowerPointVersion
if (-not $powerPointVersion) {
    # Without PowerPoint every suite would fail one by one and waste an hour proving it.
    Write-Host "FAIL: PowerPoint was not found on this machine." -ForegroundColor Red
    Write-Host "      This gate exists precisely because these tests need real PowerPoint."
    exit 1
}

$dotnetVersion = (dotnet --version).Trim()
Write-Host "PowerPoint: $powerPointVersion   .NET SDK: $dotnetVersion   Machine: $env:COMPUTERNAME"
Write-Host ""

# ---------------------------------------------------------------- build
if (-not $NoBuild) {
    Write-Host "Building solution (Release)..." -ForegroundColor Cyan

    # A stale CLI daemon holds the output DLLs and turns the build into MSB3027.
    Get-Process pptcli -ErrorAction SilentlyContinue | ForEach-Object {
        try { Stop-Process -Id $_.Id -Force } catch { }
    }
    dotnet build-server shutdown | Out-Null
    Start-Sleep -Seconds 2

    dotnet build PptMcp.sln -c Release | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "FAIL: build failed. No suite was run." -ForegroundColor Red
        exit 1
    }
    Write-Host "Build succeeded."
    Write-Host ""
}

# ---------------------------------------------------------------- suite definitions
# Mirrors the steps of the powerpoint-integration job, in the same order: smoke tests
# first so an obviously broken build fails in two minutes rather than forty.
$guard = Join-Path $PSScriptRoot 'Invoke-GuardedTest.ps1'

$suites = @(
    @{
        Name    = 'cli-smoke-workflow'
        Kind    = 'script'
        Script  = (Join-Path $PSScriptRoot 'Test-CliWorkflow.ps1')
        Args    = @()
    }
    @{
        Name    = 'mcp-smoke'
        Kind    = 'test'
        Project = 'tests/PptMcp.McpServer.Tests/PptMcp.McpServer.Tests.csproj'
        Filter  = 'FullyQualifiedName~McpServerIntegrationTests.SmokeTest_'
        Timeout = 15
    }
    @{
        Name    = 'cominterop-full'
        Kind    = 'test'
        Project = 'tests/PptMcp.ComInterop.Tests/PptMcp.ComInterop.Tests.csproj'
        Full    = $true
        Timeout = 30
    }
    @{
        Name    = 'core-integration'
        Kind    = 'test'
        Project = 'tests/PptMcp.Core.Tests/PptMcp.Core.Tests.csproj'
        Filter  = 'RunType!=OnDemand'
        Timeout = 90
    }
    @{
        Name    = 'cli-integration'
        Kind    = 'test'
        Project = 'tests/PptMcp.CLI.Tests/PptMcp.CLI.Tests.csproj'
        Full    = $true
        Timeout = 45
    }
    @{
        Name    = 'mcp-integration'
        Kind    = 'test'
        Project = 'tests/PptMcp.McpServer.Tests/PptMcp.McpServer.Tests.csproj'
        Full    = $true
        Timeout = 60
    }
)

$results = @()
$gateFailed = $false

foreach ($suite in $suites) {
    Write-Host "[$($suite.Name)]" -ForegroundColor Cyan
    $suiteStart = Get-Date
    $testCount = $null
    $output = ''

    if ($suite.Kind -eq 'script') {
        $output = & $suite.Script 2>&1 | Out-String
        $exitCode = $LASTEXITCODE

        # Test-CliWorkflow.ps1 prints "Passed: N" / "Failed: N".
        if ($output -match 'Passed:\s*(\d+)') { $testCount = [int]$Matches[1] }
        if ($output -match 'Failed:\s*([1-9]\d*)') { $exitCode = 1 }
    }
    else {
        $guardArgs = @('-Project', $suite.Project, '-NoBuild', '-LoggerFileName', "gate-$($suite.Name).trx")
        if ($suite.Filter) { $guardArgs += @('-Filter', $suite.Filter) }
        if ($suite.Full)   { $guardArgs += '-Full' }
        if ($suite.Timeout) { $guardArgs += @('-TimeoutMinutes', $suite.Timeout) }

        $output = & $guard @guardArgs 2>&1 | Out-String
        $exitCode = $LASTEXITCODE

        # Invoke-GuardedTest.ps1 already refuses to report success without a summary, so
        # this count is a transcription of its verdict rather than an independent claim.
        if ($output -match 'OK:\s*(\d+)\s+test') { $testCount = [int]$Matches[1] }
        elseif ($output -match 'FAIL:\s*(\d+)\s+test') { $testCount = [int]$Matches[1] }
    }

    $suiteDuration = [math]::Round(((Get-Date) - $suiteStart).TotalSeconds, 1)

    # No count means no evidence, whatever the exit code said.
    $passed = ($exitCode -eq 0) -and ($null -ne $testCount) -and ($testCount -gt 0)
    if (-not $passed) { $gateFailed = $true }

    $results += [pscustomobject]@{
        name            = $suite.Name
        kind            = $suite.Kind
        tests           = $testCount
        result          = if ($passed) { 'pass' } else { 'fail' }
        exitCode        = $exitCode
        durationSeconds = $suiteDuration
    }

    if ($passed) {
        Write-Host "  PASS - $testCount test(s) in ${suiteDuration}s" -ForegroundColor Green
    }
    else {
        Write-Host "  FAIL - exit $exitCode, tests=$testCount, ${suiteDuration}s" -ForegroundColor Red
        Write-Host "  ---- output ----"
        $output -split "`n" | Select-Object -Last 25 | ForEach-Object { Write-Host "  $_" }
        Write-Host "  ----------------"
    }
    Write-Host ""
}

# ---------------------------------------------------------------- manifest
if ($results.Count -eq 0) {
    Write-Host "FAIL: no suite ran. Refusing to write a manifest." -ForegroundColor Red
    exit 1
}

$totalTests = ($results | Measure-Object -Property tests -Sum).Sum
if (-not $totalTests -or $totalTests -le 0) {
    $gateFailed = $true
}

$totalDuration = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)

$manifest = [ordered]@{
    schemaVersion   = 1
    commit          = $commit
    branch          = $branch
    binding         = $binding
    result          = if ($gateFailed) { 'fail' } else { 'pass' }
    generatedAtUtc  = (Get-Date).ToUniversalTime().ToString('o')
    durationSeconds = $totalDuration
    totalTests      = $totalTests
    machine         = [ordered]@{
        name       = $env:COMPUTERNAME
        os         = [System.Environment]::OSVersion.VersionString
        dotnet     = $dotnetVersion
        powerPoint = $powerPointVersion
    }
    suites          = $results
}

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$manifest | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputPath -Encoding utf8

Write-Host "=================================="
Write-Host "Suites: $($results.Count)   Tests: $totalTests   Duration: $([math]::Round($totalDuration / 60, 1))m"
Write-Host "Manifest: $OutputPath"
Write-Host ""

if ($gateFailed) {
    Write-Host "GATE FAILED - the manifest records the failure and will not satisfy pre-push." -ForegroundColor Red
    exit 1
}

if ($binding -eq 'none') {
    Write-Host "GATE PASSED, but binding=none (dirty tree). Pre-push will still refuse." -ForegroundColor Yellow
    exit 0
}

Write-Host "GATE PASSED for commit $commit" -ForegroundColor Green
exit 0
