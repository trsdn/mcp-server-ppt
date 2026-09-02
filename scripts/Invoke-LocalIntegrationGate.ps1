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
    [switch]$NoBuild,

    # Run only the named suites, for iterating on the gate itself. A partial run cannot
    # be evidence for a commit, so this forces binding=none no matter how clean the tree
    # is - otherwise the gate could mint a passing manifest for coverage it never ran.
    [string[]]$Only = @(),

    # Re-run every suite even when its inputs are unchanged. Use when you suspect a
    # flaky or environment-dependent result rather than a code change.
    [switch]$Force,

    # Print each suite's derived input paths and content hash, then exit. Runs nothing,
    # launches nothing - for checking that the dependency derivation is sane.
    [switch]$ListInputs
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
    if (-not $AllowDirty -and -not $ListInputs) {
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

# ---------------------------------------------------------------- suite input hashing
# Every ComInterop test creates and tears down a real PowerPoint process, because session
# lifecycle is the thing under test. That is ~95 launches, and re-running them to re-prove
# a result about code that did not change is the single most disruptive thing this gate
# does to the machine it runs on.
#
# So a suite is re-run only when its inputs changed. "Inputs" is derived, not declared:
# the transitive ProjectReference closure of the test project, which is exactly the code
# that suite can exercise. A change confined to src/PptMcp.Core cannot alter the behaviour
# of a suite that never references it.
#
# The hash is built from git blob SHAs, so it describes committed content precisely. That
# is only meaningful on a clean tree, so reuse is disabled outright unless binding=commit.
function Get-ProjectClosure {
    param([string]$ProjectPath)

    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $frontier = @([System.IO.Path]::GetFullPath($ProjectPath))

    while ($frontier.Count -gt 0) {
        $next = @()
        foreach ($proj in $frontier) {
            if (-not $seen.Add($proj)) { continue }
            if (-not (Test-Path $proj)) { continue }

            $dir = Split-Path -Parent $proj
            try { [xml]$xml = Get-Content $proj -Raw } catch { continue }

            foreach ($group in $xml.Project.ItemGroup) {
                foreach ($ref in $group.ProjectReference) {
                    if ($ref -and $ref.Include) {
                        $next += [System.IO.Path]::GetFullPath((Join-Path $dir $ref.Include))
                    }
                }
            }
        }
        $frontier = $next
    }

    return @($seen)
}

function Get-SuiteInputsHash {
    param([string[]]$Paths)

    # git ls-files -s prints mode, blob SHA, stage and path for every tracked file. The
    # blob SHA is the content, so this is a precise content fingerprint of the inputs
    # without reading a single file ourselves.
    $listing = git ls-files -s -- @Paths 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $listing) { return $null }

    $joined = ($listing | Sort-Object) -join "`n"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($joined)
    $sha = [System.Security.Cryptography.SHA256]::HashData($bytes)
    return [System.Convert]::ToHexString($sha).ToLowerInvariant()
}

function Get-SuiteInputPaths {
    param([hashtable]$Suite)

    $paths = @()

    if ($Suite.Project) {
        foreach ($proj in Get-ProjectClosure (Join-Path $repoRoot $Suite.Project)) {
            $dir = Split-Path -Parent $proj
            $paths += [System.IO.Path]::GetRelativePath($repoRoot, $dir).Replace('\', '/')
        }
    }

    foreach ($extra in @($Suite.ExtraInputs)) {
        if ($extra) { $paths += $extra }
    }

    return @($paths | Sort-Object -Unique)
}

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
        Name        = 'cli-smoke-workflow'
        Kind        = 'script'
        Script      = (Join-Path $PSScriptRoot 'Test-CliWorkflow.ps1')
        Args        = @()
        # Not a test project, but it drives the built CLI, so it depends on the CLI's
        # project closure plus the workflow script itself.
        Project     = 'src/PptMcp.CLI/PptMcp.CLI.csproj'
        ExtraInputs = @('scripts/Test-CliWorkflow.ps1')
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

if ($ListInputs) {
    foreach ($suite in $suites) {
        $paths = Get-SuiteInputPaths $suite
        $hash = Get-SuiteInputsHash $paths
        Write-Host "[$($suite.Name)]" -ForegroundColor Cyan
        Write-Host "  hash: $hash"
        foreach ($p in $paths) { Write-Host "  - $p" }
        Write-Host ""
    }
    exit 0
}

if ($Only.Count -gt 0) {
    $unknown = @($Only | Where-Object { $_ -notin ($suites | ForEach-Object { $_.Name }) })
    if ($unknown.Count -gt 0) {
        Write-Host "FAIL: unknown suite name(s): $($unknown -join ', ')" -ForegroundColor Red
        Write-Host "      Known: $(($suites | ForEach-Object { $_.Name }) -join ', ')"
        exit 1
    }

    $suites = @($suites | Where-Object { $_.Name -in $Only })
    $binding = 'none'
    Write-Host "Partial run ($($suites.Count) suite(s)): binding=none, this cannot satisfy pre-push." -ForegroundColor Yellow
    Write-Host ""
}

# ---------------------------------------------------------------- prior evidence
# Load the manifest this run is about to replace, so suites whose inputs are unchanged can
# stand on their previous result instead of launching PowerPoint again to reach it.
$previousSuites = @{}
$previousCommit = $null

if (-not $Force -and $binding -eq 'commit' -and (Test-Path $OutputPath)) {
    try {
        $prev = Get-Content $OutputPath -Raw | ConvertFrom-Json
        if ($prev.schemaVersion -eq 1 -and $prev.binding -eq 'commit') {
            $previousCommit = $prev.commit
            foreach ($s in @($prev.suites)) {
                if ($s.result -eq 'pass' -and $s.inputsHash) {
                    $previousSuites[$s.name] = $s
                }
            }
        }
    }
    catch {
        # An unreadable manifest simply means no reuse. It is never a reason to fail.
        $previousSuites = @{}
    }
}

foreach ($suite in $suites) {
    Write-Host "[$($suite.Name)]" -ForegroundColor Cyan
    $suiteStart = Get-Date
    $testCount = $null
    $output = ''

    # ------------------------------------------------------------ reuse check
    $inputsHash = if ($binding -eq 'commit') { Get-SuiteInputsHash (Get-SuiteInputPaths $suite) } else { $null }

    if ($inputsHash -and $previousSuites.ContainsKey($suite.Name)) {
        $prior = $previousSuites[$suite.Name]
        if ($prior.inputsHash -eq $inputsHash) {
            $priorFrom = if ($prior.reusedFrom) { $prior.reusedFrom } else { $previousCommit }

            Write-Host "  REUSED - $($prior.tests) test(s), inputs unchanged since $($priorFrom.Substring(0, 8))" -ForegroundColor DarkGray
            Write-Host ""

            $results += [pscustomobject]@{
                name            = $suite.Name
                kind            = $suite.Kind
                tests           = [int]$prior.tests
                result          = 'pass'
                exitCode        = 0
                durationSeconds = [double]$prior.durationSeconds
                inputsHash      = $inputsHash
                reused          = $true
                reusedFrom      = $priorFrom
            }
            continue
        }
    }

    if ($suite.Kind -eq 'script') {
        # *>&1 and not 2>&1: these scripts report through Write-Host, which writes to the
        # information stream, not stdout. Capturing only stdout yielded an empty test
        # count and failed a suite that had in fact passed.
        $output = & $suite.Script *>&1 | Out-String
        $exitCode = $LASTEXITCODE

        # Test-CliWorkflow.ps1 prints "Passed: N" / "Failed: N".
        if ($output -match 'Passed:\s*(\d+)') { $testCount = [int]$Matches[1] }
        if ($output -match 'Failed:\s*([1-9]\d*)') { $exitCode = 1 }
    }
    else {
        # Hashtable splatting, not an array. With array splatting PowerShell binds the
        # token after a switch as that switch's VALUE, so '-NoBuild', '-LoggerFileName'
        # swallowed the parameter name and pushed the file name onto -MinimumTests.
        $guardParams = @{
            Project        = $suite.Project
            NoBuild        = $true
            LoggerFileName = "gate-$($suite.Name).trx"
        }
        if ($suite.Filter)  { $guardParams.Filter = $suite.Filter }
        if ($suite.Full)    { $guardParams.Full = $true }
        if ($suite.Timeout) { $guardParams.TimeoutMinutes = $suite.Timeout }

        $output = & $guard @guardParams *>&1 | Out-String
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
        inputsHash      = $inputsHash
        reused          = $false
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

# Re-check the tree. The start-of-run check alone leaves a hole: a suite takes ~10
# minutes, and anything edited under src/ or tests/ during that window would still be
# recorded against the SHA the run started on. Evidence must describe code that was
# actually tested, so a tree that moved underneath us downgrades to binding=none.
if ($binding -eq 'commit') {
    $endCommit = (git rev-parse HEAD).Trim()
    $endDirty = @(git status --porcelain -- src tests | Where-Object { $_ -ne '' })

    if ($endCommit -ne $commit -or $endDirty.Count -gt 0) {
        $binding = 'none'
        Write-Host "WARNING: src/ or tests/ changed while the gate was running." -ForegroundColor Yellow
        Write-Host "         This run no longer describes commit $($commit.Substring(0, 8)), so it is"
        Write-Host "         marked binding=none. Re-run once the tree is settled."
        Write-Host ""
    }
}

$totalTests = [int](($results | Measure-Object -Property tests -Sum).Sum)
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
$reusedCount = @($results | Where-Object { $_.reused }).Count
$ranCount = $results.Count - $reusedCount
Write-Host "Suites: $($results.Count) ($ranCount run, $reusedCount reused)   Tests: $totalTests   Duration: $([math]::Round($totalDuration / 60, 1))m"
Write-Host "Manifest: $OutputPath"
Write-Host ""

if ($gateFailed) {
    Write-Host "GATE FAILED - the manifest records the failure and will not satisfy pre-push." -ForegroundColor Red
    exit 1
}

if ($binding -eq 'none') {
    $why = if ($Only.Count -gt 0) { 'partial run' } else { 'tree does not match the commit' }
    Write-Host "GATE PASSED, but binding=none ($why). Pre-push will still refuse." -ForegroundColor Yellow
    exit 0
}

Write-Host "GATE PASSED for commit $commit" -ForegroundColor Green
exit 0
