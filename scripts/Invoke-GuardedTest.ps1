# Runs `dotnet test` and refuses to report success without evidence that tests ran.
#
# Guards, in the order they bite:
#
# 1. UNFILTERED RUNS ARE REFUSED on projects that drive real PowerPoint. Those suites
#    create a PowerPoint instance per test, so an unfiltered `dotnet test` launches
#    POWERPNT dozens to hundreds of times back to back. That is not a faster way to
#    get the same answer - every run after the first tells you nothing new, it just
#    hammers the developer's machine and the COM layer. Rule 16 already says test only
#    what you changed; this makes it structural instead of advisory.
#    Pass -Full to opt in deliberately.
#
# 2. A HARD TIMEOUT. tests\PptMcp.ComInterop.Tests does not terminate as a full
#    assembly (issue #139), and an un-timeboxed hang keeps spawning PowerPoint for as
#    long as nobody is watching. Nothing here may run unbounded.
#
# 3. STRAY POWERPOINT CLEANUP. Only processes that appear DURING the run are killed -
#    PIDs present beforehand are the developer's own PowerPoint and are left alone.
#
# 4. The original guard: `dotnet test --filter <anything>` prints "No test matches the
#    given testcase filter" and then exits 0, so a filter selecting nothing is
#    indistinguishable from a run where everything passed.
#
# ASCII only - see scripts/check-test-filters.ps1 for why.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Project,

    [string]$Filter,

    [string]$Configuration = 'Release',

    [string]$LoggerFileName,

    [switch]$NoBuild,

    [int]$MinimumTests = 1,

    # Opt in to running a whole PowerPoint-driving assembly. Required by guard 1.
    [switch]$Full,

    # Hard ceiling on wall-clock time. Guard 2.
    [int]$TimeoutMinutes = 20
)

$ErrorActionPreference = 'Stop'

# Projects that do NOT drive PowerPoint. Everything else is assumed to, because the
# safe direction for this guard is to over-trigger rather than under-trigger.
$nonPowerPointProjects = @(
    'PptMcp.SkillGeneration.Tests'
)

# -Project is given as either a directory (tests\Foo.Tests) or a project file
# (tests/Foo.Tests/Foo.Tests.csproj - which is what the workflows pass). Strip the
# extension so both spellings resolve to the same name, or the allow-list silently
# never matches and a COM-free suite gets refused.
$projectLeaf = [System.IO.Path]::GetFileNameWithoutExtension(($Project -replace '[\\/]+$', ''))
$drivesPowerPoint = $projectLeaf -notin $nonPowerPointProjects

# ---------------------------------------------------------------- guard 1
if ($drivesPowerPoint -and -not $Filter -and -not $Full) {
    Write-Host ""
    Write-Host "REFUSED: unfiltered run of a PowerPoint-driving suite." -ForegroundColor Red
    Write-Host "         Project: $Project"
    Write-Host ""
    Write-Host "These suites create a PowerPoint instance per test, so this would launch"
    Write-Host "POWERPNT once per test with no upper bound. Rule 16: test only what you"
    Write-Host "changed."
    Write-Host ""
    Write-Host "Use a filter:"
    Write-Host "  .\scripts\Invoke-GuardedTest.ps1 -Project $Project -Filter 'Feature=Slide'"
    Write-Host ""
    Write-Host "If you genuinely want the whole assembly, say so:"
    Write-Host "  .\scripts\Invoke-GuardedTest.ps1 -Project $Project -Full"
    Write-Host ""
    Write-Host "Note: tests\PptMcp.ComInterop.Tests does not terminate as a full assembly"
    Write-Host "      (issue #139). -Full will hit the timeout there."
    exit 1
}

# Snapshot pre-existing PowerPoint so the developer's own instance is never killed.
$preExistingPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

$testArgs = @('test', $Project, '-c', $Configuration)
if ($Filter)         { $testArgs += @('--filter', $Filter) }
if ($LoggerFileName) { $testArgs += @('--logger', "trx;LogFileName=$LoggerFileName") }
if ($NoBuild)        { $testArgs += '--no-build' }

Write-Host "dotnet $($testArgs -join ' ')"
if ($drivesPowerPoint) {
    Write-Host "guard: PowerPoint-driving suite, timeout ${TimeoutMinutes}m, $($preExistingPids.Count) pre-existing POWERPNT process(es) protected"
}

$stdoutFile = [System.IO.Path]::GetTempFileName()
$stderrFile = [System.IO.Path]::GetTempFileName()
$timedOut = $false
$testExit = 1
$output = @()

try {
    $proc = Start-Process -FilePath 'dotnet' -ArgumentList $testArgs -NoNewWindow -PassThru `
        -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

    # ------------------------------------------------------------ guard 2
    if (-not $proc.WaitForExit($TimeoutMinutes * 60 * 1000)) {
        $timedOut = $true
        Write-Host ""
        Write-Host "TIMEOUT: the run exceeded $TimeoutMinutes minute(s). Killing it." -ForegroundColor Red
        try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch { }
        Start-Sleep -Seconds 2
        $testExit = 124
    }
    else {
        $testExit = $proc.ExitCode
    }
}
finally {
    if (Test-Path $stdoutFile) { $output += Get-Content $stdoutFile }
    if (Test-Path $stderrFile) { $output += Get-Content $stderrFile }
    Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue

    $output | ForEach-Object { Write-Host $_ }

    # ------------------------------------------------------------ guard 3
    $strays = @(Get-Process POWERPNT -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -notin $preExistingPids })

    if ($strays.Count -gt 0) {
        Write-Host ""
        Write-Host "Cleaning up $($strays.Count) PowerPoint process(es) left behind by this run..." -ForegroundColor Yellow
        foreach ($p in $strays) {
            try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
}

$text = ($output | Out-String)

if ($timedOut) {
    Write-Host ""
    Write-Host "FAIL: run timed out after $TimeoutMinutes minute(s) and was killed." -ForegroundColor Red
    Write-Host "      A hanging suite keeps spawning PowerPoint, so it is never left running."
    exit 124
}

# ---------------------------------------------------------------- guard 4
if ($text -match 'No test matches the given testcase filter') {
    Write-Host ""
    Write-Host "FAIL: the filter selected no tests." -ForegroundColor Red
    Write-Host "      Filter: $Filter"
    Write-Host "      dotnet test exits 0 in this case, so this would otherwise be a silent no-op."
    exit 1
}

# Sum the per-assembly totals rather than trusting a single summary line.
$total = 0
$sawSummary = $false
foreach ($m in [regex]::Matches($text, 'Total:\s*(\d+)')) {
    $total += [int]$m.Groups[1].Value
    $sawSummary = $true
}

if (-not $sawSummary) {
    Write-Host ""
    Write-Host "FAIL: could not find a test summary in the output." -ForegroundColor Red
    Write-Host "      Refusing to report success without evidence that tests ran."
    exit 1
}

if ($total -lt $MinimumTests) {
    Write-Host ""
    Write-Host "FAIL: only $total test(s) ran, expected at least $MinimumTests." -ForegroundColor Red
    Write-Host "      Filter: $Filter"
    exit 1
}

if ($testExit -ne 0) {
    Write-Host ""
    Write-Host "FAIL: $total test(s) ran and dotnet test reported failures." -ForegroundColor Red
    exit $testExit
}

Write-Host ""
Write-Host "OK: $total test(s) ran and passed." -ForegroundColor Green
exit 0
