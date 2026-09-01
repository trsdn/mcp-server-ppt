# Runs `dotnet test` and fails when the filter selected no tests.
#
# `dotnet test --filter <anything>` prints "No test matches the given testcase filter"
# and then exits 0, so a filter that selects nothing is indistinguishable from a run
# where everything passed. Every CI step and every documented workflow that applies a
# filter must go through this wrapper, or a typo silently turns the step into a no-op.
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

    [int]$MinimumTests = 1
)

$ErrorActionPreference = 'Stop'

$testArgs = @('test', $Project, '-c', $Configuration)
if ($Filter)         { $testArgs += @('--filter', $Filter) }
if ($LoggerFileName) { $testArgs += @('--logger', "trx;LogFileName=$LoggerFileName") }
if ($NoBuild)        { $testArgs += '--no-build' }

Write-Host "dotnet $($testArgs -join ' ')"
$output = & dotnet @testArgs 2>&1
$testExit = $LASTEXITCODE
$output | ForEach-Object { Write-Host $_ }

$text = ($output | Out-String)

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
