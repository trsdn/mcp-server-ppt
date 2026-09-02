# Verifies that a local integration gate manifest exists and actually covers a commit.
#
# Produced by scripts/Invoke-LocalIntegrationGate.ps1, consumed by scripts/pre-push.ps1.
#
# Every check below answers one way the evidence could be a lie:
#
#   missing file        - nobody ran anything
#   wrong schema        - a manifest from a different tool or an older format
#   binding != commit   - produced from a dirty tree, so it describes no commit
#   commit mismatch     - evidence for OTHER code; the most dangerous case, because the
#                         file exists and says "pass"
#   result != pass      - the run failed and the manifest honestly says so
#   no suites / 0 tests - a gate that inspected nothing must not report success
#
# ASCII only - see scripts/check-test-filters.ps1 for why.

[CmdletBinding()]
param(
    # The commit the evidence must cover. Defaults to HEAD.
    [string]$Commit,

    [string]$ManifestPath = '.integration-evidence/manifest.json',

    # Print the per-suite breakdown even when the evidence is valid.
    [switch]$Detailed
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if (-not $Commit) {
    $Commit = (git rev-parse HEAD).Trim()
}

function Fail {
    param([string]$Reason, [string[]]$Detail = @())

    Write-Host ""
    Write-Host "FAIL: $Reason" -ForegroundColor Red
    foreach ($line in $Detail) { Write-Host "      $line" }
    Write-Host ""
    Write-Host "      Produce evidence with:"
    Write-Host "        .\scripts\Invoke-LocalIntegrationGate.ps1"
    exit 1
}

if (-not (Test-Path $ManifestPath)) {
    Fail "no integration evidence found at $ManifestPath." @(
        "Nothing has verified this commit against real PowerPoint."
    )
}

try {
    $manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
}
catch {
    Fail "the manifest at $ManifestPath is not valid JSON." @($_.Exception.Message)
}

if ($manifest.schemaVersion -ne 1) {
    Fail "unexpected manifest schemaVersion '$($manifest.schemaVersion)' (expected 1)." @(
        "Re-run the gate to regenerate it in the current format."
    )
}

if ($manifest.binding -ne 'commit') {
    Fail "the evidence is not bound to a commit (binding='$($manifest.binding)')." @(
        "It was produced from a dirty working tree with -AllowDirty, so it describes",
        "content that no commit holds. Commit, then re-run the gate."
    )
}

if ($manifest.commit -ne $Commit) {
    Fail "the evidence covers a different commit." @(
        "evidence: $($manifest.commit)",
        "required: $Commit",
        "",
        "Stale evidence is worse than none, because it looks like coverage."
    )
}

if ($manifest.result -ne 'pass') {
    Fail "the recorded gate run did not pass (result='$($manifest.result)')." @(
        "Fix the failures and re-run the gate."
    )
}

$suites = @($manifest.suites)
if ($suites.Count -eq 0) {
    Fail "the manifest records no suites." @(
        "A gate that inspected nothing must never report success."
    )
}

$failedSuites = @($suites | Where-Object { $_.result -ne 'pass' })
if ($failedSuites.Count -gt 0) {
    Fail "$($failedSuites.Count) suite(s) in the manifest did not pass." @(
        ($failedSuites | ForEach-Object { "$($_.name): $($_.result)" })
    )
}

$emptySuites = @($suites | Where-Object { -not $_.tests -or $_.tests -le 0 })
if ($emptySuites.Count -gt 0) {
    Fail "$($emptySuites.Count) suite(s) recorded zero tests." @(
        ($emptySuites | ForEach-Object { "$($_.name)" }),
        "A suite that ran no tests is not evidence."
    )
}

if (-not $manifest.totalTests -or $manifest.totalTests -le 0) {
    Fail "the manifest records zero tests in total."
}

$ageHours = [math]::Round(((Get-Date).ToUniversalTime() - [datetime]::Parse($manifest.generatedAtUtc).ToUniversalTime()).TotalHours, 1)

Write-Host "Integration evidence OK" -ForegroundColor Green
Write-Host "  commit:  $($manifest.commit)"
Write-Host "  suites:  $($suites.Count)   tests: $($manifest.totalTests)   age: ${ageHours}h"
Write-Host "  machine: $($manifest.machine.name), PowerPoint $($manifest.machine.powerPoint)"

if ($Detailed) {
    foreach ($s in $suites) {
        Write-Host "  - $($s.name): $($s.tests) test(s), $($s.durationSeconds)s"
    }
}

exit 0
