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

$reusedFrom = $null
if ($manifest.commit -ne $Commit) {
    # Evidence names a different commit, but what the suites actually exercised is the
    # content of src/ and tests/. If that content is byte-identical between the two
    # commits, the run tested exactly this code and re-running would ask PowerPoint the
    # same 787 questions for a second time to get the same answers. A docs, scripts or
    # CHANGELOG commit on top of a verified one is the common case, and forcing a fresh
    # 11-minute run there teaches people to reach for the override - which is how a
    # blocking gate ends up disabled in practice.
    #
    # This is a narrow exemption, not a softening: any difference at all under src/ or
    # tests/, or an evidence commit git can no longer resolve, still fails.
    git diff --quiet $manifest.commit $Commit -- src tests 2>$null
    $diffExit = $LASTEXITCODE

    if ($diffExit -eq 0) {
        $reusedFrom = $manifest.commit
    }
    else {
        $detail = if ($diffExit -gt 1) {
            @("git could not compare the two commits - the evidence commit may have been",
              "rebased away or never existed here.")
        }
        else {
            @("src/ or tests/ differ between them, so the run did not exercise this code.")
        }

        Fail "the evidence covers a different commit." (@(
            "evidence: $($manifest.commit)",
            "required: $Commit",
            ""
        ) + $detail + @(
            "",
            "Stale evidence is worse than none, because it looks like coverage."
        ))
    }
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

# ConvertFrom-Json already converts an ISO 8601 string into a [datetime], so calling
# [datetime]::Parse on it stringifies it in the current culture and reads it back in
# another - which turned 2026-09-02 into 2026-02-09 and reported a minutes-old manifest
# as 4921 hours stale. Only parse when it is genuinely still a string, and pin the
# culture and round-trip kind when doing so.
$generatedRaw = $manifest.generatedAtUtc
if ($generatedRaw -is [datetime]) {
    $generatedUtc = $generatedRaw.ToUniversalTime()
}
else {
    $generatedUtc = [datetime]::Parse(
        [string]$generatedRaw,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
}

$ageHours = [math]::Round(((Get-Date).ToUniversalTime() - $generatedUtc).TotalHours, 1)

Write-Host "Integration evidence OK" -ForegroundColor Green
Write-Host "  commit:  $($manifest.commit)"
if ($reusedFrom) {
    Write-Host "  reused:  src/ and tests/ are identical in $($Commit.Substring(0, 8)), so this run covers it" -ForegroundColor Yellow
}
Write-Host "  suites:  $($suites.Count)   tests: $($manifest.totalTests)   age: ${ageHours}h"
Write-Host "  machine: $($manifest.machine.name), PowerPoint $($manifest.machine.powerPoint)"

if ($Detailed) {
    foreach ($s in $suites) {
        $tag = if ($s.reused) { "  (reused from $($s.reusedFrom.Substring(0, 8)) - inputs unchanged)" } else { '' }
        Write-Host "  - $($s.name): $($s.tests) test(s), $($s.durationSeconds)s$tag"
    }
}

exit 0
