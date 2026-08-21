#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Checks that all ((dynamic)) casts in PptMcp.Core and PptMcp.ComInterop have justification comments.

.DESCRIPTION
    Every use of ((dynamic)) cast (explicit type coercion) must be preceded by a comment explaining
    why the PIA type cannot be used. Bare ((dynamic)) casts indicate potential PIA coverage gaps
    that weren't investigated.

    Valid comment prefixes (on the line immediately before the cast):
      // PIA gap: ...    - Type not in v16 Microsoft.Office.Interop.PowerPoint PIA
      // TODO: ...       - Type IS in PIA but migration not yet done (tracked for removal)
      // Reason: ...     - Other documented reason for dynamic cast

    False positives are excluded:
      - PptBatch.cs / PptSession.cs / PptShutdownService.cs (infrastructure - uses `dynamic powerpoint`)
      - Lines inside comments

.EXAMPLE
    .\check-dynamic-casts.ps1

.NOTES
    Run automatically as part of pre-commit hook.
    To add a new documented cast, place a comment ending in "// PIA gap:", "// TODO:", or "// Reason:"
    on the line immediately before the ((dynamic)) cast.
#>

param(
    [switch]$Verbose,
    # Rewrites the baseline to the current findings. Use only when the count drops.
    [switch]$UpdateBaseline
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$baselinePath = Join-Path $PSScriptRoot "dynamic-casts-baseline.txt"

$searchDirs = @(
    (Join-Path $rootDir "src\PptMcp.Core"),
    (Join-Path $rootDir "src\PptMcp.ComInterop")
)

# Files where bare dynamic casts are acceptable (infrastructure files)
$excludeFiles = @(
    "PptBatch.cs",
    "PptSession.cs",
    "PptShutdownService.cs"
)

$violations = @()
$checkedFiles = 0

foreach ($dir in $searchDirs) {
    $csFiles = Get-ChildItem -Path $dir -Filter "*.cs" -Recurse -ErrorAction SilentlyContinue
    foreach ($file in $csFiles) {
        if ($excludeFiles -contains $file.Name) {
            if ($Verbose) { Write-Host "   Skipped (infrastructure): $($file.Name)" -ForegroundColor Gray }
            continue
        }

        $checkedFiles++
        $lines = Get-Content $file.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # Check for ((dynamic)) cast pattern
            if ($line -match '\(\(dynamic\)') {
                # Skip lines that are themselves comments
                $trimmed = $line.TrimStart()
                if ($trimmed.StartsWith("//")) { continue }

                # Check if any preceding comment line (within 5 lines) has a justification comment
                $hasJustification = $false
                for ($j = $i - 1; $j -ge 0 -and $j -ge ($i - 5); $j--) {
                    $prevLine = $lines[$j].TrimStart()
                    if ([string]::IsNullOrWhiteSpace($prevLine)) { continue }

                    # Once we hit a non-comment line, stop looking
                    if (-not $prevLine.StartsWith("//")) { break }

                    if ($prevLine.StartsWith("// PIA gap:") -or
                        $prevLine.StartsWith("// TODO:") -or
                        $prevLine.StartsWith("// Reason:") -or
                        $prevLine.StartsWith("// REASON:")) {
                        $hasJustification = $true
                        break
                    }
                }

                if (-not $hasJustification) {
                    $violations += [PSCustomObject]@{
                        File = $file.FullName.Replace($rootDir, "").TrimStart("\")
                        Line = $i + 1
                        Code = $line.Trim()
                    }
                }
            }
        }
    }
}

Write-Host "Checked $checkedFiles C# files for undocumented ((dynamic)) casts" -ForegroundColor Cyan

if ($checkedFiles -eq 0) {
    # A check that inspected nothing has not passed.
    Write-Host "No C# files were inspected - the search paths are wrong." -ForegroundColor Red
    exit 1
}

# Per-file baseline. The 140 casts that predate this gate are tolerated so the hook
# stays installable, but the count may never rise: a new undocumented cast in any
# file, or a cast in a file that had none, fails the check.
$current = @{}
foreach ($v in $violations) {
    if ($current.ContainsKey($v.File)) { $current[$v.File]++ } else { $current[$v.File] = 1 }
}

if ($UpdateBaseline) {
    $lines = $current.Keys | Sort-Object | ForEach-Object { "$_=$($current[$_])" }
    Set-Content -Path $baselinePath -Value $lines -Encoding UTF8
    Write-Host "Baseline updated: $($violations.Count) casts across $($current.Count) files" -ForegroundColor Green
    exit 0
}

$baseline = @{}
if (Test-Path $baselinePath) {
    foreach ($line in Get-Content $baselinePath) {
        if ($line -match '^(?<f>.+)=(?<c>\d+)$') { $baseline[$Matches.f] = [int]$Matches.c }
    }
}

$regressions = @()
foreach ($file in $current.Keys) {
    $allowed = if ($baseline.ContainsKey($file)) { $baseline[$file] } else { 0 }
    if ($current[$file] -gt $allowed) {
        $regressions += [PSCustomObject]@{ File = $file; Was = $allowed; Now = $current[$file] }
    }
}

if ($regressions.Count -eq 0) {
    $total = $violations.Count
    $baselineTotal = ($baseline.Values | Measure-Object -Sum).Sum
    Write-Host "No new undocumented casts ($total tolerated by baseline)" -ForegroundColor Green
    if ($total -lt $baselineTotal) {
        Write-Host "Casts dropped from $baselineTotal to $total - lower the baseline with -UpdateBaseline" -ForegroundColor Yellow
    }
    exit 0
}

Write-Host ""
Write-Host "NEW UNDOCUMENTED ((dynamic)) CASTS: $($regressions.Count) file(s)" -ForegroundColor Red
Write-Host ""
Write-Host "Every new ((dynamic)) cast needs a comment on the preceding line explaining why:" -ForegroundColor Yellow
Write-Host "  // PIA gap: <type> not in Microsoft.Office.Interop.PowerPoint v16 PIA because..." -ForegroundColor Gray
Write-Host "  // TODO: <type> IS in PIA, migration tracked - left as dynamic temporarily" -ForegroundColor Gray
Write-Host "  // Reason: <explanation>" -ForegroundColor Gray
Write-Host ""

foreach ($r in $regressions) {
    Write-Host "  $($r.File): $($r.Was) -> $($r.Now)" -ForegroundColor Yellow
    foreach ($v in $violations | Where-Object { $_.File -eq $r.File }) {
        Write-Host "    line $($v.Line): $($v.Code)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Fix these before committing. See docs/PIA-COVERAGE.md for guidance." -ForegroundColor Red
exit 1
