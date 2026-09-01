# Detects COM member chains that are used inline instead of via locals.
#
# `slide.Design.SlideMaster.Name` materialises two COM proxies (Design, then
# SlideMaster) that are never bound to a variable, so ComUtilities.Release can
# never be called on them - not because someone forgot, but because there is
# nothing to pass. They are abandoned to the GC and the RCW survives until a
# non-deterministic finalizer runs, which for an out-of-process Office server
# means the PowerPoint instance is held open.
#
# check-com-leaks.ps1 cannot see these. It reasons about `dynamic` locals and
# their matching Release calls, so a file whose leaks are entirely inline is
# reported as "Proper COM cleanup" - which is exactly what it said about
# SlideCommands.cs while that file leaked on every `slide read`.
#
# 190 such chains predate this gate. They are frozen per-file in
# scripts/inline-com-chains-baseline.txt. New chains fail; when a count drops,
# lower the baseline with -UpdateBaseline.
#
# ASCII only. Windows PowerShell reads these files as cp1252, and a stray
# typographic character silently turns the whole script into a parse error,
# which would mean the gate never runs at all.

[CmdletBinding()]
param([switch]$UpdateBaseline)

$ErrorActionPreference = 'Stop'

$repoRoot     = Split-Path -Parent $PSScriptRoot
$commandsDir  = Join-Path $repoRoot 'src\PptMcp.Core\Commands'
$baselineFile = Join-Path $PSScriptRoot 'inline-com-chains-baseline.txt'

Write-Host "Scanning for inline COM member chains..."
Write-Host ""

if (-not (Test-Path $commandsDir)) {
    Write-Host "FAIL: $commandsDir does not exist." -ForegroundColor Red
    Write-Host "      A gate that inspects nothing must not report success."
    exit 1
}

# Roots that denote a live COM object in this codebase. Deliberately explicit:
# matching any lowercase identifier would sweep up result DTOs and options
# objects, whose property chains are plain managed calls and leak nothing.
$comRoots = @(
    'slide', 'slides', 'shape', 'shapes', 'pres', 'presentation',
    'design', 'designs', 'master', 'layout', 'textFrame', 'textRange',
    'table', 'chart', 'app', 'ctx\.Presentation'
)
$rootAlternation = ($comRoots -join '|')

# <comRoot>.<Prop>.<Prop> - two or more hops with no intermediate local.
$pattern = "\b($rootAlternation)\.[A-Z]\w*\.[A-Z]\w*"

$files = Get-ChildItem -Path $commandsDir -Recurse -Filter *.cs |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

if ($files.Count -eq 0) {
    Write-Host "FAIL: found no command source files to scan." -ForegroundColor Red
    Write-Host "      A gate that inspects nothing must not report success."
    exit 1
}

$current = @{}
$details = @{}
foreach ($file in $files) {
    $rel = $file.FullName.Substring($repoRoot.Length + 1)
    $lines = Get-Content $file.FullName
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^\s*//') { continue }
        $matches = [regex]::Matches($line, $pattern)
        if ($matches.Count -eq 0) { continue }
        if (-not $current.ContainsKey($rel)) {
            $current[$rel] = 0
            $details[$rel] = @()
        }
        $current[$rel] += $matches.Count
        $details[$rel] += "    line $($i + 1): $($line.Trim())"
    }
}

$total = ($current.Values | Measure-Object -Sum).Sum
if ($null -eq $total) { $total = 0 }

Write-Host "Files scanned:        $($files.Count)"
Write-Host "Inline COM chains:    $total in $($current.Count) file(s)"
Write-Host ""

if ($UpdateBaseline) {
    $out = @(
        '# Inline COM member chains per file. See check-inline-com-chains.ps1.',
        '# Counts may only go down. Regenerate with -UpdateBaseline.',
        "# Total at last update: $total"
    )
    foreach ($k in ($current.Keys | Sort-Object)) { $out += "$k=$($current[$k])" }
    Set-Content -Path $baselineFile -Value $out -Encoding ASCII
    Write-Host "Baseline written: $total chain(s) across $($current.Count) file(s)." -ForegroundColor Yellow
    exit 0
}

if (-not (Test-Path $baselineFile)) {
    Write-Host "FAIL: baseline file not found at $baselineFile" -ForegroundColor Red
    Write-Host "      Create it with: .\scripts\check-inline-com-chains.ps1 -UpdateBaseline"
    exit 1
}

$baseline = @{}
foreach ($line in (Get-Content $baselineFile)) {
    if ($line -match '^\s*#' -or $line -notmatch '=') { continue }
    $idx = $line.LastIndexOf('=')
    $baseline[$line.Substring(0, $idx)] = [int]$line.Substring($idx + 1)
}

if ($baseline.Count -eq 0) {
    Write-Host "FAIL: baseline file contains no entries." -ForegroundColor Red
    exit 1
}

$regressions = @()
$improvements = @()
foreach ($rel in ($current.Keys | Sort-Object)) {
    $was = if ($baseline.ContainsKey($rel)) { $baseline[$rel] } else { 0 }
    if ($current[$rel] -gt $was) { $regressions += [pscustomobject]@{ File = $rel; Was = $was; Now = $current[$rel] } }
    elseif ($current[$rel] -lt $was) { $improvements += "  $rel : $was -> $($current[$rel])" }
}
foreach ($rel in ($baseline.Keys | Sort-Object)) {
    if (-not $current.ContainsKey($rel) -and $baseline[$rel] -gt 0) {
        $improvements += "  $rel : $($baseline[$rel]) -> 0"
    }
}

if ($regressions.Count -gt 0) {
    Write-Host "FAIL: $($regressions.Count) file(s) gained inline COM chains." -ForegroundColor Red
    Write-Host ""
    foreach ($r in $regressions) {
        Write-Host "  $($r.File): $($r.Was) -> $($r.Now)" -ForegroundColor Red
        foreach ($d in $details[$r.File]) { Write-Host $d }
        Write-Host ""
    }
    Write-Host "An inline chain such as slide.Design.SlideMaster.Name creates COM proxies"
    Write-Host "that are never bound to a local, so they can never be released. Assign each"
    Write-Host "hop to a dynamic local and release it in a finally block (Rule 22)."
    exit 1
}

if ($improvements.Count -gt 0) {
    Write-Host "Improvements since the baseline:" -ForegroundColor Green
    $improvements | ForEach-Object { Write-Host $_ -ForegroundColor Green }
    Write-Host ""
    Write-Host "Lower the baseline: .\scripts\check-inline-com-chains.ps1 -UpdateBaseline" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "No new inline COM chains ($total total, baseline allows $(($baseline.Values | Measure-Object -Sum).Sum))." -ForegroundColor Green
exit 0
