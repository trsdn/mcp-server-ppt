<#
.SYNOPSIS
    Detects COM objects that are acquired into a `dynamic` variable and never released.

.DESCRIPTION
    This replaces a gate that could not fail. The previous implementation asked two
    whole-file yes/no questions - "does this file contain a dynamic acquisition?" and
    "does this file contain the text ComUtilities.Release?" - and reported "Proper COM
    cleanup" whenever both were true. Every one of the 33 command files contains at
    least one Release, so every one was green, permanently. Injecting a blatant
    never-released acquisition into SlideCommands.cs did not change the verdict:
    still "Proper COM cleanup", still "Leak files: 0", still exit 0.

    Rule 5 says the gate "must report 0 leaks". It did, unconditionally, which made
    the rule vacuous rather than satisfied.

    This tracks each variable instead. For every `dynamic` variable that acquires a
    COM object, there must be a matching `ComUtilities.Release(ref <name>)`.

    BORROWED REFERENCES ARE NOT LEAKS. `dynamic pres = ctx.Presentation` aliases an
    object owned by the batch context. The command borrows it for the duration of the
    call and must NOT release it - doing so would tear down the caller's presentation.
    Flagging those would produce 20 false positives, and a gate that cries wolf gets
    switched off. Ownership, not syntax, decides.

    A gate that inspects nothing must fail, so this exits non-zero if it finds no
    dynamic acquisitions at all - that means detection broke, not that the code is
    clean.

    ASCII only. Windows PowerShell reads these files as cp1252, and a stray
    typographic character silently turns the whole script into a parse error, which
    would mean the gate never runs at all.

.PARAMETER UpdateBaseline
    Rewrites the baseline to the current findings. Use only to ratchet the count DOWN.
#>
[CmdletBinding()]
param(
    [switch]$UpdateBaseline,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $repoRoot 'src'
$baselinePath = Join-Path $PSScriptRoot 'com-leaks-baseline.txt'

if (-not $Quiet) { Write-Host "Scanning for COM object leaks (per variable)..." -ForegroundColor Yellow }

# Session plumbing owns the lifetime of the objects it hands out.
$exemptFiles = 'PptBatch\.cs|PptSession\.cs'

$files = Get-ChildItem -Path $srcDir -Recurse -Filter '*.cs' -File |
    Where-Object { $_.FullName -notmatch '\\obj\\|\\bin\\' } |
    Where-Object { $_.Name -notlike '*.g.cs' } |
    Where-Object { $_.FullName -notmatch $exemptFiles }

if ($files.Count -eq 0) {
    Write-Host "FAIL: no source files found under $srcDir - the discovery glob is broken." -ForegroundColor Red
    exit 1
}

# A dynamic whose initialiser is the batch context's presentation is a borrowed
# alias, not an acquisition. Casts and whitespace vary, so match the shape.
$borrowedPattern = '^\s*\(?\s*\(\s*dynamic\s*\)\s*\)?\s*ctx\.Presentation\s*$|^\s*ctx\.Presentation\s*$'

$acquisitions = 0
$findings = @()

foreach ($file in $files) {
    $lines = Get-Content -LiteralPath $file.FullName
    $content = $lines -join "`n"
    $relative = $file.FullName.Substring($repoRoot.Length + 1)

    # name -> first line number where it is declared as dynamic
    $declared = [ordered]@{}

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $m = [regex]::Match($lines[$i], '\bdynamic\??\s+(\w+)\s*=\s*(.*?);')
        if (-not $m.Success) { continue }

        $name = $m.Groups[1].Value
        $rhs = $m.Groups[2].Value

        # `dynamic? x = null;` is the declare-then-assign-in-try pattern. The
        # acquisition happens later, so judge it by its later assignments.
        if ($rhs.Trim() -eq 'null') {
            $assign = [regex]::Match($content, '(?m)^\s*' + [regex]::Escape($name) + '\s*=\s*(?!null\s*;)(.*?);')
            if (-not $assign.Success) { continue }
            $rhs = $assign.Groups[1].Value
        }

        # Borrowed from the batch context - the command does not own it.
        if ($rhs -match $borrowedPattern) { continue }

        # Only member access acquires a new COM proxy. `dynamic n = 5;` does not.
        if ($rhs -notmatch '\.') { continue }

        if (-not $declared.Contains($name)) { $declared[$name] = $i + 1 }
    }

    foreach ($name in $declared.Keys) {
        $acquisitions++
        # ReleaseIfNotNull is a thin wrapper over ComUtilities.Release used where the
        # variable may legitimately be null; it discharges ownership just as well.
        $releasePattern = '(?:ComUtilities\.Release|ReleaseIfNotNull)\(\s*ref\s+' + [regex]::Escape($name) + '\b'
        if ($content -match $releasePattern) { continue }

        $findings += ('{0}:{1}:{2}' -f $relative, $declared[$name], $name)
    }
}

if ($acquisitions -eq 0) {
    Write-Host "FAIL: inspected $($files.Count) file(s) but found no COM acquisitions." -ForegroundColor Red
    Write-Host "      Detection is broken - this gate is not testing anything." -ForegroundColor Red
    exit 1
}

$findings = $findings | Sort-Object

if ($UpdateBaseline) {
    Set-Content -Path $baselinePath -Value $findings -Encoding ASCII
    Write-Host "Baseline updated: $($findings.Count) known leak(s) recorded." -ForegroundColor Yellow
    exit 0
}

$baseline = @()
if (Test-Path $baselinePath) {
    $baseline = Get-Content $baselinePath | Where-Object { $_ -and -not $_.StartsWith('#') }
}

$new = $findings | Where-Object { $baseline -notcontains $_ }
$fixed = $baseline | Where-Object { $findings -notcontains $_ }

if ($new.Count -gt 0) {
    Write-Host ""
    Write-Host "FAIL: $($new.Count) new COM leak(s) - acquired into a dynamic, never released." -ForegroundColor Red
    Write-Host ""
    foreach ($n in $new) {
        $parts = $n -split ':'
        $var = $parts[-1]
        $line = $parts[-2]
        $path = ($parts[0..($parts.Count - 3)]) -join ':'
        Write-Host ("  {0}" -f $var) -ForegroundColor Yellow
        Write-Host ("      {0}:{1}" -f $path, $line)
    }
    Write-Host ""
    Write-Host "Every acquired COM object needs ComUtilities.Release(ref <name>!) in a finally (Rule 22)." -ForegroundColor Red
    exit 1
}

if (-not $Quiet) {
    Write-Host ""
    Write-Host "  Dynamic COM acquisitions tracked: $acquisitions" -ForegroundColor Cyan
    Write-Host "  Known leaks (baselined):          $($findings.Count)" -ForegroundColor Yellow
    if ($fixed.Count -gt 0) {
        Write-Host ""
        Write-Host "  $($fixed.Count) baselined leak(s) fixed. Lower the baseline:" -ForegroundColor Green
        Write-Host "    scripts\check-com-leaks.ps1 -UpdateBaseline" -ForegroundColor Green
    }
    Write-Host ""
    Write-Host "No new COM object leaks." -ForegroundColor Green
}

exit 0
