#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Rule 22 / Rule 1b ratchet: no NEW catch blocks that swallow exceptions in Core commands.

.DESCRIPTION
    Rule 22 says COM cleanup belongs in a finally block, never in a catch that
    suppresses the failure. Rule 1b says a Core command must let exceptions
    propagate to batch.Execute(), which converts them into
    OperationResult { Success = false } at the correct layer.

    A catch that neither rethrows nor performs loop control silently discards the
    reason an operation failed. The caller then sees a success-shaped result, or a
    fallback value such as "(unknown)", with no record of what went wrong.

    This gate does NOT try to fix the existing population. Those are frozen in
    scripts/swallowed-catches-baseline.txt and paid down under issue #126. What it
    prevents is the population growing.

    Detection is brace-matched, not regex-matched. A catch body is located by
    counting braces from its opening brace, so nested blocks, strings containing
    braces and comments do not shift the boundary.

    Permitted (not reported):
      - any body containing throw (rethrow, or throwing a more specific exception)
      - loop control only: continue; / break;   (Rule 1b "safe patterns")

.PARAMETER UpdateBaseline
    Rewrite the baseline from the current tree. Use only when the count DROPS.

.PARAMETER Quiet
    Suppress per-finding output; still sets the exit code.

.NOTES
    Exits non-zero if it inspects nothing. A gate that cannot fail is not a gate -
    see the vacuous COM-leak gate corrected in #125.
#>
[CmdletBinding()]
param(
    [switch]$UpdateBaseline,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scanRoot = Join-Path $repoRoot 'src\PptMcp.Core\Commands'
$baselineFile = Join-Path $PSScriptRoot 'swallowed-catches-baseline.txt'

if (-not (Test-Path $scanRoot)) {
    Write-Host "FAIL: scan root not found: $scanRoot" -ForegroundColor Red
    Write-Host "      The gate inspected nothing, which is a failure, not a pass."
    exit 1
}

$files = @(Get-ChildItem -Path $scanRoot -Filter '*.cs' -Recurse -File)

if ($files.Count -eq 0) {
    Write-Host "FAIL: no C# files found under $scanRoot" -ForegroundColor Red
    Write-Host "      The gate inspected nothing, which is a failure, not a pass."
    exit 1
}

# Returns the index just past the matching close brace for the block that starts
# at $openIndex, or -1 if the braces never balance.
function Get-BlockEnd {
    param([string]$Text, [int]$OpenIndex)

    $depth = 0
    for ($i = $OpenIndex; $i -lt $Text.Length; $i++) {
        $c = $Text[$i]
        if ($c -eq '{') { $depth++ }
        elseif ($c -eq '}') {
            $depth--
            if ($depth -eq 0) { return $i }
        }
    }
    return -1
}

$findings = @()
$catchCount = 0
$probeCount = 0

# Blank out comments and string literals, preserving length and line breaks so every
# offset and line number computed afterwards still refers to the real file.
#
# Without this the scanner matches the word "catch" inside XML doc comments - the
# remarks on SlideCommands.GetSlides discuss Rule 1b and the word "catch" appears in
# prose - and reports a documentation paragraph as a swallowed exception. A gate that
# cries wolf gets switched off, so it must never report prose as code.
function Remove-CommentsAndStrings {
    param([string]$Text)

    $sb = [System.Text.StringBuilder]::new($Text)
    $i = 0
    $n = $Text.Length

    while ($i -lt $n) {
        $c = $Text[$i]
        $next = if ($i + 1 -lt $n) { $Text[$i + 1] } else { [char]0 }

        if ($c -eq '/' -and $next -eq '/') {
            while ($i -lt $n -and $Text[$i] -ne "`n") { [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++ }
        }
        elseif ($c -eq '/' -and $next -eq '*') {
            while ($i -lt $n -and -not ($Text[$i] -eq '*' -and $i + 1 -lt $n -and $Text[$i + 1] -eq '/')) {
                if ($Text[$i] -ne "`n" -and $Text[$i] -ne "`r") { [void]$sb.Replace($Text[$i], ' ', $i, 1) }
                $i++
            }
            if ($i -lt $n) { [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++ }
            if ($i -lt $n) { [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++ }
        }
        elseif ($c -eq '"') {
            [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++
            while ($i -lt $n -and $Text[$i] -ne '"' -and $Text[$i] -ne "`n") {
                if ($Text[$i] -eq '\' -and $i + 1 -lt $n) { [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++ }
                if ($i -lt $n -and $Text[$i] -ne "`n") { [void]$sb.Replace($Text[$i], ' ', $i, 1) }
                $i++
            }
            if ($i -lt $n -and $Text[$i] -eq '"') { [void]$sb.Replace($Text[$i], ' ', $i, 1); $i++ }
        }
        else { $i++ }
    }

    return $sb.ToString()
}

foreach ($file in $files) {
    $raw = Get-Content $file.FullName -Raw
    if ([string]::IsNullOrEmpty($raw)) { continue }
    $text = Remove-CommentsAndStrings -Text $raw

    $relative = $file.FullName.Substring($repoRoot.Length + 1) -replace '\\', '/'

    foreach ($m in [regex]::Matches($text, '\bcatch\b')) {
        # Find the opening brace of the catch body, skipping any (Exception ex)
        # filter and any "when (...)" clause.
        $i = $m.Index + 5
        while ($i -lt $text.Length -and $text[$i] -ne '{' -and $text[$i] -ne ';') { $i++ }
        if ($i -ge $text.Length -or $text[$i] -ne '{') { continue }

        $end = Get-BlockEnd -Text $text -OpenIndex $i
        if ($end -lt 0) { continue }

        $catchCount++

        $body = $text.Substring($i + 1, $end - $i - 1)

        # Strip comments so that the word "throw" in prose does not excuse a swallow.
        $code = [regex]::Replace($body, '//[^\r\n]*', '')
        $code = [regex]::Replace($code, '/\*.*?\*/', '', 'Singleline')

        if ($code -match '\bthrow\b') { continue }

        $stripped = ($code -replace '\s', '')
        if ($stripped -eq 'continue;' -or $stripped -eq 'break;') { continue }

        # Rule 1b explicitly sanctions the "optional property access" probe:
        #     try { info.HasTable = Convert.ToInt32(shape.HasTable) != 0; } catch { }
        # A COM property that may not exist on this shape is probed, and absence is
        # recorded as a default. That is a value being unavailable, not an operation
        # failing, and Rule 22's "never swallow" is not aimed at it.
        #
        # The discriminator is the size of the guarded try, not the catch. A try
        # holding ONE statement is a probe. A try wrapping several statements is real
        # work, and swallowing there hides a genuine failure - which both rules forbid.
        #
        # Without this distinction the gate would fail code the instructions bless, and
        # a gate that cries wolf gets switched off.
        $tryEnd = -1
        for ($j = $m.Index - 1; $j -ge 0; $j--) {
            if ($text[$j] -match '\S') {
                if ($text[$j] -eq '}') { $tryEnd = $j }
                break
            }
        }

        $isSingleStatementTry = $false
        if ($tryEnd -ge 0) {
            $depth = 0
            for ($j = $tryEnd; $j -ge 0; $j--) {
                if ($text[$j] -eq '}') { $depth++ }
                elseif ($text[$j] -eq '{') {
                    $depth--
                    if ($depth -eq 0) {
                        $tryBody = $text.Substring($j + 1, $tryEnd - $j - 1)
                        $tryCode = [regex]::Replace($tryBody, '//[^\r\n]*', '')
                        $tryCode = [regex]::Replace($tryCode, '/\*.*?\*/', '', 'Singleline')
                        # One statement, no nested block or control flow.
                        $isSingleStatementTry = ($tryCode -notmatch '[{}]') -and
                                                (([regex]::Matches($tryCode, ';')).Count -le 1)
                        break
                    }
                }
            }
        }

        if ($isSingleStatementTry) { $probeCount++; continue }

        $line = ($text.Substring(0, $m.Index) -split "`n").Count
        $findings += "${relative}:${line}"
    }
}

if ($catchCount -eq 0) {
    Write-Host "FAIL: scanned $($files.Count) file(s) but found no catch blocks at all." -ForegroundColor Red
    Write-Host "      That is implausible for this codebase and means the parser is broken."
    exit 1
}

$findings = @($findings | Sort-Object)

if ($UpdateBaseline) {
    $findings | Set-Content -Path $baselineFile -Encoding ascii
    Write-Host "Baseline updated: $($findings.Count) swallowing catch block(s) recorded."
    exit 0
}

$baseline = @()
if (Test-Path $baselineFile) {
    $baseline = @(Get-Content $baselineFile | Where-Object { $_.Trim() -ne '' })
}

$new = @($findings | Where-Object { $_ -notin $baseline })
$fixed = @($baseline | Where-Object { $_ -notin $findings })

if (-not $Quiet) {
    Write-Host "Swallowed Catch Check (Rule 22 / Rule 1b)"
    Write-Host "========================================="
    Write-Host ""
    Write-Host "Files scanned:            $($files.Count)"
    Write-Host "Catch blocks inspected:   $catchCount"
    Write-Host "Optional-property probes: $probeCount (allowed by Rule 1b)"
    Write-Host "Swallowing catches found: $($findings.Count)"
    Write-Host "Tolerated by baseline:    $($baseline.Count)"
    Write-Host ""
}

if ($new.Count -gt 0) {
    Write-Host "FAIL: $($new.Count) NEW swallowing catch block(s):" -ForegroundColor Red
    $new | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Let the exception propagate. batch.Execute() already converts it into"
    Write-Host "OperationResult { Success = false } at the correct layer (Rule 1b), and"
    Write-Host "COM cleanup belongs in finally, not catch (Rule 22)."
    exit 1
}

if ($fixed.Count -gt 0 -and -not $Quiet) {
    Write-Host "$($fixed.Count) baselined catch block(s) no longer swallow. Lower the baseline:" -ForegroundColor Yellow
    $fixed | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    Write-Host "  .\scripts\check-swallowed-catches.ps1 -UpdateBaseline"
    Write-Host ""
}

if (-not $Quiet) {
    Write-Host "OK: no new swallowing catch blocks ($($findings.Count) tolerated by baseline)." -ForegroundColor Green
}
exit 0
