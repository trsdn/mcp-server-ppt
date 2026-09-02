<#
.SYNOPSIS
    Verifies every xunit test class carries the traits used for surgical test selection.

.DESCRIPTION
    Rule 16 tells contributors to run only the tests for the feature they changed, using
    --filter "Feature=<name>". That instruction is only executable if every test class
    actually carries a Feature trait. Classes without one are unreachable by any feature
    filter: they run in the full suite (45+ minutes) or they do not run at all.

    This checks per CLASS, not per file. A file whose first class is traited can still
    hide a second, untraited class, and a per-file check would call that clean.

    Required on any class that declares [Fact] or [Theory]:
      Category, Feature

    A trait counts as present if it is declared on the class, OR on every single test
    method in that class. Both make the tests reachable by a filter, which is the
    property that matters. Only Category and Feature are required, because those are
    the only trait keys any documented filter actually selects on; demanding Layer as
    well would fail classes that are perfectly reachable.

    A gate that inspects nothing must fail, so this exits non-zero if it finds no test
    classes at all - that means the discovery glob broke, not that the repo is clean.

    Complements check-test-filters.ps1, which runs the opposite direction: that gate
    proves every documented Feature filter resolves to a real trait, this one proves
    every test is reachable by some Feature filter.

.PARAMETER TestsPath
    Root of the test projects. Defaults to the repository's tests directory.

.PARAMETER MinimumClasses
    Lower bound on discovered test classes. Guards against a silent discovery failure.
#>
[CmdletBinding()]
param(
    [string]$TestsPath,
    [int]$MinimumClasses = 40
)

$ErrorActionPreference = 'Stop'

if (-not $TestsPath) {
    $TestsPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'tests'
}

if (-not (Test-Path $TestsPath)) {
    Write-Host "FAIL: tests directory not found: $TestsPath" -ForegroundColor Red
    exit 1
}

$requiredTraits = @('Category', 'Feature')

$files = Get-ChildItem -Path $TestsPath -Filter '*.cs' -Recurse -File |
    Where-Object { $_.FullName -notmatch '\\obj\\|\\bin\\' }

if ($files.Count -eq 0) {
    Write-Host "FAIL: no C# files found under $TestsPath - the discovery glob is broken." -ForegroundColor Red
    exit 1
}

$classDeclPattern = '(?m)^\s*(?:public|internal)\s+(?:sealed\s+|static\s+|abstract\s+|partial\s+)*class\s+(\w+)'

$totalClasses = 0
$violations = @()

foreach ($file in $files) {
    $lines = Get-Content -LiteralPath $file.FullName

    # Locate every class declaration with its line index.
    $decls = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $m = [regex]::Match($lines[$i], $classDeclPattern)
        if ($m.Success) {
            $decls += [pscustomobject]@{ Name = $m.Groups[1].Value; Line = $i }
        }
    }

    if ($decls.Count -eq 0) { continue }

    for ($d = 0; $d -lt $decls.Count; $d++) {
        $decl = $decls[$d]

        # Body runs from this declaration to the next one (or end of file).
        $bodyEnd = if ($d + 1 -lt $decls.Count) { $decls[$d + 1].Line - 1 } else { $lines.Count - 1 }
        $body = ($lines[$decl.Line..$bodyEnd]) -join "`n"

        if ($body -notmatch '\[Fact\b|\[Theory\b') { continue }

        $totalClasses++

        # Walk backwards over the contiguous attribute/comment block above the class.
        $attributes = @()
        for ($j = $decl.Line - 1; $j -ge 0; $j--) {
            $line = $lines[$j].Trim()
            if ($line -eq '') { continue }
            if ($line.StartsWith('[')) { $attributes += $line; continue }
            if ($line.StartsWith('///') -or $line.StartsWith('//')) { continue }
            break
        }
        $attributeBlock = $attributes -join "`n"

        $missing = @()
        foreach ($trait in $requiredTraits) {
            $token = [regex]::Escape("[Trait(`"$trait`"")
            if ($attributeBlock -match $token) { continue }

            # Not on the class. It still counts if every test method declares it,
            # because those tests remain reachable by a filter.
            $bodyLines = $lines[$decl.Line..$bodyEnd]
            $testCount = 0
            $traitedCount = 0

            for ($k = 0; $k -lt $bodyLines.Count; $k++) {
                if ($bodyLines[$k] -notmatch '\[Fact\b|\[Theory\b') { continue }
                $testCount++

                # Attributes may sit above or below [Fact]; scan the contiguous run.
                $start = $k
                while ($start -gt 0 -and $bodyLines[$start - 1].Trim().StartsWith('[')) { $start-- }
                $end = $k
                while ($end -lt $bodyLines.Count - 1 -and $bodyLines[$end + 1].Trim().StartsWith('[')) { $end++ }

                if ((($bodyLines[$start..$end]) -join "`n") -match $token) { $traitedCount++ }
            }

            if ($testCount -gt 0 -and $traitedCount -eq $testCount) { continue }

            $missing += $trait
        }

        if ($missing.Count -gt 0) {
            $relative = $file.FullName.Replace((Split-Path $PSScriptRoot -Parent), '').TrimStart('\')
            $violations += [pscustomobject]@{
                Class    = $decl.Name
                Location = "${relative}:$($decl.Line + 1)"
                Missing  = ($missing -join ', ')
            }
        }
    }
}

if ($totalClasses -eq 0) {
    Write-Host "FAIL: inspected $($files.Count) file(s) but found no test classes." -ForegroundColor Red
    Write-Host "      Class or attribute detection is broken - this gate is not testing anything." -ForegroundColor Red
    exit 1
}

if ($totalClasses -lt $MinimumClasses) {
    Write-Host "FAIL: found only $totalClasses test class(es), expected at least $MinimumClasses." -ForegroundColor Red
    Write-Host "      Either tests were deleted or discovery regressed. Lower -MinimumClasses deliberately." -ForegroundColor Red
    exit 1
}

if ($violations.Count -gt 0) {
    Write-Host "FAIL: $($violations.Count) test class(es) are missing required traits." -ForegroundColor Red
    Write-Host ""
    foreach ($v in $violations) {
        Write-Host ("  {0}" -f $v.Class) -ForegroundColor Yellow
        Write-Host ("      {0}" -f $v.Location)
        Write-Host ("      missing: {0}" -f $v.Missing) -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Every test class needs $($requiredTraits -join ' and ') traits, on the class or on every test method." -ForegroundColor Red
    Write-Host "Without a Feature trait the class cannot be reached by --filter Feature=<name> (Rule 16)." -ForegroundColor Red
    exit 1
}

Write-Host "PASS: all $totalClasses test class(es) carry $($requiredTraits -join ' and ') traits." -ForegroundColor Green
exit 0
