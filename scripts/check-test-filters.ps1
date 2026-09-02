# Verifies that every Feature=<Name> test filter shown in the documentation
# resolves to at least one real [Trait("Feature", "<Name>")] in the test tree.
#
# Why this gate exists: `dotnet test --filter Feature=Shape` prints
# "No test matches the given testcase filter" and then exits 0. A zero-match run
# is indistinguishable from a passing run, so an agent following the documented
# surgical-testing workflow can run nothing and report the change as tested.
# The documentation listed Shape, Text, Chart, VBA, Table and Animation; none of
# those trait values have ever existed in this fork.
#
# ASCII only. Windows PowerShell reads these files as cp1252, and a stray
# typographic character silently turns the whole script into a parse error,
# which would mean the gate never runs at all.

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

Write-Host "Checking documented test filters..."
Write-Host ""

# 1. Collect the Feature trait values that actually exist.
$testFiles = Get-ChildItem -Path (Join-Path $repoRoot 'tests') -Recurse -Include *.cs -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

$actual = @{}
foreach ($file in $testFiles) {
    foreach ($m in [regex]::Matches((Get-Content $file.FullName -Raw), '\[Trait\(\s*"Feature"\s*,\s*"([^"]+)"\s*\)\]')) {
        $actual[$m.Groups[1].Value] = $true
    }
}

if ($actual.Count -eq 0) {
    Write-Host "FAIL: found no [Trait(\"Feature\", ...)] declarations at all." -ForegroundColor Red
    Write-Host "      A gate that inspects nothing must not report success."
    exit 1
}

# 2. Collect the Feature filters the documentation tells people to run.
$docPaths = @(
    '.github\copilot-instructions.md',
    'AGENTS.md',
    'tests\README.md'
) + (Get-ChildItem -Path (Join-Path $repoRoot '.github\instructions') -Filter *.md -ErrorAction SilentlyContinue |
        ForEach-Object { $_.FullName.Substring($repoRoot.Length + 1) })

$documented = @{}
foreach ($rel in $docPaths) {
    $full = Join-Path $repoRoot $rel
    if (-not (Test-Path $full)) { continue }

    $lines = Get-Content $full
    for ($i = 0; $i -lt $lines.Count; $i++) {
        # Documentation legitimately needs to *name* the non-existent traits in
        # order to warn about them. Such a line must opt out explicitly, so that
        # the escape hatch is visible in review rather than inferred by the gate.
        if ($lines[$i] -match 'ghost-filter-ok') { continue }

        foreach ($m in [regex]::Matches($lines[$i], 'Feature\s*(?:=|!=)\s*([A-Za-z0-9_]+)')) {
            $name = $m.Groups[1].Value
            # Placeholders in prose, not real trait values.
            if ($name -in @('name', 'feature')) { continue }
            if (-not $documented.ContainsKey($name)) { $documented[$name] = @() }
            $documented[$name] += "$rel`:$($i + 1)"
        }
    }
}

if ($documented.Count -eq 0) {
    Write-Host "FAIL: found no documented Feature= filters to verify." -ForegroundColor Red
    Write-Host "      Either the docs stopped prescribing filters or this check is looking in the wrong place."
    exit 1
}

Write-Host "Documented Feature filters: $($documented.Count)"
Write-Host "Feature traits in tests:    $($actual.Count)"
Write-Host ""

# 3. Every documented filter must resolve to at least one test.
$ghosts = $documented.Keys | Where-Object { -not $actual.ContainsKey($_) } | Sort-Object

if ($ghosts.Count -gt 0) {
    Write-Host "FAIL: $($ghosts.Count) documented filter(s) match zero tests." -ForegroundColor Red
    Write-Host ""
    foreach ($g in $ghosts) {
        Write-Host "  Feature=$g" -ForegroundColor Red
        foreach ($site in ($documented[$g] | Sort-Object -Unique)) {
            Write-Host "    cited at $site"
        }
    }
    Write-Host ""
    Write-Host "'dotnet test --filter' exits 0 when nothing matches, so these instructions"
    Write-Host "produce a green run that executed no tests. Either correct the filter or add"
    Write-Host "the trait to the tests it is meant to select."
    Write-Host ""
    Write-Host "Trait values that do exist:"
    Write-Host "  $((($actual.Keys) | Sort-Object) -join ', ')"
    exit 1
}

Write-Host "All $($documented.Count) documented Feature filter(s) resolve to real tests." -ForegroundColor Green
exit 0
