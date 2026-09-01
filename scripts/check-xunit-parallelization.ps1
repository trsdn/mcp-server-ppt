#Requires -Version 5.1
<#
.SYNOPSIS
    Verifies every test project's xunit.runner.json actually reaches the build output.

.DESCRIPTION
    PptMcp drives a single out-of-process PowerPoint instance over COM. Running test
    collections concurrently inside one assembly means several sessions competing for
    that instance, which produces failures that vanish on re-run in isolation.

    Every test project therefore ships an xunit.runner.json that sets
    parallelizeTestCollections to false and maxParallelThreads to 1.

    xunit v2 only honours that file when it sits next to the test assembly. Unlike
    xunit v3, the v2 package does not copy it for you - the project must carry an
    explicit None/CopyToOutputDirectory item. A project can hold a correct-looking
    xunit.runner.json in source, have it silently omitted from bin, and run fully
    parallel anyway. That is what this check exists to catch: three of the four
    projects were in exactly that state.

    Fails when a project declares the config but does not copy it, and fails when it
    finds no test projects at all - a check that inspects nothing must not report
    success.

.EXAMPLE
    .\scripts\check-xunit-parallelization.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$testsRoot = Join-Path $repoRoot 'tests'

if (-not (Test-Path $testsRoot)) {
    Write-Host "FAIL: tests directory not found at $testsRoot" -ForegroundColor Red
    exit 1
}

$configs = Get-ChildItem -Path $testsRoot -Filter 'xunit.runner.json' -Recurse |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }

if ($configs.Count -eq 0) {
    Write-Host "FAIL: no xunit.runner.json files found under tests/." -ForegroundColor Red
    Write-Host "      This check inspected nothing, so it cannot report success." -ForegroundColor Red
    exit 1
}

$failures = @()
$checked = 0

foreach ($config in $configs) {
    $projectDir = $config.Directory
    $csproj = Get-ChildItem -Path $projectDir.FullName -Filter '*.csproj' | Select-Object -First 1

    if ($null -eq $csproj) {
        $failures += "$($config.FullName): no .csproj alongside xunit.runner.json"
        continue
    }

    $checked++
    $projectName = $csproj.BaseName
    $content = Get-Content $csproj.FullName -Raw

    # The item may be written as <None Update="..."> or <None Include="...">, with the
    # CopyToOutputDirectory either as a child element or an attribute.
    $copies = $content -match '<None\s+(Update|Include)="xunit\.runner\.json"[^>]*?(/>|>)' -and
              $content -match 'xunit\.runner\.json[\s\S]{0,200}?CopyToOutputDirectory'

    if (-not $copies) {
        $failures += "$projectName does not copy xunit.runner.json to the output directory"
        continue
    }

    # The config must actually serialize collections; a copied file with the wrong
    # contents is no better than an absent one.
    $json = Get-Content $config.FullName -Raw | ConvertFrom-Json

    if ($json.parallelizeTestCollections -ne $false) {
        $failures += "$projectName sets parallelizeTestCollections to '$($json.parallelizeTestCollections)'; expected false"
    }
    if ($json.maxParallelThreads -ne 1) {
        $failures += "$projectName sets maxParallelThreads to '$($json.maxParallelThreads)'; expected 1"
    }
}

if ($checked -eq 0) {
    Write-Host "FAIL: found xunit.runner.json files but could not resolve any project." -ForegroundColor Red
    exit 1
}

if ($failures.Count -gt 0) {
    Write-Host "FAIL: xunit parallelization config is not reaching the build output." -ForegroundColor Red
    Write-Host ""
    foreach ($f in $failures) {
        Write-Host "  - $f" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "  Add this to the project so xunit v2 can read the file at run time:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host '    <ItemGroup>' -ForegroundColor Yellow
    Write-Host '      <None Update="xunit.runner.json">' -ForegroundColor Yellow
    Write-Host '        <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>' -ForegroundColor Yellow
    Write-Host '      </None>' -ForegroundColor Yellow
    Write-Host '    </ItemGroup>' -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host "OK: $checked test project(s) copy xunit.runner.json and serialize test collections." -ForegroundColor Green
exit 0
