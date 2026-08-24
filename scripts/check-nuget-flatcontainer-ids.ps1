<#
.SYNOPSIS
    Verifies every NuGet flat container URL uses an all-lowercase package id.

.DESCRIPTION
    The NuGet v3 "package base address" (flat container) API requires the package id
    to be lowercased. Any other casing returns 404:
      GET {@id}/{LOWER_ID}/index.json   -- "The package ID, lowercased"
      https://learn.microsoft.com/nuget/api/package-base-address-resource

    This has bitten the repository twice, and both times it failed silently:
      * the release workflow polled a mixed-case readme URL and burned 20 minutes
        of retries on every release before falling through, and
      * the shipped CLI and MCP Server update checks requested a mixed-case index
        URL, so every check 404'd into a catch-all and users were never told that
        a newer version existed.

    Because the failure mode is a silent 404 rather than a crash, it cannot be
    caught by running the code. This check reads the source instead.

.NOTES
    ASCII output only - the pre-commit hook runs under consoles without UTF-8.
#>

$ErrorActionPreference = "Stop"

$rootDir = Split-Path -Parent $PSScriptRoot
$violations = @()

# CHANGELOG.md documents the historical mixed-case values on purpose.
$excludedFiles = @("CHANGELOG.md")
$excludedDirs = @("\bin\", "\obj\", "\node_modules\", "\.git\", "\artifacts\", "\packages\")

$files = Get-ChildItem -Path $rootDir -Recurse -File -Include "*.cs", "*.yml", "*.yaml", "*.ps1", "*.md" |
    Where-Object {
        $path = $_.FullName
        if ($excludedFiles -contains $_.Name) { return $false }
        foreach ($dir in $excludedDirs) {
            if ($path -like "*$dir*") { return $false }
        }
        return $true
    }

foreach ($file in $files) {
    # @() keeps single-line files as an array - indexing a bare string yields characters.
    $lines = @(Get-Content $file.FullName)
    $relativePath = $file.FullName.Substring($rootDir.Length + 1)

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $lineNumber = $i + 1

        # A literal id directly in a flat container URL.
        foreach ($match in [regex]::Matches($line, 'v3-flatcontainer/([^/\s"'')]+)')) {
            $id = $match.Groups[1].Value

            # Skip interpolated values - the constant itself is checked below.
            if ($id -match '[\$\{\<]') { continue }

            if ($id -cmatch '[A-Z]') {
                $violations += [PSCustomObject]@{
                    File = $relativePath
                    Line = $lineNumber
                    Found = $id
                    Expected = $id.ToLowerInvariant()
                    Text = $line.Trim()
                }
            }
        }

        # The constant that gets interpolated into the URL.
        foreach ($match in [regex]::Matches($line, 'PackageId\s*=\s*"([^"]+)"')) {
            $id = $match.Groups[1].Value

            if ($id -cmatch '[A-Z]') {
                $violations += [PSCustomObject]@{
                    File = $relativePath
                    Line = $lineNumber
                    Found = $id
                    Expected = $id.ToLowerInvariant()
                    Text = $line.Trim()
                }
            }
        }
    }
}

if ($violations.Count -gt 0) {
    Write-Host ""
    Write-Host "Mixed-case NuGet flat container package id(s) found:" -ForegroundColor Red
    Write-Host ""

    foreach ($violation in $violations) {
        Write-Host ("  {0}:{1}" -f $violation.File, $violation.Line) -ForegroundColor Yellow
        Write-Host ("    {0}" -f $violation.Text) -ForegroundColor Gray
        Write-Host ("    found '{0}' - expected '{1}'" -f $violation.Found, $violation.Expected) -ForegroundColor Red
        Write-Host ""
    }

    Write-Host "The flat container API returns 404 for any casing other than lowercase," -ForegroundColor Red
    Write-Host "and callers here swallow that 404, so the failure is invisible at runtime." -ForegroundColor Red
    Write-Host "See https://learn.microsoft.com/nuget/api/package-base-address-resource" -ForegroundColor Red
    exit 1
}

Write-Host ("Flat container id check passed - {0} files scanned" -f $files.Count) -ForegroundColor Green
exit 0
