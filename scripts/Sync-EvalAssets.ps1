param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Import", "Export")]
    [string]$Direction,

    [Parameter(Mandatory = $true)]
    [string]$PrivateRepoRoot,

    [string]$PublicRepoRoot = (Join-Path $PSScriptRoot ".."),

    [switch]$Input,
    [switch]$Output,
    [switch]$Results,
    [switch]$Data
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-NormalizedRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$FullPath
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $fullItemPath = [System.IO.Path]::GetFullPath($FullPath)

    if (-not $baseFullPath.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $baseFullPath = "$baseFullPath$([System.IO.Path]::DirectorySeparatorChar)"
    }

    $relativePath = $fullItemPath.Substring($baseFullPath.Length)
    return $relativePath.Replace('/', '\')
}

function Should-SkipEvalAsset {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,

        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    if ($Section -ne "input") {
        return $false
    }

    $normalizedPath = $RelativePath.Replace('/', '\')
    $leafName = Split-Path -Path $normalizedPath -Leaf

    if ($normalizedPath -ieq "master.pdf") {
        return $true
    }

    return $leafName -like "extract*"
}

function Copy-EvalSection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,

        [Parameter(Mandatory = $true)]
        [string]$SourceEvalRoot,

        [Parameter(Mandatory = $true)]
        [string]$TargetEvalRoot
    )

    $sourceSectionRoot = Join-Path $SourceEvalRoot $Section
    $targetSectionRoot = Join-Path $TargetEvalRoot $Section

    if (-not (Test-Path -LiteralPath $sourceSectionRoot)) {
        Write-Host "Skipping missing source section: $sourceSectionRoot"
        return
    }

    Get-ChildItem -LiteralPath $sourceSectionRoot -Recurse -File | ForEach-Object {
        $relativePath = Get-NormalizedRelativePath -BasePath $sourceSectionRoot -FullPath $_.FullName
        if (Should-SkipEvalAsset -Section $Section -RelativePath $relativePath) {
            Write-Host "Skipping excluded asset: eval\$Section\$relativePath"
            return
        }

        $targetPath = Join-Path $targetSectionRoot $relativePath
        $targetDirectory = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
        }

        Copy-Item -LiteralPath $_.FullName -Destination $targetPath -Force
        Write-Host "Copied eval\$Section\$relativePath"
    }
}

$resolvedPrivateRepoRoot = Resolve-ExistingPath -Path $PrivateRepoRoot
$resolvedPublicRepoRoot = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $PublicRepoRoot).Path)

$privateEvalRoot = Join-Path $resolvedPrivateRepoRoot "eval"
$publicEvalRoot = Join-Path $resolvedPublicRepoRoot "eval"

if (-not (Test-Path -LiteralPath $privateEvalRoot)) {
    throw "Private repo does not contain an eval directory: $privateEvalRoot"
}

if (-not (Test-Path -LiteralPath $publicEvalRoot)) {
    throw "Public repo does not contain an eval directory: $publicEvalRoot"
}

$selectedSections = @()
if ($Input) { $selectedSections += "input" }
if ($Output) { $selectedSections += "output" }
if ($Results) { $selectedSections += "results" }
if ($Data) { $selectedSections += "data" }
if ($selectedSections.Count -eq 0) {
    $selectedSections = @("input", "output", "results", "data")
}

switch ($Direction) {
    "Import" {
        $sourceEvalRoot = $privateEvalRoot
        $targetEvalRoot = $publicEvalRoot
    }
    "Export" {
        $sourceEvalRoot = $publicEvalRoot
        $targetEvalRoot = $privateEvalRoot
    }
}

Write-Host "Direction: $Direction"
Write-Host "Source eval root: $sourceEvalRoot"
Write-Host "Target eval root: $targetEvalRoot"
Write-Host "Sections: $($selectedSections -join ', ')"
Write-Host "Excluded input assets: master.pdf, extract*"

foreach ($section in $selectedSections) {
    Copy-EvalSection -Section $section -SourceEvalRoot $sourceEvalRoot -TargetEvalRoot $targetEvalRoot
}
