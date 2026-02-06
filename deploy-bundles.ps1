<#
.SYNOPSIS
    Copies build output into the corresponding Civil 3D bundle folders.

.DESCRIPTION
    For each version (2023, 2024, 2025), copies DLLs and .config files from
    the project's bin output into bundles/GeoTableReports.<year>.bundle/Contents/Windows/<year>/

.PARAMETER Configuration
    Build configuration to deploy. Default: Release

.EXAMPLE
    .\deploy-bundles.ps1
    .\deploy-bundles.ps1 -Configuration Debug
#>

param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$versions = @("2023", "2024", "2025")

Write-Host "Deploying $Configuration builds to bundles..." -ForegroundColor Cyan
Write-Host ""

foreach ($year in $versions) {
    $source = Join-Path (Join-Path (Join-Path $scriptDir "GeoTableReports-$year") "bin") $Configuration
    $target = Join-Path (Join-Path (Join-Path (Join-Path (Join-Path $scriptDir "bundles") "GeoTableReports.$year.bundle") "Contents") "Windows") $year

    Write-Host "[$year] Source: $source" -ForegroundColor DarkGray
    Write-Host "[$year] Target: $target" -ForegroundColor DarkGray

    # Validate source exists and has DLLs
    if (-not (Test-Path $source)) {
        Write-Host "[$year] SKIPPED - Build output not found. Run 'msbuild /p:Configuration=$Configuration' for GeoTableReports-$year first." -ForegroundColor Yellow
        Write-Host ""
        continue
    }

    $dlls = Get-ChildItem $source -Filter "*.dll" -File
    if ($dlls.Count -eq 0) {
        Write-Host "[$year] SKIPPED - No DLLs found in build output." -ForegroundColor Yellow
        Write-Host ""
        continue
    }

    # Create target directory if needed
    if (-not (Test-Path $target)) {
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        Write-Host "[$year] Created target directory." -ForegroundColor DarkGray
    }

    # Copy DLLs
    $dllCount = 0
    foreach ($dll in $dlls) {
        Copy-Item $dll.FullName -Destination $target -Force
        $dllCount++
    }

    # Copy .config files (e.g. GeoTableReports.2025.dll.config)
    $configCount = 0
    $configs = Get-ChildItem $source -Filter "*.dll.config" -File
    foreach ($cfg in $configs) {
        Copy-Item $cfg.FullName -Destination $target -Force
        $configCount++
    }

    $totalSize = [math]::Round(((Get-ChildItem $target -File | Measure-Object -Property Length -Sum).Sum / 1MB), 1)
    Write-Host "[$year] Copied $dllCount DLLs + $configCount config files ($totalSize MB)" -ForegroundColor Green
    Write-Host ""
}

Write-Host "Bundle deployment complete." -ForegroundColor Cyan
