# GeoTable Reports - Distribution Package Creator
# This script creates a ZIP file for distribution to end users

Write-Host "================================" -ForegroundColor Cyan
Write-Host "GeoTable Reports - Package Creator" -ForegroundColor Cyan
Write-Host "================================`n" -ForegroundColor Cyan

# Check if bundle exists in AppData
$bundlePath = "$env:APPDATA\Autodesk\ApplicationPlugins\GeoTableReports.bundle"
$projectRoot = $PSScriptRoot
$outputZip = Join-Path $projectRoot "GeoTableReports_Distribution.zip"

Write-Host "Checking for bundle..." -ForegroundColor Yellow

if (Test-Path $bundlePath) {
    Write-Host "✓ Bundle found at: $bundlePath`n" -ForegroundColor Green

    # Remove old distribution package if exists
    if (Test-Path $outputZip) {
        Write-Host "Removing old distribution package..." -ForegroundColor Yellow
        Remove-Item $outputZip -Force
    }

    # Create new distribution package
    Write-Host "Creating distribution package..." -ForegroundColor Yellow

    try {
        Compress-Archive -Path $bundlePath -DestinationPath $outputZip -Force
        Write-Host "`n✓ Distribution package created successfully!" -ForegroundColor Green
        Write-Host "   Location: $outputZip`n" -ForegroundColor White

        # Get file size
        $fileSize = (Get-Item $outputZip).Length / 1MB
        Write-Host "   Package size: $([math]::Round($fileSize, 2)) MB`n" -ForegroundColor White

        # Show contents summary
        Write-Host "Package Contents:" -ForegroundColor Cyan
        Write-Host "  └─ GeoTableReports.bundle\" -ForegroundColor White
        Write-Host "      ├─ PackageContents.xml" -ForegroundColor White
        Write-Host "      └─ Contents\Windows\2025\" -ForegroundColor White
        Write-Host "          ├─ GeoTableReports.dll" -ForegroundColor White
        Write-Host "          └─ Dependencies (iText7, etc.)`n" -ForegroundColor White

        Write-Host "Next Steps:" -ForegroundColor Cyan
        Write-Host "  1. Share '$outputZip' with end users" -ForegroundColor White
        Write-Host "  2. Users extract to: %AppData%\Autodesk\ApplicationPlugins\" -ForegroundColor White
        Write-Host "  3. Restart Civil 3D`n" -ForegroundColor White

        Write-Host "Press any key to open the output folder..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        # Open folder in Explorer
        Start-Process explorer.exe -ArgumentList "/select,`"$outputZip`""

    } catch {
        Write-Host "`n✗ Error creating package: $_" -ForegroundColor Red
        exit 1
    }

} else {
    Write-Host "✗ Bundle not found!" -ForegroundColor Red
    Write-Host "`nThe bundle needs to be built first." -ForegroundColor Yellow
    Write-Host "`nTo build the bundle:" -ForegroundColor Cyan
    Write-Host "  1. Close Civil 3D (if running)" -ForegroundColor White
    Write-Host "  2. Run: dotnet build GeoTableReports.csproj" -ForegroundColor White
    Write-Host "  3. Re-run this script`n" -ForegroundColor White

    Read-Host "Press Enter to exit"
    exit 1
}
