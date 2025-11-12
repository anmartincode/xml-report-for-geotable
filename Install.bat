@echo off
REM GeoTable Reports - Quick Installer
REM This batch file installs the add-in to Civil 3D

color 0A
title GeoTable Reports - Installer

echo.
echo ========================================
echo  GeoTable Reports - Installer
echo ========================================
echo.

REM Check if running from extracted bundle folder
if not exist "GeoTableReports.bundle" (
    echo ERROR: GeoTableReports.bundle folder not found!
    echo.
    echo Please make sure you:
    echo   1. Extracted the ZIP file completely
    echo   2. Are running this script from the extracted folder
    echo.
    pause
    exit /b 1
)

echo Checking installation path...
set "DEST=%APPDATA%\Autodesk\ApplicationPlugins\GeoTableReports.bundle"

echo.
echo Installation will copy files to:
echo %DEST%
echo.

REM Check if already installed
if exist "%DEST%" (
    echo WARNING: GeoTable Reports is already installed.
    echo.
    choice /C YN /M "Do you want to overwrite the existing installation"
    if errorlevel 2 goto :cancel
    echo.
    echo Removing old installation...
    rmdir /S /Q "%DEST%" 2>nul
)

echo.
echo Installing GeoTable Reports...

REM Create destination directory
if not exist "%APPDATA%\Autodesk\ApplicationPlugins\" mkdir "%APPDATA%\Autodesk\ApplicationPlugins\"

REM Copy bundle
echo Copying files...
xcopy /E /I /Y /Q "GeoTableReports.bundle" "%DEST%\"

if errorlevel 1 (
    echo.
    echo ERROR: Installation failed!
    echo Please check permissions and try again.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Installation Complete!
echo ========================================
echo.
echo The GeoTable Reports add-in has been installed.
echo.
echo Next Steps:
echo   1. Close this window
echo   2. Restart Civil 3D (if currently open)
echo   3. The add-in will load automatically
echo.
echo Commands available in Civil 3D:
echo   - GEOTABLE_PANEL : Open the dockable panel
echo   - GEOTABLE       : Quick report generation
echo   - GEOTABLE_BATCH : Batch process all alignments
echo.
pause
exit /b 0

:cancel
echo.
echo Installation cancelled.
echo.
pause
exit /b 0
