@echo off
REM Civil 3D Geotable XML Report Generator
REM Environment Setup Script for Windows
REM
REM This script helps set up the working environment for the project

echo ========================================
echo Civil 3D Geotable XML Report Generator
echo Environment Setup
echo ========================================
echo.

REM Check if running from project directory
if not exist "civil3d_geotable_extractor.py" (
    echo ERROR: This script must be run from the project root directory!
    echo Please navigate to the xml-report-for-geotable directory and try again.
    pause
    exit /b 1
)

echo [1/5] Checking project structure...
echo.

REM Create output directory if it doesn't exist
if not exist "output" (
    echo Creating output directory...
    mkdir output
    echo   [OK] Created: output\
) else (
    echo   [OK] Found: output\
)

REM Create backup directory
if not exist "backup" (
    echo Creating backup directory...
    mkdir backup
    echo   [OK] Created: backup\
) else (
    echo   [OK] Found: backup\
)

echo.
echo [2/5] Verifying core files...
echo.

REM Check for required Python scripts
set MISSING_FILES=0

if exist "civil3d_geotable_extractor.py" (
    echo   [OK] civil3d_geotable_extractor.py
) else (
    echo   [MISSING] civil3d_geotable_extractor.py
    set MISSING_FILES=1
)

if exist "xml_report_generator.py" (
    echo   [OK] xml_report_generator.py
) else (
    echo   [MISSING] xml_report_generator.py
    set MISSING_FILES=1
)

if exist "GeoTableReport.dyn" (
    echo   [OK] GeoTableReport.dyn
) else (
    echo   [MISSING] GeoTableReport.dyn
    set MISSING_FILES=1
)

if exist "config.json" (
    echo   [OK] config.json
) else (
    echo   [MISSING] config.json
    set MISSING_FILES=1
)

if %MISSING_FILES%==1 (
    echo.
    echo WARNING: Some required files are missing!
    echo Please ensure you have all files from the repository.
    pause
)

echo.
echo [3/5] Checking Civil 3D installation...
echo.

REM Common Civil 3D installation paths
set C3D_FOUND=0
set C3D_VERSIONS=2024 2023 2022 2021 2020

for %%V in (%C3D_VERSIONS%) do (
    if exist "C:\Program Files\Autodesk\AutoCAD %%V\C3D" (
        echo   [OK] Found Civil 3D %%V
        set C3D_FOUND=1
    )
)

if %C3D_FOUND%==0 (
    echo   [WARNING] Civil 3D installation not found in standard locations
    echo   Please verify Civil 3D is installed before using this tool
)

echo.
echo [4/5] Configuration summary...
echo.

REM Display configuration
echo   Project Directory: %CD%
echo   Output Directory:  %CD%\output
echo   Backup Directory:  %CD%\backup
echo.

echo [5/5] Setup complete!
echo.

REM Display next steps
echo ========================================
echo Next Steps:
echo ========================================
echo.
echo 1. Open Autodesk Civil 3D with a drawing containing alignments
echo.
echo 2. Launch Dynamo (Manage tab -^> Dynamo button)
echo.
echo 3. Open GeoTableReport.dyn in Dynamo
echo.
echo 4. Copy the Python scripts into the Python Script nodes:
echo    - Node 1: Copy content from civil3d_geotable_extractor.py
echo    - Node 2: Copy content from xml_report_generator.py
echo.
echo 5. Configure your input parameters:
echo    - Alignment name (or leave empty for all)
echo    - Station interval (e.g., 10.0)
echo    - Output file path (e.g., %CD%\output\report.xml)
echo.
echo 6. Run the Dynamo script and check the output directory
echo.
echo ========================================
echo For detailed instructions, see:
echo    - README.md (comprehensive guide)
echo    - QUICKSTART.md (quick start guide)
echo ========================================
echo.

REM Create a sample config if needed
if not exist "output\README.txt" (
    echo Creating output directory README...
    (
        echo XML reports will be generated in this directory.
        echo.
        echo Generated files will follow this naming pattern:
        echo   {project_name}_{alignment_name}_geotable_{timestamp}.xml
        echo.
        echo You can customize the output location in config.json
        echo or by specifying a different path in Dynamo.
    ) > "output\README.txt"
)

echo Setup script completed successfully!
echo.
pause



