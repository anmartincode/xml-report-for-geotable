# GeoTable Reports - Installation Guide

This guide explains how to install and distribute the GeoTable Reports add-in for Civil 3D.

## For End Users

### Method 1: Bundle Installation (Recommended - Auto-loads)

This method automatically loads the add-in every time Civil 3D starts.

#### Installation Steps:

1. **Download** the `GeoTableReports.bundle.zip` file

2. **Extract the ZIP** to get the `GeoTableReports.bundle` folder

3. **Copy the bundle folder** to your ApplicationPlugins directory:
   ```
   C:\Users\[YourUsername]\AppData\Roaming\Autodesk\ApplicationPlugins\
   ```

   **Quick Access**:
   - Press `Windows + R`
   - Type: `%AppData%\Autodesk\ApplicationPlugins`
   - Press Enter
   - Paste the `GeoTableReports.bundle` folder here

4. **Restart Civil 3D**

5. **Verify installation**: The add-in should load automatically. You'll see a message in the command line:
   ```
   GeoTable Reports loaded. Type GEOTABLE_PANEL to open the panel or GEOTABLE for quick generation.
   ```

#### Folder Structure After Installation:
```
C:\Users\[YourUsername]\AppData\Roaming\Autodesk\ApplicationPlugins\
└── GeoTableReports.bundle\
    ├── PackageContents.xml
    └── Contents\
        └── Windows\
            └── 2025\
                ├── GeoTableReports.dll
                ├── itext.kernel.dll
                ├── itext.layout.dll
                ├── BouncyCastle.Cryptography.dll
                └── (other dependency DLLs)
```

---

### Method 2: Manual NETLOAD (Testing/Temporary Use)

Use this method for quick testing without permanent installation.

#### Installation Steps:

1. **Download** the DLLs (can be found in `bin\Debug\` folder)

2. **Extract to a local folder**, for example:
   ```
   C:\CustomPlugins\GeoTableReports\
   ```

3. **Open Civil 3D**

4. **Type** `NETLOAD` in the command line

5. **Browse** to `GeoTableReports.dll` and click Open

6. **Use the add-in** - Available until you close Civil 3D

⚠️ **Note**: You must run NETLOAD every time you start Civil 3D with this method.

---

### Method 3: Network Deployment (For IT Administrators)

Deploy the add-in to multiple users via network share.

#### Setup Steps:

1. **Place bundle on network share**:
   ```
   \\YourServer\CADStandards\Civil3DPlugins\GeoTableReports.bundle\
   ```

2. **Configure Civil 3D on each workstation** (one-time setup):
   - Open Civil 3D
   - Type `OPTIONS`
   - Go to **Files** tab
   - Expand **Support File Search Path**
   - Click **Add** and browse to: `\\YourServer\CADStandards\Civil3DPlugins\`
   - Click **OK**

3. **Restart Civil 3D** - The add-in loads automatically from the network

**Benefits**:
- Central management - update once, affects all users
- No local installation on each machine
- Consistent version across organization

---

## For Developers/Distributors

### Creating a Distribution Package

#### Option A: Build and Deploy to Bundles (Recommended)

Build each version, then run the deploy script to populate the bundle folders:

```powershell
# Build all versions
cd GeoTableReports-2023; dotnet build -c Release; cd ..
cd GeoTableReports-2024; dotnet build -c Release; cd ..
cd GeoTableReports-2025; dotnet build -c Release; cd ..

# Deploy build output to bundle folders
.\deploy-bundles.ps1
# Or for Debug builds:
.\deploy-bundles.ps1 -Configuration Debug
```

The `deploy-bundles.ps1` script copies DLLs and config files from each version's `bin/{Configuration}/` into `bundles/GeoTableReports.{year}.bundle/Contents/Windows/{year}/`.

To create a distributable ZIP from a bundle:

```powershell
Compress-Archive -Path "bundles\GeoTableReports.2025.bundle" -DestinationPath ".\GeoTableReports.2025.bundle.zip" -Force
```

#### Option B: Installer Script

Create a simple batch installer:

```batch
@echo off
echo Installing GeoTable Reports for Civil 3D...

set "DEST=%APPDATA%\Autodesk\ApplicationPlugins\GeoTableReports.bundle"

if not exist "%DEST%" mkdir "%DEST%"

echo Copying files...
xcopy /E /I /Y "GeoTableReports.bundle\*" "%DEST%\"

echo.
echo Installation complete!
echo Please restart Civil 3D.
pause
```

---

## Using the Add-in

### Available Commands:

1. **GEOTABLE_PANEL** - Opens the dockable panel interface
   - Configure settings
   - Select report type (Vertical/Horizontal)
   - Choose output formats (PDF, TXT, XML)
   - Generate reports or batch process

2. **GEOTABLE** - Quick report generation via dialog
   - Faster for one-off reports
   - Dialog-based interface

3. **GEOTABLE_BATCH** - Batch process all alignments
   - Processes all alignments in the drawing
   - Generates reports for each alignment

### Report Types:

- **Vertical Alignment Reports**: Station, elevation, grades, PVI points, curves, K-values
- **Horizontal Alignment Reports**: Coordinates, bearings, tangents, curves, spirals

### Output Formats:

- **PDF**: Professional, printable format
- **TXT**: Plain text, editable
- **XML**: Structured data for integration

---

## Troubleshooting

### "Could not load file or assembly" Error

**Cause**: Missing dependency DLLs

**Solution**: Ensure ALL DLLs are present in the bundle Contents folder, not just GeoTableReports.dll

### Add-in Doesn't Auto-load

**Cause**: Bundle not in correct location or PackageContents.xml missing

**Solution**:
1. Verify bundle location: `%AppData%\Autodesk\ApplicationPlugins\`
2. Check PackageContents.xml exists in bundle root
3. Restart Civil 3D

### Commands Not Recognized

**Cause**: Add-in failed to load

**Solution**:
1. Check command line for error messages
2. Try NETLOAD manually to see specific error
3. Verify Civil 3D version (2023, 2024, or 2025)

### "This assembly is built by a runtime newer than..."

**Cause**: Civil 3D version mismatch

**Solution**: Ensure you're using the correct bundle version (2023, 2024, or 2025)

---

## Version Compatibility

- **Civil 3D 2023**: Requires .NET Framework 4.8
- **Civil 3D 2024**: Requires .NET Framework 4.8
- **Civil 3D 2025**: Requires .NET 8.0

Current build targets: **Civil 3D 2025** (.NET 8.0)

---

## Support

For issues, feature requests, or questions:
- Check the troubleshooting section above
- Review the README.md for usage examples
- Contact your CAD administrator

---

## Uninstallation

### To Remove the Add-in:

1. **Close Civil 3D**

2. **Delete the bundle folder**:
   ```
   C:\Users\[YourUsername]\AppData\Roaming\Autodesk\ApplicationPlugins\GeoTableReports.bundle\
   ```

3. **Restart Civil 3D** (if needed)

---

## License

This add-in is provided as-is. See LICENSE file for details.
