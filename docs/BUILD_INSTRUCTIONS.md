# Build Instructions

## Dependencies Status

The iText7 Bouncy Castle adapter dependency is already configured in the project file (`GeoTableReports.csproj`):

```xml
<PackageReference Include="itext7" Version="8.0.5" />
<PackageReference Include="itext7.bouncy-castle-adapter" Version="8.0.5" />
```

## Building the Project

To build this project on Windows:

### Option 1: Using Visual Studio
1. Open `xml-report-for-geotable.sln` in Visual Studio
2. Right-click on the solution and select **Restore NuGet Packages**
3. Build the solution (Ctrl+Shift+B or Build → Build Solution)
4. The post-build event will automatically copy the DLL and dependencies to the Civil 3D plugins folder

### Option 2: Using .NET CLI
Open PowerShell or Command Prompt in the project directory and run:

```powershell
# Restore NuGet packages
dotnet restore

# Build the project
dotnet build --configuration Release
```

## Verification

After building, verify that the following DLLs are present in the output directory:
- `GeoTableReports.dll` (main assembly)
- `itext.kernel.dll`
- `itext.layout.dll`
- `itext.io.dll`
- `BouncyCastle.Cryptography.dll` (from the bouncy-castle-adapter package)

## Post-Build Deployment

The project includes a post-build event that automatically copies the compiled DLL and all dependencies to:
```
%AppData%\Autodesk\ApplicationPlugins\GeoTableReports.bundle\Contents\Windows\2025\
```

This ensures Civil 3D can load the plugin with all required dependencies.

### Deploying to Bundle Folders

To populate the `bundles/` directory with build output for all three versions, run the deploy script from the repository root:

```powershell
.\deploy-bundles.ps1                        # copies Release builds (default)
.\deploy-bundles.ps1 -Configuration Debug   # copies Debug builds
```

The script copies `*.dll` and `*.dll.config` files from each version's `bin/{Configuration}/` into the corresponding `bundles/GeoTableReports.{year}.bundle/Contents/Windows/{year}/` folder. Any version whose build output doesn't exist is skipped with a warning.

## Troubleshooting

If you encounter the error:
```
NotSupportedException - Either com.itextpdf.bouncy-castle-adapter or
com.itextpdf.bouncy-castle-fips-adapter dependency must be added
```

This typically means:
1. NuGet packages haven't been restored
2. The Bouncy Castle adapter DLL isn't being deployed with the main assembly
3. The package restore failed

**Solution:**
- Clean the solution: `dotnet clean` or Build → Clean Solution
- Delete the `bin` and `obj` folders
- Restore packages: `dotnet restore`
- Rebuild the solution
- Verify all DLLs are copied to the Civil 3D plugins folder
