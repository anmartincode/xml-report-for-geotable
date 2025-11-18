# GeoTable Reports - Multi-Target Build (Option 1)

This folder contains a **single project file** that builds for all three Civil3D versions (2023, 2024, 2025) simultaneously.

## How It Works

The multi-target build uses MSBuild's multi-targeting feature to compile the same source code against different target frameworks:

- **net48** → Used for Civil3D 2023 and 2024
- **net8.0-windows** → Used for Civil3D 2025

## Build Commands

### Build All Versions at Once
```bash
dotnet build GeoTableReports.csproj
```

This will produce:
- `bin/Debug/net48/GeoTableReports.dll` → For Civil3D 2023/2024
- `bin/Debug/net8.0-windows/GeoTableReports.dll` → For Civil3D 2025

### Build Specific Version
```bash
# For 2023/2024 (.NET Framework 4.8)
dotnet build GeoTableReports.csproj -f net48

# For 2025 (.NET 8.0)
dotnet build GeoTableReports.csproj -f net8.0-windows
```

## Important Notes

### ⚠️ Limitation: net48 Builds Both 2023 and 2024
Since both Civil3D 2023 and 2024 use .NET Framework 4.8, the multi-target approach cannot distinguish between them in a single build. You have two options:

**Option A: Use Separate Folders (Recommended)**
- Use the separate project folders: `GeoTableReports-2023/`, `GeoTableReports-2024/`, `GeoTableReports-2025/`
- Each builds independently with version-specific configurations

**Option B: Post-Build Copying**
- Build the net48 target once
- Manually copy and rename the output to separate folders for 2023 and 2024
- Update assembly names using IL rewriting or build scripts

## Advantages of Multi-Target Build

✅ **Single Source of Truth** - One project file maintains all versions
✅ **Less Maintenance** - Changes apply to all versions automatically
✅ **Consistent Dependencies** - Same NuGet packages across versions
✅ **Faster Development** - One build command for all versions

## Disadvantages

❌ **Cannot distinguish 2023 vs 2024** - Both use net48
❌ **More complex project file** - Requires conditional logic
❌ **Harder to debug** - Version-specific issues harder to isolate
❌ **Assembly naming** - Cannot easily have separate assembly names per version with net48

## Recommended Approach

For this project, **Option 3 (Separate Packages)** from the original analysis is recommended because:

1. Civil3D 2023 and 2024 both use net48 - can't distinguish in multi-target
2. Each version needs unique assembly names (GeoTableReports.2023.dll, etc.)
3. Version-specific configurations are clearer in separate projects
4. Easier to test and troubleshoot version-specific issues

## Converting This to Work

To make this work properly, you would need to:

1. Use build parameters: `-p:Civil3DVersion=2023` or `-p:Civil3DVersion=2024`
2. Add post-build scripts to rename assemblies
3. Use separate output directories for each version

Example:
```bash
dotnet build -f net48 -p:Civil3DVersion=2023
dotnet build -f net48 -p:Civil3DVersion=2024
dotnet build -f net8.0-windows
```

## Files in This Folder

- `GeoTableReports.csproj` - Multi-target project file
- `PackageContents.2023.xml` - Metadata for Civil3D 2023
- `PackageContents.2024.xml` - Metadata for Civil3D 2024
- `PackageContents.2025.xml` - Metadata for Civil3D 2025
- `dotnet_starter_template.cs` - Shared source code
- `app.config` - Application configuration

## See Also

- **GeoTableReports-2023/** - Separate project for 2023 (Recommended)
- **GeoTableReports-2024/** - Separate project for 2024 (Recommended)
- **GeoTableReports-2025/** - Separate project for 2025 (Recommended)
