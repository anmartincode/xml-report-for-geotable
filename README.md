# GeoTable Reports for Civil 3D

GeoTable Reports is a Civil 3D add-in that generates professional horizontal and vertical alignment reports in multiple output formats. It reads alignment geometry directly from a Civil 3D drawing and produces GeoTable-formatted documents in Excel (`.xlsx`) and PDF, as well as traditional alignment reports in PDF and XML. The output follows the InRoads-style GeoTable specification used in rail and highway survey documentation.

The add-in ships as three independent builds, one for each supported Civil 3D release: 2023, 2024, and 2025. Each build targets the corresponding AutoCAD API series and .NET runtime so that it loads natively without compatibility shims.

---

## Repository Layout

The repository is organized into version-specific project folders and shared resources.

**GeoTableReports-2023**, **GeoTableReports-2024**, and **GeoTableReports-2025** each contain a self-contained Visual Studio project (`GeoTableReports.csproj`) along with its own `PackageContents.xml` manifest and build output directories. The core source file in every version is `civil3d_report_app.cs`, which houses all command logic, report generation methods, the dialog form, and helper utilities in a single compilation unit. The three files are functionally identical; only the class name suffixes (e.g., `ReportCommands2023` vs `ReportCommands2025`) and target framework differ.

**bundles/** holds the pre-built Autodesk ApplicationPlugins bundle folders (`GeoTableReports.2023.bundle`, `GeoTableReports.2024.bundle`, `GeoTableReports.2025.bundle`) ready for deployment. Each bundle contains a `PackageContents.xml` and a `Contents/Windows/{year}/` directory with the compiled DLL and all runtime dependencies.

**docs/** contains supplementary documentation including build instructions, installation guides, and reference material for the GeoTable Excel sheet structure.

**reference/** contains the `TrackGeo_Condensed_V11.xsl` stylesheet used as a formatting reference for the GeoTable layout.

**test/** contains an alignment API explorer project (`AlignmentApiExplorer.csproj`) used during development to inspect Civil 3D alignment properties and spiral geometry.

---

## Supported Civil 3D Versions

| Version | Target Framework | AutoCAD Series | Assembly Name |
|---------|-----------------|----------------|---------------|
| Civil 3D 2023 | .NET Framework 4.8 | R24.2 | GeoTableReports.2023.dll |
| Civil 3D 2024 | .NET Framework 4.8 | R24.3 | GeoTableReports.2024.dll |
| Civil 3D 2025 | .NET 8.0 (Windows) | R25.0 | GeoTableReports.2025.dll |

---

## Commands

The add-in registers two commands with Civil 3D:

**GEOTABLE** is the primary command. It prompts the user to select an alignment in the current drawing, then opens a dialog where the user chooses which report formats to generate (Alignment PDF, Alignment XML, GeoTable PDF, GeoTable Excel), sets the output base path, and clicks OK. The add-in extracts alignment entities (tangents, curves, spirals, and SCS compound entities), computes geometry, and writes the selected output files. A progress window displays status during generation, and optionally the output folder or files are opened when complete.

**GEOTABLE_BATCH** processes every alignment in the active Civil 3D document without manual selection. It creates a `GeoTable_Batch` folder in the user's Documents directory and generates a horizontal alignment PDF for each alignment found.

---

## Report Formats

### GeoTable Excel

Produces a `.xlsx` workbook with a single sheet named after the alignment. The sheet follows the standard GeoTable column layout: Element, Curve No., Point, Station, Bearing, Northing, Easting, and a four-column Data region (columns H through K). Each alignment entity is written as a block of rows:

Tangent rows display only the bearing and a `L = {length}` data value, with no point label, station, or coordinates, matching the condensed tangent format used in printed GeoTables.

Curve rows use neighbor-aware point labels. A curve preceded by a spiral shows `SC` at its start instead of `PC`, and a curve followed by a spiral shows `CS` at its end instead of `PT`. The data region contains delta, degree of curvature, radius, arc length, design speed, actual and equilibrium superelevation, tangent distance, external distance, and chord center coordinates across two rows.

Spiral rows show a single point label (`TS` for entry spirals, `ST` for exit spirals) with station and coordinates on the first row, and spiral parameters across two data rows: deflection angle, spiral length, long tangent, short tangent on the first, and X/Y coordinates, throw, and short tangent K on the second.

SCS (Spiral-Curve-Spiral) compound entities are expanded into their individual spiral and arc components and written using the same rules above.

### GeoTable PDF

Produces a formatted PDF document with the same GeoTable content rendered in a fixed-width tabular layout suitable for printing and plan sheet inclusion.

### Alignment PDF

A detailed printable report of the alignment geometry with section headers, entity-by-entity breakdowns, and coordinate data.

### Alignment XML

A structured XML export of the alignment data for use in interoperability workflows or downstream processing tools.

---

## Building from Source

### Prerequisites

Civil 3D must be installed on the build machine. The `.csproj` file auto-detects the installation path (e.g., `C:\Program Files\Autodesk\AutoCAD 2025\`) and resolves the AutoCAD and Civil 3D managed API assemblies (`AcCoreMgd.dll`, `AcDbMgd.dll`, `AcMgd.dll`, `AecBaseMgd.dll`, `AeccDbMgd.dll`) from that location. These references are marked `Private=False` so they are not copied to the output directory.

The following NuGet packages are restored automatically during build:

| Package | Version | Purpose |
|---------|---------|---------|
| itext7 | 8.0.5 | PDF document generation |
| itext7.bouncy-castle-adapter | 8.0.5 | Cryptography support required by iText |
| EPPlus | 7.5.2 | Excel workbook generation |
| NETStandard.Library | 2.0.3 | .NET Standard compatibility shim |

### Building with Visual Studio

Open the solution file for the desired version (or the root `xml-report-for-geotable.sln`), restore NuGet packages, and build. The post-build target in each `.csproj` automatically copies the output DLL and all dependency DLLs to the ApplicationPlugins bundle directory under `%AppData%`.

### Building with the .NET CLI

Navigate to the version-specific project folder and run:

```
dotnet restore
dotnet build --configuration Debug
```

For the 2023 and 2024 projects, ensure the .NET Framework 4.8 targeting pack is installed. For the 2025 project, the .NET 8.0 SDK is required.

The compiled output lands in `bin/Debug/` within each project folder. The post-build step copies everything to `%AppData%\Autodesk\ApplicationPlugins\GeoTableReports.bundle\Contents\Windows\{year}\`.

---

## Installation and Deployment

### Bundle Installation (Recommended)

This is the standard Autodesk ApplicationPlugins deployment method. Civil 3D scans the `ApplicationPlugins` folder at startup and automatically loads any valid `.bundle` directory it finds.

1. Build the project for the Civil 3D version you need (2023, 2024, or 2025), or use the pre-built bundles in the `bundles/` directory.

2. Copy the appropriate `.bundle` folder to the Autodesk ApplicationPlugins directory. The path is:

   ```
   C:\Users\<username>\AppData\Roaming\Autodesk\ApplicationPlugins\
   ```

   A quick way to navigate there is to press `Win+R`, type `%AppData%\Autodesk\ApplicationPlugins`, and press Enter. If the `ApplicationPlugins` folder does not exist, create it.

3. The copied folder must retain its `.bundle` extension and internal structure. For example, after copying the 2025 bundle, the directory tree should look like:

   ```
   %AppData%\Autodesk\ApplicationPlugins\
     GeoTableReports.2025.bundle\
       PackageContents.xml
       Contents\
         Windows\
           2025\
             GeoTableReports.2025.dll
             itext.kernel.dll
             itext.layout.dll
             itext.io.dll
             itext.commons.dll
             BouncyCastle.Cryptography.dll
             EPPlus.dll
             (additional dependency DLLs)
   ```

4. Start (or restart) Civil 3D. A confirmation message will appear in the command line indicating the add-in has loaded. Type `GEOTABLE` to begin generating reports.

Only install the bundle that matches your Civil 3D version. The 2025 bundle will not load in Civil 3D 2024 because the AutoCAD series numbers will not match, and vice versa.

### Manual NETLOAD

For one-time testing without permanent installation, open Civil 3D, type `NETLOAD` at the command line, and browse to the compiled `GeoTableReports.{year}.dll` in the project's `bin/Debug/` folder. The add-in will remain available only for the current session. You will need to repeat this step every time Civil 3D is restarted.

### Network Deployment

For multi-user environments, place the `.bundle` folder on a shared network path (e.g., `\\server\CADStandards\Civil3DPlugins\GeoTableReports.2025.bundle\`). On each workstation, open Civil 3D, type `OPTIONS`, navigate to the Files tab, expand Trusted Locations, and add the network path. Also add the network parent folder to the Support File Search Path. After restarting Civil 3D, the add-in will load from the network share. This allows a single update point that takes effect for all users on next launch.

---

## PackageContents.xml

Each bundle includes a `PackageContents.xml` file that tells Civil 3D how to discover and load the add-in. The key attributes are:

**AutodeskProduct** is set to `Civil3D` to restrict loading to Civil 3D sessions only.

**SeriesMin / SeriesMax** must match the AutoCAD internal version for the target release. Civil 3D 2023 uses `R24.2`, 2024 uses `R24.3`, and 2025 uses `R25.0`. If these values are wrong, Civil 3D will silently skip the bundle.

**ModuleName** is the relative path from the bundle root to the compiled DLL (e.g., `./Contents/Windows/2025/GeoTableReports.2025.dll`).

**LoadOnAutoCADStartup** is set to `True` so the add-in loads automatically when Civil 3D starts. The commands are also registered for on-demand loading via `LoadOnCommandInvocation`.

---

## Dialog Options

When the `GEOTABLE` command is invoked, a dialog window appears with the following controls:

**Report Type** selects between Horizontal and Vertical alignment processing. GeoTable PDF and GeoTable Excel outputs are only available for horizontal alignments.

**Base Output Path** sets the file location and base name for all generated files. The add-in appends suffixes like `_Alignment_Report.pdf`, `_GeoTable.pdf`, and `_GeoTable.xlsx` to this base path automatically.

**Alignment Reports** group contains checkboxes for Alignment PDF and Alignment XML generation.

**GeoTable Reports** group contains checkboxes for GeoTable PDF and GeoTable Excel generation.

**Remember my selections** persists the dialog state to a preferences file in `%AppData%\GeoTableReports\dialog_prefs.json` so the same options are pre-selected next time.

**Open folder after creation** and **Open files after creation** control whether the output directory or individual files are launched automatically when generation completes.

The dialog also adapts to the current Civil 3D color theme (light or dark) for visual consistency.

---

## Troubleshooting

If Civil 3D does not recognize the bundle after copying it to ApplicationPlugins, verify the following:

The `.bundle` folder name must end with `.bundle` (e.g., `GeoTableReports.2025.bundle`, not `GeoTableReports.2025`).

The `PackageContents.xml` must be at the root of the `.bundle` folder, not nested inside a subfolder.

The `SeriesMin` and `SeriesMax` values in `PackageContents.xml` must match your Civil 3D version exactly. An incorrect series number is the most common reason for silent load failure.

The `ModuleName` path in `PackageContents.xml` must point to the actual DLL location relative to the bundle root. Verify the DLL file exists at that path.

All dependency DLLs (iText, BouncyCastle, EPPlus, and their transitive dependencies) must be in the same folder as the main DLL. If any are missing, the add-in will fail to load or will throw errors at runtime.

If the add-in loads but `GEOTABLE` does not work, type `NETLOAD` and manually load the DLL to see if a more specific error message is displayed. Check the Civil 3D command line for .NET assembly binding errors.

---

## Dependencies

The add-in references the following Autodesk managed assemblies (resolved from the local Civil 3D installation, not distributed with the bundle):

AcCoreMgd.dll, AcDbMgd.dll, AcMgd.dll (AutoCAD core managed APIs), AecBaseMgd.dll (Architecture/Engineering base), AeccDbMgd.dll (Civil 3D database), and AeccPressurePipesMgd.dll.

Third-party libraries distributed with the bundle include iText 7 (kernel, layout, IO, commons, bouncy castle adapter) for PDF generation and EPPlus for Excel generation.

---

## License

Copyright 2025 A. Martinez. See individual project files for version-specific details.
