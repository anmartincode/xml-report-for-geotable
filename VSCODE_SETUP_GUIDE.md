# Build Civil 3D .NET Add-in with VS Code

**Goal**: Develop Civil 3D .NET add-in using VS Code instead of Visual Studio

## Why VS Code?

✅ **Free and lightweight**
✅ **Cross-platform** (develop on any OS)
✅ **Fast and modern** UI
✅ **Great for .NET development** with C# Dev Kit
✅ **Integrated terminal** for builds
✅ **Git integration** built-in

---

## Prerequisites

### 1. Install .NET SDK

**Download**: https://dotnet.microsoft.com/download

```bash
# Verify installation
dotnet --version
# Should show: 6.0 or higher
```

### 2. Install VS Code

**Download**: https://code.visualstudio.com/

### 3. Install VS Code Extensions

Open VS Code and install these extensions:

1. **C# Dev Kit** (Microsoft) - Essential for C# development
2. **C#** (Microsoft) - Language support
3. **.NET Install Tool** (Microsoft) - Manages .NET SDKs

**Quick install via command line:**
```bash
code --install-extension ms-dotnettools.csdevkit
code --install-extension ms-dotnettools.csharp
code --install-extension ms-dotnettools.vscode-dotnet-runtime
```

### 4. Civil 3D Installation

You need Civil 3D installed for the API DLLs:
- `AcCoreMgd.dll`
- `AcDbMgd.dll`
- `AeccDbMgd.dll`

Located in: `C:\Program Files\Autodesk\AutoCAD 2025\` (or your version)

---

## Project Setup

### Step 1: Create Project Structure

```bash
# Navigate to your project folder
cd /home/amartinez/.local/local/xml-report-for-geotable

# Create .NET class library project
dotnet new classlib -n GeoTableReports -f net48

# Navigate to project
cd GeoTableReports
```

### Step 2: Update Project File

Edit `GeoTableReports.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <PlatformTarget>x64</PlatformTarget>
    <LangVersion>latest</LangVersion>
    <OutputType>Library</OutputType>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>
  </PropertyGroup>

  <PropertyGroup>
    <Civil3DPath Condition="Exists('C:\Program Files\Autodesk\AutoCAD 2025\')">C:\Program Files\Autodesk\AutoCAD 2025\</Civil3DPath>
    <Civil3DPath Condition="Exists('C:\Program Files\Autodesk\AutoCAD 2024\')">C:\Program Files\Autodesk\AutoCAD 2024\</Civil3DPath>
  </PropertyGroup>

  <ItemGroup>
    <!-- Civil 3D API References -->
    <Reference Include="AcCoreMgd">
      <HintPath>$(Civil3DPath)AcCoreMgd.dll</HintPath>
      <Private>False</Private>
    </Reference>
    <Reference Include="AcDbMgd">
      <HintPath>$(Civil3DPath)AcDbMgd.dll</HintPath>
      <Private>False</Private>
    </Reference>
    <Reference Include="AcMgd">
      <HintPath>$(Civil3DPath)AcMgd.dll</HintPath>
      <Private>False</Private>
    </Reference>
    <Reference Include="AeccDbMgd">
      <HintPath>$(Civil3DPath)C3D\AeccDbMgd.dll</HintPath>
      <Private>False</Private>
    </Reference>
    <Reference Include="AeccPressurePipesMgd">
      <HintPath>$(Civil3DPath)C3D\AeccPressurePipesMgd.dll</HintPath>
      <Private>False</Private>
    </Reference>
  </ItemGroup>

  <ItemGroup>
    <!-- System References -->
    <Reference Include="System.Windows.Forms" />
    <Reference Include="System.Drawing" />
  </ItemGroup>

  <ItemGroup>
    <!-- NuGet Packages -->
    <PackageReference Include="iTextSharp.LGPLv2.Core" Version="3.4.4" />
  </ItemGroup>

  <Target Name="PostBuild" AfterTargets="PostBuildEvent">
    <!-- Copy DLL to Civil 3D plugin folder for testing -->
    <Exec Command="xcopy /Y /D &quot;$(TargetPath)&quot; &quot;$(AppData)\Autodesk\ApplicationPlugins\GeoTableReports.bundle\Contents\Windows\$(Civil3DVersion)\&quot;" />
  </Target>

</Project>
```

---

## Building the Project

### Method 1: VS Code Tasks (Recommended)

Create `.vscode/tasks.json`:

```json
{
    "version": "2.0.0",
    "tasks": [
        {
            "label": "build",
            "command": "dotnet",
            "type": "process",
            "args": [
                "build",
                "${workspaceFolder}/GeoTableReports/GeoTableReports.csproj",
                "/property:GenerateFullPaths=true",
                "/consoleloggerparameters:NoSummary"
            ],
            "problemMatcher": "$msCompile",
            "group": {
                "kind": "build",
                "isDefault": true
            }
        },
        {
            "label": "clean",
            "command": "dotnet",
            "type": "process",
            "args": [
                "clean",
                "${workspaceFolder}/GeoTableReports/GeoTableReports.csproj"
            ],
            "problemMatcher": "$msCompile"
        },
        {
            "label": "rebuild",
            "dependsOn": ["clean", "build"],
            "group": "build"
        }
    ]
}
```

**Build with**: `Ctrl+Shift+B` (Windows) or `Cmd+Shift+B` (Mac)

### Method 2: Command Line

```bash
# Build project
dotnet build

# Build in Release mode
dotnet build -c Release

# Clean build
dotnet clean
```

### Method 3: VS Code Terminal

```bash
# Open integrated terminal: Ctrl+`
cd GeoTableReports
dotnet build
```

---

## Debugging with Civil 3D

### Option 1: Attach to Process (Simple)

1. Start Civil 3D manually
2. In VS Code: `Run > Start Debugging` (F5)
3. Select "Attach to Process"
4. Find and attach to `acad.exe` or `civil3d.exe`
5. Load your DLL with `NETLOAD` command

### Option 2: Launch Configuration (Advanced)

Create `.vscode/launch.json`:

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Debug with Civil 3D",
            "type": "coreclr",
            "request": "launch",
            "preLaunchTask": "build",
            "program": "C:/Program Files/Autodesk/AutoCAD 2025/acad.exe",
            "args": [],
            "cwd": "${workspaceFolder}",
            "stopAtEntry": false,
            "console": "internalConsole"
        },
        {
            "name": "Attach to Civil 3D",
            "type": "coreclr",
            "request": "attach",
            "processName": "acad.exe"
        }
    ]
}
```

**Usage:**
1. Press F5
2. Civil 3D will launch
3. In Civil 3D: `NETLOAD` → Browse to your DLL
4. Set breakpoints in VS Code
5. Run your command (e.g., `GEOTABLE`)

---

## Project Structure in VS Code

```
GeoTableReports/
├── .vscode/
│   ├── tasks.json          # Build tasks
│   ├── launch.json         # Debug configuration
│   └── settings.json       # VS Code settings
├── Commands/
│   └── ReportCommands.cs   # Main commands
├── Core/
│   └── AlignmentExtractor.cs
├── UI/
│   └── ReportDialog.cs
├── Properties/
│   └── AssemblyInfo.cs
├── GeoTableReports.csproj  # Project file
└── README.md
```

---

## VS Code Workspace Settings

Create `.vscode/settings.json`:

```json
{
    "files.exclude": {
        "**/bin": true,
        "**/obj": true
    },
    "omnisharp.enableRoslynAnalyzers": true,
    "omnisharp.enableEditorConfigSupport": true,
    "omnisharp.enableImportCompletion": true,
    "editor.formatOnSave": true,
    "csharp.format.enable": true
}
```

---

## Quick Start Workflow

### 1. Create New Project

```bash
# In your project folder
dotnet new classlib -n GeoTableReports -f net48
cd GeoTableReports
code .
```

### 2. Add Code

Copy [dotnet_starter_template.cs](dotnet_starter_template.cs:1) to `Commands/ReportCommands.cs`

### 3. Build

Press `Ctrl+Shift+B` or run:
```bash
dotnet build
```

### 4. Test in Civil 3D

```
1. Open Civil 3D
2. Type: NETLOAD
3. Browse to: bin\Debug\net48\GeoTableReports.dll
4. Type: GEOTABLE
```

### 5. Debug

```
1. Press F5 in VS Code
2. Attach to acad.exe process
3. Set breakpoints
4. Run command in Civil 3D
```

---

## Advantages of VS Code vs Visual Studio

| Feature | VS Code | Visual Studio |
|---------|---------|---------------|
| **Size** | ~200 MB | ~10 GB |
| **Speed** | Very fast | Slower |
| **Cost** | Free | Free (Community) or $$$ |
| **Build Time** | Same | Same |
| **IntelliSense** | Good | Excellent |
| **Debugging** | Good | Excellent |
| **UI Designer** | No | Yes |
| **Git Integration** | Excellent | Good |
| **Extensions** | Many | Many |

**Verdict**: VS Code is perfect for this project!

---

## Troubleshooting

### Issue: "Cannot find Civil 3D DLLs"

**Solution**: Update `Civil3DPath` in `.csproj` to match your installation:
```xml
<Civil3DPath>C:\Program Files\Autodesk\AutoCAD 2025\</Civil3DPath>
```

### Issue: "Build fails with 'SDK not found'"

**Solution**: Ensure .NET Framework 4.8 Developer Pack is installed:
```
Download: https://dotnet.microsoft.com/download/dotnet-framework/net48
```

### Issue: "IntelliSense not working"

**Solution**:
1. Install C# Dev Kit extension
2. Reload VS Code
3. Open `.csproj` file
4. Wait for OmniSharp to load (see bottom status bar)

### Issue: "Debugging doesn't hit breakpoints"

**Solution**:
1. Build in Debug mode (not Release)
2. Ensure PDB files are generated
3. Attach debugger AFTER loading DLL with NETLOAD

---

## Building Release Version

```bash
# Build optimized release
dotnet build -c Release

# Output in: bin/Release/net48/GeoTableReports.dll
```

---

## Packaging for Distribution

### Create .bundle Structure

```bash
# Create bundle folder
mkdir -p GeoTableReports.bundle/Contents/Windows/2025

# Copy DLL
cp bin/Release/net48/GeoTableReports.dll GeoTableReports.bundle/Contents/Windows/2025/

# Create PackageContents.xml
# (See DOTNET_ADDIN_GUIDE.md for XML content)

# Zip for distribution
zip -r GeoTableReports.bundle.zip GeoTableReports.bundle/
```

---

## Development Workflow

### Daily Development

1. **Code** in VS Code
2. **Build** with `Ctrl+Shift+B`
3. **Test** in Civil 3D with `NETLOAD`
4. **Debug** by attaching to process (F5)
5. **Commit** changes with Source Control tab

### Before Release

1. Build in Release mode
2. Test on clean machine
3. Create .bundle package
4. Test installation
5. Create installer (optional)

---

## Next Steps

### Immediate (Today)

1. ✅ Install VS Code + C# Dev Kit
2. ✅ Install .NET SDK
3. ✅ Create project with `dotnet new`
4. ✅ Copy starter template code
5. ✅ Build and test

### This Week

1. Add basic command functionality
2. Test with real Civil 3D drawing
3. Add error handling
4. Create simple UI dialog

### Next Month

1. Implement full report generation
2. Add PDF support
3. Create ribbon UI
4. Package for distribution

---

## Resources

- **VS Code C# Guide**: https://code.visualstudio.com/docs/languages/csharp
- **.NET CLI**: https://docs.microsoft.com/en-us/dotnet/core/tools/
- **Civil 3D API**: https://help.autodesk.com/view/OARX/2025/ENU/
- **Debugging in VS Code**: https://code.visualstudio.com/docs/csharp/debugging

---

## Comparison: VS Code vs Visual Studio

**Use VS Code if:**
- ✅ You want lightweight and fast
- ✅ You prefer command-line builds
- ✅ You're comfortable with JSON configs
- ✅ You work on multiple OSes
- ✅ You want modern editor experience

**Use Visual Studio if:**
- ✅ You need visual form designer
- ✅ You want more powerful debugging
- ✅ You prefer GUI for everything
- ✅ You're building complex WinForms

**For this project**: VS Code is perfect! The add-in is relatively simple and doesn't need heavy Visual Studio features.

---

**Ready to start? Let's create the project!**
