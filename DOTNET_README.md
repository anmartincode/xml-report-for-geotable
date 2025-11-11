# GeoTable Reports - .NET Add-in for Civil 3D

Native Civil 3D add-in for generating alignment reports

## Quick Start with VS Code

### 1. Install Prerequisites

```bash
# Install .NET SDK (if not already installed)
# Download: https://dotnet.microsoft.com/download

# Verify
dotnet --version
```

**VS Code Extensions** (install these):
- C# Dev Kit (ms-dotnettools.csdevkit)
- C# (ms-dotnettools.csharp)

### 2. Open Project in VS Code

```bash
cd /home/amartinez/.local/local/xml-report-for-geotable
code .
```

### 3. Build the Project

**Option A: Keyboard shortcut**
- Press `Ctrl+Shift+B` (Windows) or `Cmd+Shift+B` (Mac)

**Option B: VS Code Terminal**
```bash
dotnet build
```

**Option C: Command Palette**
- Press `Ctrl+Shift+P`
- Type: "Run Build Task"
- Select: "build"

### 4. Test in Civil 3D

1. Open Civil 3D
2. Type command: `NETLOAD`
3. Browse to: `bin/Debug/GeoTableReports.dll`
4. Click Load
5. Type command: `GEOTABLE`

### 5. Debug (Optional)

1. Set breakpoint in code (click left of line number)
2. Start Civil 3D
3. Load DLL with `NETLOAD`
4. In VS Code: Press `F5`
5. Select "Attach to Civil 3D"
6. Run command in Civil 3D

---

## Project Structure

```
xml-report-for-geotable/
├── .vscode/                  # VS Code configuration
│   ├── tasks.json           # Build tasks (Ctrl+Shift+B)
│   ├── launch.json          # Debug configuration (F5)
│   └── settings.json        # Editor settings
├── Commands/                 # Command classes (COMING SOON)
├── Core/                     # Business logic (COMING SOON)
├── UI/                       # Dialogs and forms (COMING SOON)
├── GeoTableReports.csproj   # Project file ✓
├── dotnet_starter_template.cs  # Starter code ✓
├── VSCODE_SETUP_GUIDE.md    # Full setup guide ✓
└── DOTNET_README.md         # This file ✓
```

---

## Commands Available

After building and loading in Civil 3D:

| Command | Description |
|---------|-------------|
| `GEOTABLE` | Generate alignment report |
| `GEOTABLE_BATCH` | Batch process multiple alignments |

*(Implement using dotnet_starter_template.cs)*

---

## Build Configurations

### Debug Build (default)
```bash
dotnet build
```
- Includes debug symbols
- No optimizations
- Located in: `bin/Debug/`

### Release Build
```bash
dotnet build -c Release
```
- Optimized code
- Smaller file size
- Located in: `bin/Release/`

---

## Development Workflow

### Daily Development Cycle

1. **Edit code** in VS Code
2. **Build** with `Ctrl+Shift+B`
3. **Test** in Civil 3D:
   - Type: `NETLOAD`
   - Browse to DLL
   - Run command
4. **Debug** if needed (F5 → Attach)
5. **Repeat**

### Before Deploying to Team

1. Build in Release mode:
   ```bash
   dotnet build -c Release
   ```

2. Test on clean machine

3. Create .bundle folder structure

4. Distribute .bundle to team

---

## Troubleshooting

### Error: "Cannot find Civil 3D DLLs"

**Check**: Civil 3D installation path in [GeoTableReports.csproj](GeoTableReports.csproj:13)

```xml
<Civil3DPath>C:\Program Files\Autodesk\AutoCAD 2025\</Civil3DPath>
```

### Error: "SDK not found"

**Install**: .NET Framework 4.8 Developer Pack
- Download: https://dotnet.microsoft.com/download/dotnet-framework/net48

### Build Fails: "Reference not found"

**Solution**:
1. Check Civil 3D is installed
2. Verify paths in .csproj
3. Clean and rebuild:
   ```bash
   dotnet clean
   dotnet build
   ```

### IntelliSense Not Working

**Solution**:
1. Reload VS Code window
2. Wait for OmniSharp to load (bottom status bar)
3. Check C# Dev Kit extension is installed

---

## Next Steps

### Phase 1: Basic Structure (Week 1)
- [ ] Create Command classes
- [ ] Implement alignment selection
- [ ] Basic data extraction
- [ ] Simple text report output

### Phase 2: Report Generation (Week 2)
- [ ] PDF generation with iTextSharp
- [ ] Format matching examples
- [ ] Error handling
- [ ] Progress indicators

### Phase 3: UI (Week 3)
- [ ] Selection dialog
- [ ] Settings dialog
- [ ] Batch processing form
- [ ] Ribbon integration

### Phase 4: Polish (Week 4)
- [ ] Testing
- [ ] Documentation
- [ ] Packaging
- [ ] Deployment

---

## Resources

- **Full Guide**: [VSCODE_SETUP_GUIDE.md](VSCODE_SETUP_GUIDE.md)
- **Architecture**: [DOTNET_ADDIN_GUIDE.md](DOTNET_ADDIN_GUIDE.md)
- **Starter Code**: [dotnet_starter_template.cs](dotnet_starter_template.cs)
- **Civil 3D API Docs**: https://help.autodesk.com/view/OARX/2025/ENU/

---

## Current Status

✅ **Project configured** for VS Code
✅ **Build system** ready (.csproj, tasks.json)
✅ **Starter template** created
⏳ **Implementation** pending

**Ready to code? Open in VS Code and press `Ctrl+Shift+B` to build!**
