# GeoTable Reports - Quick Start Guide

## For Developers: Building the Distribution

### One Command Build
```powershell
.\Build-MultiVersionBundle.ps1
```

**Output:** `dist\GeoTableReports-MultiVersion-v1.0.0.zip`

**What it does:**
1. ✅ Builds GeoTableReports-2023 (.NET Framework 4.8)
2. ✅ Builds GeoTableReports-2024 (.NET Framework 4.8)
3. ✅ Builds GeoTableReports-2025 (.NET 8.0)
4. ✅ Creates unified bundle structure
5. ✅ Generates multi-version PackageContents.xml
6. ✅ Packages everything into distributable ZIP

---

## For IT: Deploying to Department

### Option 1: Network Share (Recommended)
```powershell
# Extract bundle to network location
Expand-Archive GeoTableReports-MultiVersion-v1.0.0.zip -DestinationPath "\\Server\CAD-Standards\Plugins\"

# Users add to Civil3D support paths:
\\Server\CAD-Standards\Plugins\
```

**Advantages:** Single update location, automatic for all users

### Option 2: Individual Installation
```powershell
# Distribute ZIP to users
# Users run:
.\Install-GeoTableReports.ps1
```

**Advantages:** No network dependency, works offline

---

## For Users: Installing

### Easy Install
1. Extract ZIP file
2. Run `Install-GeoTableReports.ps1`
3. Restart Civil3D
4. Use commands: `GEOTABLEEXCEL` or `GEOTABLEPDF`

### Manual Install
1. Extract `GeoTableReports.bundle` folder
2. Copy to: `%AppData%\Autodesk\ApplicationPlugins\`
3. Restart Civil3D

---

## Verification

In Civil3D, type:
```
GEOTABLEEXCEL
```

If command is recognized, installation successful! ✅

---

## How It Works

The bundle contains **all three versions**:
- `/2023/GeoTableReports.2023.dll`
- `/2024/GeoTableReports.2024.dll`
- `/2025/GeoTableReports.2025.dll`

Civil3D **automatically loads the correct one** based on its version.

**Users don't need to choose** - it's automatic!

---

## Need Help?

📖 See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md) for detailed instructions

🐛 **Troubleshooting:**
- Plugin not loading? Check `%AppData%\Autodesk\ApplicationPlugins\GeoTableReports.bundle`
- Wrong version? Check Civil3D version with `ABOUT` command
- Network issues? Verify permissions on network path

---

## Project Structure

```
xml-report-for-geotable/
├── GeoTableReports-2023/     ← Build for Civil3D 2023
├── GeoTableReports-2024/     ← Build for Civil3D 2024
├── GeoTableReports-2025/     ← Build for Civil3D 2025
├── GeoTableReports-MultiTarget/ ← Alternative approach
├── Build-MultiVersionBundle.ps1  ← Main build script
├── Install-GeoTableReports.ps1   ← User installer
└── DEPLOYMENT_GUIDE.md       ← Full documentation
```

---

## Common Commands

### Build Everything
```powershell
.\Build-MultiVersionBundle.ps1
```

### Install Locally
```powershell
.\Install-GeoTableReports.ps1
```

### Uninstall
```powershell
.\Install-GeoTableReports.ps1 -Uninstall
```

### Build Single Version
```powershell
cd GeoTableReports-2025
dotnet build -c Release
```

---

## Tips for Department Deployment

**✅ DO:**
- Use network deployment for teams >5 users
- Test on one machine before deploying to all
- Keep old version backed up
- Notify users before updates

**❌ DON'T:**
- Mix network and local installations
- Deploy without testing
- Update during critical project deadlines
- Forget to verify permissions on network shares

---

**Quick Reference Card:**

| Task | Command |
|------|---------|
| Build all versions | `.\Build-MultiVersionBundle.ps1` |
| Install for user | `.\Install-GeoTableReports.ps1` |
| Uninstall | `.\Install-GeoTableReports.ps1 -Uninstall` |
| Check installation | `%AppData%\Autodesk\ApplicationPlugins\` |
| Use in Civil3D | `GEOTABLEEXCEL` or `GEOTABLEPDF` |

---

Need more details? See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md)
