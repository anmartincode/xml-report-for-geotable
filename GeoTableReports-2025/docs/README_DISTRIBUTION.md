# GeoTable Reports for Civil 3D

**Automated Alignment Report Generation**

Generate professional PDF, TXT, and XML reports for horizontal and vertical alignments with a single click.

---

## 📋 Quick Start

### Installation (Automatic)

1. **Extract** the downloaded ZIP file
2. **Double-click** `Install.bat`
3. **Restart** Civil 3D
4. **Done!** Type `GEOTABLE_PANEL` in Civil 3D to start

### Installation (Manual)

See [INSTALLATION.md](INSTALLATION.md) for detailed instructions.

---

## ✨ Features

### Report Types
- **Vertical Alignment Reports**
  - Station and elevation data
  - Tangent grades and lengths
  - PVI, PVC, PVT points
  - Parabolic curve details (K-values, rate of change)
  - Sight distance calculations

- **Horizontal Alignment Reports**
  - Point coordinates (Northing/Easting)
  - Tangent bearings and lengths
  - Circular curves (radius, delta, chord, etc.)
  - Spiral transitions
  - Design speed and superelevation

### Output Formats
- **PDF** - Professional, printable reports
- **TXT** - Plain text for easy editing
- **XML** - Structured data for integration

### Workflow Options
- **Dockable Panel** - Configure and generate reports from persistent UI
- **Quick Dialog** - Fast generation with dialog interface
- **Batch Processing** - Process all alignments in drawing at once

---

## 🚀 Usage

### Commands

| Command | Description |
|---------|-------------|
| `GEOTABLE_PANEL` | Open the dockable panel interface |
| `GEOTABLE` | Quick report generation via dialog |
| `GEOTABLE_BATCH` | Batch process all alignments |

### Basic Workflow

1. **Open** a drawing with alignments in Civil 3D
2. **Type** `GEOTABLE_PANEL` to open the panel
3. **Select** report type (Vertical or Horizontal)
4. **Choose** output formats (PDF, TXT, XML - or all three!)
5. **Browse** to select output location
6. **Click** "Generate Report"
7. **Select** an alignment in the drawing
8. **Done!** Reports are generated

---

## 📊 Sample Reports

### Vertical Alignment Report
```
VERTICAL ALIGNMENT REPORT
Alignment: Main Road
Generated: 11/12/2025 10:30:00 AM

VERTICAL POINTS
┌────────────┬──────────┬────────┬──────────┬──────────┐
│ Type       │ Station  │ Elev   │ Grade In │ Grade Out│
├────────────┼──────────┼────────┼──────────┼──────────┤
│ POB        │ 0+00.00  │ 100.00 │          │          │
│ PVI        │ 5+00.00  │ 105.00 │  1.00%   │ -0.50%   │
│ PVC        │ 4+00.00  │ 104.00 │          │          │
│ PVT        │ 6+00.00  │ 104.75 │          │          │
└────────────┴──────────┴────────┴──────────┴──────────┘
```

### Horizontal Alignment Report
Note: Legacy Civil 3D sample label `POT` has been updated to InRoads-convention `POB` (Point of Beginning). Spiral transition labels (TS/SC/CS/ST) remain only when true spirals exist.
```
HORIZONTAL ALIGNMENT REPORT
Alignment: Main Road
Generated: 11/12/2025 10:30:00 AM

HORIZONTAL POINTS
┌────────────┬──────────┬────────────┬────────────┐
│ Type       │ Station  │ Northing   │ Easting    │
├────────────┼──────────┼────────────┼────────────┤
│ POB        │ 0+00.00  │ 1000.0000  │ 2000.0000  │
│ PC         │ 2+50.00  │ 1050.0000  │ 2200.0000  │
│ PT         │ 5+00.00  │ 1100.0000  │ 2450.0000  │
└────────────┴──────────┴────────────┴────────────┘
```

---

## 🔧 System Requirements

- **Civil 3D 2025** (or newer)
- **.NET 8.0** Runtime (included with Civil 3D 2025)
- **Windows 10/11** (64-bit)
- ~30 MB disk space for installation

---

## 📖 Documentation

- [INSTALLATION.md](INSTALLATION.md) - Detailed installation instructions
- Inline tooltips in the panel interface
- Command line help: Type command name followed by `?`

---

## 🐛 Troubleshooting

### Add-in doesn't load
- Verify Civil 3D version (2025 or newer)
- Check installation path: `%AppData%\Autodesk\ApplicationPlugins\`
- Restart Civil 3D

### "Could not load file or assembly" error
- Ensure all DLL files are present in the bundle
- Don't delete any dependency DLLs
- Reinstall using `Install.bat`

### Reports not generating
- Ensure alignment is selected when prompted
- Check output path is writable
- Verify at least one output format is selected

### See [INSTALLATION.md](INSTALLATION.md) for more troubleshooting

---

## 📝 License

This software is provided as-is. See LICENSE file for details.

---

## 🤝 Support

For issues, questions, or feature requests, contact your CAD administrator or software provider.

---

## 🔄 Version History

### Version 1.0.0 (Current)
- Initial release
- Vertical and horizontal alignment reports
- Multi-format output (PDF, TXT, XML)
- Dockable panel interface
- Batch processing support
- Enhanced error handling and validation

---

**© 2025 GeoTable Reports. All rights reserved.**
