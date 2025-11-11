# Quick Start Guide
## Civil 3D Geotable XML Report Generator

Get up and running in 5 minutes!

## Prerequisites Check

- [ ] Civil 3D 2020+ installed
- [ ] Dynamo for Civil 3D available
- [ ] Drawing with at least one alignment open

## Steps

### 1. Open Your Civil 3D Drawing

Open a drawing that contains rail alignments in Civil 3D.

### 2. Launch Dynamo

In Civil 3D:
- Click the **Manage** tab
- Click **Dynamo** button
- Or type `DYNAMO` in command line

### 3. Set Up the Dynamo Graph

**Option A: Use the Template (Quick)**

1. In Dynamo: `File` → `Open`
2. Navigate to `GeoTableReport.dyn`
3. **IMPORTANT**: Edit both Python Script nodes:
   - Click on first Python node
   - Delete placeholder text
   - Copy **entire content** of `civil3d_geotable_extractor.py`
   - Paste into the Python node
   - Repeat for second Python node with `xml_report_generator.py`

**Option B: Create From Scratch (Custom)**

1. Create new Dynamo graph
2. Add two `Python Script` nodes (from Core → Scripting)
3. Copy scripts as described in Option A above

### 4. Configure Inputs

Edit the Code Block nodes or add input nodes:

```python
// Node 3: Extraction Configuration
alignment_name = "";        // Leave empty for ALL alignments
                           // Or specify: "Main_Rail_Line"
station_interval = 10.0;   // Interval in drawing units (10m or 10ft)
```

```python
// Node 4: XML Output Configuration  
output_path = "C:\\Output\\geotable_report.xml";  // Where to save XML
pretty_print = true;                              // Formatted output
```

**Quick Path Setup:**
- Create an `output` folder in your project directory
- Use path like: `C:\Projects\MyRail\output\report.xml`

### 5. Run the Script

1. Click the **Run** button (▶) in Dynamo
2. Wait for execution (typically 5-30 seconds)
3. Check the **Watch** node for results

### 6. Review Output

**Success Messages:**
- Watch node shows: `"XML report saved to: [path]"`
- Check the output file exists at specified location

**If You See Errors:**
- "Error retrieving alignments" → No alignments in drawing
- "Error generating XML" → Check script was pasted correctly
- See main README.md Troubleshooting section

### 7. Open the XML Report

Navigate to your output path and open the XML file in:
- Web browser (for formatted view)
- Text editor (Notepad++, VS Code)
- XML viewer application
- GIS/CAD software that accepts XML

## Example Output Location

```
C:\Projects\RailProject\
├── MyDrawing.dwg
└── output\
    └── geotable_report.xml  ← Your generated report
```

## What's in the XML?

Your report contains:

✅ **Project Information**
- Project name, timestamp, report type

✅ **Alignment Properties**
- Name, description, length
- Start and end stations

✅ **Station Data** (every N meters/feet)
- X, Y, Z coordinates
- Direction/bearing
- Offset from centerline

✅ **Geometric Elements**
- Lines (tangent sections)
- Arcs (curves with radius)
- Spirals (transition curves)

✅ **Superelevation** (if defined)
- Critical stations
- Left/right slopes
- Transition types

## Quick Customization

### Change Station Interval

For **dense sampling** (more detail):
```python
station_interval = 5.0;   // Every 5 units
```

For **sparse sampling** (smaller files):
```python
station_interval = 25.0;  // Every 25 units
```

### Extract Single Alignment

```python
alignment_name = "Main_Rail_Line_01";  // Specific alignment only
```

### Extract All Alignments

```python
alignment_name = "";  // Empty = all alignments
```

## Next Steps

### Validate Your XML

Use online validator or command line:
```bash
xmllint --schema geotable_schema.xsd output/geotable_report.xml
```

### Import to GIS

1. Open QGIS/ArcGIS
2. Import XML or convert to shapefile
3. Use coordinates for spatial analysis

### Batch Process

Save your Dynamo graph and run it on multiple drawings:
1. Open drawing #1
2. Run Dynamo script
3. Change output path
4. Open drawing #2
5. Repeat

### Automate Further

Create a Dynamo Player script:
1. Save graph with inputs exposed
2. Use Dynamo Player to run without opening Dynamo
3. Perfect for repetitive tasks

## Common Workflows

### Workflow 1: Single Alignment Report
```
1. Set alignment_name = "YourAlignment"
2. Set station_interval = 10.0
3. Set output_path with alignment name
4. Run → Get single alignment XML
```

### Workflow 2: Full Project Export
```
1. Set alignment_name = "" (empty)
2. Set station_interval = 20.0 (broader view)
3. Set output_path = "Project_AllAlignments.xml"
4. Run → Get all alignments in one XML
```

### Workflow 3: Detailed Curve Analysis
```
1. Set station_interval = 2.0 (very detailed)
2. Extract alignment with complex curves
3. Analyze geometric entities in XML
4. Review spiral parameters, curve radii
```

## Tips for Success

💡 **Tip 1**: Start with default settings (station_interval = 10.0)

💡 **Tip 2**: Test on a simple alignment first

💡 **Tip 3**: Keep output paths simple (no special characters)

💡 **Tip 4**: Use the Watch node to debug issues

💡 **Tip 5**: Save your Dynamo graph for reuse

💡 **Tip 6**: Back up your XML reports

## Help & Support

**Still having issues?**

1. ✓ Check that Civil 3D drawing has alignments
2. ✓ Verify Python scripts are fully pasted
3. ✓ Ensure output directory exists
4. ✓ Review Dynamo console for errors
5. ✓ See full README.md for detailed troubleshooting

**Where to get help:**
- Main README.md (comprehensive guide)
- sample_output.xml (see expected format)
- Civil 3D API documentation
- Dynamo forums

## You're Ready! 🚀

You should now have:
- ✅ XML report of your alignment(s)
- ✅ Structured geotable data
- ✅ Reusable Dynamo script

**Next**: Explore the XML, integrate with your tools, or customize the extraction!

---

**Questions?** Check the main README.md file for detailed documentation.

