# Civil 3D Rail Alignment Geotable XML Report Generator

A comprehensive Dynamo script solution for extracting geotable data from Civil 3D rail alignments and generating formatted XML reports for downstream use.

## Quick Start (NEW in v2.0)

The XML generator now accepts **raw geotable data** directly! No need for complex API calls.

**Simple Usage in Dynamo:**
1. Get your geotable data as a list (from Civil 3D geotable view)
2. Feed it to `xml_report_generator.py`
3. Get structured XML output

**Example:**
```python
# IN[0]: Your geotable data (automatically detected format)
# IN[1]: Output file path (optional)
# IN[2]: Pretty print (True/False)
```

See the [Updated Usage Section](#method-1-direct-geotable-input-new-in-v20) for details.

## Overview

This tool automates the extraction of rail alignment geotable data from Autodesk Civil 3D and converts it into structured XML reports suitable for:
- Data interchange with other railway design systems
- GIS applications
- Asset management databases
- Custom reporting and analysis tools
- Quality assurance and documentation

## Features

- **Comprehensive Data Extraction**: 
  - Alignment properties (name, length, start/end stations)
  - Station-by-station coordinates and geometry
  - Horizontal curve data (tangents, arcs, spirals)
  - Superelevation/cant information
  - Rail-specific geometric properties

- **Flexible Configuration**:
  - Customizable station sampling intervals
  - Filter alignments by name or properties
  - Configurable output precision
  - Multiple export formats

- **XML Output**:
  - Well-formatted, human-readable XML
  - Schema-validated structure
  - Namespace-compliant
  - Suitable for automated processing

## Project Structure

```
xml-report-for-geotable/
├── civil3d_geotable_extractor.py   # Main extraction script
├── xml_report_generator.py          # XML formatting module
├── GeoTableReport.dyn                # Dynamo graph file
├── config.json                       # Configuration settings
├── geotable_schema.xsd              # XML validation schema
├── sample_output.xml                # Example output
├── README.md                        # This file
└── output/                          # Generated reports (created automatically)
```

## Requirements

### Software
- Autodesk Civil 3D 2020 or later
- Dynamo for Civil 3D (typically bundled with Civil 3D)
- Python 3.x (CPython3 engine in Dynamo)

### Civil 3D Objects Required
- At least one alignment object in the active drawing
- Alignments should have valid geometry (tangents, curves, and/or spirals)
- Optional: Superelevation data for enhanced output

## Installation

1. **Clone or Download** this repository to a local directory:
   ```bash
   git clone <repository-url>
   cd xml-report-for-geotable
   ```

2. **Verify Civil 3D Installation**: Ensure Civil 3D is installed with Dynamo support.

3. **Update File Paths**: Edit `GeoTableReport.dyn` and update the script directory path:
   ```python
   sys.path.append(r'C:\Path\To\Script\Directory')  # UPDATE THIS PATH
   ```

4. **Configure Settings** (Optional): Edit `config.json` to customize extraction parameters.

## Usage

### Method 1: Direct Geotable Input (NEW in v2.0)

**This is the simplest method if you have geotable data from Civil 3D.**

1. **Open Civil 3D** with your alignment

2. **Get Geotable Data** (as a list from Civil 3D or Dynamo node)

3. **Add Python Script Node in Dynamo**:
   - Copy content from `xml_report_generator.py`
   - Configure inputs:
     - `IN[0]`: Your geotable data (list format)
     - `IN[1]`: Output file path (e.g., `C:\Reports\alignment.xml`) or leave empty for XML string
     - `IN[2]`: `True` for pretty-printed XML

4. **Run** and get your XML report!

**Input Format Example:**
```python
geotable_data = [
    ['Project Name:', 'Codman_Final'],
    ['Horizontal Alignment Name:', 'Prop_AshNB'],
    ['', 'STATION', 'ELEVATION'],
    ['Element: Linear', '', ''],
    ['POB', '641+44.67', '42.97'],
    ['PVI', '641+95.67', '41.74'],
    ['Tangent Grade:', '-2.411', ''],
    ['Element: Parabola', '', ''],
    ['PVC', '642+07.29', '41.46'],
    # ... etc
]
```

The script automatically:
- Detects the format
- Converts station values (e.g., "641+44.67" → 64144.67)
- Parses element types and properties
- Generates structured XML

### Method 2: Using Dynamo Graph (Original Method)

1. **Open Civil 3D** with a drawing containing rail alignments

2. **Launch Dynamo**: 
   - In Civil 3D ribbon: `Manage` tab → `Dynamo` button
   - Or type `DYNAMO` in the command line

3. **Open the Graph**:
   - File → Open
   - Navigate to `GeoTableReport.dyn`

4. **Configure the Python Script Nodes**:
   - **Node 1 (Extractor)**: Copy the entire content of `civil3d_geotable_extractor.py` into the Python node
   - **Node 2 (XML Generator)**: Copy the entire content of `xml_report_generator.py` into the Python node

5. **Set Input Parameters**:
   - **Alignment Name**: Leave empty for all alignments, or specify a specific name
   - **Station Interval**: Set sampling interval (e.g., 10.0 for 10m/ft intervals)
   - **Output Path**: Specify where to save the XML file (e.g., `C:\Output\report.xml`)

6. **Run the Script**:
   - Click `Run` button in Dynamo
   - View results in the Watch node
   - Check the output directory for the generated XML file

### Method 3: Using Python Scripts Directly in Dynamo (API Extraction)

1. **Create a New Dynamo Graph** or add to existing workflow

2. **Add Python Script Node** (First Node - Data Extraction):
   - Add node: `Core → Scripting → Python Script`
   - Copy content from `civil3d_geotable_extractor.py`
   - Configure inputs:
     - `IN[0]`: Alignment name (String or empty)
     - `IN[1]`: Station interval (Number)

3. **Add Python Script Node** (Second Node - XML Generation):
   - Add another Python Script node
   - Copy content from `xml_report_generator.py`
   - Configure inputs:
     - `IN[0]`: Connect to output of first node
     - `IN[1]`: Output file path (String)
     - `IN[2]`: Pretty print (Boolean, true/false)

4. **Add Input Nodes**:
   - Code Block nodes for static values
   - String/Number nodes for dynamic inputs
   - File Path node for output location

5. **Run and Review Results**

## Configuration

Edit `config.json` to customize behavior:

### Key Settings

```json
{
  "extraction_settings": {
    "station_interval": 10.0,           // Sampling interval
    "coordinate_precision": 6,          // Decimal places for coordinates
    "station_precision": 3              // Decimal places for stations
  },
  
  "rail_specific_settings": {
    "track_gauge": 1435.0,             // Standard gauge (mm)
    "extract_cant_data": true          // Include superelevation
  },
  
  "output_settings": {
    "default_output_directory": "./output",
    "filename_template": "{project_name}_{alignment_name}_geotable_{timestamp}.xml"
  }
}
```

## Output Format

The generated XML follows this structure:

```xml
<GeotableReport version="1.0">
  <ProjectInfo>
    <ProjectName>...</ProjectName>
    <GeneratedDate>...</GeneratedDate>
    <ReportType>Rail Alignment Geotable</ReportType>
  </ProjectInfo>
  
  <Alignments count="N">
    <Alignment name="..." id="...">
      <Properties>
        <Description>...</Description>
        <Length>...</Length>
        <StartStation>...</StartStation>
        <EndStation>...</EndStation>
      </Properties>
      
      <Stations count="N">
        <Station value="0.000">
          <Coordinates>
            <X>...</X>
            <Y>...</Y>
            <Z>...</Z>
          </Coordinates>
          <Direction>...</Direction>
          <Offset>...</Offset>
        </Station>
        <!-- More stations... -->
      </Stations>
      
      <GeometricEntities count="N">
        <Entity type="Line|Arc|Spiral" index="0">
          <!-- Entity-specific properties -->
        </Entity>
        <!-- More entities... -->
      </GeometricEntities>
      
      <Superelevation count="N">
        <!-- Superelevation data if available -->
      </Superelevation>
    </Alignment>
  </Alignments>
</GeotableReport>
```

See `sample_output.xml` for a complete example.

## Data Elements

### Station Data
- **Station Value**: Chainage/station along alignment
- **Coordinates**: X, Y, Z in drawing units
- **Direction**: Bearing/azimuth in radians
- **Offset**: Cross-sectional offset (typically 0 for centerline)

### Geometric Entities
- **Line**: Tangent sections with bearing
- **Arc**: Circular curves with radius, chord length, direction
- **Spiral**: Transition curves with in/out radii and type

### Superelevation (if available)
- **Critical Stations**: Key points in cant transitions
- **Transition Types**: Begin/End normal crown, full super, etc.
- **Slopes**: Left/right cross slopes as decimal grades

## Validation

The XML output can be validated against the provided schema:

```bash
xmllint --schema geotable_schema.xsd output/report.xml
```

Or use any XML validation tool that supports XSD 1.0.

## Troubleshooting

### Common Issues

**1. "AttributeError: 'list' object has no attribute 'get'"** ⚠️ FIXED IN v2.0
- **Solution**: Update to v2.0 - the script now handles list input automatically
- The script detects and converts list format to dictionary format
- Both input formats are now supported

**2. "Error retrieving alignments"**
- Ensure the drawing contains alignment objects
- Check that alignments are not corrupted
- Verify Civil 3D API references are loaded

**3. "Module not found" errors**
- Verify Python script paths are correct
- Ensure all required Civil 3D API DLLs are referenced
- Check Dynamo Python engine is set to CPython3

**4. Empty or incomplete XML**
- Check alignment has valid geometry
- Verify station interval is appropriate (not too large)
- Review Dynamo console for warnings

**5. Station data missing or not parsed**
- Ensure geotable data includes "STATION" and "ELEVATION" columns
- Station values should be in format "NNN+DD.DD" or numeric
- Check for "Element:" headers in the data
- Verify elevation values are numeric

**6. Element properties not appearing**
- Ensure element headers use format "Element: Type" (e.g., "Element: Parabola")
- Properties should come after element header
- Check property values are not empty

**7. Superelevation data not appearing**
- Verify superelevation is defined in Civil 3D
- Check alignment has associated superelevation view
- Ensure proper assembly/corridor setup (if applicable)

### Debug Mode

Enable detailed logging by adding to the Python scripts:

```python
import sys
sys.stdout = sys.__stdout__  # Enable console output
print("Debug information here")
```

## Advanced Usage

### Filtering Alignments

Extract specific alignments using name patterns:

```python
# In config.json
"filters": {
  "alignment_name_pattern": "Main.*",  // Regex pattern
  "min_alignment_length": 100.0        // Minimum length filter
}
```

### Custom Station Intervals

Use variable intervals for different sections:

```python
# Modify extractor to use adaptive intervals
# Dense sampling on curves, sparse on tangents
```

### Batch Processing

Process multiple drawings:

```python
# Create a batch script that:
# 1. Opens each drawing
# 2. Runs the Dynamo script
# 3. Saves XML output
# 4. Closes drawing
```

### Integration with GIS

Convert coordinate system for GIS import:

```python
# Add coordinate transformation in extractor
from Autodesk.Aec.DatabaseServices import *
# Transform to WGS84 or local projection
```

## API Reference

### GeoTableDataExtractor Class

#### Methods

**`__init__(document=None)`**
- Initialize extractor with Civil 3D document
- Parameters: Optional document object (uses active if None)

**`get_all_alignments()`**
- Returns: List of alignment dictionaries with basic properties

**`extract_alignment_stations(alignment, interval=None)`**
- Extract station data along alignment
- Parameters: 
  - `alignment`: Alignment object
  - `interval`: Station sampling interval (default: 10.0)
- Returns: List of station dictionaries

**`extract_alignment_curves(alignment)`**
- Extract geometric entity data
- Returns: List of entity dictionaries

**`extract_superelevation_data(alignment)`**
- Extract superelevation/cant data
- Returns: List of superelevation dictionaries

**`extract_geotable_data(alignment_name=None, station_interval=10.0)`**
- Main extraction method
- Returns: Complete geotable data dictionary

### XMLReportGenerator Class

#### Methods

**`__init__(geotable_data)`**
- Initialize with extracted data dictionary

**`generate_xml()`**
- Generate XML element tree
- Returns: XML root element

**`to_xml_string(pretty_print=True)`**
- Convert to formatted XML string
- Returns: XML string

**`save_to_file(file_path, pretty_print=True)`**
- Save XML to file
- Returns: File path

## Performance Considerations

- **Large Alignments**: For alignments > 10km, use larger station intervals (20-50m)
- **Multiple Alignments**: Process individually if memory constrained
- **Complex Geometry**: Spirals and compound curves may slow extraction
- **File Size**: XML files grow linearly with station count

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

### Development Guidelines

- Follow PEP 8 for Python code
- Add docstrings to all functions
- Update schema if adding new XML elements
- Test with Civil 3D 2020-2024

## License

This project is provided as-is for use with Autodesk Civil 3D. Please review your Civil 3D license agreement for API usage terms.

## Support

For issues or questions:

1. Check the Troubleshooting section
2. Review sample files and configuration
3. Consult Civil 3D API documentation
4. Open an issue on the repository

## Version History

### Version 2.0.0 (2025-11-10)
- **MAJOR UPDATE**: Added support for raw geotable list input
- Automatic format detection (dictionary vs list)
- Station format conversion (e.g., "641+44.67" → 64144.67)
- Enhanced geometric element parsing (Linear, Parabola, Spiral)
- Detailed property extraction (grades, sight distances, K values)
- Station point type tracking (POB, PVI, PVC, PVT)
- Improved error handling and debug output
- Backward compatible with v1.0 dictionary format

### Version 1.0.0 (2025-10-29)
- Initial release
- Support for alignment extraction via Civil 3D API
- Station, curve, and superelevation data
- XML output with schema validation
- Configurable parameters
- Dynamo graph template

## Acknowledgments

- Autodesk Civil 3D API documentation
- Dynamo community
- Railway engineering best practices

## Related Resources

- [Autodesk Civil 3D Developer's Guide](https://help.autodesk.com/view/CIV3D/2024/ENU/)
- [Dynamo Primer](https://primer.dynamobim.org/)
- [Railway Geometric Design Standards](https://railsystem.net/)
- [XML Schema Documentation](https://www.w3.org/XML/Schema)

## Future Enhancements

Planned features for future releases:

- [ ] Profile data extraction (vertical alignment)
- [ ] Cross-section data
- [ ] Cant deficiency calculations
- [ ] Track quality index (TQI) analysis
- [ ] Export to other formats (CSV, JSON, GeoJSON)
- [ ] Integration with BIM 360
- [ ] Automated quality checks
- [ ] Visualization dashboard
- [ ] Multi-language support

---

**Note**: This tool is designed for rail alignment applications but can be adapted for highway, pipeline, or other linear infrastructure projects by modifying the extraction logic and XML schema.

For the latest updates and documentation, visit the project repository.









