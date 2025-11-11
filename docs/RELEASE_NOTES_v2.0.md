# Release Notes - Version 2.0.0

**Release Date**: November 10, 2025

## What's New

### Major Features

#### 1. Raw Geotable List Input Support
The XML generator now accepts **raw geotable data** directly as a list of lists, eliminating the need for complex API extraction in many cases.

**Before (v1.0)**: Required Civil 3D API extraction script
```python
# Required complex API calls to extract data
extractor = GeoTableDataExtractor()
data = extractor.extract_geotable_data()
```

**Now (v2.0)**: Direct list input works automatically
```python
# Just pass your geotable data as-is
geotable_data = [
    ['Project Name:', 'Codman_Final'],
    ['Horizontal Alignment Name:', 'Prop_AshNB'],
    ['Element: Linear', '', ''],
    ['POB', '641+44.67', '42.97'],
    # ...
]
# Script auto-detects and processes it
```

#### 2. Automatic Format Detection
The script intelligently detects whether input is:
- Dictionary format (from API extractor)
- List format (from raw geotable)
- And converts accordingly

#### 3. Civil 3D Station Format Conversion
Stations in Civil 3D format are automatically converted:
- `"641+44.67"` → `64144.67`
- `"100.50"` → `100.50`
- Handles both formats seamlessly

#### 4. Enhanced Geometric Element Parsing
Now extracts detailed element data including:
- **Linear Elements**: POB, PVI points with tangent grades and lengths
- **Parabola Elements**: PVC, PVI, PVT points with:
  - Entrance/Exit grades
  - Sight distances
  - K values
  - Middle ordinate
- **Spiral Elements**: Transition curve data

#### 5. Station Point Type Tracking
All station points are now categorized by type:
- POB (Point of Beginning)
- PVI (Point of Vertical Intersection)
- PVC (Point of Vertical Curvature)
- PVT (Point of Vertical Tangency)
- Custom point types from geotable

## Bug Fixes

### Fixed: "AttributeError: 'list' object has no attribute 'get'"
**Issue**: Script crashed when receiving list input instead of dictionary
**Root Cause**: Original code expected only dictionary format from API extractor
**Fix**: Added format detection and conversion layer that handles both inputs

### Fixed: Station Values Not Being Parsed
**Issue**: Station values like "641+44.67" were being stored as strings
**Fix**: Implemented `parse_station_value()` function to convert Civil 3D format

### Fixed: Empty Elements in XML Output
**Issue**: Element sections were empty even when data existed
**Fix**: Enhanced parsing logic to properly capture element properties

## Enhancements

### Improved Error Handling
- Better debug output showing data types at each processing stage
- Clearer error messages
- Graceful fallbacks when data is malformed

### Enhanced XML Structure
The XML output now includes:
- Project-level metadata
- Horizontal/Vertical alignment names
- Style information
- Detailed element properties with proper nesting
- Station point types as attributes

### Backward Compatibility
**v2.0 is fully backward compatible with v1.0**
- Existing Dynamo scripts using API extractor will continue to work
- Dictionary input format still supported
- No breaking changes to existing workflows

## Updated Files

### Modified Files
- **xml_report_generator.py**:
  - Added `parse_station_value()` function
  - Added `convert_geotable_list_to_dict()` function
  - Enhanced `parse_geotable_rows()` with detailed element parsing
  - Updated `add_alignment_to_xml()` with new properties
  - Added `add_elements_to_xml()` method
  - Enhanced `add_stations_to_xml()` with point type support

- **README.md**:
  - Added Quick Start section
  - Updated usage instructions
  - Added troubleshooting for list format
  - Updated version history

### New Files
- **test_geotable_parser.py**: Test script with sample data
- **test_output.xml**: Sample XML output
- **RELEASE_NOTES_v2.0.md**: This file

## Migration Guide

### From v1.0 to v2.0

**No migration required!** v2.0 is drop-in compatible.

#### If Using API Extractor (Dictionary Format)
Your existing setup will continue to work without changes:
```python
# This still works exactly as before
extractor = GeoTableDataExtractor()
data = extractor.extract_geotable_data()
generator = XMLReportGenerator(data)
xml = generator.to_xml_string()
```

#### If Switching to List Format
Simply pass your geotable list instead:
```python
# New capability in v2.0
geotable_list = [...]  # Your raw geotable data
generator = XMLReportGenerator(geotable_list)  # Works!
xml = generator.to_xml_string()
```

Or in Dynamo:
```
IN[0]: geotable_list (instead of dictionary)
IN[1]: output_path
IN[2]: pretty_print
```

## Testing

All new features have been tested with:
- Sample geotable data from Civil 3D 2024
- Multiple alignment types (Linear, Parabola)
- Various station formats
- Different property combinations

Run the test script to validate:
```bash
python3 test_geotable_parser.py
```

Expected output:
- Parsed alignment properties
- 7 station points (POB, PVI, PVC, PVT)
- 3 geometric elements (2 Linear, 1 Parabola)
- Valid XML file at `test_output.xml`

## Performance

v2.0 maintains the same performance characteristics as v1.0:
- Linear time complexity with respect to data size
- Minimal memory overhead for format conversion
- No performance regression

## Known Limitations

1. **X/Y Coordinates**: When using list input, X/Y coordinates default to 0.0 (only elevation is captured from geotable). Use API extractor for full coordinates.

2. **Element Property Naming**: Property names include colons from Civil 3D (e.g., "Tangent Grade:"). This matches the source data format.

3. **Duplicate Stations**: When stations appear in multiple elements (e.g., PVI between two tangents), they are included in both the overall station list and each element's station list.

## Future Enhancements

Based on this release, potential future improvements:
- Option to remove duplicate stations
- X/Y coordinate extraction from geotable (if available)
- Property name normalization (remove trailing colons)
- Support for additional element types (Clothoid, Compound curves)
- Station equation handling

## Getting Help

If you encounter issues:
1. Check the updated [Troubleshooting section](README.md#troubleshooting)
2. Run `test_geotable_parser.py` to validate installation
3. Review debug output in Dynamo console
4. Report issues with sample data and error messages

## Contributors

- Enhanced parsing logic
- Station format conversion
- Automatic format detection
- Comprehensive testing

## Acknowledgments

Thanks to users who reported the list format issue that led to this enhancement.

---

**Download**: [xml_report_generator.py v2.0](xml_report_generator.py)

**Documentation**: [README.md](README.md)

**Issues**: Report bugs or request features via GitHub issues
