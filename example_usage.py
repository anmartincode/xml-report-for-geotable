"""
Example Usage Script for Civil 3D Geotable XML Report Generator

This file demonstrates various ways to use the extraction and XML generation
modules. These examples can be adapted for use in Dynamo or standalone scripts.

NOTE: This script is for reference only. In Dynamo, you'll copy the content
of the main scripts into Python Script nodes rather than importing them.
"""

# =============================================================================
# EXAMPLE 1: Basic Usage in Dynamo
# =============================================================================

"""
In Dynamo Python Script Node:

# Copy entire content of civil3d_geotable_extractor.py here
# Then use the main() function:

alignment_name = IN[0]  # "Main_Rail_Line" or "" for all
station_interval = IN[1]  # 10.0

OUT = main(alignment_name, station_interval)
"""

# =============================================================================
# EXAMPLE 2: Extract All Alignments with Default Interval
# =============================================================================

"""
# In Dynamo - Python Script Node 1 (Extractor)

# [Paste civil3d_geotable_extractor.py content here]

# At the bottom:
alignment_name = IN[0] if IN[0] else None
station_interval = 10.0  # Default to 10 units

OUT = main(alignment_name, station_interval)
"""

# =============================================================================
# EXAMPLE 3: Generate XML and Save to File
# =============================================================================

"""
# In Dynamo - Python Script Node 2 (XML Generator)

# [Paste xml_report_generator.py content here]

# At the bottom:
geotable_data = IN[0]  # From extractor node
output_path = IN[1]  # "C:\\Output\\report.xml"
pretty_print = True

OUT = main(geotable_data, output_path, pretty_print)
"""

# =============================================================================
# EXAMPLE 4: Extract Specific Alignment with Custom Interval
# =============================================================================

"""
# Configuration Code Block in Dynamo:

alignment_name = "Main_Rail_Line_Track_01";
station_interval = 5.0;  // Every 5 meters/feet
output_path = "C:\\Projects\\Rail\\output\\detailed_report.xml";

# Connect these to the Python Script nodes
"""

# =============================================================================
# EXAMPLE 5: Multiple Alignments, Separate Files
# =============================================================================

"""
# For processing multiple alignments into separate files,
# create a loop in Dynamo using List nodes:

# Code Block:
alignment_names = ["Track_01", "Track_02", "Track_03"];

# Python Script (with List Levels):
# Process each alignment separately
# Output will be a list of XML files

# This requires using Dynamo's list management features
"""

# =============================================================================
# EXAMPLE 6: Conditional Extraction Based on Length
# =============================================================================

"""
# Modified extractor logic:

def extract_long_alignments_only(min_length=500.0):
    extractor = GeoTableDataExtractor()
    all_alignments = extractor.get_all_alignments()
    
    # Filter by length
    filtered = [a for a in all_alignments if a['length'] >= min_length]
    
    # Extract only filtered alignments
    result = {
        'project_name': extractor.civil_doc.Name,
        'timestamp': DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        'alignments': []
    }
    
    for alignment_info in filtered:
        # Extract data for each long alignment
        alignment_data = extract_single_alignment(alignment_info, 10.0)
        result['alignments'].append(alignment_data)
    
    return result

# Use in Dynamo:
OUT = extract_long_alignments_only(min_length=1000.0)
"""

# =============================================================================
# EXAMPLE 7: Custom Station Interval Per Alignment
# =============================================================================

"""
# Different intervals for different alignments:

config = {
    "Main_Track": 10.0,      # Main track every 10m
    "Siding_Track": 25.0,    # Sidings every 25m
    "Yard_Track": 50.0       # Yard tracks every 50m
}

# In extractor:
alignment_name = IN[0]
interval = config.get(alignment_name, 10.0)  # Default to 10.0

OUT = main(alignment_name, interval)
"""

# =============================================================================
# EXAMPLE 8: Dense Sampling on Curves Only
# =============================================================================

"""
# Adaptive interval based on entity type:

def extract_with_adaptive_interval(alignment, base_interval=20.0, curve_interval=5.0):
    stations_data = []
    curves = extract_alignment_curves(alignment)
    
    for curve in curves:
        if curve['type'] in ['Arc', 'Spiral']:
            # Dense sampling on curves
            interval = curve_interval
        else:
            # Sparse sampling on tangents
            interval = base_interval
        
        # Extract stations for this segment
        # Add to stations_data
    
    return stations_data
"""

# =============================================================================
# EXAMPLE 9: Include Project Metadata
# =============================================================================

"""
# Enhanced metadata in XML:

def add_project_metadata(geotable_data):
    # Add custom metadata
    geotable_data['project_metadata'] = {
        'project_code': 'RAIL-2025-01',
        'designer': 'Engineering Team',
        'review_status': 'For Review',
        'units': 'Metric',
        'coordinate_system': 'Local Grid',
        'datum': 'Project Datum',
        'notes': 'Preliminary design alignment'
    }
    return geotable_data

# Use before XML generation:
data = main(alignment_name, station_interval)
data_with_metadata = add_project_metadata(data)
OUT = data_with_metadata
"""

# =============================================================================
# EXAMPLE 10: Batch Export with Timestamp
# =============================================================================

"""
# Generate unique filename with timestamp:

import datetime

# Code Block in Dynamo:
base_path = "C:\\Output\\";
alignment_name = "Main_Track";
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S");
filename = alignment_name + "_" + timestamp + ".xml";
full_path = base_path + filename;

# Result: C:\Output\Main_Track_20251029_143015.xml
"""

# =============================================================================
# EXAMPLE 11: Validation and Error Handling
# =============================================================================

"""
# Robust extraction with validation:

def extract_with_validation(alignment_name, station_interval):
    try:
        # Input validation
        if station_interval <= 0:
            return "Error: Station interval must be positive"
        
        if station_interval > 1000:
            return "Warning: Very large interval may miss details"
        
        # Extract data
        extractor = GeoTableDataExtractor()
        result = extractor.extract_geotable_data(alignment_name, station_interval)
        
        # Validate result
        if not result.get('alignments'):
            return "Error: No alignments found"
        
        # Check data completeness
        for alignment in result['alignments']:
            if not alignment.get('stations'):
                return f"Warning: No stations in {alignment['name']}"
        
        return result
        
    except Exception as e:
        return f"Error during extraction: {str(e)}"

# In Dynamo:
OUT = extract_with_validation(IN[0], IN[1])
"""

# =============================================================================
# EXAMPLE 12: XML String Output for Preview
# =============================================================================

"""
# Generate XML as string for preview in Dynamo Watch node:

# Python Script Node 2:
geotable_data = IN[0]
output_path = ""  # Empty = return string instead of saving
pretty_print = True

OUT = main(geotable_data, output_path, pretty_print)

# Result: Watch node shows formatted XML string
# Good for debugging before saving to file
"""

# =============================================================================
# EXAMPLE 13: Coordinate Transformation (Advanced)
# =============================================================================

"""
# Transform coordinates to different system:

def transform_coordinates(stations_data, transform_function):
    # Apply transformation to each station
    for station in stations_data:
        x, y = transform_function(station['x'], station['y'])
        station['x'] = x
        station['y'] = y
    return stations_data

# Define transformation (example: shift origin)
def shift_to_project_origin(x, y):
    project_origin_x = 1000000.0
    project_origin_y = 2000000.0
    return x - project_origin_x, y - project_origin_y

# Apply to extracted data
"""

# =============================================================================
# EXAMPLE 14: Export Summary Statistics
# =============================================================================

"""
# Add statistics to XML output:

def calculate_alignment_stats(alignment_data):
    stats = {
        'total_length': alignment_data['length'],
        'total_stations': len(alignment_data['stations']),
        'num_curves': sum(1 for e in alignment_data['curves'] if e['type'] == 'Arc'),
        'num_spirals': sum(1 for e in alignment_data['curves'] if e['type'] == 'Spiral'),
        'num_tangents': sum(1 for e in alignment_data['curves'] if e['type'] == 'Line'),
        'min_radius': min([e.get('radius', float('inf')) 
                          for e in alignment_data['curves'] 
                          if e['type'] == 'Arc'] or [0]),
    }
    alignment_data['statistics'] = stats
    return alignment_data
"""

# =============================================================================
# EXAMPLE 15: Integration with Dynamo Player
# =============================================================================

"""
# For Dynamo Player automation:

1. Create input nodes in Dynamo graph:
   - String Input: "Alignment Name"
   - Number Slider: "Station Interval" (range: 1-100)
   - File Path: "Output Location"

2. Mark these as inputs (right-click → "Is Input")

3. Save the graph

4. In Dynamo Player:
   - Select the graph
   - Set input values in the dialog
   - Click "Run" - no need to open full Dynamo UI

Perfect for non-technical users!
"""

# =============================================================================
# TIPS FOR DYNAMO USAGE
# =============================================================================

"""
1. **Copy Full Scripts**: Always copy the ENTIRE script content into Dynamo
   Python nodes, not just portions

2. **Input Validation**: Add checks for empty strings, None values, negative numbers

3. **Error Display**: Use Watch nodes to display errors clearly

4. **File Paths**: Use raw strings in Python or double backslashes:
   - r"C:\Output\file.xml" OR "C:\\Output\\file.xml"

5. **Testing**: Start with small alignments and default settings

6. **Performance**: For large projects, increase station interval to 20-50m

7. **Reusability**: Save your configured Dynamo graph as a template

8. **Documentation**: Add notes to Dynamo nodes explaining settings

9. **Version Control**: Keep backups of working Dynamo graphs

10. **Batch Processing**: Use Dynamo Player for running on multiple drawings
"""

# =============================================================================
# COMMON PITFALLS TO AVOID
# =============================================================================

"""
❌ DON'T: Use very small intervals (<1m) on long alignments (>10km)
✅ DO: Use 10-20m intervals for most cases

❌ DON'T: Forget to create output directory before running
✅ DO: Ensure output path exists or create it

❌ DON'T: Leave alignment_name as None when you want specific alignment
✅ DO: Use "" (empty string) for all, or specific name for one

❌ DON'T: Try to import scripts as modules in Dynamo
✅ DO: Copy the full script content into Python Script nodes

❌ DON'T: Ignore error messages in Watch nodes
✅ DO: Read errors carefully and check troubleshooting guide

❌ DON'T: Run on production drawings without testing first
✅ DO: Test on sample/copy drawings first
"""

# =============================================================================
# SUPPORT
# =============================================================================

"""
For more examples and help:
- See README.md for comprehensive documentation
- Check QUICKSTART.md for rapid deployment
- Review sample_output.xml for expected format
- Consult Civil 3D API documentation for advanced customization
"""

print("Example usage reference loaded. See comments for usage patterns.")

