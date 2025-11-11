"""
Direct GeoTable Report Formatter for Civil 3D
Generates formatted text reports directly from geotable data
Bypasses XML → XSLT → XLSM pipeline

Usage in Dynamo:
- Import as Python Script node
- Connect geotable data input
- Outputs formatted text report matching MicroStation/InRoads format
"""

from datetime import datetime


class GeoTableReportFormatter:
    """Format geotable data into text-based reports"""

    def __init__(self, alignment_data):
        """Initialize with parsed alignment data"""
        self.alignment = alignment_data

    def format_vertical_alignment_report(self):
        """
        Generate vertical alignment report
        Format matches: Codman_Final example with STATION/ELEVATION columns
        """
        lines = []

        # Header
        lines.append(f"Project Name: {self.alignment.get('project_name', 'Unknown')}")
        lines.append(f"   Description: {self.alignment.get('description', '')}")

        if self.alignment.get('horizontal_alignment'):
            lines.append(f"Horizontal Alignment Name: {self.alignment.get('horizontal_alignment')}")
            lines.append(f"   Description:")
            lines.append(f"        Style: {self.alignment.get('style', 'Default')}")

        if self.alignment.get('vertical_alignment'):
            lines.append(f"Vertical Alignment Name: {self.alignment.get('vertical_alignment')}")
            lines.append(f"   Description:")
            lines.append(f"        Style: {self.alignment.get('style', 'Default')}")

        # Column headers
        lines.append(f"{'':40s} {'STATION':>15s} {'ELEVATION':>15s}")
        lines.append("")

        # Elements
        for element in self.alignment.get('elements', []):
            element_type = element.get('type', 'Unknown')
            lines.append(f"Element: {element_type}")

            # Station points with elevations
            for station in element.get('stations', []):
                point_type = station.get('point_type', '')
                station_str = self._format_station(station.get('station', 0))
                elevation = station.get('elevation', 0)

                lines.append(f"    {point_type:20s} {station_str:>15s} {elevation:>15.2f}")

            # Element properties
            props = element.get('properties', {})
            for prop_name, prop_value in props.items():
                # Format property lines
                lines.append(f"    {prop_name:20s} {prop_value:>15s}")

            lines.append("")  # Blank line between elements

        return '\n'.join(lines)

    def format_horizontal_alignment_report(self):
        """
        Generate horizontal alignment report
        Format matches: SCR Complete Align example with STATION/NORTHING/EASTING
        """
        lines = []

        # Header
        lines.append(f"Project Name: {self.alignment.get('project_name', 'Unknown')}")
        lines.append(f"   Description: {self.alignment.get('description', '')}")
        lines.append(f"Horizontal Alignment Name: {self.alignment.get('horizontal_alignment', self.alignment.get('name'))}")
        lines.append(f"   Description:")
        lines.append(f"        Style: {self.alignment.get('style', 'Default')}")

        # Column headers (for station/coordinate data)
        lines.append(f"{'':40s} {'STATION':>15s} {'NORTHING':>15s} {'EASTING':>15s}")
        lines.append("")

        # Elements
        for element in self.alignment.get('elements', []):
            element_type = element.get('type', 'Unknown')
            lines.append(f"Element: {element_type}")

            # Station points with coordinates
            for station in element.get('stations', []):
                point_type = station.get('point_type', '')
                station_str = self._format_station(station.get('station', 0))
                northing = station.get('y', 0)  # Y = Northing
                easting = station.get('x', 0)   # X = Easting

                lines.append(f"    {point_type:15s} ({')':1s} {station_str:>15s} {northing:>15.4f} {easting:>15.4f}")

            # Element properties
            props = element.get('properties', {})
            for prop_name, prop_value in props.items():
                # Format property lines with proper indentation
                if ':' in prop_name or 'Direction' in prop_name:
                    lines.append(f"    {prop_name:25s} {prop_value}")
                else:
                    lines.append(f"    {prop_name:20s} {prop_value:>15s}")

            lines.append("")  # Blank line between elements

        return '\n'.join(lines)

    def format_combined_report(self):
        """
        Generate combined report with both vertical and horizontal data
        """
        lines = []

        # Determine if we have vertical or horizontal data
        has_elevations = any(
            station.get('elevation') is not None
            for element in self.alignment.get('elements', [])
            for station in element.get('stations', [])
        )

        has_coordinates = any(
            station.get('x', 0) != 0 or station.get('y', 0) != 0
            for element in self.alignment.get('elements', [])
            for station in element.get('stations', [])
        )

        if has_elevations and not has_coordinates:
            return self.format_vertical_alignment_report()
        elif has_coordinates:
            return self.format_horizontal_alignment_report()
        else:
            return self.format_vertical_alignment_report()

    def _format_station(self, station_value):
        """
        Format station value back to Civil 3D format
        64144.67 -> "641+44.67"
        Ensures proper formatting with leading zeros
        """
        if station_value >= 100:
            hundreds = int(station_value / 100)
            remainder = station_value % 100
            # Ensure remainder is formatted with leading zero if needed (e.g., 07.29)
            return f"{hundreds}+{remainder:05.2f}"
        else:
            return f"{station_value:.2f}"

    def save_to_file(self, file_path, report_type='auto'):
        """
        Save report to text file

        Parameters:
        - file_path: Output file path
        - report_type: 'vertical', 'horizontal', or 'auto' (default)
        """
        if report_type == 'vertical':
            content = self.format_vertical_alignment_report()
        elif report_type == 'horizontal':
            content = self.format_horizontal_alignment_report()
        else:
            content = self.format_combined_report()

        with open(file_path, 'w') as f:
            f.write(content)

        return file_path


# ============= HELPER FUNCTIONS =============

def parse_station_value(station_str):
    """
    Convert Civil 3D station format to float
    Examples: "641+44.67" -> 64144.67, "100.50" -> 100.50
    """
    try:
        station_str = str(station_str).strip()
        if '+' in station_str:
            parts = station_str.split('+')
            if len(parts) == 2:
                return float(parts[0]) * 100 + float(parts[1])
        return float(station_str)
    except:
        return 0.0


def parse_geotable_to_alignment(geotable_rows):
    """
    Parse raw geotable rows into alignment dictionary
    Simplified version focusing on report formatting
    """
    alignment = {
        'name': 'Unknown',
        'project_name': '',
        'description': '',
        'horizontal_alignment': '',
        'vertical_alignment': '',
        'style': '',
        'elements': []
    }

    current_element = None

    for row in geotable_rows:
        if not isinstance(row, list) or len(row) < 2:
            continue

        prop_name = str(row[0]).strip()
        prop_value = str(row[1]).strip() if len(row) > 1 else ''
        third_value = str(row[2]).strip() if len(row) > 2 else ''

        prop_name_lower = prop_name.lower()

        # Skip header rows
        if prop_name_lower == '' and prop_value.upper() in ['STATION', 'NORTHING']:
            continue

        # Project properties
        if 'project name' in prop_name_lower:
            alignment['project_name'] = prop_value
        elif 'horizontal alignment name' in prop_name_lower:
            alignment['horizontal_alignment'] = prop_value
            if not alignment['name'] or alignment['name'] == 'Unknown':
                alignment['name'] = prop_value
        elif 'vertical alignment name' in prop_name_lower:
            alignment['vertical_alignment'] = prop_value
        elif prop_name_lower == 'style':
            alignment['style'] = prop_value
        elif prop_name_lower == 'description':
            alignment['description'] = prop_value

        # Element headers
        elif prop_name_lower.startswith('element:'):
            if current_element is not None:
                alignment['elements'].append(current_element)

            element_type = prop_name.split(':')[1].strip()
            current_element = {
                'type': element_type,
                'stations': [],
                'properties': {}
            }

        # Station points
        elif current_element is not None and prop_value:
            # Check if this is a station point (has numeric station value)
            station_val = parse_station_value(prop_value)

            if station_val > 0 and third_value:
                # This is a station with coordinate/elevation
                try:
                    numeric_third = float(third_value)
                    current_element['stations'].append({
                        'point_type': prop_name,
                        'station': station_val,
                        'elevation': numeric_third,
                        'x': 0.0,
                        'y': 0.0
                    })
                except ValueError:
                    # Not a coordinate, treat as property
                    current_element['properties'][prop_name] = prop_value
            else:
                # Property line
                current_element['properties'][prop_name] = prop_value

    # Save last element
    if current_element is not None:
        alignment['elements'].append(current_element)

    return alignment


# ============= HELPER: SANITIZE .NET OBJECTS =============

def sanitize_data(obj):
    """
    Recursively convert .NET types to Python native types
    This is critical when data comes from Dynamo nodes
    """
    if obj is None:
        return None

    # Check if already a Python dict
    if isinstance(obj, dict):
        py_dict = {}
        for k, v in obj.items():
            py_dict[str(k)] = sanitize_data(v)
        return py_dict

    # Check if already a Python list
    if isinstance(obj, list):
        return [sanitize_data(item) for item in obj]

    # Check if already a Python tuple
    if isinstance(obj, tuple):
        return tuple(sanitize_data(item) for item in obj)

    # Handle booleans (before numbers!)
    if isinstance(obj, bool):
        return bool(obj)

    # Handle strings
    if isinstance(obj, str):
        return str(obj)

    # Handle numbers (int and float)
    if isinstance(obj, (int, float)) and not isinstance(obj, bool):
        try:
            return float(str(obj))
        except:
            return obj

    # Handle .NET dict-like objects (have both 'keys' and 'items' methods)
    if hasattr(obj, 'keys') and hasattr(obj, 'items'):
        try:
            if callable(getattr(obj, 'items')) and callable(getattr(obj, 'keys')):
                py_dict = {}
                for k, v in obj.items():
                    py_dict[str(k)] = sanitize_data(v)
                return py_dict
        except:
            pass

    # Handle .NET list-like objects (iterable but not string/dict)
    if not hasattr(obj, 'keys'):
        try:
            # Try to iterate
            py_list = []
            for item in obj:
                py_list.append(sanitize_data(item))
            return py_list
        except (TypeError, AttributeError):
            pass

    # Last resort: convert to string
    return str(obj)


# ============= DYNAMO ENTRY POINT =============

def main(geotable_data, output_path=None, report_type='auto'):
    """
    Main function for Dynamo

    Inputs:
    - IN[0]: geotable_data - Raw geotable list or parsed alignment dict
    - IN[1]: output_path - File path to save report (optional)
    - IN[2]: report_type - 'vertical', 'horizontal', or 'auto' (default)

    Output:
    - Formatted text report string or file path
    """
    try:
        # CRITICAL: Sanitize input to convert .NET objects to Python types
        geotable_data = sanitize_data(geotable_data)

        # Debug: Show what we received
        debug_info = []
        debug_info.append(f"Input type: {type(geotable_data)}")

        # Handle different input types
        if isinstance(geotable_data, list):
            # Check if it's already a parsed list (list of dicts) or raw geotable (list of lists)
            if len(geotable_data) > 0 and isinstance(geotable_data[0], list):
                # Raw geotable list - parse it
                debug_info.append("Parsing raw geotable list")
                alignment_data = parse_geotable_to_alignment(geotable_data)
            elif len(geotable_data) > 0 and isinstance(geotable_data[0], dict):
                # This is a list of alignment dicts - use first one
                debug_info.append("Using first alignment from list")
                alignment_data = geotable_data[0]
            else:
                return "Error: Empty or invalid list format\n" + "\n".join(debug_info)

        elif isinstance(geotable_data, dict):
            # Already parsed - check if it's wrapped or direct alignment
            if 'alignments' in geotable_data:
                alignments_list = geotable_data['alignments']
                debug_info.append(f"Found {len(alignments_list)} alignments in dict")

                if len(alignments_list) == 0:
                    return "Error: No alignments in data\n" + "\n".join(debug_info)

                first_alignment = alignments_list[0]
                debug_info.append(f"First alignment type: {type(first_alignment)}")

                # Check if alignments are still lists (not parsed yet)
                if isinstance(first_alignment, list):
                    debug_info.append("Alignments are still in list format - parsing now")
                    alignment_data = parse_geotable_to_alignment(first_alignment)
                else:
                    alignment_data = first_alignment
            else:
                # Direct alignment dict
                debug_info.append("Using direct alignment dict")
                alignment_data = geotable_data
        else:
            return "Error: Invalid input type. Expected list or dict, got: " + str(type(geotable_data)) + "\n" + "\n".join(debug_info)

        # Verify alignment_data is a dict
        if not isinstance(alignment_data, dict):
            return f"Error: alignment_data is {type(alignment_data)}, expected dict\n" + "\n".join(debug_info)

        debug_info.append(f"Final alignment_data type: {type(alignment_data)}")
        debug_info.append(f"Alignment name: {alignment_data.get('name', 'Unknown')}")

        # Create formatter
        formatter = GeoTableReportFormatter(alignment_data)

        # Generate report
        if output_path and output_path != "":
            # Save to file
            result_path = formatter.save_to_file(output_path, report_type)
            return f"Report saved to: {result_path}"
        else:
            # Return as string
            if report_type == 'vertical':
                return formatter.format_vertical_alignment_report()
            elif report_type == 'horizontal':
                return formatter.format_horizontal_alignment_report()
            else:
                return formatter.format_combined_report()

    except Exception as e:
        import traceback
        error_msg = f"Error generating report: {str(e)}\n{traceback.format_exc()}"
        if 'debug_info' in locals():
            error_msg += "\n\nDEBUG INFO:\n" + "\n".join(debug_info)
        return error_msg


# Dynamo execution
if 'IN' in dir():
    geotable_data = IN[0] if len(IN) > 0 else []
    output_path = IN[1] if len(IN) > 1 and IN[1] != "" else None
    report_type = IN[2] if len(IN) > 2 else 'auto'

    OUT = main(geotable_data, output_path, report_type)
else:
    OUT = "Script loaded successfully. Use in Dynamo context."
