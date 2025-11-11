"""
XML Report Generator for Civil 3D Geotable Data
Converts extracted geotable data into formatted XML reports

Usage in Dynamo:
- Import as second Python Script node
- Connect output from civil3d_geotable_extractor.py
- Outputs formatted XML string or file
"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
import json
import sys
from datetime import datetime


class XMLReportGenerator:
    """Generate XML reports from geotable data"""
    
    def __init__(self, geotable_data):
        """Initialize with geotable data dictionary"""
        self.data = geotable_data
        self.root = None
        
    def create_xml_structure(self):
        """Create the main XML structure"""
        # Create root element
        self.root = ET.Element('GeotableReport')
        self.root.set('version', '1.0')
        self.root.set('xmlns', 'http://civil3d.autodesk.com/geotable')
        
        # Add project information
        project_info = ET.SubElement(self.root, 'ProjectInfo')
        ET.SubElement(project_info, 'ProjectName').text = self.data.get('project_name', 'Unknown')
        ET.SubElement(project_info, 'GeneratedDate').text = self.data.get('timestamp', 
                                                                          datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        ET.SubElement(project_info, 'ReportType').text = 'Rail Alignment Geotable'
        
        # Add alignments section
        alignments_section = ET.SubElement(self.root, 'Alignments')
        alignments_section.set('count', str(len(self.data.get('alignments', []))))
        
        return alignments_section
    
    def add_alignment_to_xml(self, alignment_data, parent_element):
        """Add a single alignment's data to XML"""
        alignment_elem = ET.SubElement(parent_element, 'Alignment')
        alignment_elem.set('name', alignment_data.get('name', 'Unnamed'))
        alignment_elem.set('id', alignment_data.get('id', ''))

        # Basic properties
        properties = ET.SubElement(alignment_elem, 'Properties')
        ET.SubElement(properties, 'Description').text = alignment_data.get('description', '')

        # Add project name if available
        if alignment_data.get('project_name'):
            ET.SubElement(properties, 'ProjectName').text = alignment_data.get('project_name', '')

        # Add horizontal/vertical alignment names if available
        if alignment_data.get('horizontal_alignment'):
            ET.SubElement(properties, 'HorizontalAlignment').text = alignment_data.get('horizontal_alignment', '')
        if alignment_data.get('vertical_alignment'):
            ET.SubElement(properties, 'VerticalAlignment').text = alignment_data.get('vertical_alignment', '')
        if alignment_data.get('style'):
            ET.SubElement(properties, 'Style').text = alignment_data.get('style', '')

        ET.SubElement(properties, 'Length').text = str(round(alignment_data.get('length', 0), 3))
        ET.SubElement(properties, 'StartStation').text = str(round(alignment_data.get('start_station', 0), 3))
        ET.SubElement(properties, 'EndStation').text = str(round(alignment_data.get('end_station', 0), 3))

        # Station data
        self.add_stations_to_xml(alignment_data.get('stations', []), alignment_elem)

        # Elements data (detailed geometric elements from geotable)
        if alignment_data.get('elements'):
            self.add_elements_to_xml(alignment_data.get('elements', []), alignment_elem)

        # Curve data (from extractor)
        if alignment_data.get('curves'):
            self.add_curves_to_xml(alignment_data.get('curves', []), alignment_elem)

        # Superelevation data
        if alignment_data.get('superelevation'):
            self.add_superelevation_to_xml(alignment_data.get('superelevation', []), alignment_elem)
    
    def add_stations_to_xml(self, stations_data, parent_element):
        """Add station data to XML"""
        stations_elem = ET.SubElement(parent_element, 'Stations')
        stations_elem.set('count', str(len(stations_data)))

        for station in stations_data:
            station_elem = ET.SubElement(stations_elem, 'Station')
            station_elem.set('value', str(round(station.get('station', 0), 3)))

            # Add point type if available (POB, PVI, PVC, PVT, etc.)
            if station.get('point_type'):
                station_elem.set('type', station.get('point_type'))

            # Coordinates
            coords = ET.SubElement(station_elem, 'Coordinates')
            ET.SubElement(coords, 'X').text = str(round(station.get('x', 0), 6))
            ET.SubElement(coords, 'Y').text = str(round(station.get('y', 0), 6))
            ET.SubElement(coords, 'Z').text = str(round(station.get('z', 0), 6))

            # Elevation (if different from Z or explicitly provided)
            if station.get('elevation') is not None:
                ET.SubElement(station_elem, 'Elevation').text = str(round(station.get('elevation', 0), 2))

            # Direction/Bearing
            ET.SubElement(station_elem, 'Direction').text = str(round(station.get('direction', 0), 6))
            ET.SubElement(station_elem, 'Offset').text = str(round(station.get('offset', 0), 3))
    
    def add_elements_to_xml(self, elements_data, parent_element):
        """Add detailed geometric elements from geotable to XML"""
        elements_elem = ET.SubElement(parent_element, 'GeometricElements')
        elements_elem.set('count', str(len(elements_data)))

        for idx, element in enumerate(elements_data):
            element_elem = ET.SubElement(elements_elem, 'Element')
            element_elem.set('type', element.get('type', 'Unknown'))
            element_elem.set('index', str(idx))

            # Add element stations (POB, PVI, PVC, PVT, etc.)
            if element.get('stations'):
                stations_elem = ET.SubElement(element_elem, 'Stations')
                for station in element.get('stations', []):
                    station_elem = ET.SubElement(stations_elem, 'Station')
                    station_elem.set('type', station.get('point_type', 'Unknown'))
                    station_elem.set('value', str(round(station.get('station', 0), 2)))

                    if station.get('elevation') is not None:
                        ET.SubElement(station_elem, 'Elevation').text = str(round(station.get('elevation', 0), 2))

            # Add element properties
            if element.get('properties'):
                props_elem = ET.SubElement(element_elem, 'ElementProperties')
                for prop_name, prop_value in element.get('properties', {}).items():
                    prop_elem = ET.SubElement(props_elem, 'Property')
                    prop_elem.set('name', prop_name)
                    prop_elem.text = str(prop_value)

    def add_curves_to_xml(self, curves_data, parent_element):
        """Add curve/entity data to XML"""
        curves_elem = ET.SubElement(parent_element, 'GeometricEntities')
        curves_elem.set('count', str(len(curves_data)))

        for curve in curves_data:
            entity_elem = ET.SubElement(curves_elem, 'Entity')
            entity_elem.set('type', curve.get('type', 'Unknown'))
            entity_elem.set('index', str(curve.get('index', 0)))

            # Common properties
            ET.SubElement(entity_elem, 'StartStation').text = str(round(curve.get('start_station', 0), 3))
            ET.SubElement(entity_elem, 'EndStation').text = str(round(curve.get('end_station', 0), 3))
            ET.SubElement(entity_elem, 'Length').text = str(round(curve.get('length', 0), 3))

            # Type-specific properties
            if curve.get('type') == 'Arc':
                arc_props = ET.SubElement(entity_elem, 'ArcProperties')
                ET.SubElement(arc_props, 'Radius').text = str(round(curve.get('radius', 0), 3))
                ET.SubElement(arc_props, 'ChordLength').text = str(round(curve.get('chord_length', 0), 3))
                ET.SubElement(arc_props, 'Direction').text = curve.get('direction', 'CW')

            elif curve.get('type') == 'Spiral':
                spiral_props = ET.SubElement(entity_elem, 'SpiralProperties')
                ET.SubElement(spiral_props, 'RadiusIn').text = str(round(curve.get('radius_in', 0), 3))
                ET.SubElement(spiral_props, 'RadiusOut').text = str(round(curve.get('radius_out', 0), 3))
                ET.SubElement(spiral_props, 'SpiralType').text = curve.get('spiral_type', 'Unknown')

            elif curve.get('type') == 'Line':
                line_props = ET.SubElement(entity_elem, 'LineProperties')
                ET.SubElement(line_props, 'Bearing').text = str(round(curve.get('bearing', 0), 6))
    
    def add_superelevation_to_xml(self, superelevation_data, parent_element):
        """Add superelevation data to XML"""
        se_elem = ET.SubElement(parent_element, 'Superelevation')
        se_elem.set('count', str(len(superelevation_data)))
        
        for se in superelevation_data:
            critical_elem = ET.SubElement(se_elem, 'CriticalStation')
            critical_elem.set('station', str(round(se.get('station', 0), 3)))
            
            ET.SubElement(critical_elem, 'TransitionType').text = se.get('type', 'Unknown')
            ET.SubElement(critical_elem, 'LeftSlope').text = str(round(se.get('left_slope', 0), 4))
            ET.SubElement(critical_elem, 'RightSlope').text = str(round(se.get('right_slope', 0), 4))
    
    def generate_xml(self):
        """Generate complete XML structure"""
        alignments_section = self.create_xml_structure()
        
        # Add each alignment
        for alignment in self.data.get('alignments', []):
            self.add_alignment_to_xml(alignment, alignments_section)
        
        return self.root
    
    def to_xml_string(self, pretty_print=True):
        """Convert to XML string"""
        xml_root = self.generate_xml()
        
        if pretty_print:
            # Pretty print with proper indentation
            rough_string = ET.tostring(xml_root, encoding='utf-8')
            reparsed = minidom.parseString(rough_string)
            return reparsed.toprettyxml(indent="  ", encoding='utf-8').decode('utf-8')
        else:
            return ET.tostring(xml_root, encoding='utf-8').decode('utf-8')
    
    def save_to_file(self, file_path, pretty_print=True):
        """Save XML to file"""
        xml_string = self.to_xml_string(pretty_print)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(xml_string)
        
        return file_path


# ============= DYNAMO SCRIPT ENTRY POINT =============

def sanitize_data(obj):
    """Force convert any object to Python native types"""
    if obj is None:
        return None
    
    # Check if already a Python dict (most important check first!)
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
    
    # Now handle .NET dict-like objects (have both 'keys' and 'items' methods)
    # Check BOTH to ensure it's really dict-like and not just iterable
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
    # But first make sure it's NOT dict-like!
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


def convert_geotable_list_to_dict(geotable_list):
    """
    Convert raw geotable list format to dictionary format

    Input: List of lists like [['Name, Text'], ['Description, Text'], ...]
    Output: Dictionary structure compatible with XMLReportGenerator
    """
    result = {
        'project_name': 'Civil 3D Geotable Export',
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'alignments': []
    }

    # Check if this is a list of geotables (multiple alignments)
    if isinstance(geotable_list, list) and len(geotable_list) > 0:

        # Case 1: Single geotable (list of rows)
        if isinstance(geotable_list[0], list) and len(geotable_list[0]) == 2:
            # This is a single geotable with rows like ['Property', 'Value']
            alignment_dict = parse_geotable_rows(geotable_list)
            if alignment_dict:
                result['alignments'].append(alignment_dict)

        # Case 2: Multiple geotables (list of geotables)
        else:
            for item in geotable_list:
                if isinstance(item, list):
                    alignment_dict = parse_geotable_rows(item)
                    if alignment_dict:
                        result['alignments'].append(alignment_dict)

    return result


def parse_station_value(station_str):
    """
    Convert Civil 3D station format to float
    Examples: "641+44.67" -> 64144.67, "100.50" -> 100.50
    """
    try:
        station_str = str(station_str).strip()
        if '+' in station_str:
            # Format: "641+44.67"
            parts = station_str.split('+')
            if len(parts) == 2:
                return float(parts[0]) * 100 + float(parts[1])
        # Direct numeric value
        return float(station_str)
    except:
        return 0.0


def parse_geotable_rows(rows):
    """
    Parse geotable rows into alignment dictionary

    Input: List of [property, value] pairs from Civil 3D geotable
    Output: Alignment dictionary with detailed geometric data
    """
    alignment = {
        'name': 'Unknown',
        'id': '',
        'description': '',
        'length': 0.0,
        'start_station': 0.0,
        'end_station': 0.0,
        'project_name': '',
        'horizontal_alignment': '',
        'vertical_alignment': '',
        'style': '',
        'stations': [],
        'curves': [],
        'elements': [],
        'superelevation': []
    }

    current_element = None
    current_element_type = None
    station_data = []

    for i, row in enumerate(rows):
        if not isinstance(row, list) or len(row) < 2:
            continue

        # Get property name and value
        prop_name = str(row[0]).strip()
        prop_value = str(row[1]).strip() if len(row) > 1 else ''

        # Check if there's a third column (e.g., ELEVATION)
        third_value = str(row[2]).strip() if len(row) > 2 else ''

        prop_name_lower = prop_name.lower()

        # Skip header rows
        if prop_name_lower == '' and prop_value.upper() == 'STATION':
            continue

        # Project-level properties
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

        # Element headers (Linear, Parabola, Spiral, etc.)
        elif prop_name_lower.startswith('element:'):
            # Save previous element if exists
            if current_element is not None:
                alignment['elements'].append(current_element)

            # Start new element
            element_type = prop_name.split(':')[1].strip()
            current_element_type = element_type
            current_element = {
                'type': element_type,
                'properties': {},
                'stations': []
            }

        # Station points with elevation (POB, PVI, PVC, PVT, etc.)
        elif prop_value and third_value and current_element is not None:
            # Try to parse as station+elevation pair
            station_val = parse_station_value(prop_value)
            try:
                elevation_val = float(third_value)

                if station_val > 0 or elevation_val != 0:
                    current_element['stations'].append({
                        'point_type': prop_name,
                        'station': station_val,
                        'elevation': elevation_val
                    })

                    # Also track for overall alignment
                    station_data.append({
                        'station': station_val,
                        'elevation': elevation_val,
                        'point_type': prop_name,
                        'x': 0.0,
                        'y': 0.0,
                        'z': elevation_val,
                        'direction': 0.0,
                        'offset': 0.0
                    })
                else:
                    # Not a valid station/elevation, treat as property
                    if current_element is not None:
                        current_element['properties'][prop_name] = prop_value
            except ValueError:
                # Not a numeric elevation, treat as property
                if current_element is not None:
                    current_element['properties'][prop_name] = prop_value

        # Element properties
        elif current_element is not None and prop_value and prop_name:
            current_element['properties'][prop_name] = prop_value

        # Fallback: generic property mapping
        elif 'length' in prop_name_lower and current_element is None:
            try:
                alignment['length'] = float(prop_value)
            except:
                pass
        elif 'start' in prop_name_lower and 'station' in prop_name_lower:
            try:
                alignment['start_station'] = parse_station_value(prop_value)
            except:
                pass
        elif 'end' in prop_name_lower and 'station' in prop_name_lower:
            try:
                alignment['end_station'] = parse_station_value(prop_value)
            except:
                pass
        elif 'id' in prop_name_lower:
            alignment['id'] = prop_value

    # Save last element
    if current_element is not None:
        alignment['elements'].append(current_element)

    # Add collected station data
    alignment['stations'] = station_data

    # Calculate start/end stations if not set
    if station_data:
        if alignment['start_station'] == 0.0:
            alignment['start_station'] = min(s['station'] for s in station_data)
        if alignment['end_station'] == 0.0:
            alignment['end_station'] = max(s['station'] for s in station_data)

    return alignment


def main(geotable_data, output_path=None, pretty_print=True):
    """
    Main function called by Dynamo

    Inputs (IN[0], IN[1], IN[2]):
    - geotable_data: Dictionary or List - Output from civil3d_geotable_extractor.py or raw geotable
    - output_path: String - File path to save XML (optional, None for string output)
    - pretty_print: Boolean - Whether to format XML with indentation

    Output (OUT):
    - XML string or file path (if saved)
    """
    try:
        # Validate input type
        if geotable_data is None:
            return "Error: No data received from extractor"

        # DEBUG: Show what we received
        debug_info = []
        debug_info.append(f"Received type: {type(geotable_data)}")
        debug_info.append(f"Type string: {str(type(geotable_data))}")

        # Handle error input
        if isinstance(geotable_data, str):
            if geotable_data.startswith("Error"):
                return geotable_data
            else:
                return "Error: Received string instead of data dictionary: " + str(geotable_data)[:500]

        # CRITICAL: Sanitize data again here in case Dynamo converted it during transfer
        geotable_data = sanitize_data(geotable_data)

        debug_info.append(f"After sanitization type: {type(geotable_data)}")

        # Convert list format to dictionary if needed
        if isinstance(geotable_data, list):
            debug_info.append("Converting list format to dictionary structure...")
            geotable_data = convert_geotable_list_to_dict(geotable_data)
            debug_info.append(f"After conversion - Alignments count: {len(geotable_data.get('alignments', []))}")

        # Verify it's a dictionary
        if not isinstance(geotable_data, dict):
            return "Error: Expected dictionary but received: " + str(type(geotable_data)) + "\n" + "\n".join(debug_info)

        # Check if dictionary has required keys
        if 'alignments' not in geotable_data:
            return "Error: Data missing 'alignments' key. Keys found: " + str(list(geotable_data.keys())) + "\n" + "\n".join(debug_info)

        # DEBUG: Check alignments type
        alignments = geotable_data.get('alignments', [])
        debug_info.append(f"Alignments type: {type(alignments)}")
        debug_info.append(f"Alignments count: {len(alignments) if hasattr(alignments, '__len__') else 'N/A'}")

        if len(alignments) > 0:
            first_alignment = alignments[0]
            debug_info.append(f"First alignment type: {type(first_alignment)}")

            # Additional check - if alignments still contain lists, try to convert them
            if isinstance(first_alignment, list):
                debug_info.append("WARNING: Alignments still in list format after conversion. Attempting individual conversion...")
                converted_alignments = []
                for idx, alignment in enumerate(alignments):
                    if isinstance(alignment, list):
                        converted = parse_geotable_rows(alignment)
                        converted_alignments.append(converted)
                        if idx == 0:
                            debug_info.append(f"First converted alignment: {str(converted)[:200]}")
                    else:
                        converted_alignments.append(alignment)
                geotable_data['alignments'] = converted_alignments
                debug_info.append(f"Converted {len(converted_alignments)} alignments from list to dict format")
            else:
                debug_info.append(f"First alignment value (first 200 chars): {str(first_alignment)[:200]}")

        # Generate XML
        generator = XMLReportGenerator(geotable_data)

        if output_path and output_path != "":
            # Save to file
            result_path = generator.save_to_file(output_path, pretty_print)
            return "XML report saved to: " + result_path
        else:
            # Return XML string
            return generator.to_xml_string(pretty_print)

    except Exception as e:
        import traceback
        error_msg = "Error generating XML: " + str(e) + "\n" + traceback.format_exc()
        if 'debug_info' in locals():
            error_msg += "\n\nDEBUG INFO:\n" + "\n".join(debug_info)
        return error_msg


# Check if running in Dynamo context
if 'IN' in dir():
    # Get inputs from Dynamo
    geotable_data = IN[0] if len(IN) > 0 else {}
    output_path = IN[1] if len(IN) > 1 and IN[1] != "" else None
    pretty_print = IN[2] if len(IN) > 2 else True
    
    # Execute and return output
    OUT = main(geotable_data, output_path, pretty_print)
else:
    # Standalone execution for testing
    OUT = "Script loaded successfully. Use in Dynamo context."

