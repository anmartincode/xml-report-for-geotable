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
        ET.SubElement(properties, 'Length').text = str(round(alignment_data.get('length', 0), 3))
        ET.SubElement(properties, 'StartStation').text = str(round(alignment_data.get('start_station', 0), 3))
        ET.SubElement(properties, 'EndStation').text = str(round(alignment_data.get('end_station', 0), 3))
        
        # Station data
        self.add_stations_to_xml(alignment_data.get('stations', []), alignment_elem)
        
        # Curve data
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
            
            # Coordinates
            coords = ET.SubElement(station_elem, 'Coordinates')
            ET.SubElement(coords, 'X').text = str(round(station.get('x', 0), 6))
            ET.SubElement(coords, 'Y').text = str(round(station.get('y', 0), 6))
            ET.SubElement(coords, 'Z').text = str(round(station.get('z', 0), 6))
            
            # Direction/Bearing
            ET.SubElement(station_elem, 'Direction').text = str(round(station.get('direction', 0), 6))
            ET.SubElement(station_elem, 'Offset').text = str(round(station.get('offset', 0), 3))
    
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


def main(geotable_data, output_path=None, pretty_print=True):
    """
    Main function called by Dynamo
    
    Inputs (IN[0], IN[1], IN[2]):
    - geotable_data: Dictionary - Output from civil3d_geotable_extractor.py
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
            if isinstance(first_alignment, dict):
                debug_info.append(f"First alignment keys: {list(first_alignment.keys())}")
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


