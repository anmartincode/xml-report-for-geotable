"""
Civil 3D GeoTable PDF Report Generator
Generates PDF reports matching MicroStation/InRoads format from geotable data or XML

Requirements:
    pip install reportlab

Usage in Dynamo:
- IN[0]: geotable data (list) or XML file path (string)
- IN[1]: output PDF file path
- IN[2]: report_type: 'auto', 'vertical', or 'horizontal'
- OUT: PDF file path or error message
"""

from datetime import datetime
import xml.etree.ElementTree as ET
import os

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class GeoTablePDFGenerator:
    """Generate PDF reports from geotable data"""

    def __init__(self, alignment_data):
        """Initialize with parsed alignment data (dict format)"""
        self.alignment = alignment_data

    def generate_pdf(self, output_path, report_type='auto'):
        """
        Generate PDF report

        Parameters:
        - output_path: Output PDF file path
        - report_type: 'vertical', 'horizontal', or 'auto'
        """
        if not REPORTLAB_AVAILABLE:
            raise ImportError("ReportLab is required. Install with: pip install reportlab")

        # Determine report type
        if report_type == 'auto':
            report_type = self._detect_report_type()

        # Create directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # Create PDF document
        if report_type == 'horizontal':
            pagesize = landscape(letter)
        else:
            pagesize = letter

        doc = SimpleDocTemplate(
            output_path,
            pagesize=pagesize,
            rightMargin=0.5*inch,
            leftMargin=0.5*inch,
            topMargin=0.5*inch,
            bottomMargin=0.5*inch
        )

        # Build content
        story = []

        if report_type == 'vertical':
            story = self._build_vertical_report()
        else:
            story = self._build_horizontal_report()

        # Build PDF
        doc.build(story)
        return output_path

    def _detect_report_type(self):
        """Auto-detect vertical vs horizontal report"""
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

        return 'horizontal' if has_coordinates else 'vertical'

    def _build_vertical_report(self):
        """Build vertical alignment report (Station/Elevation)"""
        story = []
        styles = getSampleStyleSheet()

        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Normal'],
            fontSize=10,
            leading=12,
            leftIndent=0
        )

        # Header information
        story.append(Paragraph(f"<b>Project Name:</b> {self.alignment.get('project_name', 'Unknown')}", title_style))
        story.append(Paragraph(f"<b>   Description:</b> {self.alignment.get('description', '')}", title_style))

        if self.alignment.get('horizontal_alignment'):
            story.append(Paragraph(f"<b>Horizontal Alignment Name:</b> {self.alignment.get('horizontal_alignment')}", title_style))
            story.append(Paragraph(f"<b>   Description:</b>", title_style))
            story.append(Paragraph(f"<b>        Style:</b> {self.alignment.get('style', 'Default')}", title_style))

        if self.alignment.get('vertical_alignment'):
            story.append(Paragraph(f"<b>Vertical Alignment Name:</b> {self.alignment.get('vertical_alignment')}", title_style))
            story.append(Paragraph(f"<b>   Description:</b>", title_style))
            story.append(Paragraph(f"<b>        Style:</b> {self.alignment.get('style', 'Default')}", title_style))

        story.append(Spacer(1, 0.2*inch))

        # Column headers
        header_data = [['', 'STATION', 'ELEVATION']]
        header_table = Table(header_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch])
        header_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Courier-Bold', 9),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 0.1*inch))

        # Elements
        for element in self.alignment.get('elements', []):
            element_type = element.get('type', 'Unknown')
            story.append(Paragraph(f"<b>Element: {element_type}</b>", title_style))

            # Station data
            element_data = []
            for station in element.get('stations', []):
                point_type = station.get('point_type', '')
                station_str = self._format_station(station.get('station', 0))
                elevation = station.get('elevation', 0)
                element_data.append([f"    {point_type}", station_str, f"{elevation:.2f}"])

            if element_data:
                station_table = Table(element_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch])
                station_table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'Courier', 8),
                    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ]))
                story.append(station_table)

            # Properties
            props = element.get('properties', {})
            prop_data = []
            for prop_name, prop_value in props.items():
                prop_data.append([f"    {prop_name}", prop_value, ''])

            if prop_data:
                prop_table = Table(prop_data, colWidths=[3.5*inch, 1.5*inch, 1.5*inch])
                prop_table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'Courier', 8),
                    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ]))
                story.append(prop_table)

            story.append(Spacer(1, 0.15*inch))

        return story

    def _build_horizontal_report(self):
        """Build horizontal alignment report (Station/Northing/Easting)"""
        story = []
        styles = getSampleStyleSheet()

        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Normal'],
            fontSize=10,
            leading=12,
            leftIndent=0
        )

        # Header information
        story.append(Paragraph(f"<b>Project Name:</b> {self.alignment.get('project_name', 'Unknown')}", title_style))
        story.append(Paragraph(f"<b>   Description:</b> {self.alignment.get('description', '')}", title_style))
        story.append(Paragraph(f"<b>Horizontal Alignment Name:</b> {self.alignment.get('horizontal_alignment', self.alignment.get('name'))}", title_style))
        story.append(Paragraph(f"<b>   Description:</b>", title_style))
        story.append(Paragraph(f"<b>        Style:</b> {self.alignment.get('style', 'Default')}", title_style))
        story.append(Spacer(1, 0.2*inch))

        # Column headers
        header_data = [['', 'STATION', 'NORTHING', 'EASTING']]
        header_table = Table(header_data, colWidths=[3*inch, 1.5*inch, 1.75*inch, 1.75*inch])
        header_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Courier-Bold', 9),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 0.1*inch))

        # Elements
        for element in self.alignment.get('elements', []):
            element_type = element.get('type', 'Unknown')
            story.append(Paragraph(f"<b>Element: {element_type}</b>", title_style))

            # Station data
            element_data = []
            for station in element.get('stations', []):
                point_type = station.get('point_type', '')
                station_str = self._format_station(station.get('station', 0))
                northing = station.get('y', 0)
                easting = station.get('x', 0)
                element_data.append([f"    {point_type} ( )", station_str, f"{northing:.4f}", f"{easting:.4f}"])

            if element_data:
                station_table = Table(element_data, colWidths=[3*inch, 1.5*inch, 1.75*inch, 1.75*inch])
                station_table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'Courier', 8),
                    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ]))
                story.append(station_table)

            # Properties
            props = element.get('properties', {})
            prop_data = []
            for prop_name, prop_value in props.items():
                # Some properties span multiple columns
                if ':' in prop_name or 'Direction' in prop_name:
                    prop_data.append([f"    {prop_name}", prop_value, '', ''])
                else:
                    prop_data.append([f"    {prop_name}", prop_value, '', ''])

            if prop_data:
                prop_table = Table(prop_data, colWidths=[3*inch, 1.5*inch, 1.75*inch, 1.75*inch])
                prop_table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'Courier', 8),
                    ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ]))
                story.append(prop_table)

            story.append(Spacer(1, 0.15*inch))

        return story

    def _format_station(self, station_value):
        """Format station value to Civil 3D format"""
        if station_value >= 100:
            hundreds = int(station_value / 100)
            remainder = station_value % 100
            return f"{hundreds}+{remainder:05.2f}"
        else:
            return f"{station_value:.2f}"


# ============= HELPER FUNCTIONS =============

def parse_station_value(station_str):
    """Convert Civil 3D station format to float"""
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
    """Parse raw geotable rows into alignment dictionary"""
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
            # Check if this is a station point
            station_val = parse_station_value(prop_value)

            if station_val > 0 and third_value:
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
                    current_element['properties'][prop_name] = prop_value
            else:
                current_element['properties'][prop_name] = prop_value

    # Save last element
    if current_element is not None:
        alignment['elements'].append(current_element)

    return alignment


def parse_xml_to_alignment(xml_path):
    """Parse XML file to alignment dictionary"""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Get namespace if present
    ns = {'geo': 'http://civil3d.autodesk.com/geotable'}

    # Get project info
    project_info = root.find('.//geo:ProjectInfo', ns) or root.find('.//ProjectInfo')
    project_name = ''
    if project_info is not None:
        proj_name_elem = project_info.find('.//geo:ProjectName', ns) or project_info.find('.//ProjectName')
        if proj_name_elem is not None:
            project_name = proj_name_elem.text or ''

    # Get first alignment
    alignment_elem = root.find('.//geo:Alignment', ns) or root.find('.//Alignment')
    if alignment_elem is None:
        raise ValueError("No alignment found in XML")

    alignment = {
        'name': alignment_elem.get('name', 'Unknown'),
        'project_name': project_name,
        'description': '',
        'horizontal_alignment': alignment_elem.get('name', ''),
        'vertical_alignment': '',
        'style': 'Default',
        'elements': []
    }

    # Parse elements
    elements = alignment_elem.findall('.//geo:GeometricElement', ns) or alignment_elem.findall('.//GeometricElement')
    for elem in elements:
        element_type = elem.get('type', 'Unknown')
        element = {
            'type': element_type,
            'stations': [],
            'properties': {}
        }

        # Parse stations
        stations = elem.findall('.//geo:Station', ns) or elem.findall('.//Station')
        for sta in stations:
            station_val = float(sta.get('value', 0))
            point_type = sta.get('pointType', '')
            x = float(sta.get('x', 0))
            y = float(sta.get('y', 0))
            z = float(sta.get('z', 0))

            element['stations'].append({
                'point_type': point_type,
                'station': station_val,
                'elevation': z,
                'x': x,
                'y': y
            })

        # Parse properties
        props = elem.findall('.//geo:Property', ns) or elem.findall('.//Property')
        for prop in props:
            prop_name = prop.get('name', '')
            prop_value = prop.get('value', '')
            element['properties'][prop_name] = prop_value

        alignment['elements'].append(element)

    return alignment


def sanitize_data(obj):
    """Recursively convert .NET types to Python native types"""
    if obj is None:
        return None

    if isinstance(obj, dict):
        py_dict = {}
        for k, v in obj.items():
            py_dict[str(k)] = sanitize_data(v)
        return py_dict

    if isinstance(obj, list):
        return [sanitize_data(item) for item in obj]

    if isinstance(obj, tuple):
        return tuple(sanitize_data(item) for item in obj)

    if isinstance(obj, bool):
        return bool(obj)

    if isinstance(obj, str):
        return str(obj)

    if isinstance(obj, (int, float)) and not isinstance(obj, bool):
        try:
            return float(str(obj))
        except:
            return obj

    # Handle .NET dict-like objects
    if hasattr(obj, 'keys') and hasattr(obj, 'items'):
        try:
            if callable(getattr(obj, 'items')) and callable(getattr(obj, 'keys')):
                py_dict = {}
                for k, v in obj.items():
                    py_dict[str(k)] = sanitize_data(v)
                return py_dict
        except:
            pass

    # Handle .NET list-like objects
    if not hasattr(obj, 'keys'):
        try:
            py_list = []
            for item in obj:
                py_list.append(sanitize_data(item))
            return py_list
        except (TypeError, AttributeError):
            pass

    return str(obj)


# ============= DYNAMO ENTRY POINT =============

def main(input_data, output_path, report_type='auto'):
    """
    Main function for Dynamo

    Inputs:
    - IN[0]: input_data - Geotable data (list) or XML file path (string)
    - IN[1]: output_path - Output PDF file path
    - IN[2]: report_type - 'vertical', 'horizontal', or 'auto' (default)

    Output:
    - PDF file path or error message
    """
    try:
        if not REPORTLAB_AVAILABLE:
            return "ERROR: ReportLab is required. Install with: pip install reportlab"

        # Sanitize input
        input_data = sanitize_data(input_data)

        # Determine input type
        if isinstance(input_data, str) and input_data.endswith('.xml'):
            # XML file path
            if not os.path.exists(input_data):
                return f"Error: XML file not found: {input_data}"
            alignment_data = parse_xml_to_alignment(input_data)

        elif isinstance(input_data, list):
            # Raw geotable data
            if len(input_data) > 0 and isinstance(input_data[0], list):
                alignment_data = parse_geotable_to_alignment(input_data)
            elif len(input_data) > 0 and isinstance(input_data[0], dict):
                alignment_data = input_data[0]
            else:
                return "Error: Invalid list format"

        elif isinstance(input_data, dict):
            # Already parsed alignment data
            if 'alignments' in input_data:
                alignments_list = input_data['alignments']
                if len(alignments_list) == 0:
                    return "Error: No alignments in data"
                first_alignment = alignments_list[0]
                if isinstance(first_alignment, list):
                    alignment_data = parse_geotable_to_alignment(first_alignment)
                else:
                    alignment_data = first_alignment
            else:
                alignment_data = input_data

        else:
            return f"Error: Invalid input type: {type(input_data)}"

        # Verify alignment data
        if not isinstance(alignment_data, dict):
            return f"Error: alignment_data is {type(alignment_data)}, expected dict"

        # Generate PDF
        generator = GeoTablePDFGenerator(alignment_data)
        result_path = generator.generate_pdf(output_path, report_type)

        return f"PDF report saved to: {result_path}"

    except Exception as e:
        import traceback
        return f"Error generating PDF: {str(e)}\n{traceback.format_exc()}"


# Dynamo execution
if 'IN' in dir():
    input_data = IN[0] if len(IN) > 0 else []
    output_path = IN[1] if len(IN) > 1 and IN[1] != "" else None
    report_type = IN[2] if len(IN) > 2 else 'auto'

    if output_path is None:
        OUT = "Error: Output PDF file path required (IN[1])"
    else:
        OUT = main(input_data, output_path, report_type)
else:
    OUT = "Script loaded successfully. Use in Dynamo context with ReportLab installed."
