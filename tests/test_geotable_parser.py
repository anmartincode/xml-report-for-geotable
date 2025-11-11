"""
Test script to validate geotable parsing with sample data
This simulates the data format from Civil 3D geotables
"""

# Import the generator
from xml_report_generator import parse_geotable_rows, convert_geotable_list_to_dict, XMLReportGenerator

# Sample geotable data matching the format from the error screenshot
sample_geotable = [
    ['Project Name:', 'Codman_Final'],
    ['Description:', ''],
    ['Horizontal Alignment Name:', 'Prop_AshNB'],
    ['Description:', ''],
    ['Style:', 'Default'],
    ['Vertical Alignment Name:', 'Prop_AshNB'],
    ['Description:', ''],
    ['Style:', 'Default'],
    ['', 'STATION', 'ELEVATION'],
    ['Element: Linear', '', ''],
    ['POB', '641+44.67', '42.97'],
    ['PVI', '641+95.67', '41.74'],
    ['Tangent Grade:', '-2.411', ''],
    ['Tangent Length:', '51.01', ''],
    ['Element: Linear', '', ''],
    ['PVI', '641+95.67', '41.74'],
    ['PVC', '642+07.29', '41.46'],
    ['Tangent Grade:', '-2.411', ''],
    ['Tangent Length:', '11.62', ''],
    ['Element: Parabola', '', ''],
    ['PVC', '642+07.29', '41.46'],
    ['PVI', '642+57.29', '40.25'],
    ['PVT', '643+07.29', '40.21'],
    ['Length:', '100.00', ''],
    ['Headlight Sight Distance:', '540.41', ''],
    ['Entrance Grade:', '-2.411', ''],
    ['Exit Grade:', '-0.075', ''],
    ['r = ( g2 - g1 ) / L:', '2.336', ''],
    ['K = 1 / ( g2 - g1 ):', '42.804', ''],
    ['Middle Ordinate:', '0.29', ''],
]

def test_parser():
    """Test the geotable parser"""
    print("=" * 80)
    print("Testing Geotable Parser")
    print("=" * 80)

    # Parse the sample data
    alignment = parse_geotable_rows(sample_geotable)

    # Print results
    print("\n1. Alignment Properties:")
    print(f"   Project Name: {alignment['project_name']}")
    print(f"   Alignment Name: {alignment['name']}")
    print(f"   Horizontal: {alignment['horizontal_alignment']}")
    print(f"   Vertical: {alignment['vertical_alignment']}")
    print(f"   Style: {alignment['style']}")

    print("\n2. Stations Found:")
    for station in alignment['stations']:
        print(f"   {station['point_type']:8s} - Station: {station['station']:10.2f}, Elevation: {station['elevation']:7.2f}")

    print(f"\n3. Elements Found: {len(alignment['elements'])}")
    for idx, element in enumerate(alignment['elements']):
        print(f"\n   Element {idx + 1}: {element['type']}")
        print(f"      Stations in element: {len(element['stations'])}")
        for station in element['stations']:
            print(f"         {station['point_type']:8s} - {station['station']:10.2f} @ {station['elevation']:7.2f}")

        print(f"      Properties:")
        for prop_name, prop_value in element['properties'].items():
            print(f"         {prop_name}: {prop_value}")

    # Test XML generation
    print("\n" + "=" * 80)
    print("Testing XML Generation")
    print("=" * 80)

    # Convert to full data structure
    full_data = {
        'project_name': 'Test Project',
        'timestamp': '2025-11-10 12:00:00',
        'alignments': [alignment]
    }

    # Generate XML
    generator = XMLReportGenerator(full_data)
    xml_string = generator.to_xml_string(pretty_print=True)

    # Print first 2000 characters
    print("\nGenerated XML (first 2000 chars):")
    print(xml_string[:2000])

    # Save to file for inspection
    output_file = '/home/amartinez/.local/local/xml-report-for-geotable/test_output.xml'
    with open(output_file, 'w') as f:
        f.write(xml_string)
    print(f"\nFull XML saved to: {output_file}")

    print("\n" + "=" * 80)
    print("Test Complete!")
    print("=" * 80)

if __name__ == '__main__':
    test_parser()
