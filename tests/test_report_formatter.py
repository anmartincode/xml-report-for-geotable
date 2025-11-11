"""
Test the direct report formatter
"""

from geotable_report_formatter import parse_geotable_to_alignment, GeoTableReportFormatter

# Sample geotable data matching Codman_Final example
sample_vertical_geotable = [
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
    ['Element: Linear', '', ''],
    ['PVT', '643+07.29', '40.21'],
    ['PVC', '645+66.48', '40.02'],
    ['Tangent Grade:', '-0.075', ''],
    ['Tangent Length:', '259.19', ''],
    ['Element: Parabola', '', ''],
    ['PVC', '645+66.48', '40.02'],
    ['PVI', '646+78.98', '39.94'],
    ['PVT', '647+91.48', '38.22'],
    ['Length:', '225.00', ''],
    ['Stopping Sight Distance:', '571.52', ''],
    ['Entrance Grade:', '-0.075', ''],
    ['Exit Grade:', '-1.523', ''],
    ['r = ( g2 - g1 ) / L:', '-0.643', ''],
    ['K = 1 / ( g2 - g1 ):', '155.406', ''],
]

def test_formatter():
    print("=" * 80)
    print("Testing Direct Report Formatter")
    print("=" * 80)

    # Parse the data
    alignment = parse_geotable_to_alignment(sample_vertical_geotable)

    print("\n1. Parsed Alignment:")
    print(f"   Project: {alignment['project_name']}")
    print(f"   Name: {alignment['name']}")
    print(f"   Elements: {len(alignment['elements'])}")

    # Generate formatted report
    formatter = GeoTableReportFormatter(alignment)
    report = formatter.format_vertical_alignment_report()

    print("\n2. Generated Report:")
    print("=" * 80)
    print(report)
    print("=" * 80)

    # Save to file
    output_file = '/home/amartinez/.local/local/xml-report-for-geotable/test_formatted_report.txt'
    formatter.save_to_file(output_file)
    print(f"\nReport saved to: {output_file}")

    print("\n" + "=" * 80)
    print("Test Complete!")
    print("=" * 80)

if __name__ == '__main__':
    test_formatter()
