"""
Test the PDF report generator
"""

import os
from geotable_pdf_generator import parse_geotable_to_alignment, GeoTablePDFGenerator

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

def test_pdf_generator():
    print("=" * 80)
    print("Testing PDF Report Generator")
    print("=" * 80)

    try:
        from reportlab.lib.pagesizes import letter
        print("\n✓ ReportLab is installed")
    except ImportError:
        print("\n✗ ReportLab is NOT installed")
        print("  Install with: pip install reportlab")
        print("  OR: python3 -m pip install reportlab")
        return

    # Parse the data
    alignment = parse_geotable_to_alignment(sample_vertical_geotable)

    print("\n1. Parsed Alignment:")
    print(f"   Project: {alignment['project_name']}")
    print(f"   Name: {alignment['name']}")
    print(f"   Elements: {len(alignment['elements'])}")

    # Generate PDF
    generator = GeoTablePDFGenerator(alignment)
    output_file = '/home/amartinez/.local/local/xml-report-for-geotable/test_report.pdf'

    try:
        result = generator.generate_pdf(output_file, 'vertical')
        print(f"\n2. PDF Generated:")
        print(f"   ✓ Saved to: {result}")
        print(f"   ✓ File exists: {os.path.exists(output_file)}")

        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"   ✓ File size: {file_size:,} bytes")
    except Exception as e:
        print(f"\n✗ Error generating PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return

    print("\n" + "=" * 80)
    print("Test Complete!")
    print("=" * 80)
    print("\nNext steps:")
    print("1. Open test_report.pdf to verify formatting")
    print("2. Compare to your MicroStation example reports")
    print("3. Use in Dynamo with geotable_pdf_generator.py")

if __name__ == '__main__':
    test_pdf_generator()
