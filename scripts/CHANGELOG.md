# Changelog

All notable changes to the Civil 3D Geotable XML Report Generator will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-10-29

### Added
- Initial release of Civil 3D Geotable XML Report Generator
- Core extraction module (`civil3d_geotable_extractor.py`)
  - Extract alignment properties (name, description, length, stations)
  - Extract station-by-station coordinate data
  - Extract geometric entities (lines, arcs, spirals)
  - Extract superelevation/cant data
  - Configurable station sampling intervals
- XML generation module (`xml_report_generator.py`)
  - Convert geotable data to formatted XML
  - Pretty-print formatting option
  - Save to file or return as string
  - Namespace-compliant XML structure
- Dynamo graph template (`GeoTableReport.dyn`)
  - Pre-configured workflow for Dynamo
  - Input nodes for configuration
  - Watch nodes for result visualization
- Configuration system (`config.json`)
  - Extraction settings
  - XML formatting options
  - Rail-specific parameters
  - Output configuration
  - Filter options
- XML Schema Definition (`geotable_schema.xsd`)
  - Complete XSD 1.0 schema for validation
  - Supports all data elements
  - Type definitions and constraints
- Sample output (`sample_output.xml`)
  - Example XML report with realistic data
  - Demonstrates all data elements
  - Reference for expected output format
- Comprehensive documentation
  - README.md with full usage guide
  - QUICKSTART.md for rapid deployment
  - API reference documentation
  - Troubleshooting section
  - Configuration guide

### Features
- Extract single or multiple alignments
- Customizable station intervals for data density control
- Automatic coordinate and geometry extraction
- Support for standard rail geometric elements
- Schema-validated XML output
- Configurable precision for coordinates and stations
- Error handling and validation
- Compatible with Civil 3D 2020+
- Integration with Dynamo for Civil 3D

### Technical Details
- Python 3.x compatible
- Uses Civil 3D .NET API
- ElementTree XML generation
- Transaction-based data extraction for safety
- Support for locked documents

### Documentation
- Installation instructions
- Usage examples
- Configuration reference
- Troubleshooting guide
- API documentation
- Quick start guide

## [Unreleased]

### Planned Features
- Profile (vertical alignment) data extraction
- Cross-section data export
- Cant deficiency calculations
- Track quality index (TQI) metrics
- Additional export formats (CSV, JSON, GeoJSON)
- BIM 360 integration
- Automated validation and quality checks
- Interactive visualization dashboard
- Batch processing capabilities
- Multi-language support

### Under Consideration
- Support for corridor data
- Assembly extraction
- Material quantities
- Construction staging information
- Integration with project management tools
- Cloud storage integration
- Real-time collaboration features
- Mobile viewing capabilities

---

## Version History Summary

| Version | Date       | Status  | Description                |
|---------|------------|---------|----------------------------|
| 1.0.0   | 2025-10-29 | Current | Initial public release     |

---

## Migration Guides

### Upgrading to Future Versions

When new versions are released, migration guides will be provided here to help users upgrade smoothly.

---

## Breaking Changes

### Version 1.0.0
- Initial release - no breaking changes

---

## Known Issues

### Version 1.0.0
- Superelevation data may not extract if alignment lacks proper superelevation view
- Very large alignments (>50km) with small intervals (<1m) may cause performance issues
- Some spiral types may not be fully characterized depending on Civil 3D version
- Coordinate system transformations not included (uses drawing coordinates as-is)

### Workarounds
- For superelevation: Ensure alignment has associated superelevation defined in Civil 3D
- For performance: Use larger station intervals (20-50m) for long alignments
- For spirals: Verify spiral definitions in Civil 3D before extraction
- For coordinates: Apply transformations in post-processing if needed

---

## Support & Feedback

To report bugs, request features, or provide feedback:
1. Check existing documentation for solutions
2. Review known issues above
3. Submit issue on repository
4. Include Civil 3D version, Dynamo version, and sample data if possible

---

## Release Notes Format

Future releases will follow this format:

### [X.Y.Z] - YYYY-MM-DD

#### Added
- New features

#### Changed
- Changes in existing functionality

#### Deprecated
- Features that will be removed in future versions

#### Removed
- Removed features

#### Fixed
- Bug fixes

#### Security
- Security patches or updates

---

**Note**: This project follows semantic versioning:
- **Major** version (X.0.0): Incompatible API changes
- **Minor** version (0.X.0): Backwards-compatible new features
- **Patch** version (0.0.X): Backwards-compatible bug fixes



