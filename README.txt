Alignment API Explorer for Civil 3D
====================================

This plugin allows you to export Civil 3D alignment geometry data to JSON format.

INSTALLATION:
-------------
This plugin is automatically installed to:
%APPDATA%\Autodesk\ApplicationPlugins\AlignmentApiExplorer.bundle

After building the project, restart Civil 3D to load the plugin.

USAGE:
------
Command: EXPORT_ALIGNMENT_GEOMETRY_JSON

1. Type the command in Civil 3D
2. Select an alignment when prompted
3. The alignment geometry will be exported to a JSON file on your desktop
4. The file will be named: AlignmentExport_[AlignmentName].json

OUTPUT:
-------
The JSON file contains detailed geometry information for each segment:
- Line segments: Start/End points, Length, PI points
- Curve segments: Radius, Delta, Length, PC/PT/PI points, Center point
- Spiral segments: Spiral type, Length, Start/End radius, Points

For support, contact: support@yourcompany.com







