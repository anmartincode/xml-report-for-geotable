# Testing Spiral Export with AlignmentSubEntitySpiral

## What Changed

The code has been updated to properly expose **Clothoid spiral properties** from the `AlignmentSubEntitySpiral` class.

### Key Improvements:

1. **Better Error Handling**: Wrapped the entire spiral case in try-catch to capture any errors
2. **Basic Spiral Properties** exported from `AlignmentSpiral`:
   - `StartStation`, `EndStation`, `Length`
   - `StartPoint`, `EndPoint` (3D coordinates)
   - `SpiralDefinition` (type: Clothoid, Cubic Parabola, etc.)
   - `RadiusIn`, `RadiusOut` (incoming/outgoing radius)
   - `A` (spiral parameter)
   - `K` (spiral constant = A²/R)
   - `Direction`

3. **AlignmentSubEntitySpiral Properties**:
   - `TotalX`, `TotalY` (tangent offsets)
   - `Delta` (total deflection angle)
   - **Dynamic property discovery** using reflection

4. **Reflection-Based Discovery**: The code now uses .NET reflection to automatically discover and export ALL available properties on the `AlignmentSubEntitySpiral` class, including:
   - Primitive types (int, double, etc.)
   - Enums (converted to strings)
   - Point2d and Point3d types (converted to {X, Y, Z} objects)

## How to Test

### Step 1: Reload the DLL in AutoCAD

In AutoCAD command line:
```
NETUNLOAD
```
Select `AlignmentApiExplorer` from the list

Then rebuild:
```powershell
cd test
dotnet build AlignmentApiExplorer.csproj
```

### Step 2: Load the New DLL

In AutoCAD command line:
```
NETLOAD
```
Browse to: `C:\Users\anmartinez\.vscode\local\xml-report-for-geotable\test\bin\Debug\AlignmentApiExplorer.dll`

### Step 3: Run the Export Command

```
EXPORT_ALIGNMENT_GEOMETRY_JSON
```

Select the alignment with spiral (Clothoid) elements.

### Step 4: Check the Results

The JSON file will be saved to your Desktop as `AlignmentExport_<AlignmentName>.json`

Look for spiral entries that now contain:
- All basic properties (stations, length, radii, etc.)
- SubEntity properties (TotalX, TotalY, Delta)
- Any additional properties discovered via reflection

### Expected Output for Spirals

Instead of empty spiral entries like:
```json
{
  "Index": 4,
  "EntityType": "Spiral"
}
```

You should now see detailed information like:
```json
{
  "Index": 4,
  "EntityType": "Spiral",
  "StartStation": 281.048...,
  "EndStation": 311.048...,
  "Length": 30.0,
  "StartPoint": { "X": ..., "Y": ..., "Z": ... },
  "EndPoint": { "X": ..., "Y": ..., "Z": ... },
  "SpiralDefinition": "Clothoid",
  "RadiusIn": 1.7976931348623157E+308,
  "RadiusOut": 250.0,
  "A": 86.60254037844387,
  "K": 30.0,
  "Direction": "DirectionRight",
  "TotalX": 29.985...,
  "TotalY": 1.799...,
  "Delta": 0.12,
  ... (additional discovered properties)
}
```

## Troubleshooting

### If spirals are still empty:

1. Check the AutoCAD command line for error messages when running the export
2. The JSON may contain an "Error" field with details:
   ```json
   {
     "Index": 4,
     "EntityType": "Spiral",
     "Error": "Unable to read spiral data: <specific error message>"
   }
   ```

3. Check if `SubEntityCount` is 0 for spirals - this would mean no AlignmentSubEntitySpiral is available

### If you see "_Error" properties:

The reflection code logs property access failures as `<PropertyName>_Error` fields. These indicate which properties exist but couldn't be accessed (e.g., due to permissions or invalid states).

## Next Steps

After successful export, review the discovered properties to:
1. Identify which properties are most useful for your GeoTable reports
2. Document any undocumented properties found via reflection
3. Add specific property calculations if needed (e.g., tangent lengths, spiral angles, etc.)
