# Spiral (Clothoid) Properties Now Exposed

Based on the Civil 3D UI requirements shown in the screenshots, the following properties are now being extracted for spiral elements:

## Core Properties

| Property | Source | Description |
|----------|--------|-------------|
| `Index` | Loop counter | Element index in alignment |
| `EntityType` | AlignmentEntity | "Spiral" |
| `StartStation` | AlignmentSpiral | Starting station |
| `EndStation` | AlignmentSpiral | Ending station |
| `Length` | AlignmentSpiral | Spiral length |
| `StartPoint` | PointLocation() | 3D coordinates {X, Y, Z} |
| `EndPoint` | PointLocation() | 3D coordinates {X, Y, Z} |

## Spiral Definition Properties

| Property | Source | Description |
|----------|--------|-------------|
| `SpiralDefinition` | AlignmentSpiral.SpiralDefinition | Type: "Clothoid", "CubicParabola", etc. |
| `SpiralType` | Calculated from EntityBefore/After | "Simple" or "Compound" |
| `Direction` | AlignmentSpiral.Direction | Spiral direction |

## Radius Properties

| Property | Source | Description |
|----------|--------|-------------|
| `RadiusIn` | AlignmentSpiral.RadiusIn | Incoming radius (Infinity for tangent) |
| `RadiusOut` | AlignmentSpiral.RadiusOut | Outgoing radius |
| `A` | AlignmentSpiral.A | Spiral parameter A |

## Calculated Geometric Properties

| Property | Formula | Description |
|----------|---------|-------------|
| `K` | L² / R | Spiral constant |
| `P` | L² / (24·R) | Offset distance parameter |
| `ChordLength` | √[(X₂-X₁)² + (Y₂-Y₁)²] | Direct distance start to end |
| `ChordDirection` | atan2(Y₂-Y₁, X₂-X₁) | Azimuth of chord |

## Tangent Offset Properties (from AlignmentSubEntitySpiral)

| Property | Source | Description |
|----------|--------|-------------|
| `TotalX` | AlignmentSubEntitySpiral.TotalX | Tangent offset X |
| `TotalY` | AlignmentSubEntitySpiral.TotalY | Tangent offset Y |
| `Delta` | AlignmentSubEntitySpiral.Delta | Total deflection angle |

## Spiral PI Properties

| Property | Calculation | Description |
|----------|-------------|-------------|
| `Spiral_PI_Station` | StartStation + Length/2 | Station at spiral midpoint |
| `Spiral_PI_Northing` | PointLocation(PI_Station).Y | Y coordinate at PI |
| `Spiral_PI_Easting` | PointLocation(PI_Station).X | X coordinate at PI |
| `Spiral_PI_Included_Angle` | L / (2·R) | Deflection angle at PI |

## Additional Properties (via Reflection)

The code also uses .NET reflection to automatically discover and export any additional properties available on `AlignmentSubEntitySpiral` that are not explicitly documented.

## Example JSON Output

```json
{
  "Index": 4,
  "EntityType": "Spiral",
  "StartStation": 281.048,
  "EndStation": 311.048,
  "Length": 30.0,
  "StartPoint": { "X": 774340.2422, "Y": 2928747.3795, "Z": 0.0 },
  "EndPoint": { "X": 774350.9831, "Y": 2928736.0550, "Z": 0.0 },
  "SpiralDefinition": "Clothoid",
  "RadiusIn": 1.7976931348623157E+308,
  "RadiusOut": 250.0,
  "A": 86.60254037844387,
  "Direction": "DirectionRight",
  "SpiralType": "Simple",
  "K": 3.6,
  "P": 0.15,
  "ChordLength": 15.618,
  "ChordDirection": 2.9287,
  "TotalX": 29.985,
  "TotalY": 1.799,
  "Delta": 0.12,
  "Spiral_PI_Station": 296.048,
  "Spiral_PI_Northing": 2928741.717,
  "Spiral_PI_Easting": 774345.613,
  "Spiral_PI_Included_Angle": 0.06
}
```

## Notes

1. **RadiusIn = Infinity**: When a spiral comes from a tangent (line), RadiusIn will be `1.7976931348623157E+308` (double.MaxValue representing infinity in C#)

2. **K Constant**: Uses the formula K = L²/R which is standard for clothoid spirals

3. **P Parameter**: Uses the approximation P = L²/(24·R) which is valid for small deflection angles typical in highway design

4. **Spiral PI**: The PI (Point of Intersection) for a spiral is calculated at the midpoint of the spiral length, which is where the maximum offset typically occurs

5. **Included Angle**: The total angular change of the spiral, calculated as L/(2·R)

## Testing

After rebuilding and reloading in AutoCAD:
1. Run `NETUNLOAD` to unload the existing DLL
2. Rebuild: `dotnet build AlignmentApiExplorer.csproj`
3. Run `NETLOAD` and select the new DLL
4. Run `EXPORT_ALIGNMENT_GEOMETRY_JSON`
5. Check the JSON output for complete spiral data
