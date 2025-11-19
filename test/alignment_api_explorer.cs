using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using static Autodesk.Civil.DatabaseServices.AlignmentSubEntity;
using static Autodesk.Civil.DatabaseServices.AlignmentSubEntitySpiral;
using Newtonsoft.Json;

namespace C3DAlignmentExporter
{
    public class Commands
    {
        [CommandMethod("EXPORT_ALIGNMENT_GEOMETRY_JSON")] 
        public void ExportAlignmentGeometryJson()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            CivilDocument civDoc = CivilDocument.GetCivilDocument(db);

            try
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect an Alignment: ");
                peo.SetRejectMessage("\nObject must be an Alignment.");
                peo.AddAllowedClass(typeof(Alignment), true);
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                    return;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Alignment alignment = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        ed.WriteMessage("\nFailed to read alignment.");
                        return;
                    }

                    var exportData = new List<object>();

                    AlignmentEntityCollection ents = alignment.Entities;

                    for (int i = 0; i < ents.Count; i++)
                    {
                        AlignmentEntity ent = ents[i];
                        if (ent == null) continue;
                        
                        var segmentData = new Dictionary<string, object>();

                        segmentData["Index"] = i;
                        segmentData["EntityType"] = ent.EntityType.ToString();

                        switch (ent.EntityType)
                        {
                            case AlignmentEntityType.Line:
                                AlignmentLine line = ent as AlignmentLine;
                                if (line != null)
                                {
                                    segmentData["StartStation"] = line.StartStation;
                                    segmentData["EndStation"] = line.EndStation;
                                    segmentData["Length"] = line.Length;
                                    
                                    // Get 3D coordinates using PointLocation
                                    double x1 = 0, y1 = 0, z1 = 0;
                                    double x2 = 0, y2 = 0, z2 = 0;
                                    alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                    alignment.PointLocation(line.EndStation, 0, 0, ref x2, ref y2, ref z2);
                                    
                                    segmentData["StartPoint"] = new { X = x1, Y = y1, Z = z1 };
                                    segmentData["EndPoint"] = new { X = x2, Y = y2, Z = z2 };
                                    segmentData["Direction"] = line.Direction;
                                }
                                break;

                            case AlignmentEntityType.Arc:
                                AlignmentArc arc = ent as AlignmentArc;
                                if (arc != null)
                                {
                                    segmentData["StartStation"] = arc.StartStation;
                                    segmentData["EndStation"] = arc.EndStation;
                                    segmentData["Radius"] = arc.Radius;
                                    segmentData["Delta"] = arc.Delta;

                                    // Calculate Degree of Curvature (Chord definition)
                                    // D = 2 * arcsin(chord_length / (2 * R)) * (180 / π)
                                    // Using standard 100-foot chord
                                    double standardChord = 100.0;
                                    double degreeOfCurvatureChord = 2.0 * Math.Asin(standardChord / (2.0 * arc.Radius)) * (180.0 / Math.PI);
                                    segmentData["DegreeOfCurvature_Chord"] = degreeOfCurvatureChord;

                                    // Calculate ChordLength (full chord from PC to PT)
                                    // Chord = 2 * R * sin(Delta / 2)
                                    double chordLength = 2.0 * arc.Radius * Math.Sin(arc.Delta / 2.0);
                                    segmentData["ChordLength"] = chordLength;

                                    // Calculate MidOrdinate (distance from midpoint of chord to midpoint of arc)
                                    // M = R * (1 - cos(Delta / 2))
                                    double midOrdinate = arc.Radius * (1.0 - Math.Cos(arc.Delta / 2.0));
                                    segmentData["MidOrdinate"] = midOrdinate;

                                    // Calculate External Tangent (distance from PI to midpoint of arc)
                                    // E = R * (1 / cos(Delta / 2) - 1) = R * (sec(Delta / 2) - 1)
                                    double externalTangent = arc.Radius * ((1.0 / Math.Cos(arc.Delta / 2.0)) - 1.0);
                                    segmentData["ExternalTangent"] = externalTangent;

                                    // Calculate External Secant (distance from PI to PC or PT along tangent)
                                    // T = R * tan(Delta / 2)
                                    double externalSecant = arc.Radius * Math.Tan(arc.Delta / 2.0);
                                    segmentData["ExternalSecant"] = externalSecant;

                                    segmentData["Length"] = arc.Length;
                                    segmentData["Clockwise"] = arc.Clockwise;
                                    
                                    // Get 3D coordinates using PointLocation
                                    double ax1 = 0, ay1 = 0, az1 = 0;
                                    double ax2 = 0, ay2 = 0, az2 = 0;
                                    alignment.PointLocation(arc.StartStation, 0, 0, ref ax1, ref ay1, ref az1);
                                    alignment.PointLocation(arc.EndStation, 0, 0, ref ax2, ref ay2, ref az2);
                                    
                                    segmentData["StartPoint"] = new { X = ax1, Y = ay1, Z = az1 };
                                    segmentData["EndPoint"] = new { X = ax2, Y = ay2, Z = az2 };
                                    
                                    // Add PC = startpoint of arc
                                    segmentData["PC"] = new { X = ax1, Y = ay1, Z = az1 };
                                    
                                    // Add PT = endpoint of arc
                                    segmentData["PT"] = new { X = ax2, Y = ay2, Z = az2 };
                                    
                                    // Add center point of circular curve (2D)
                                    var center = arc.CenterPoint;
                                    segmentData["CenterPoint"] = new { X = center.X, Y = center.Y, Z = 0.0 };
                                    
                                    // Get PIStation from arc
                                    try
                                    {
                                        double piStation = arc.PIStation;
                                        segmentData["PIStation"] = piStation;
                                        
                                        // Get PIPoint from sub-entity
                                        if (arc.SubEntityCount > 0)
                                        {
                                            var subEntity = arc[0];
                                            if (subEntity is AlignmentSubEntityArc)
                                            {
                                                var subEntityArc = subEntity as AlignmentSubEntityArc;
                                                var piPoint = subEntityArc.PIPoint;
                                                
                                                // Get Z coordinate at PI station (use exact X,Y from PIPoint, get Z from alignment)
                                                double piX = piPoint.X;
                                                double piY = piPoint.Y;
                                                double piZ = 0;
                                                alignment.PointLocation(piStation, 0, 0, ref piX, ref piY, ref piZ);
                                                
                                                // Use exact X,Y from PIPoint (not modified by PointLocation)
                                                segmentData["PIPoint"] = new { X = piPoint.X, Y = piPoint.Y, Z = piZ };
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                break;

                            case AlignmentEntityType.Spiral:
                                try
                                {
                                    // Access spiral through SubEntity
                                    if (ent.SubEntityCount > 0)
                                    {
                                        var subEntity = ent[0];
                                        if (subEntity is AlignmentSubEntitySpiral)
                                        {
                                            var spiralSubEntity = subEntity as AlignmentSubEntitySpiral;
                                            
                                            // Get basic properties
                                            double startStation = spiralSubEntity.StartStation;
                                            double endStation = spiralSubEntity.EndStation;
                                            double length = spiralSubEntity.Length;
                                            
                                            segmentData["StartStation"] = startStation;
                                            segmentData["EndStation"] = endStation;
                                            segmentData["Length"] = length;

                                            // Get 3D coordinates using PointLocation
                                            double sx1 = 0, sy1 = 0, sz1 = 0;
                                            double sx2 = 0, sy2 = 0, sz2 = 0;
                                            alignment.PointLocation(startStation, 0, 0, ref sx1, ref sy1, ref sz1);
                                            alignment.PointLocation(endStation, 0, 0, ref sx2, ref sy2, ref sz2);

                                            segmentData["StartPoint"] = new { X = sx1, Y = sy1, Z = sz1 };
                                            segmentData["EndPoint"] = new { X = sx2, Y = sy2, Z = sz2 };

                                            // Get SubEntity properties
                                            try { segmentData["TotalX"] = spiralSubEntity.TotalX; } catch { }
                                            try { segmentData["TotalY"] = spiralSubEntity.TotalY; } catch { }
                                            try { segmentData["Delta"] = spiralSubEntity.Delta; } catch { }

                                            // Use reflection to get all properties
                                            var spiralType = spiralSubEntity.GetType();
                                            
                                            // Try to get specific properties via reflection
                                            try
                                            {
                                                var radiusInProp = spiralType.GetProperty("RadiusIn");
                                                var radiusOutProp = spiralType.GetProperty("RadiusOut");
                                                var aProp = spiralType.GetProperty("A");
                                                var defProp = spiralType.GetProperty("SpiralDefinition");
                                                var dirProp = spiralType.GetProperty("Direction");
                                                var beforeProp = spiralType.GetProperty("EntityBefore");
                                                var afterProp = spiralType.GetProperty("EntityAfter");
                                                
                                                double radiusIn = 0, radiusOut = 0, spiralA = 0;
                                                
                                                if (radiusInProp != null) 
                                                {
                                                    radiusIn = (double)radiusInProp.GetValue(spiralSubEntity);
                                                    segmentData["RadiusIn"] = radiusIn;
                                                }
                                                if (radiusOutProp != null) 
                                                {
                                                    radiusOut = (double)radiusOutProp.GetValue(spiralSubEntity);
                                                    segmentData["RadiusOut"] = radiusOut;
                                                }
                                                if (aProp != null) 
                                                {
                                                    spiralA = (double)aProp.GetValue(spiralSubEntity);
                                                    segmentData["A"] = spiralA;
                                                }
                                                if (defProp != null) segmentData["SpiralDefinition"] = defProp.GetValue(spiralSubEntity).ToString();
                                                if (dirProp != null) segmentData["Direction"] = dirProp.GetValue(spiralSubEntity).ToString();
                                                
                                                // Get SPI properties from API
                                                var spiStationProp = spiralType.GetProperty("SPIStation");
                                                var spiProp = spiralType.GetProperty("SPI");
                                                
                                                if (spiStationProp != null)
                                                {
                                                    double spiStation = (double)spiStationProp.GetValue(spiralSubEntity);
                                                    segmentData["SPIStation"] = spiStation;
                                                }
                                                
                                                if (spiProp != null)
                                                {
                                                    var spiPoint = (Autodesk.AutoCAD.Geometry.Point3d)spiProp.GetValue(spiralSubEntity);
                                                    segmentData["SPI_Point"] = new { X = spiPoint.X, Y = spiPoint.Y, Z = spiPoint.Z };
                                                    segmentData["SPI_Northing"] = spiPoint.Y;
                                                    segmentData["SPI_Easting"] = spiPoint.X;
                                                }
                                                
                                                // Determine spiral type
                                                if (beforeProp != null && afterProp != null)
                                                {
                                                    int entityBefore = (int)beforeProp.GetValue(spiralSubEntity);
                                                    int entityAfter = (int)afterProp.GetValue(spiralSubEntity);
                                                    var beforeType = (AlignmentEntityType)entityBefore;
                                                    var afterType = (AlignmentEntityType)entityAfter;
                                                    
                                                    if ((beforeType == AlignmentEntityType.Line && afterType == AlignmentEntityType.Arc) ||
                                                        (beforeType == AlignmentEntityType.Arc && afterType == AlignmentEntityType.Line))
                                                        segmentData["SpiralType"] = "Simple";
                                                    else if (beforeType == AlignmentEntityType.Arc && afterType == AlignmentEntityType.Arc)
                                                        segmentData["SpiralType"] = "Compound";
                                                    else
                                                        segmentData["SpiralType"] = "Simple";
                                                }
                                                
                                                // Calculate K (spiral constant)
                                                double radiusValue = radiusIn == 0 || double.IsInfinity(radiusIn) ? radiusOut : radiusIn;
                                                if (radiusValue != 0 && !double.IsInfinity(radiusValue))
                                                {
                                                    segmentData["K"] = length * length / radiusValue;
                                                    segmentData["P"] = (length * length) / (24.0 * radiusValue);
                                                    segmentData["Spiral_PI_Included_Angle"] = length / (2.0 * radiusValue);
                                                    
                                                    // Add manual calculation as fallback if API properties not available
                                                    if (!segmentData.ContainsKey("SPIStation"))
                                                    {
                                                        double piStation = startStation + (length / 2.0);
                                                        segmentData["Spiral_PI_Station_Calculated"] = piStation;
                                                        
                                                        double piX = 0, piY = 0, piZ = 0;
                                                        alignment.PointLocation(piStation, 0, 0, ref piX, ref piY, ref piZ);
                                                        segmentData["Spiral_PI_Northing_Calculated"] = piY;
                                                        segmentData["Spiral_PI_Easting_Calculated"] = piX;
                                                    }
                                                }
                                            }
                                            catch { }

                                            // Chord length and direction
                                            try
                                            {
                                                double chordLength = Math.Sqrt(Math.Pow(sx2 - sx1, 2) + Math.Pow(sy2 - sy1, 2));
                                                segmentData["ChordLength"] = chordLength;
                                                segmentData["ChordDirection"] = Math.Atan2(sy2 - sy1, sx2 - sx1);
                                            }
                                            catch { }

                                            // Discover all additional properties via reflection
                                            var props = spiralType.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                                            foreach (var prop in props)
                                            {
                                                if (segmentData.ContainsKey(prop.Name))
                                                    continue;

                                                try
                                                {
                                                    var value = prop.GetValue(spiralSubEntity);
                                                    if (value != null)
                                                    {
                                                        if (prop.PropertyType.IsPrimitive || prop.PropertyType == typeof(string) || 
                                                            prop.PropertyType == typeof(decimal) || prop.PropertyType == typeof(double))
                                                        {
                                                            segmentData[prop.Name] = value;
                                                        }
                                                        else if (prop.PropertyType.IsEnum)
                                                        {
                                                            segmentData[prop.Name] = value.ToString();
                                                        }
                                                        else if (prop.PropertyType.Name.Contains("Point"))
                                                        {
                                                            try
                                                            {
                                                                var xProp = value.GetType().GetProperty("X");
                                                                var yProp = value.GetType().GetProperty("Y");
                                                                if (xProp != null && yProp != null)
                                                                {
                                                                    segmentData[prop.Name] = new { X = xProp.GetValue(value), Y = yProp.GetValue(value), Z = 0.0 };
                                                                }
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                        else
                                        {
                                            segmentData["Error"] = "SubEntity is not AlignmentSubEntitySpiral";
                                        }
                                    }
                                    else
                                    {
                                        segmentData["Error"] = "No SubEntities found";
                                    }
                                }
                                catch (System.Exception spiralEx)
                                {
                                    segmentData["Error"] = $"Unable to read spiral data: {spiralEx.Message}";
                                    ed.WriteMessage($"\nSpiral at index {i} error: {spiralEx.Message}");
                                }
                                break;
                        }

                        exportData.Add(segmentData);
                    }

                    string json = JsonConvert.SerializeObject(exportData, Formatting.Indented);

                    string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string path = Path.Combine(folder, $"AlignmentExport_{alignment.Name}.json");
                    File.WriteAllText(path, json);

                    tr.Commit();

                    ed.WriteMessage($"\nExported to: {path}");
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Error: {ex.Message}");
            }
        }
    }
}
