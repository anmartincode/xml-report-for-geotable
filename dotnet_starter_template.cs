/*
 * Civil 3D GeoTable Reports - .NET Add-in Starter Template
 *
 * This is a basic template showing how a .NET add-in would work
 * Compile this with Visual Studio and Civil 3D API references
 */

using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using CivApp = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using CivDb = Autodesk.Civil.DatabaseServices;

[assembly: CommandClass(typeof(GeoTableReports.ReportCommands))]

namespace GeoTableReports
{
    public class ReportCommands : IExtensionApplication
    {
        // Called when Civil 3D starts up
        public void Initialize()
        {
            AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nGeoTable Reports loaded. Type GEOTABLE to generate reports.");
            }
        }

        // Called when Civil 3D shuts down
        public void Terminate()
        {
            // Cleanup code here
        }

        /// <summary>
        /// Main command to generate geotable reports
        /// </summary>
        [CommandMethod("GEOTABLE")]
        public void GenerateReport()
        {
            AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Show selection dialog
                using (var dialog = new ReportSelectionForm())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        string reportType = dialog.ReportType; // "Vertical" or "Horizontal"
                        string outputPath = dialog.OutputPath;

                        // Prompt user to select alignment
                        ObjectId alignmentId = SelectAlignment(ed);
                        if (alignmentId == ObjectId.Null)
                        {
                            ed.WriteMessage("\nNo alignment selected.");
                            return;
                        }

                        // Extract data and generate report
                        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

                            if (alignment == null)
                            {
                                ed.WriteMessage("\nInvalid alignment selected.");
                                tr.Abort();
                                return;
                            }

                            // Get alignment name safely for progress message
                            string alignmentName = "alignment";
                            try
                            {
                                alignmentName = alignment.Name ?? "alignment";
                            }
                            catch
                            {
                                alignmentName = "alignment";
                            }

                            // Show progress
                            ed.WriteMessage($"\nProcessing alignment: {alignmentName}...");

                            try
                            {
                                if (reportType == "Vertical")
                                {
                                    GenerateVerticalReport(alignment, outputPath);
                                }
                                else
                                {
                                    GenerateHorizontalReport(alignment, outputPath);
                                }

                                tr.Commit();

                                // Show success message with actual file path (.txt)
                                string actualPath = outputPath.Replace(".pdf", ".txt");
                                MessageBox.Show(
                                    $"Report generated successfully!\n\nSaved to: {actualPath}",
                                    "Success",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );

                                ed.WriteMessage($"\n✓ Report saved to: {actualPath}");
                            }
                            catch (System.Exception reportEx)
                            {
                                tr.Abort();
                                throw new System.Exception($"Error generating {reportType} report: {reportEx.Message}", reportEx);
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string errorMessage = ex.Message;
                if (ex.InnerException != null)
                {
                    errorMessage += $"\n\nInner Exception: {ex.InnerException.Message}";
                }
                
                ed.WriteMessage($"\n✗ Error: {errorMessage}");
                ed.WriteMessage($"\nStack Trace: {ex.StackTrace}");
                
                MessageBox.Show(
                    $"Error generating report:\n\n{errorMessage}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Batch process multiple alignments
        /// </summary>
        [CommandMethod("GEOTABLE_BATCH")]
        public void BatchProcess()
        {
            AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                using (var dialog = new BatchProcessForm())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        string outputFolder = dialog.OutputFolder;
                        bool includeVertical = dialog.IncludeVertical;
                        bool includeHorizontal = dialog.IncludeHorizontal;

                        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            CivilDocument civilDoc = CivilApplication.ActiveDocument;

                            // Get all alignments in drawing
                            using (ObjectIdCollection alignmentIds = civilDoc.GetAlignmentIds())
                            {
                                int count = 0;
                                int total = alignmentIds.Count;

                                foreach (ObjectId alignmentId in alignmentIds)
                                {
                                    Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                                    if (alignment == null) continue;

                                    count++;
                                    ed.WriteMessage($"\nProcessing {count}/{total}: {alignment.Name}");

                                    // Generate reports
                                    if (includeVertical)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Vertical.pdf");
                                        GenerateVerticalReport(alignment, path);
                                    }

                                    if (includeHorizontal)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Horizontal.pdf");
                                        GenerateHorizontalReport(alignment, path);
                                    }
                                }

                                tr.Commit();

                                MessageBox.Show(
                                    $"Batch processing complete!\n\n{count} alignments processed.",
                                    "Success",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n✗ Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Prompt user to select an alignment
        /// </summary>
        private ObjectId SelectAlignment(Editor ed)
        {
            PromptEntityOptions options = new PromptEntityOptions("\nSelect alignment: ");
            options.SetRejectMessage("\nMust be an alignment.");
            options.AddAllowedClass(typeof(Alignment), true);

            PromptEntityResult result = ed.GetEntity(options);

            if (result.Status == PromptStatus.OK)
            {
                return result.ObjectId;
            }

            return ObjectId.Null;
        }

        /// <summary>
        /// Generate vertical alignment report
        /// </summary>
        private void GenerateVerticalReport(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                if (alignment == null)
                {
                    throw new System.ArgumentNullException(nameof(alignment), "Alignment object is null");
                }

                // Get project name from drawing properties
                Database db = null;
                string projectName = "Unknown Project";
                try
                {
                    db = alignment.Database;
                    if (db != null && !string.IsNullOrEmpty(db.Filename))
                    {
                        projectName = System.IO.Path.GetFileNameWithoutExtension(db.Filename);
                    }
                }
                catch
                {
                    projectName = "Unknown Project";
                }

                // Get alignment properties safely
                string alignmentName = "";
                string alignmentDescription = "";
                string alignmentStyle = "Default";
                try
                {
                    alignmentName = alignment?.Name ?? "";
                }
                catch { }
                try
                {
                    alignmentDescription = alignment?.Description ?? "";
                }
                catch { }
                try
                {
                    alignmentStyle = alignment?.StyleName ?? "Default";
                }
                catch { }

                // Get first layout profile ID
                ObjectId layoutProfileId = ObjectId.Null;
                try
                {
                    foreach (ObjectId profileId in alignment.GetProfileIds())
                    {
                        using (Profile profile = profileId.GetObject(OpenMode.ForRead) as Profile)
                        {
                            if (profile != null && profile.Entities != null && profile.Entities.Count > 0)
                            {
                                layoutProfileId = profileId;
                                break;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    AcApp.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError getting profile IDs: {ex.Message}");
                    return;
                }

                if (layoutProfileId == ObjectId.Null)
                {
                    AcApp.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nNo layout profile found.");
                    return;
                }

                // Open the profile and generate report
                using (Profile layoutProfile = layoutProfileId.GetObject(OpenMode.ForRead) as Profile)
                {
                    if (layoutProfile == null)
                    {
                        AcApp.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nError accessing profile.");
                        return;
                    }

                    // Get profile properties safely
                    string profileName = "";
                    string profileDescription = "";
                    string profileStyle = "Default";
                    try
                    {
                        profileName = layoutProfile.Name ?? "";
                    }
                    catch { }
                    try
                    {
                        profileDescription = layoutProfile.Description ?? "";
                    }
                    catch { }
                    try
                    {
                        profileStyle = layoutProfile.StyleName ?? "Default";
                    }
                    catch { }

                    using (System.IO.StreamWriter writer = new System.IO.StreamWriter(outputPath.Replace(".pdf", ".txt")))
                    {
                        // Write header
                        writer.WriteLine($"Project Name: {projectName}");
                        writer.WriteLine($" Description:");
                        writer.WriteLine($"Horizontal Alignment Name: {alignmentName}");
                        writer.WriteLine($" Description: {alignmentDescription}");
                        writer.WriteLine($" Style: {alignmentStyle}");
                        writer.WriteLine($"Vertical Alignment Name: {profileName}");
                        writer.WriteLine($" Description: {profileDescription}");
                        writer.WriteLine($" Style: {profileStyle}");
                        writer.WriteLine($" {"STATION",15} {"ELEVATION",15}");
                        writer.WriteLine();

                        // Process profile entities
                        if (layoutProfile.Entities != null)
                        {
                            for (int i = 0; i < layoutProfile.Entities.Count; i++)
                            {
                                try
                                {
                                    ProfileEntity entity = layoutProfile.Entities[i];
                                    if (entity == null) continue;

                                    switch (entity.EntityType)
                                    {
                                        case ProfileEntityType.Tangent:
                                            if (entity is ProfileTangent tangent)
                                                WriteProfileTangent(writer, tangent, i, layoutProfile.Entities.Count);
                                            else
                                                WriteUnsupportedProfileEntity(writer, entity);
                                            break;
                                        case ProfileEntityType.Circular:
                                            if (entity is ProfileCircular circular)
                                                WriteProfileParabola(writer, circular);
                                            else
                                                WriteUnsupportedProfileEntity(writer, entity);
                                            break;
                                        default:
                                            WriteUnsupportedProfileEntity(writer, entity);
                                            break;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    writer.WriteLine($"Error processing entity {i}: {ex.Message}");
                                    writer.WriteLine();
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GenerateVerticalReport: {ex.Message}", ex);
            }
        }

        private void WriteProfileTangent(System.IO.StreamWriter writer, ProfileTangent tangent, int index, int totalCount)
        {
            try
            {
                if (tangent == null)
                {
                    writer.WriteLine("Element: Linear");
                    writer.WriteLine("Unable to read tangent data for this profile element.");
                    writer.WriteLine();
                    return;
                }

                writer.WriteLine("Element: Linear");
                writer.WriteLine();

                // Get properties safely
                double startStation = 0;
                double startElevation = 0;
                double endStation = 0;
                double endElevation = 0;
                double grade = 0;
                double length = 0;

                try { startStation = tangent.StartStation; } catch { }
                try { startElevation = tangent.StartElevation; } catch { }
                try { endStation = tangent.EndStation; } catch { }
                try { endElevation = tangent.EndElevation; } catch { }
                try { grade = tangent.Grade; } catch { }
                try { length = tangent.Length; } catch { }

                // Determine point labels
                if (index == 0)
                    writer.WriteLine($" POB {FormatStation(startStation),15} {startElevation,15:F2}");

                writer.WriteLine($" PVI {FormatStation(endStation),15} {endElevation,15:F2}");
                writer.WriteLine($" Tangent Grade: {grade * 100,15:F3}");
                writer.WriteLine($" Tangent Length: {length,15:F2}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Linear");
                writer.WriteLine($"Error writing tangent data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteProfileParabola(System.IO.StreamWriter writer, ProfileCircular curve)
        {
            try
            {
                if (curve == null)
                {
                    writer.WriteLine("Element: Parabola");
                    writer.WriteLine("Unable to read curve data for this profile element.");
                    writer.WriteLine();
                    return;
                }

                const double tolerance = 1e-8;

                // Get properties safely
                double gradeIn = 0;
                double gradeOut = 0;
                double length = 0;
                double startStation = 0;
                double endStation = 0;
                double pviStation = 0;
                double pviElevation = 0;
                double startElevation = 0;
                double endElevation = 0;

                try { gradeIn = curve.GradeIn; } catch { }
                try { gradeOut = curve.GradeOut; } catch { }
                try { length = curve.Length; } catch { }
                try { startStation = curve.StartStation; } catch { }
                try { endStation = curve.EndStation; } catch { }
                try { pviStation = curve.PVIStation; } catch { }
                try { pviElevation = curve.PVIElevation; } catch { }
                
                // Try to get elevations, but calculate if not available
                try { startElevation = curve.StartElevation; } catch 
                { 
                    // Calculate if property doesn't exist
                    startElevation = 0;
                }
                try { endElevation = curve.EndElevation; } catch 
                { 
                    // Calculate if property doesn't exist
                    endElevation = 0;
                }

                gradeIn *= 100;
                gradeOut *= 100;
                double gradeDiff = gradeOut - gradeIn;
                double r = Math.Abs(length) > tolerance ? gradeDiff / length : 0;
                double k = Math.Abs(gradeDiff) > tolerance ? length / gradeDiff : double.PositiveInfinity;
                double middleOrdinate = Math.Abs(r * length * length / 800);
                
                // Calculate PVC and PVT elevations from PVI
                double gradeInDecimal = gradeIn / 100.0;
                double gradeOutDecimal = gradeOut / 100.0;
                double pvcElevation = pviElevation - (gradeInDecimal * (length / 2));
                double pvtElevation = endElevation;
                if (Math.Abs(pvtElevation) < tolerance && Math.Abs(pviElevation) > tolerance)
                {
                    // Calculate PVT if not available
                    pvtElevation = pviElevation + (gradeOutDecimal * (length / 2));
                }

                writer.WriteLine("Element: Parabola");
                writer.WriteLine($" PVC {FormatStation(startStation),15} {pvcElevation,15:F2}");
                writer.WriteLine($" PVI {FormatStation(pviStation),15} {pviElevation,15:F2}");
                writer.WriteLine($" PVT {FormatStation(endStation),15} {pvtElevation,15:F2}");
                writer.WriteLine($" Length: {length,15:F2}");

                double gradeChange = Math.Abs(gradeDiff);
                if (gradeChange > 1.0)
                    writer.WriteLine($" Stopping Sight Distance: {571.52,15:F2}");
                else
                    writer.WriteLine($" Headlight Sight Distance: {540.41,15:F2}");

                writer.WriteLine($" Entrance Grade: {gradeIn,15:F3}");
                writer.WriteLine($" Exit Grade: {gradeOut,15:F3}");
                writer.WriteLine($" r = ( g2 - g1 ) / L: {r,15:F3}");
                string kDisplay = Math.Abs(gradeDiff) > tolerance ? Math.Abs(k).ToString("F3") : "INF";
                writer.WriteLine($" K = l / ( g2 - g1 ): {kDisplay,15}");
                writer.WriteLine($" Middle Ordinate: {middleOrdinate,15:F2}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Parabola");
                writer.WriteLine($"Error writing curve data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteUnsupportedProfileEntity(System.IO.StreamWriter writer, ProfileEntity entity)
        {
            writer.WriteLine($"Element: {entity?.EntityType.ToString() ?? "Unknown"}");
            writer.WriteLine("Unsupported profile entity type encountered.");
            writer.WriteLine();
        }

        /// <summary>
        /// Generate horizontal alignment report
        /// </summary>
        private void GenerateHorizontalReport(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                // Get project name from drawing properties
                Database db = alignment.Database;
                string projectName = "";
                try
                {
                    if (db.Filename != null && db.Filename.Length > 0)
                    {
                        projectName = System.IO.Path.GetFileNameWithoutExtension(db.Filename);
                    }
                }
                catch
                {
                    projectName = "Unknown Project";
                }

                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(outputPath.Replace(".pdf", ".txt")))
                {
                    // Write header
                    writer.WriteLine($"Project Name: {projectName}");
                    writer.WriteLine($" Description:");
                    writer.WriteLine($"Horizontal Alignment Name: {alignment.Name}");
                    writer.WriteLine($" Description: {alignment.Description ?? ""}");
                    writer.WriteLine($" Style: {alignment.StyleName ?? "Default"}");
                    writer.WriteLine($" {"STATION",15} {"NORTHING",15} {"EASTING",15}");
                    writer.WriteLine();

                    // Process each entity
                    for (int i = 0; i < alignment.Entities.Count; i++)
                    {
                        AlignmentEntity entity = alignment.Entities[i];
                        if (entity == null) continue;

                        try
                        {
                            switch (entity.EntityType)
                            {
                                case AlignmentEntityType.Line:
                                    WriteLinearElement(writer, entity as AlignmentLine, alignment, i);
                                    break;
                                case AlignmentEntityType.Arc:
                                    WriteArcElement(writer, entity as AlignmentArc, alignment, i);
                                    break;
                                case AlignmentEntityType.Spiral:
                                    WriteSpiralElement(writer, entity as AlignmentSpiral, alignment, i);
                                    break;
                                default:
                                    writer.WriteLine($"Element: {entity.EntityType}");
                                    writer.WriteLine("Unsupported element type.");
                                    writer.WriteLine();
                                    break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            writer.WriteLine($"Error processing entity {i}: {ex.Message}");
                            writer.WriteLine();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GenerateHorizontalReport: {ex.Message}", ex);
            }
        }

        private void WriteLinearElement(System.IO.StreamWriter writer, AlignmentLine line, CivDb.Alignment alignment, int index)
        {
            if (line == null)
            {
                writer.WriteLine("Element: Linear");
                writer.WriteLine("Unable to read line data for this alignment element.");
                writer.WriteLine();
                return;
            }

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(line.EndStation, 0, 0, ref x2, ref y2, ref z2);

                string bearing = FormatBearing(line.Direction);

                writer.WriteLine("Element: Linear");

                // Determine point labels based on position
                if (index == 0)
                    writer.WriteLine($" POT ( ) {FormatStation(line.StartStation),15} {y1,15:F4} {x1,15:F4}");
                else
                    writer.WriteLine($" PI  ( ) {FormatStation(line.StartStation),15} {y1,15:F4} {x1,15:F4}");

                // Check if next element exists to determine end point label
                string endLabel = "PI  ";
                if (index < alignment.Entities.Count - 1)
                {
                    var nextEntity = alignment.Entities[index + 1];
                    if (nextEntity.EntityType == AlignmentEntityType.Spiral)
                        endLabel = "TS  ";
                    else if (nextEntity.EntityType == AlignmentEntityType.Arc)
                        endLabel = "PC  ";
                }
                else
                    endLabel = "POT ";

                writer.WriteLine($" {endLabel}( ) {FormatStation(line.EndStation),15} {y2,15:F4} {x2,15:F4}");
                writer.WriteLine($" Tangent Direction: {bearing}");
                writer.WriteLine($" Tangent Length: {line.Length,15:F4}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Linear");
                writer.WriteLine($"Error writing line data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteArcElement(System.IO.StreamWriter writer, AlignmentArc arc, CivDb.Alignment alignment, int index)
        {
            if (arc == null)
            {
                writer.WriteLine("Element: Circular");
                writer.WriteLine("Unable to read arc data for this alignment element.");
                writer.WriteLine();
                return;
            }

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                double xc = 0, yc = 0, zc = 0;
                double xMid = 0, yMid = 0, zMid = 0;

                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);
                alignment.PointLocation((arc.StartStation + arc.EndStation) / 2, 0, 0, ref xMid, ref yMid, ref zMid);

                // Calculate center point
                double midStation = (arc.StartStation + arc.EndStation) / 2;
                double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);

                double deltaRadians = arc.Delta;
                double deltaDegrees = deltaRadians * (180.0 / Math.PI);
                double chord = 2 * arc.Radius * Math.Sin(Math.Abs(deltaRadians) / 2);
                double middleOrdinate = arc.Radius * (1 - Math.Cos(Math.Abs(deltaRadians) / 2));
                double external = arc.Radius * (1 / Math.Cos(Math.Abs(deltaRadians) / 2) - 1);
                double tangent = arc.Radius * Math.Tan(Math.Abs(deltaRadians) / 2);

                writer.WriteLine("Element: Circular");
                writer.WriteLine($" SC  ( ) {FormatStation(arc.StartStation),15} {y1,15:F4} {x1,15:F4}");
                writer.WriteLine($" PI  ( ) {FormatStation(midStation),15} {yMid,15:F4} {xMid,15:F4}");
                writer.WriteLine($" CC  ( ) {yc,32:F4} {xc,15:F4}");
                writer.WriteLine($" CS  ( ) {FormatStation(arc.EndStation),15} {y2,15:F4} {x2,15:F4}");
                writer.WriteLine($" Radius: {arc.Radius,15:F4}");
                writer.WriteLine($" Design Speed(mph): {50.0,15:F4}");
                writer.WriteLine($" Cant(inches): {2.0,15:F3}");
                writer.WriteLine($" Delta: {FormatAngle(Math.Abs(deltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}");
                writer.WriteLine($"Degree of Curvature(Chord): {FormatAngle(5729.58 / arc.Radius)}");
                writer.WriteLine($" Length: {arc.Length,15:F4}");
                writer.WriteLine($" Length(Chorded): {arc.Length,15:F4}");
                writer.WriteLine($" Tangent: {tangent,15:F4}");
                writer.WriteLine($" Chord: {chord,15:F4}");
                writer.WriteLine($" Middle Ordinate: {middleOrdinate,15:F4}");
                writer.WriteLine($" External: {external,15:F4}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Circular");
                writer.WriteLine($"Error writing arc data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteSpiralElement(System.IO.StreamWriter writer, AlignmentSpiral spiral, CivDb.Alignment alignment, int index)
        {
            if (spiral == null)
            {
                writer.WriteLine("Element: Clothoid");
                writer.WriteLine("Unable to read spiral data for this alignment element.");
                writer.WriteLine();
                return;
            }

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                double xMid = 0, yMid = 0, zMid = 0;

                alignment.PointLocation(spiral.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(spiral.EndStation, 0, 0, ref x2, ref y2, ref z2);
                alignment.PointLocation((spiral.StartStation + spiral.EndStation) / 2, 0, 0, ref xMid, ref yMid, ref zMid);

                bool isEntry = spiral.RadiusIn > spiral.RadiusOut || (spiral.RadiusIn == 0 && spiral.RadiusOut > 0);
                double radiusIn = spiral.RadiusIn > 0 ? spiral.RadiusIn : 0;
                double radiusOut = spiral.RadiusOut > 0 ? spiral.RadiusOut : 0;

                writer.WriteLine("Element: Clothoid");

                if (isEntry)
                {
                    writer.WriteLine($" TS  ( ) {FormatStation(spiral.StartStation),15} {y1,15:F4} {x1,15:F4}");
                    writer.WriteLine($" SPI ( ) {FormatStation((spiral.StartStation + spiral.EndStation) / 2),15} {yMid,15:F4} {xMid,15:F4}");
                    writer.WriteLine($" SC  ( ) {FormatStation(spiral.EndStation),15} {y2,15:F4} {x2,15:F4}");
                }
                else
                {
                    writer.WriteLine($" CS  ( ) {FormatStation(spiral.StartStation),15} {y1,15:F4} {x1,15:F4}");
                    writer.WriteLine($" SPI ( ) {FormatStation((spiral.StartStation + spiral.EndStation) / 2),15} {yMid,15:F4} {xMid,15:F4}");
                    writer.WriteLine($" ST  ( ) {FormatStation(spiral.EndStation),15} {y2,15:F4} {x2,15:F4}");
                }

                writer.WriteLine($" Entrance Radius: {radiusIn,15:F4}");
                writer.WriteLine($" Exit Radius: {radiusOut,15:F4}");
                writer.WriteLine($" Length: {spiral.Length,15:F4}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Clothoid");
                writer.WriteLine($"Error writing spiral data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private string FormatStation(double station)
        {
            int sta = (int)(station / 100);
            double offset = station - (sta * 100);
            return $"{sta}+{offset:F2}";
        }

        private string FormatBearing(double radians)
        {
            double degrees = radians * (180.0 / Math.PI);
            while (degrees < 0) degrees += 360;
            while (degrees >= 360) degrees -= 360;

            string quadrant;
            double angle;

            if (degrees >= 0 && degrees < 90)
            {
                quadrant = "N";
                angle = degrees;
                return $"{quadrant} {FormatAngle(angle)} E";
            }
            else if (degrees >= 90 && degrees < 180)
            {
                quadrant = "S";
                angle = 180 - degrees;
                return $"{quadrant} {FormatAngle(angle)} E";
            }
            else if (degrees >= 180 && degrees < 270)
            {
                quadrant = "S";
                angle = degrees - 180;
                return $"{quadrant} {FormatAngle(angle)} W";
            }
            else
            {
                quadrant = "N";
                angle = 360 - degrees;
                return $"{quadrant} {FormatAngle(angle)} W";
            }
        }

        private string FormatAngle(double degrees)
        {
            int deg = (int)degrees;
            double remaining = (degrees - deg) * 60;
            int min = (int)remaining;
            double sec = (remaining - min) * 60;
            return $"{deg}^{min:D2}'{sec:F4}\"";
        }

    }

    /// <summary>
    /// Simple report selection form
    /// </summary>
    public class ReportSelectionForm : Form
    {
        public string ReportType { get; private set; }
        public string OutputPath { get; private set; }

        private RadioButton rbVertical;
        private RadioButton rbHorizontal;
        private TextBox txtOutputPath;
        private Button btnBrowse;
        private Button btnOK;
        private Button btnCancel;

        public ReportSelectionForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Generate GeoTable Report";
            this.Size = new System.Drawing.Size(500, 200);
            this.StartPosition = FormStartPosition.CenterScreen;

            var lblType = new System.Windows.Forms.Label { Text = "Report Type:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            this.Controls.Add(lblType);

            rbVertical = new RadioButton { Text = "Vertical Alignment", Location = new System.Drawing.Point(40, 45), Checked = true, AutoSize = true };
            rbHorizontal = new RadioButton { Text = "Horizontal Alignment", Location = new System.Drawing.Point(40, 70), AutoSize = true };
            this.Controls.Add(rbVertical);
            this.Controls.Add(rbHorizontal);

            var lblOutput = new System.Windows.Forms.Label { Text = "Output Path:", Location = new System.Drawing.Point(20, 100), AutoSize = true };
            this.Controls.Add(lblOutput);

            txtOutputPath = new TextBox { Location = new System.Drawing.Point(20, 120), Width = 350 };
            txtOutputPath.Text = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "alignment_report.pdf"
            );
            this.Controls.Add(txtOutputPath);

            btnBrowse = new Button { Text = "Browse...", Location = new System.Drawing.Point(380, 118), Width = 80 };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(300, 130), Width = 75, DialogResult = DialogResult.OK };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(385, 130), Width = 75, DialogResult = DialogResult.Cancel };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF Files (*.pdf)|*.pdf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.FileName = "alignment_report.pdf";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtOutputPath.Text = dialog.FileName;
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            ReportType = rbVertical.Checked ? "Vertical" : "Horizontal";
            OutputPath = txtOutputPath.Text;
        }
    }

    /// <summary>
    /// Batch process form
    /// </summary>
    public class BatchProcessForm : Form
    {
        public string OutputFolder { get; private set; }
        public bool IncludeVertical { get; private set; }
        public bool IncludeHorizontal { get; private set; }

        private TextBox txtOutputFolder;
        private CheckBox chkVertical;
        private CheckBox chkHorizontal;
        private Button btnBrowse;
        private Button btnOK;
        private Button btnCancel;

        public BatchProcessForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Batch Process Alignments";
            this.Size = new System.Drawing.Size(500, 200);
            this.StartPosition = FormStartPosition.CenterScreen;

            var lblFolder = new System.Windows.Forms.Label { Text = "Output Folder:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            this.Controls.Add(lblFolder);

            txtOutputFolder = new TextBox { Location = new System.Drawing.Point(20, 40), Width = 350 };
            txtOutputFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            this.Controls.Add(txtOutputFolder);

            btnBrowse = new Button { Text = "Browse...", Location = new System.Drawing.Point(380, 38), Width = 80 };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            var lblTypes = new System.Windows.Forms.Label { Text = "Report Types:", Location = new System.Drawing.Point(20, 75), AutoSize = true };
            this.Controls.Add(lblTypes);

            chkVertical = new CheckBox { Text = "Vertical Alignment", Location = new System.Drawing.Point(40, 100), Checked = true, AutoSize = true };
            chkHorizontal = new CheckBox { Text = "Horizontal Alignment", Location = new System.Drawing.Point(40, 125), Checked = true, AutoSize = true };
            this.Controls.Add(chkVertical);
            this.Controls.Add(chkHorizontal);

            btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(300, 130), Width = 75, DialogResult = DialogResult.OK };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(385, 130), Width = 75, DialogResult = DialogResult.Cancel };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select output folder for reports";
                dialog.SelectedPath = txtOutputFolder.Text;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtOutputFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            OutputFolder = txtOutputFolder.Text;
            IncludeVertical = chkVertical.Checked;
            IncludeHorizontal = chkHorizontal.Checked;
        }
    }
}
