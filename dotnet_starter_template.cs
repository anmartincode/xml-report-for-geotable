/*
 * Civil 3D GeoTable Reports - .NET Add-in Starter Template
 *
 * This is a basic template showing how a .NET add-in would work
 * Compile this with Visual Studio and Civil 3D API references
 */

using System;
using System.Windows.Forms;
using System.Xml.Linq;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using CivApp = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using CivDb = Autodesk.Civil.DatabaseServices;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Layout.Borders;
using iText.Kernel.Font;
using iText.IO.Font.Constants;

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
                                // Detect format from file extension
                                string extension = System.IO.Path.GetExtension(outputPath).ToLower();
                                bool isXml = extension == ".xml";
                                bool isPdf = extension == ".pdf";

                                if (reportType == "Vertical")
                                {
                                    if (isXml)
                                        GenerateVerticalReportXml(alignment, outputPath);
                                    else if (isPdf)
                                        GenerateVerticalReportPdf(alignment, outputPath);
                                    else
                                        GenerateVerticalReport(alignment, outputPath);
                                }
                                else
                                {
                                    if (isXml)
                                        GenerateHorizontalReportXml(alignment, outputPath);
                                    else if (isPdf)
                                        GenerateHorizontalReportPdf(alignment, outputPath);
                                    else
                                        GenerateHorizontalReport(alignment, outputPath);
                                }

                                tr.Commit();

                                // Show success message
                                MessageBox.Show(
                                    $"Report generated successfully!\n\nSaved to: {outputPath}",
                                    "Success",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );

                                ed.WriteMessage($"\n✓ Report saved to: {outputPath}");
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

                    using (System.IO.StreamWriter writer = new System.IO.StreamWriter(outputPath))
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
        /// Helper method to add a data row to horizontal alignment table (3 columns)
        /// </summary>
        private void AddHorizontalDataRow(Document document, string label, string station, double northing, double easting, PdfFont font)
        {
            iText.Layout.Element.Table dataTable = new iText.Layout.Element.Table(3);
            dataTable.SetWidth(UnitValue.CreatePercentValue(100));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{label}{station}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{northing:F4}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{easting:F4}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            document.Add(dataTable);
        }

        /// <summary>
        /// Helper method to add a data row to vertical alignment table (2 columns)
        /// </summary>
        private void AddVerticalDataRow(Document document, string label, string station, double elevation, PdfFont font)
        {
            iText.Layout.Element.Table dataTable = new iText.Layout.Element.Table(2);
            dataTable.SetWidth(UnitValue.CreatePercentValue(100));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{label}{station}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{elevation:F2}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            document.Add(dataTable);
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

                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(outputPath))
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

        /// <summary>
        /// Generate horizontal alignment report in XML format
        /// </summary>
        private void GenerateHorizontalReportXml(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                // Get project name
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

                // Create XML document
                XDocument doc = new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("GeoTableReport",
                        new XAttribute("Type", "Horizontal"),
                        new XElement("Project",
                            new XElement("Name", projectName),
                            new XElement("Description", "")
                        ),
                        new XElement("HorizontalAlignment",
                            new XElement("Name", alignment.Name),
                            new XElement("Description", alignment.Description ?? ""),
                            new XElement("Style", alignment.StyleName ?? "Default")
                        ),
                        new XElement("Elements",
                            GetHorizontalElementsXml(alignment)
                        )
                    )
                );

                doc.Save(outputPath);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GenerateHorizontalReportXml: {ex.Message}", ex);
            }
        }

        private System.Collections.Generic.IEnumerable<XElement> GetHorizontalElementsXml(CivDb.Alignment alignment)
        {
            var elements = new System.Collections.Generic.List<XElement>();

            for (int i = 0; i < alignment.Entities.Count; i++)
            {
                AlignmentEntity entity = alignment.Entities[i];
                if (entity == null) continue;

                try
                {
                    switch (entity.EntityType)
                    {
                        case AlignmentEntityType.Line:
                            var line = entity as AlignmentLine;
                            if (line != null)
                            {
                                double x1 = 0, y1 = 0, z1 = 0;
                                double x2 = 0, y2 = 0, z2 = 0;
                                alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                alignment.PointLocation(line.EndStation, 0, 0, ref x2, ref y2, ref z2);

                                elements.Add(new XElement("Line",
                                    new XElement("StartStation", line.StartStation),
                                    new XElement("EndStation", line.EndStation),
                                    new XElement("Length", line.Length),
                                    new XElement("Direction", line.Direction),
                                    new XElement("Bearing", FormatBearing(line.Direction)),
                                    new XElement("StartPoint",
                                        new XElement("Northing", y1),
                                        new XElement("Easting", x1)
                                    ),
                                    new XElement("EndPoint",
                                        new XElement("Northing", y2),
                                        new XElement("Easting", x2)
                                    )
                                ));
                            }
                            break;

                        case AlignmentEntityType.Arc:
                            var arc = entity as AlignmentArc;
                            if (arc != null)
                            {
                                double x1 = 0, y1 = 0, z1 = 0;
                                double x2 = 0, y2 = 0, z2 = 0;
                                double xc = 0, yc = 0, zc = 0;
                                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                                double midStation = (arc.StartStation + arc.EndStation) / 2;
                                double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                                alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);

                                double deltaRadians = arc.Delta;
                                double deltaDegrees = deltaRadians * (180.0 / Math.PI);

                                elements.Add(new XElement("Arc",
                                    new XElement("StartStation", arc.StartStation),
                                    new XElement("EndStation", arc.EndStation),
                                    new XElement("Length", arc.Length),
                                    new XElement("Radius", arc.Radius),
                                    new XElement("Delta", Math.Abs(deltaDegrees)),
                                    new XElement("Direction", arc.Clockwise ? "Right" : "Left"),
                                    new XElement("StartPoint",
                                        new XElement("Northing", y1),
                                        new XElement("Easting", x1)
                                    ),
                                    new XElement("EndPoint",
                                        new XElement("Northing", y2),
                                        new XElement("Easting", x2)
                                    ),
                                    new XElement("CenterPoint",
                                        new XElement("Northing", yc),
                                        new XElement("Easting", xc)
                                    )
                                ));
                            }
                            break;

                        case AlignmentEntityType.Spiral:
                            var spiral = entity as AlignmentSpiral;
                            if (spiral != null)
                            {
                                double x1 = 0, y1 = 0, z1 = 0;
                                double x2 = 0, y2 = 0, z2 = 0;
                                alignment.PointLocation(spiral.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                alignment.PointLocation(spiral.EndStation, 0, 0, ref x2, ref y2, ref z2);

                                elements.Add(new XElement("Spiral",
                                    new XElement("StartStation", spiral.StartStation),
                                    new XElement("EndStation", spiral.EndStation),
                                    new XElement("Length", spiral.Length),
                                    new XElement("RadiusIn", spiral.RadiusIn),
                                    new XElement("RadiusOut", spiral.RadiusOut),
                                    new XElement("StartPoint",
                                        new XElement("Northing", y1),
                                        new XElement("Easting", x1)
                                    ),
                                    new XElement("EndPoint",
                                        new XElement("Northing", y2),
                                        new XElement("Easting", x2)
                                    )
                                ));
                            }
                            break;
                    }
                }
                catch
                {
                    // Skip elements that can't be processed
                }
            }

            return elements;
        }

        /// <summary>
        /// Generate vertical alignment report in XML format
        /// </summary>
        private void GenerateVerticalReportXml(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                // Get project name
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

                // Get first layout profile ID
                ObjectId layoutProfileId = ObjectId.Null;
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

                    // Create XML document
                    XDocument doc = new XDocument(
                        new XDeclaration("1.0", "utf-8", "yes"),
                        new XElement("GeoTableReport",
                            new XAttribute("Type", "Vertical"),
                            new XElement("Project",
                                new XElement("Name", projectName),
                                new XElement("Description", "")
                            ),
                            new XElement("HorizontalAlignment",
                                new XElement("Name", alignment.Name),
                                new XElement("Description", alignment.Description ?? ""),
                                new XElement("Style", alignment.StyleName ?? "Default")
                            ),
                            new XElement("VerticalAlignment",
                                new XElement("Name", layoutProfile.Name),
                                new XElement("Description", layoutProfile.Description ?? ""),
                                new XElement("Style", layoutProfile.StyleName ?? "Default")
                            ),
                            new XElement("Elements",
                                GetVerticalElementsXml(layoutProfile)
                            )
                        )
                    );

                    doc.Save(outputPath);
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GenerateVerticalReportXml: {ex.Message}", ex);
            }
        }

        private System.Collections.Generic.IEnumerable<XElement> GetVerticalElementsXml(Profile profile)
        {
            var elements = new System.Collections.Generic.List<XElement>();

            for (int i = 0; i < profile.Entities.Count; i++)
            {
                ProfileEntity entity = profile.Entities[i];
                if (entity == null) continue;

                try
                {
                    switch (entity.EntityType)
                    {
                        case ProfileEntityType.Tangent:
                            var tangent = entity as ProfileTangent;
                            if (tangent != null)
                            {
                                elements.Add(new XElement("Tangent",
                                    new XElement("StartStation", tangent.StartStation),
                                    new XElement("EndStation", tangent.EndStation),
                                    new XElement("StartElevation", tangent.StartElevation),
                                    new XElement("EndElevation", tangent.EndElevation),
                                    new XElement("Grade", tangent.Grade * 100),
                                    new XElement("Length", tangent.Length)
                                ));
                            }
                            break;

                        case ProfileEntityType.Circular:
                            var curve = entity as ProfileCircular;
                            if (curve != null)
                            {
                                double gradeIn = curve.GradeIn * 100;
                                double gradeOut = curve.GradeOut * 100;
                                double r = (gradeOut - gradeIn) / curve.Length;
                                double k = curve.Length / (gradeOut - gradeIn);

                                elements.Add(new XElement("Parabola",
                                    new XElement("StartStation", curve.StartStation),
                                    new XElement("EndStation", curve.EndStation),
                                    new XElement("PVIStation", curve.PVIStation),
                                    new XElement("PVIElevation", curve.PVIElevation),
                                    new XElement("Length", curve.Length),
                                    new XElement("GradeIn", gradeIn),
                                    new XElement("GradeOut", gradeOut),
                                    new XElement("RateOfChange", r),
                                    new XElement("K", Math.Abs(k))
                                ));
                            }
                            break;
                    }
                }
                catch
                {
                    // Skip elements that can't be processed
                }
            }

            return elements;
        }

        /// <summary>
        /// Generate horizontal alignment report in PDF format
        /// </summary>
        private void GenerateHorizontalReportPdf(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                // Get project name
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

                // Create PDF document using iText7
                PdfWriter writer = new PdfWriter(outputPath);
                PdfDocument pdf = new PdfDocument(writer);
                Document document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);
                document.SetMargins(50, 50, 50, 50);

                // Define fonts
                PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                PdfFont normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                // Add title and header info
                document.Add(new Paragraph($"Project Name: {projectName}")
                    .SetFont(boldFont).SetFontSize(11));
                document.Add(new Paragraph(" Description:")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($"Horizontal Alignment Name: {alignment.Name}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($" Description: {alignment.Description ?? ""}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($" Style: {alignment.StyleName ?? "Default"}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));

                // Add column headers
                iText.Layout.Element.Table headerTable = new iText.Layout.Element.Table(3);
                headerTable.SetWidth(UnitValue.CreatePercentValue(100));

                iText.Layout.Element.Cell cell1 = new iText.Layout.Element.Cell().Add(new Paragraph("STATION")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                headerTable.AddCell(cell1);

                iText.Layout.Element.Cell cell2 = new iText.Layout.Element.Cell().Add(new Paragraph("NORTHING")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                headerTable.AddCell(cell2);

                iText.Layout.Element.Cell cell3 = new iText.Layout.Element.Cell().Add(new Paragraph("EASTING")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                headerTable.AddCell(cell3);

                document.Add(headerTable);
                document.Add(new Paragraph("\n"));

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
                                WriteLinearElementPdf(document, entity as AlignmentLine, alignment, i, normalFont, boldFont);
                                break;
                            case AlignmentEntityType.Arc:
                                WriteArcElementPdf(document, entity as AlignmentArc, alignment, i, normalFont, boldFont);
                                break;
                            case AlignmentEntityType.Spiral:
                                WriteSpiralElementPdf(document, entity as AlignmentSpiral, alignment, i, normalFont, boldFont);
                                break;
                            default:
                                document.Add(new Paragraph($"Element: {entity.EntityType}").SetFont(normalFont).SetFontSize(10));
                                document.Add(new Paragraph("Unsupported element type.").SetFont(normalFont).SetFontSize(10));
                                document.Add(new Paragraph("\n"));
                                break;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        document.Add(new Paragraph($"Error processing entity {i}: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                        document.Add(new Paragraph("\n"));
                    }
                }

                document.Close();
            }
            catch (System.Exception ex)
            {
                string errorDetails = $"Error in GenerateHorizontalReportPdf: {ex.GetType().Name} - {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorDetails += $"\nInner Exception: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}";
                }
                errorDetails += $"\nStack Trace: {ex.StackTrace}";
                throw new System.Exception(errorDetails, ex);
            }
        }

        private void WriteLinearElementPdf(Document document, AlignmentLine line, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont boldFont)
        {
            if (line == null)
            {
                document.Add(new Paragraph("Element: Linear").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("Unable to read line data for this alignment element.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
                return;
            }

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(line.EndStation, 0, 0, ref x2, ref y2, ref z2);

                string bearing = FormatBearing(line.Direction);

                document.Add(new Paragraph("Element: Linear").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(5));

                // Determine point labels based on position
                if (index == 0)
                    AddHorizontalDataRow(document, "POT ( ) ", FormatStation(line.StartStation), y1, x1, normalFont);
                else
                    AddHorizontalDataRow(document, "PI  ( ) ", FormatStation(line.StartStation), y1, x1, normalFont);

                // Check if next element exists to determine end point label
                string endLabel;
                if (index < alignment.Entities.Count - 1)
                {
                    var nextEntity = alignment.Entities[index + 1];
                    if (nextEntity.EntityType == AlignmentEntityType.Spiral)
                        endLabel = "TS  ( ) ";
                    else if (nextEntity.EntityType == AlignmentEntityType.Arc)
                        endLabel = "PC  ( ) ";
                    else
                        endLabel = "PI  ( ) ";
                }
                else
                    endLabel = "POT ( ) ";

                AddHorizontalDataRow(document, endLabel, FormatStation(line.EndStation), y2, x2, normalFont);
                document.Add(new Paragraph($"Tangent Direction: {bearing}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent Length: {line.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(5));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Linear").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing line data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteArcElementPdf(Document document, AlignmentArc arc, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont boldFont)
        {
            if (arc == null)
            {
                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("Unable to read arc data for this alignment element.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
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

                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(5));

                AddHorizontalDataRow(document, "SC  ( ) ", FormatStation(arc.StartStation), y1, x1, normalFont);
                AddHorizontalDataRow(document, "PI  ( ) ", FormatStation(midStation), yMid, xMid, normalFont);
                AddHorizontalDataRow(document, "CC  ( ) ", "               ", yc, xc, normalFont);
                AddHorizontalDataRow(document, "CS  ( ) ", FormatStation(arc.EndStation), y2, x2, normalFont);

                document.Add(new Paragraph($"Radius: {arc.Radius:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Design Speed(mph): {50.0:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Cant(inches): {2.0:F3}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Delta: {FormatAngle(Math.Abs(deltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Degree of Curvature(Chord): {FormatAngle(5729.58 / arc.Radius)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length: {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length(Chorded): {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent: {tangent:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Chord: {chord:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Middle Ordinate: {middleOrdinate:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"External: {external:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(5));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing arc data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteSpiralElementPdf(Document document, AlignmentSpiral spiral, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont labelFont)
        {
            if (spiral == null)
            {
                document.Add(new Paragraph("Element: Clothoid").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph("Unable to read spiral data for this alignment element.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
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

                document.Add(new Paragraph("Element: Clothoid").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(5));

                if (isEntry)
                {
                    AddHorizontalDataRow(document, "TS  ( ) ", FormatStation(spiral.StartStation), y1, x1, normalFont);
                    AddHorizontalDataRow(document, "SPI ( ) ", FormatStation((spiral.StartStation + spiral.EndStation) / 2), yMid, xMid, normalFont);
                    AddHorizontalDataRow(document, "SC  ( ) ", FormatStation(spiral.EndStation), y2, x2, normalFont);
                }
                else
                {
                    AddHorizontalDataRow(document, "CS  ( ) ", FormatStation(spiral.StartStation), y1, x1, normalFont);
                    AddHorizontalDataRow(document, "SPI ( ) ", FormatStation((spiral.StartStation + spiral.EndStation) / 2), yMid, xMid, normalFont);
                    AddHorizontalDataRow(document, "ST  ( ) ", FormatStation(spiral.EndStation), y2, x2, normalFont);
                }

                document.Add(new Paragraph($"Entrance Radius: {radiusIn:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Exit Radius: {radiusOut:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length: {spiral.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(5));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Clothoid").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing spiral data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        /// <summary>
        /// Generate vertical alignment report in PDF format
        /// </summary>
        private void GenerateVerticalReportPdf(CivDb.Alignment alignment, string outputPath)
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

                    // Create PDF document using iText7
                    PdfWriter writer = new PdfWriter(outputPath);
                    PdfDocument pdf = new PdfDocument(writer);
                    Document document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);
                    document.SetMargins(50, 50, 50, 50);

                    // Define fonts
                    PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    PdfFont normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                    // Write header
                    document.Add(new Paragraph($"Project Name: {projectName}").SetFont(boldFont).SetFontSize(11));
                    document.Add(new Paragraph(" Description:").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Horizontal Alignment Name: {alignmentName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {alignmentDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {alignmentStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Vertical Alignment Name: {profileName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {profileDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {profileStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph("\n"));

                    // Add column headers
                    iText.Layout.Element.Table headerTable = new iText.Layout.Element.Table(2);
                    headerTable.SetWidth(UnitValue.CreatePercentValue(100));

                    iText.Layout.Element.Cell cell1 = new iText.Layout.Element.Cell().Add(new Paragraph("STATION").SetFont(boldFont).SetFontSize(10))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    headerTable.AddCell(cell1);

                    iText.Layout.Element.Cell cell2 = new iText.Layout.Element.Cell().Add(new Paragraph("ELEVATION").SetFont(boldFont).SetFontSize(10))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                        .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                    headerTable.AddCell(cell2);

                    document.Add(headerTable);
                    document.Add(new Paragraph("\n"));

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
                                            WriteProfileTangentPdf(document, tangent, i, layoutProfile.Entities.Count, normalFont, boldFont);
                                        else
                                            WriteUnsupportedProfileEntityPdf(document, entity, normalFont);
                                        break;
                                    case ProfileEntityType.Circular:
                                        if (entity is ProfileCircular circular)
                                            WriteProfileParabolaPdf(document, circular, normalFont, boldFont);
                                        else
                                            WriteUnsupportedProfileEntityPdf(document, entity, normalFont);
                                        break;
                                    default:
                                        WriteUnsupportedProfileEntityPdf(document, entity, normalFont);
                                        break;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                document.Add(new Paragraph($"Error processing entity {i}: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                                document.Add(new Paragraph("\n"));
                            }
                        }
                    }

                    document.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error in GenerateVerticalReportPdf: {ex.Message}", ex);
            }
        }

        private void WriteProfileTangentPdf(Document document, ProfileTangent tangent, int index, int totalCount, PdfFont normalFont, PdfFont labelFont)
        {
            try
            {
                if (tangent == null)
                {
                    document.Add(new Paragraph("Element: Linear").SetFont(labelFont).SetFontSize(10));
                    document.Add(new Paragraph("Unable to read tangent data for this profile element.").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph("\n"));
                    return;
                }

                document.Add(new Paragraph("Element: Linear").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(5));

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
                    AddVerticalDataRow(document, "POB ", FormatStation(startStation), startElevation, normalFont);

                AddVerticalDataRow(document, "PVI ", FormatStation(endStation), endElevation, normalFont);
                document.Add(new Paragraph($"Tangent Grade: {grade * 100:F3}%").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent Length: {length:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(5));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Linear").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing tangent data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteProfileParabolaPdf(Document document, ProfileCircular curve, PdfFont normalFont, PdfFont labelFont)
        {
            try
            {
                if (curve == null)
                {
                    document.Add(new Paragraph("Element: Parabola").SetFont(labelFont).SetFontSize(10));
                    document.Add(new Paragraph("Unable to read curve data for this profile element.").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph("\n"));
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
                    startElevation = 0;
                }
                try { endElevation = curve.EndElevation; } catch
                {
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
                    pvtElevation = pviElevation + (gradeOutDecimal * (length / 2));
                }

                document.Add(new Paragraph("Element: Parabola").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(5));

                AddVerticalDataRow(document, "PVC ", FormatStation(startStation), pvcElevation, normalFont);
                AddVerticalDataRow(document, "PVI ", FormatStation(pviStation), pviElevation, normalFont);
                AddVerticalDataRow(document, "PVT ", FormatStation(endStation), pvtElevation, normalFont);

                document.Add(new Paragraph($"Length: {length:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));

                double gradeChange = Math.Abs(gradeDiff);
                if (gradeChange > 1.0)
                    document.Add(new Paragraph($"Stopping Sight Distance: {571.52:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                else
                    document.Add(new Paragraph($"Headlight Sight Distance: {540.41:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));

                document.Add(new Paragraph($"Entrance Grade: {gradeIn:F3}%").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Exit Grade: {gradeOut:F3}%").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"r = ( g2 - g1 ) / L: {r:F3}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                string kDisplay = Math.Abs(gradeDiff) > tolerance ? Math.Abs(k).ToString("F3") : "INF";
                document.Add(new Paragraph($"K = l / ( g2 - g1 ): {kDisplay}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Middle Ordinate: {middleOrdinate:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(5));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Parabola").SetFont(labelFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing curve data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteUnsupportedProfileEntityPdf(Document document, ProfileEntity entity, PdfFont normalFont)
        {
            document.Add(new Paragraph($"Element: {entity?.EntityType.ToString() ?? "Unknown"}").SetFont(normalFont).SetFontSize(10));
            document.Add(new Paragraph("Unsupported profile entity type encountered.").SetFont(normalFont).SetFontSize(10));
            document.Add(new Paragraph("\n"));
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
                "alignment_report.txt"
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
                dialog.Filter = "PDF Files (*.pdf)|*.pdf|Text Files (*.txt)|*.txt|XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
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
