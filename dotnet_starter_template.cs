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
using Autodesk.AutoCAD.Windows;
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
using iText.Kernel.Colors;

[assembly: CommandClass(typeof(GeoTableReports.ReportCommands))]

namespace GeoTableReports
{
    public class ReportCommands : IExtensionApplication
    {
        private static PaletteSet paletteSet;
        private static ReportPanelControl reportPanel;

        // Called when Civil 3D starts up
        public void Initialize()
        {
            AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nGeoTable Reports loaded. Type GEOTABLE_PANEL to open the panel or GEOTABLE for quick generation.");
            }
        }

        // Called when Civil 3D shuts down
        public void Terminate()
        {
            // Cleanup palette
            if (paletteSet != null)
            {
                paletteSet.Close();
                paletteSet = null;
            }
        }

        /// <summary>
        /// Command to show the GeoTable Reports panel
        /// </summary>
        [CommandMethod("GEOTABLE_PANEL")]
        public void ShowReportPanel()
        {
            if (paletteSet == null)
            {
                // Create the palette set
                paletteSet = new PaletteSet("GeoTable Reports", new System.Guid("A1B2C3D4-E5F6-4A5B-9C8D-1E2F3A4B5C6D"));

                // Create the user control
                reportPanel = new ReportPanelControl();

                // Add the control to the palette
                paletteSet.Add("Report Generator", reportPanel);

                // Set palette properties
                paletteSet.MinimumSize = new System.Drawing.Size(350, 500);
                paletteSet.DockEnabled = (DockSides.Left | DockSides.Right);
                paletteSet.Style = PaletteSetStyles.ShowCloseButton |
                                   PaletteSetStyles.ShowAutoHideButton |
                                   PaletteSetStyles.Snappable;
            }

            // Show the palette
            paletteSet.Visible = true;
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
                        string baseOutputPath = dialog.OutputPath;
                        bool generatePDF = dialog.GeneratePDF;
                        bool generateTXT = dialog.GenerateTXT;
                        bool generateXML = dialog.GenerateXML;

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
                                var generatedFiles = new System.Collections.Generic.List<string>();

                                // Generate PDF if requested
                                if (generatePDF)
                                {
                                    string pdfPath = baseOutputPath + ".pdf";
                                    ed.WriteMessage($"\nGenerating PDF: {pdfPath}");

                                    if (reportType == "Vertical")
                                        GenerateVerticalReportPdf(alignment, pdfPath);
                                    else
                                        GenerateHorizontalReportPdf(alignment, pdfPath);

                                    generatedFiles.Add(pdfPath);
                                    ed.WriteMessage($"\n✓ PDF saved successfully");
                                }

                                // Generate TXT if requested
                                if (generateTXT)
                                {
                                    string txtPath = baseOutputPath + ".txt";
                                    ed.WriteMessage($"\nGenerating TXT: {txtPath}");

                                    if (reportType == "Vertical")
                                        GenerateVerticalReport(alignment, txtPath);
                                    else
                                        GenerateHorizontalReport(alignment, txtPath);

                                    generatedFiles.Add(txtPath);
                                    ed.WriteMessage($"\n✓ TXT saved successfully");
                                }

                                // Generate XML if requested
                                if (generateXML)
                                {
                                    string xmlPath = baseOutputPath + ".xml";
                                    ed.WriteMessage($"\nGenerating XML: {xmlPath}");

                                    if (reportType == "Vertical")
                                        GenerateVerticalReportXml(alignment, xmlPath);
                                    else
                                        GenerateHorizontalReportXml(alignment, xmlPath);

                                    generatedFiles.Add(xmlPath);
                                    ed.WriteMessage($"\n✓ XML saved successfully");
                                }

                                tr.Commit();

                                // Show success message
                                string successMessage = $"Report(s) generated successfully!\n\n{generatedFiles.Count} file(s) created:\n";
                                foreach (string file in generatedFiles)
                                {
                                    successMessage += $"\n• {System.IO.Path.GetFileName(file)}";
                                }
                                successMessage += $"\n\nLocation: {System.IO.Path.GetDirectoryName(generatedFiles[0])}";

                                MessageBox.Show(
                                    successMessage,
                                    "Success",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );

                                ed.WriteMessage($"\n✓ All reports generated successfully!");
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

                                    // Generate reports - FIXED: Using PDF generation methods
                                    if (includeVertical)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Vertical.pdf");
                                        GenerateVerticalReportPdf(alignment, path);
                                    }

                                    if (includeHorizontal)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Horizontal.pdf");
                                        GenerateHorizontalReportPdf(alignment, path);
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
                MessageBox.Show(
                    $"Error during batch processing:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
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
            // 4-column layout with a subtle divider before numeric columns:
            // [ LABEL | STATION | NORTHING | EASTING ]
            iText.Layout.Element.Table dataTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 35f, 25f, 20f, 20f }));
            dataTable.SetWidth(UnitValue.CreatePercentValue(100));
            dataTable.SetMarginLeft(10);

            // LABEL (left-aligned)
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(label).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            // STATION (right-aligned under the STATION header)
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(station).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            // NORTHING with left-divider
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph($"{northing:F4}").SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            // EASTING
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
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

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

                // Add column headers with subtle divider (left border) before numeric columns
                iText.Layout.Element.Table headerTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 35f, 25f, 20f, 20f }));
                headerTable.SetWidth(UnitValue.CreatePercentValue(100));
                headerTable.SetMarginLeft(10);

                // LABEL header placeholder (blank, aligns with labels like POT/PI)
                iText.Layout.Element.Cell labelHeader = new iText.Layout.Element.Cell().Add(new Paragraph(" ")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                headerTable.AddCell(labelHeader);

                // STATION header (second column)
                iText.Layout.Element.Cell cell1 = new iText.Layout.Element.Cell().Add(new Paragraph("STATION")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER);
                headerTable.AddCell(cell1);

                // NORTHING header
                iText.Layout.Element.Cell cell2 = new iText.Layout.Element.Cell().Add(new Paragraph("NORTHING")
                    .SetFont(boldFont).SetFontSize(10))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                    .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f));
                headerTable.AddCell(cell2);

                // EASTING header
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
    /// Enhanced report selection form with multiple format options
    /// </summary>
    public class ReportSelectionForm : Form
    {
        public string ReportType { get; private set; }
        public string OutputPath { get; private set; }
        public bool GeneratePDF { get; private set; }
        public bool GenerateTXT { get; private set; }
        public bool GenerateXML { get; private set; }

        private RadioButton rbVertical;
        private RadioButton rbHorizontal;
        private TextBox txtOutputPath;
        private Button btnBrowse;
        private CheckBox chkPDF;
        private CheckBox chkTXT;
        private CheckBox chkXML;
        private Button btnSelectAll;
        private Button btnOK;
        private Button btnCancel;
        private GroupBox grpReportType;
        private GroupBox grpOutputFormat;

        public ReportSelectionForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Generate GeoTable Report";
            this.Size = new System.Drawing.Size(550, 380);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Report Type Group
            grpReportType = new GroupBox
            {
                Text = "Report Type",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(490, 80)
            };
            this.Controls.Add(grpReportType);

            rbVertical = new RadioButton
            {
                Text = "Vertical Alignment",
                Location = new System.Drawing.Point(20, 25),
                Checked = true,
                AutoSize = true
            };
            rbHorizontal = new RadioButton
            {
                Text = "Horizontal Alignment",
                Location = new System.Drawing.Point(20, 50),
                AutoSize = true
            };
            grpReportType.Controls.Add(rbVertical);
            grpReportType.Controls.Add(rbHorizontal);

            // Output Format Group
            grpOutputFormat = new GroupBox
            {
                Text = "Output Format",
                Location = new System.Drawing.Point(20, 110),
                Size = new System.Drawing.Size(490, 120)
            };
            this.Controls.Add(grpOutputFormat);

            chkPDF = new CheckBox
            {
                Text = "PDF (Portable Document Format)",
                Location = new System.Drawing.Point(20, 25),
                Checked = true,
                AutoSize = true
            };
            chkTXT = new CheckBox
            {
                Text = "TXT (Plain Text)",
                Location = new System.Drawing.Point(20, 50),
                Checked = false,
                AutoSize = true
            };
            chkXML = new CheckBox
            {
                Text = "XML (Extensible Markup Language)",
                Location = new System.Drawing.Point(20, 75),
                Checked = false,
                AutoSize = true
            };
            grpOutputFormat.Controls.Add(chkPDF);
            grpOutputFormat.Controls.Add(chkTXT);
            grpOutputFormat.Controls.Add(chkXML);

            btnSelectAll = new Button
            {
                Text = "Select All Formats",
                Location = new System.Drawing.Point(350, 10),
                Width = 140,
                Height = 30
            };
            btnSelectAll.Click += BtnSelectAll_Click;
            grpOutputFormat.Controls.Add(btnSelectAll);

            // Output Path
            var lblOutput = new System.Windows.Forms.Label
            {
                Text = "Output Location:",
                Location = new System.Drawing.Point(20, 245),
                AutoSize = true
            };
            this.Controls.Add(lblOutput);

            txtOutputPath = new TextBox
            {
                Location = new System.Drawing.Point(20, 270),
                Width = 390,
                Height = 25
            };
            txtOutputPath.Text = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "alignment_report"
            );
            this.Controls.Add(txtOutputPath);

            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(420, 268),
                Width = 90,
                Height = 28
            };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            // Action Buttons
            btnOK = new Button
            {
                Text = "Generate",
                Location = new System.Drawing.Point(320, 300),
                Width = 85,
                Height = 30,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(420, 300),
                Width = 90,
                Height = 28,
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // Add tooltips
            var toolTip = new ToolTip();
            toolTip.SetToolTip(chkPDF, "Generate report in PDF format (professional, printable)");
            toolTip.SetToolTip(chkTXT, "Generate report in plain text format (simple, editable)");
            toolTip.SetToolTip(chkXML, "Generate report in XML format (structured data, machine-readable)");
            toolTip.SetToolTip(btnSelectAll, "Select all output formats to generate all three files at once");
            toolTip.SetToolTip(txtOutputPath, "Base path for output files. Extension will be added based on selected formats.");
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            chkPDF.Checked = true;
            chkTXT.Checked = true;
            chkXML.Checked = true;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog())
            {
                dialog.Filter = "All Formats|*.*|PDF Files (*.pdf)|*.pdf|Text Files (*.txt)|*.txt|XML Files (*.xml)|*.xml";
                dialog.FilterIndex = 1;
                dialog.FileName = "alignment_report";
                dialog.Title = "Select Output Location";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    // Remove extension if provided, we'll add appropriate ones based on selection
                    txtOutputPath.Text = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(dialog.FileName),
                        System.IO.Path.GetFileNameWithoutExtension(dialog.FileName)
                    );
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // Validate at least one format is selected
            if (!chkPDF.Checked && !chkTXT.Checked && !chkXML.Checked)
            {
                MessageBox.Show(
                    "Please select at least one output format.",
                    "No Format Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                this.DialogResult = DialogResult.None;
                return;
            }

            // Validate output path
            if (string.IsNullOrWhiteSpace(txtOutputPath.Text))
            {
                MessageBox.Show(
                    "Please specify an output location.",
                    "No Output Location",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                this.DialogResult = DialogResult.None;
                return;
            }

            ReportType = rbVertical.Checked ? "Vertical" : "Horizontal";
            OutputPath = txtOutputPath.Text;
            GeneratePDF = chkPDF.Checked;
            GenerateTXT = chkTXT.Checked;
            GenerateXML = chkXML.Checked;
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

    /// <summary>
    /// User control for the report panel
    /// </summary>
    public class ReportPanelControl : UserControl
    {
        private GroupBox grpReportType;
        private RadioButton rbVertical;
        private RadioButton rbHorizontal;

        private GroupBox grpOutputFormat;
        private CheckBox chkPDF;
        private CheckBox chkTXT;
        private CheckBox chkXML;
        private Button btnSelectAllFormats;

        private GroupBox grpOutputLocation;
        private TextBox txtOutputPath;
        private Button btnBrowse;

        private GroupBox grpActions;
        private Button btnGenerate;
        private Button btnBatchProcess;

        private System.Windows.Forms.Label lblStatus;
        private ProgressBar progressBar;

        public ReportPanelControl()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.SuspendLayout();

            // Set control properties
            this.BackColor = System.Drawing.SystemColors.Control;
            this.AutoScroll = true;
            this.Padding = new Padding(10);

            int yPos = 10;

            // Report Type Group
            grpReportType = new GroupBox
            {
                Text = "Report Type",
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 80
            };

            rbVertical = new RadioButton
            {
                Text = "Vertical Alignment",
                Location = new System.Drawing.Point(15, 25),
                Checked = true,
                Width = 280
            };

            rbHorizontal = new RadioButton
            {
                Text = "Horizontal Alignment",
                Location = new System.Drawing.Point(15, 50),
                Width = 280
            };

            grpReportType.Controls.Add(rbVertical);
            grpReportType.Controls.Add(rbHorizontal);
            this.Controls.Add(grpReportType);

            yPos += 90;

            // Output Format Group
            grpOutputFormat = new GroupBox
            {
                Text = "Output Format",
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 130
            };

            chkPDF = new CheckBox
            {
                Text = "PDF (Portable Document Format)",
                Location = new System.Drawing.Point(15, 25),
                Checked = true,
                Width = 290
            };

            chkTXT = new CheckBox
            {
                Text = "TXT (Plain Text)",
                Location = new System.Drawing.Point(15, 50),
                Width = 290
            };

            chkXML = new CheckBox
            {
                Text = "XML (Extensible Markup Language)",
                Location = new System.Drawing.Point(15, 75),
                Width = 290
            };

            btnSelectAllFormats = new Button
            {
                Text = "Select All",
                Location = new System.Drawing.Point(200, 100),
                Width = 100,
                Height = 25
            };
            btnSelectAllFormats.Click += BtnSelectAllFormats_Click;

            grpOutputFormat.Controls.Add(chkPDF);
            grpOutputFormat.Controls.Add(chkTXT);
            grpOutputFormat.Controls.Add(chkXML);
            grpOutputFormat.Controls.Add(btnSelectAllFormats);
            this.Controls.Add(grpOutputFormat);

            yPos += 140;

            // Output Location Group
            grpOutputLocation = new GroupBox
            {
                Text = "Output Location",
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 80
            };

            txtOutputPath = new TextBox
            {
                Location = new System.Drawing.Point(15, 25),
                Width = 290,
                Text = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "alignment_report"
                )
            };

            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(15, 50),
                Width = 100,
                Height = 25
            };
            btnBrowse.Click += BtnBrowse_Click;

            grpOutputLocation.Controls.Add(txtOutputPath);
            grpOutputLocation.Controls.Add(btnBrowse);
            this.Controls.Add(grpOutputLocation);

            yPos += 90;

            // Actions Group
            grpActions = new GroupBox
            {
                Text = "Actions",
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 90
            };

            btnGenerate = new Button
            {
                Text = "Generate Report",
                Location = new System.Drawing.Point(15, 25),
                Width = 290,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnGenerate.Click += BtnGenerate_Click;

            btnBatchProcess = new Button
            {
                Text = "Batch Process All Alignments",
                Location = new System.Drawing.Point(15, 58),
                Width = 290,
                Height = 25
            };
            btnBatchProcess.Click += BtnBatchProcess_Click;

            grpActions.Controls.Add(btnGenerate);
            grpActions.Controls.Add(btnBatchProcess);
            this.Controls.Add(grpActions);

            yPos += 100;

            // Status Label
            lblStatus = new System.Windows.Forms.Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 20,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblStatus);

            yPos += 25;

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(10, yPos),
                Width = 320,
                Height = 20,
                Visible = false
            };
            this.Controls.Add(progressBar);

            // Set tooltip
            var toolTip = new ToolTip();
            toolTip.SetToolTip(chkPDF, "Professional, printable PDF format");
            toolTip.SetToolTip(chkTXT, "Simple, editable plain text format");
            toolTip.SetToolTip(chkXML, "Structured data format for processing");
            toolTip.SetToolTip(btnGenerate, "Select an alignment and generate report(s)");
            toolTip.SetToolTip(btnBatchProcess, "Process all alignments in the drawing");

            this.ResumeLayout(false);
        }

        private void BtnSelectAllFormats_Click(object sender, EventArgs e)
        {
            chkPDF.Checked = true;
            chkTXT.Checked = true;
            chkXML.Checked = true;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog())
            {
                dialog.Filter = "All Formats|*.*";
                dialog.FilterIndex = 1;
                dialog.FileName = "alignment_report";
                dialog.Title = "Select Output Location";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtOutputPath.Text = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(dialog.FileName),
                        System.IO.Path.GetFileNameWithoutExtension(dialog.FileName)
                    );
                }
            }
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            // Validate at least one format is selected
            if (!chkPDF.Checked && !chkTXT.Checked && !chkXML.Checked)
            {
                MessageBox.Show(
                    "Please select at least one output format.",
                    "No Format Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // Validate output path
            if (string.IsNullOrWhiteSpace(txtOutputPath.Text))
            {
                MessageBox.Show(
                    "Please specify an output location.",
                    "No Output Location",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            try
            {
                lblStatus.Text = "Prompting for alignment selection...";
                AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                // Prompt user to select alignment
                PromptEntityOptions options = new PromptEntityOptions("\nSelect alignment: ");
                options.SetRejectMessage("\nMust be an alignment.");
                options.AddAllowedClass(typeof(Alignment), true);
                PromptEntityResult result = ed.GetEntity(options);

                if (result.Status != PromptStatus.OK)
                {
                    lblStatus.Text = "Cancelled - no alignment selected";
                    return;
                }

                ObjectId alignmentId = result.ObjectId;
                string reportType = rbVertical.Checked ? "Vertical" : "Horizontal";
                string baseOutputPath = txtOutputPath.Text;

                lblStatus.Text = "Generating report(s)...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;

                // Generate reports
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    if (alignment == null)
                    {
                        lblStatus.Text = "Error: Invalid alignment";
                        progressBar.Visible = false;
                        return;
                    }

                    var generatedFiles = new System.Collections.Generic.List<string>();

                    // Generate PDF if requested
                    if (chkPDF.Checked)
                    {
                        string pdfPath = baseOutputPath + ".pdf";
                        if (reportType == "Vertical")
                            GenerateVerticalReportPdf(alignment, pdfPath);
                        else
                            GenerateHorizontalReportPdf(alignment, pdfPath);
                        generatedFiles.Add(pdfPath);
                    }

                    // Generate TXT if requested
                    if (chkTXT.Checked)
                    {
                        string txtPath = baseOutputPath + ".txt";
                        if (reportType == "Vertical")
                            GenerateVerticalReport(alignment, txtPath);
                        else
                            GenerateHorizontalReport(alignment, txtPath);
                        generatedFiles.Add(txtPath);
                    }

                    // Generate XML if requested
                    if (chkXML.Checked)
                    {
                        string xmlPath = baseOutputPath + ".xml";
                        if (reportType == "Vertical")
                            GenerateVerticalReportXml(alignment, xmlPath);
                        else
                            GenerateHorizontalReportXml(alignment, xmlPath);
                        generatedFiles.Add(xmlPath);
                    }

                    tr.Commit();

                    progressBar.Visible = false;
                    lblStatus.Text = $"Success! Generated {generatedFiles.Count} file(s)";

                    string successMessage = $"Report(s) generated successfully!\n\n{generatedFiles.Count} file(s) created:\n";
                    foreach (string file in generatedFiles)
                    {
                        successMessage += $"\n• {System.IO.Path.GetFileName(file)}";
                    }
                    successMessage += $"\n\nLocation: {System.IO.Path.GetDirectoryName(generatedFiles[0])}";

                    MessageBox.Show(successMessage, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "Error occurred";
                MessageBox.Show($"Error generating report:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnBatchProcess_Click(object sender, EventArgs e)
        {
            using (var dialog = new BatchProcessForm())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        lblStatus.Text = "Processing batch...";
                        progressBar.Visible = true;
                        progressBar.Style = ProgressBarStyle.Marquee;

                        AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
                        Editor ed = doc.Editor;

                        string outputFolder = dialog.OutputFolder;
                        bool includeVertical = dialog.IncludeVertical;
                        bool includeHorizontal = dialog.IncludeHorizontal;

                        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            CivilDocument civilDoc = CivilApplication.ActiveDocument;

                            using (ObjectIdCollection alignmentIds = civilDoc.GetAlignmentIds())
                            {
                                int count = 0;
                                int total = alignmentIds.Count;

                                foreach (ObjectId alignmentId in alignmentIds)
                                {
                                    Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                                    if (alignment == null) continue;

                                    count++;
                                    lblStatus.Text = $"Processing {count}/{total}: {alignment.Name}";
                                    System.Windows.Forms.Application.DoEvents();

                                    if (includeVertical)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Vertical.pdf");
                                        GenerateVerticalReportPdf(alignment, path);
                                    }

                                    if (includeHorizontal)
                                    {
                                        string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Horizontal.pdf");
                                        GenerateHorizontalReportPdf(alignment, path);
                                    }
                                }

                                tr.Commit();

                                progressBar.Visible = false;
                                lblStatus.Text = $"Batch complete! Processed {count} alignments";

                                MessageBox.Show(
                                    $"Batch processing complete!\n\n{count} alignments processed.",
                                    "Success",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        progressBar.Visible = false;
                        lblStatus.Text = "Error during batch processing";
                        MessageBox.Show($"Error:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Helper methods to call the report generation methods
        private void GenerateVerticalReport(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateVerticalReport",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }

        private void GenerateVerticalReportPdf(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateVerticalReportPdf",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }

        private void GenerateVerticalReportXml(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateVerticalReportXml",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }

        private void GenerateHorizontalReport(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateHorizontalReport",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }

        private void GenerateHorizontalReportPdf(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateHorizontalReportPdf",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }

        private void GenerateHorizontalReportXml(Alignment alignment, string path)
        {
            var commands = new ReportCommands();
            typeof(ReportCommands).GetMethod("GenerateHorizontalReportXml",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(commands, new object[] { alignment, path });
        }
    }
}
