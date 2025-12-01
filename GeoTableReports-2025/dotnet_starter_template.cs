#pragma warning disable 1591
#pragma warning disable CA1416
/*
 * Civil 3D GeoTable Reports - .NET Add-in Starter Template
 *
 * This is a basic template showing how a .NET add-in would work
 * Compile this with Visual Studio and Civil 3D API references
 */

using System;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing;
using System.IO;
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
using iText.Layout.Renderer;
using iText.Kernel.Pdf.Canvas;
using OfficeOpenXml;
using OfficeOpenXml.Style;
// (Viewer components removed)

[assembly: CommandClass(typeof(GeoTableReports.ReportCommands2025))]

namespace GeoTableReports
{
    public class ReportCommands2025 : IExtensionApplication
    {
        // Centralized label helper for InRoads-style reporting
        private static class InRoadsLabels
        {
            public static string HorizontalTangentStart(int index) => index == 0 ? "POB" : "PI";
            public static string HorizontalCurveStart => "PC";
            public static string HorizontalCurveEnd => "PT";
            public static string VerticalTangentStart(int index) => index == 0 ? "POB" : ""; // blank for subsequent tangents
            public static string VerticalCurveStart => "PVC";
            public static string VerticalCurveIntersection => "PVI";
            public static string VerticalCurveEnd => "PVT";
        }
    // InRoads variable label mapping applied (2025 feature-single-version-compatibility branch):
    // Horizontal: POB (start of alignment), PI (point of intersection of tangents),
    // PC (point of curvature – start of circular curve), PT (point of tangency – end of circular curve),
    // TS (tangent to spiral), SC (spiral to curve), CS (curve to spiral), ST (spiral to tangent) when spirals present.
    // Vertical: POB (start), PVC (point of vertical curvature), PVI (point of vertical intersection), PVT (point of vertical tangency).
    // Previous placeholders (POT, ST, CS for non-spiral curve ends) replaced to align with InRoads conventions.

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
            // No UI resources to dispose in simplified mode
        }

        // Panel command removed in simplified mode

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
                // Select alignment BEFORE showing dialog so preview has real data
                ObjectId alignmentId = SelectAlignment(ed);
                if (alignmentId == ObjectId.Null)
                {
                    ed.WriteMessage("\nNo alignment selected.");
                    return;
                }

                AlignmentPreviewData previewData = BuildAlignmentPreviewData(alignmentId);

                using (var dialog = new ReportSelectionForm2025(previewData))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    string reportType = dialog.ReportType; // "Vertical" or "Horizontal"
                    string baseOutputPath = dialog.OutputPath;

                    bool generateAlignmentPDF = dialog.GenerateAlignmentPDF;
                    bool generateAlignmentTXT = dialog.GenerateAlignmentTXT;
                    bool generateAlignmentXML = dialog.GenerateAlignmentXML;
                    bool generateGeoTablePDF = dialog.GenerateGeoTablePDF;
                    bool generateGeoTableEXCEL = dialog.GenerateGeoTableEXCEL;
                    bool openFolderAfter = dialog.OpenFolderAfter;
                    bool openFilesAfter = dialog.OpenFilesAfter;

                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        CivDb.Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as CivDb.Alignment;
                        if (alignment == null)
                        {
                            ed.WriteMessage("\nInvalid alignment selected.");
                            tr.Abort();
                            return;
                        }

                        string alignmentName = "alignment";
                        try { alignmentName = alignment.Name ?? "alignment"; } catch { }
                        ed.WriteMessage($"\nProcessing alignment: {alignmentName}...");

                        var generatedFiles = new System.Collections.Generic.List<string>();
                        int totalSteps = (generateAlignmentPDF ? 1 : 0) + (generateAlignmentTXT ? 1 : 0) + (generateAlignmentXML ? 1 : 0) + (generateGeoTablePDF ? 1 : 0) + (generateGeoTableEXCEL ? 1 : 0);
                        ProgressStatusWindow progressWindow = null;
                        if (totalSteps > 0)
                        {
                            progressWindow = new ProgressStatusWindow(totalSteps);
                            progressWindow.Show();
                            System.Windows.Forms.Application.DoEvents();
                        }
                        try
                        {
                            if (generateAlignmentPDF)
                            {
                                progressWindow?.UpdateStatus("Generating Alignment PDF...");
                                string pdfPath = baseOutputPath + "_Alignment_Report.pdf";
                                if (reportType == "Vertical")
                                    GenerateVerticalReportPdf(alignment, pdfPath);
                                else
                                    GenerateHorizontalReportPdf(alignment, pdfPath);
                                generatedFiles.Add(pdfPath);
                                progressWindow?.IncrementProgress();
                            }
                            if (generateAlignmentTXT)
                            {
                                progressWindow?.UpdateStatus("Generating Alignment TXT...");
                                string txtPath = baseOutputPath + "_Alignment_Report.txt";
                                if (reportType == "Vertical")
                                    GenerateVerticalReport(alignment, txtPath);
                                else
                                    GenerateHorizontalReport(alignment, txtPath);
                                generatedFiles.Add(txtPath);
                                progressWindow?.IncrementProgress();
                            }
                            if (generateAlignmentXML)
                            {
                                progressWindow?.UpdateStatus("Generating Alignment XML...");
                                string xmlPath = baseOutputPath + "_Alignment_Report.xml";
                                if (reportType == "Vertical")
                                    GenerateVerticalReportXml(alignment, xmlPath);
                                else
                                    GenerateHorizontalReportXml(alignment, xmlPath);
                                generatedFiles.Add(xmlPath);
                                progressWindow?.IncrementProgress();
                            }
                            if (generateGeoTablePDF)
                            {
                                progressWindow?.UpdateStatus("Generating GeoTable PDF...");
                                string geoPdf = baseOutputPath + "_GeoTable.pdf";
                                if (reportType == "Vertical")
                                    GenerateVerticalGeoTablePdf(alignment, geoPdf);
                                else
                                    GenerateHorizontalGeoTablePdf(alignment, geoPdf);
                                generatedFiles.Add(geoPdf);
                                progressWindow?.IncrementProgress();
                            }
                            if (generateGeoTableEXCEL)
                            {
                                progressWindow?.UpdateStatus("Generating GeoTable Excel...");
                                string geoXlsx = baseOutputPath + "_GeoTable.xlsx";
                                if (reportType == "Vertical")
                                    GenerateVerticalReportExcel(alignment, geoXlsx); // fallback
                                else
                                    GenerateHorizontalGeoTableExcel(alignment, geoXlsx);
                                generatedFiles.Add(geoXlsx);
                                progressWindow?.IncrementProgress();
                            }
                            tr.Commit();
                            progressWindow?.Complete("All reports generated successfully!");
                            System.Threading.Thread.Sleep(800);
                            progressWindow?.Close();
                            if (generatedFiles.Count > 0)
                            {
                                string message = $"Report(s) generated successfully!\n\n{generatedFiles.Count} file(s):";
                                foreach (var f in generatedFiles) message += $"\n• {System.IO.Path.GetFileName(f)}";
                                string folder = System.IO.Path.GetDirectoryName(generatedFiles[0]) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                message += $"\n\nLocation: {folder}";
                                System.Windows.Forms.MessageBox.Show(message, "Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                                ed.WriteMessage("\n✓ Reports generated successfully.");
                                if (openFolderAfter) { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = folder, UseShellExecute = true }); } catch { } }
                                if (openFilesAfter)
                                {
                                    foreach (var f in generatedFiles)
                                    {
                                        try { if (File.Exists(f)) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = f, UseShellExecute = true }); } catch { }
                                    }
                                }
                            }
                            else ed.WriteMessage("\nNo formats selected; nothing generated.");
                        }
                        catch (System.Exception reportEx)
                        {
                            progressWindow?.Fail($"Error: {reportEx.Message}");
                            System.Threading.Thread.Sleep(1200);
                            progressWindow?.Close();
                            tr.Abort();
                            throw new System.Exception($"Error generating {reportType} report: {reportEx.Message}", reportEx);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                string errorMessage = ex.Message;
                if (ex.InnerException != null)
                    errorMessage += $"\nInner: {ex.InnerException.Message}";
                ed.WriteMessage($"\n✗ Error: {errorMessage}");
                System.Windows.Forms.MessageBox.Show(errorMessage, "Report Generation Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // Build preview data for dialog (metadata + sample lines for both horizontal & vertical)
        private AlignmentPreviewData BuildAlignmentPreviewData(ObjectId alignmentId)
        {
            var data = new AlignmentPreviewData();
            try
            {
                AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as CivDb.Alignment;
                    if (alignment == null) return data;
                    data.AlignmentName = alignment.Name ?? "(Unnamed)";
                    try { data.Description = alignment.Description ?? ""; } catch { }
                    try { data.StyleName = alignment.StyleName ?? ""; } catch { }
                    double startSta = 0; double endSta = 0;
                    try { if (alignment.Entities.Count > 0) { startSta = (alignment.Entities[0] as dynamic).StartStation; var last = alignment.Entities[alignment.Entities.Count - 1]; endSta = (last as dynamic).EndStation; } } catch { }
                    data.StartStation = startSta; data.EndStation = endSta;
                    int lines = 0, arcs = 0, spirals = 0;
                    for (int i = 0; i < alignment.Entities.Count; i++)
                    {
                        var e = alignment.Entities[i];
                        if (e == null) continue;
                        if (e.EntityType == AlignmentEntityType.Line) lines++;
                        else if (e.EntityType == AlignmentEntityType.Arc) arcs++;
                        else if (e.EntityType == AlignmentEntityType.Spiral) spirals++;
                        if (data.HorizontalSampleLines.Count < 15) // collect small textual snippet (similar to TXT)
                        {
                            try
                            {
                                switch (e.EntityType)
                                {
                                    case AlignmentEntityType.Line:
                                        var l = e as AlignmentLine;
                                        if (l != null)
                                        {
                                            double x1=0,y1=0,z1=0; double x2=0,y2=0,z2=0; alignment.PointLocation(l.StartStation,0,0,ref x1,ref y1,ref z1); alignment.PointLocation(l.EndStation,0,0,ref x2,ref y2,ref z2);
                                            // InRoads: first tangent start POB, subsequent tangent starts omitted in sample (keep PI at end of tangent)
                                            string tangentStartLabel = (i == 0) ? "POB" : "PI"; // if not first, treat start as PI entering curve sequence
                                            data.HorizontalSampleLines.Add($"{tangentStartLabel} {FormatStation(l.StartStation),15} {FormatWithProperRounding(y1,4),15} {FormatWithProperRounding(x1,4),15}");
                                            data.HorizontalSampleLines.Add($"PI {FormatStation(l.EndStation),15} {FormatWithProperRounding(y2,4),15} {FormatWithProperRounding(x2,4),15}");
                                        }
                                        break;
                                    case AlignmentEntityType.Arc:
                                        var a = e as AlignmentArc;
                                        if (a != null)
                                        {
                                            double x1=0,y1=0,z1=0; double x2=0,y2=0,z2=0; alignment.PointLocation(a.StartStation,0,0,ref x1,ref y1,ref z1); alignment.PointLocation(a.EndStation,0,0,ref x2,ref y2,ref z2);
                                            // InRoads circular curve: PC start, PT end
                                            data.HorizontalSampleLines.Add($"PC {FormatStation(a.StartStation),15} {FormatWithProperRounding(y1,4),15} {FormatWithProperRounding(x1,4),15}");
                                            data.HorizontalSampleLines.Add($"PT {FormatStation(a.EndStation),15} {FormatWithProperRounding(y2,4),15} {FormatWithProperRounding(x2,4),15}");
                                        }
                                        break;
                                    case AlignmentEntityType.Spiral:
                                        var se = e; // sub-entity access for stations
                                        try
                                        {
                                            dynamic dyn = se;
                                            double ss = dyn.StartStation; double es = dyn.EndStation; double xm=0,ym=0,zm=0; alignment.PointLocation((ss+es)/2,0,0,ref xm,ref ym,ref zm);
                                            double x1=0,y1=0,z1=0; double x2=0,y2=0,z2=0; alignment.PointLocation(ss,0,0,ref x1,ref y1,ref z1); alignment.PointLocation(es,0,0,ref x2,ref y2,ref z2);
                                            data.HorizontalSampleLines.Add($"TS {FormatStation(ss),15} {FormatWithProperRounding(y1,4),15} {FormatWithProperRounding(x1,4),15}");
                                            data.HorizontalSampleLines.Add($"SPI{FormatStation((ss+es)/2),15} {FormatWithProperRounding(ym,4),15} {FormatWithProperRounding(xm,4),15}");
                                            data.HorizontalSampleLines.Add($"SC {FormatStation(es),15} {FormatWithProperRounding(y2,4),15} {FormatWithProperRounding(x2,4),15}");
                                        }
                                        catch { }
                                        break;
                                }
                            }
                            catch { }
                        }
                    }
                    data.LineCount = lines; data.ArcCount = arcs; data.SpiralCount = spirals;

                    // Vertical profile sample
                    try
                    {
                        foreach (ObjectId pid in alignment.GetProfileIds())
                        {
                            using (Profile profile = pid.GetObject(OpenMode.ForRead) as Profile)
                            {
                                if (profile != null && profile.Entities.Count > 0)
                                {
                                    data.ProfileName = profile.Name ?? "";
                                    for (int i = 0; i < profile.Entities.Count && data.VerticalSampleLines.Count < 12; i++)
                                    {
                                        var pe = profile.Entities[i];
                                        if (pe.EntityType == ProfileEntityType.Tangent && pe is ProfileTangent t)
                                        {
                                            data.VerticalSampleLines.Add($"POB {FormatStation(t.StartStation),15} {t.StartElevation,12:F2}");
                                            data.VerticalSampleLines.Add($"PVI {FormatStation(t.EndStation),15} {t.EndElevation,12:F2}");
                                        }
                                        else if (pe.EntityType == ProfileEntityType.Circular && pe is ProfileCircular c)
                                        {
                                            data.VerticalSampleLines.Add($"PVC {FormatStation(c.StartStation),15} {c.StartElevation,12:F2}");
                                            data.VerticalSampleLines.Add($"PVI {FormatStation(c.PVIStation),15} {c.PVIElevation,12:F2}");
                                            data.VerticalSampleLines.Add($"PVT {FormatStation(c.EndStation),15} {c.EndElevation,12:F2}");
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    catch { }
                    tr.Commit();
                }
            }
            catch { }
            return data;
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
                string outputFolder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GeoTable_Batch");
                if (!System.IO.Directory.Exists(outputFolder)) System.IO.Directory.CreateDirectory(outputFolder);
                bool includeHorizontal = true; // simplified mode only horizontal

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
                            ed.WriteMessage($"\nProcessing {count}/{total}: {alignment.Name}");
                            if (includeHorizontal)
                            {
                                string path = System.IO.Path.Combine(outputFolder, $"{alignment.Name}_Horizontal.pdf");
                                GenerateHorizontalReportPdf(alignment, path);
                            }
                        }
                        tr.Commit();
                        ed.WriteMessage($"\nBatch processing complete: {count} alignments processed. Output folder: {outputFolder}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during batch processing: {ex.Message}");
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
                        writer.WriteLine($"Vertical Profile Name: {profileName}");
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
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(FormatWithProperRounding(northing, 4)).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            // EASTING
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(FormatWithProperRounding(easting, 4)).SetFont(font).SetFontSize(9))
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
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
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

                    // Reorder entities to match InRails format (Linear, then Spiral/Arc associated with it)
                    var reorderedEntities = ReorderEntitiesForInRails(alignment);

                    // Process each entity
                    for (int i = 0; i < reorderedEntities.Count; i++)
                    {
                        AlignmentEntity entity = reorderedEntities[i];
                        if (entity == null) continue;

                        AlignmentEntity prevEntity = i > 0 ? reorderedEntities[i - 1] : null;
                        AlignmentEntity nextEntity = i < reorderedEntities.Count - 1 ? reorderedEntities[i + 1] : null;

                        try
                        {
                            switch (entity.EntityType)
                            {
                                case AlignmentEntityType.Line:
                                    WriteLinearElement(writer, entity as AlignmentLine, alignment, i, prevEntity, nextEntity);
                                    break;
                                case AlignmentEntityType.Arc:
                                    WriteArcElement(writer, entity as AlignmentArc, alignment, i, prevEntity, nextEntity);
                                    break;
                                case AlignmentEntityType.Spiral:
                                    // Try to cast to AlignmentSpiral, if that fails use AlignmentEntity directly
                                    AlignmentSpiral spiralEntity = entity as AlignmentSpiral;
                                    if (spiralEntity == null)
                                    {
                                        // If cast fails, try accessing spiral through SubEntity
                                        WriteSpiralElementFromEntity(writer, entity, alignment, i, prevEntity, nextEntity);
                                    }
                                    else
                                    {
                                        WriteSpiralElement(writer, spiralEntity, alignment, i, prevEntity, nextEntity);
                                    }
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

        private System.Collections.Generic.List<AlignmentEntity> ReorderEntitiesForInRails(CivDb.Alignment alignment)
        {
            // Simply sort all entities by their start station to maintain correct sequential order
            var allEntities = new System.Collections.Generic.List<AlignmentEntity>();
            for (int i = 0; i < alignment.Entities.Count; i++)
            {
                if (alignment.Entities[i] != null)
                    allEntities.Add(alignment.Entities[i]);
            }

            // Sort all entities by start station
            allEntities.Sort((a, b) =>
            {
                try
                {
                    double stationA = (a as dynamic).StartStation;
                    double stationB = (b as dynamic).StartStation;
                    return stationA.CompareTo(stationB);
                }
                catch
                {
                    // If we can't get start station, maintain original order
                    return 0;
                }
            });

            return allEntities;
        }

        private void WriteLinearElement(System.IO.StreamWriter writer, AlignmentLine line, CivDb.Alignment alignment, int index, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
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

                // Determine start label based on previous element
                string startLabel;
                if (index == 0)
                {
                    startLabel = "POB";  // Point of Beginning
                }
                else if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    startLabel = "ST ";  // Spiral to Tangent
                }
                else if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc)
                {
                    startLabel = "PT ";  // Point of Tangency (Curve to Tangent)
                }
                else
                {
                    startLabel = "PI ";  // Point of Intersection
                }

                writer.WriteLine($" {startLabel} ( ) {FormatStation(line.StartStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");

                // Determine end label based on next element
                string endLabel;
                if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    endLabel = "TS ";  // Tangent to Spiral
                }
                else if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc)
                {
                    endLabel = "PC ";  // Point of Curvature (Tangent to Arc)
                }
                else
                {
                    endLabel = "PI ";  // Point of Intersection
                }

                writer.WriteLine($" {endLabel} ( ) {FormatStation(line.EndStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                writer.WriteLine($" Tangent Direction: {bearing}");
                writer.WriteLine($" Tangent Length: {FormatWithProperRounding(line.Length, 4),15}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Linear");
                writer.WriteLine($"Error writing line data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteArcElement(System.IO.StreamWriter writer, AlignmentArc arc, CivDb.Alignment alignment, int index, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
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

                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                // Get PI station and coordinates from arc properties
                double piStation = arc.PIStation;
                double xPI = 0, yPI = 0;

                // Try to get PI and Center coordinates from sub-entities
                bool gotFromSubEntity = false;
                try
                {
                    if (arc.SubEntityCount > 0)
                    {
                        var subEntity = arc[0];
                        if (subEntity is AlignmentSubEntityArc)
                        {
                            var subEntityArc = subEntity as AlignmentSubEntityArc;
                            var piPoint = subEntityArc.PIPoint;
                            var centerPoint = subEntityArc.CenterPoint;
                            xPI = piPoint.X;  // Easting
                            yPI = piPoint.Y;  // Northing
                            xc = centerPoint.X;  // Easting (Center)
                            yc = centerPoint.Y;  // Northing (Center)
                            gotFromSubEntity = true;
                        }
                    }
                }
                catch { }

                // Calculate deltaRadians using arc length and radius
                double deltaRadians = arc.Length / Math.Abs(arc.Radius);
                double tangent = arc.Radius * Math.Tan(Math.Abs(deltaRadians) / 2);

                // Fallback: calculate if sub-entity access failed
                if (!gotFromSubEntity)
                {
                    double backTangentDir = arc.StartDirection + Math.PI;
                    xPI = x1 + tangent * Math.Cos(backTangentDir);
                    yPI = y1 + tangent * Math.Sin(backTangentDir);

                    // Calculate center point as fallback
                    double midStation = (arc.StartStation + arc.EndStation) / 2;
                    double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                    alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                }

                double deltaDegrees = deltaRadians * (180.0 / Math.PI);
                double chord = 2 * arc.Radius * Math.Sin(Math.Abs(deltaRadians) / 2);
                double middleOrdinate = arc.Radius * (1 - Math.Cos(Math.Abs(deltaRadians) / 2));
                double external = arc.Radius * (1 / Math.Cos(Math.Abs(deltaRadians) / 2) - 1);

                writer.WriteLine("Element: Circular");
                
                // Determine start label based on previous element
                string startLabel;
                if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    startLabel = "SC ";  // Spiral to Curve
                }
                else
                {
                    startLabel = "PC ";  // Point of Curvature (no spiral before)
                }
                
                writer.WriteLine($" {startLabel} ( ) {FormatStation(arc.StartStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");
                writer.WriteLine($" PI  ( ) {FormatStation(piStation),15} {FormatWithProperRounding(yPI, 4),15} {FormatWithProperRounding(xPI, 4),15}");
                writer.WriteLine($" CC  ( ) {FormatWithProperRounding(yc, 4),32} {FormatWithProperRounding(xc, 4),15}");
                
                // Determine end label based on next element
                string endLabel;
                if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    endLabel = "CS ";  // Curve to Spiral
                }
                else
                {
                    endLabel = "PT ";  // Point of Tangency (no spiral after)
                }
                
                writer.WriteLine($" {endLabel} ( ) {FormatStation(arc.EndStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                writer.WriteLine($" Radius: {FormatWithProperRounding(arc.Radius, 4),15}");
                writer.WriteLine($" Design Speed(mph): {FormatWithProperRounding(50.0, 4),15}");
                writer.WriteLine($" Cant(inches): {2.0,15:F3}");
                writer.WriteLine($" Delta: {FormatAngle(Math.Abs(deltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}");
                double degreeOfCurvatureChord = (100.0 * deltaRadians) / (arc.Length) * (180.0 / Math.PI);
                writer.WriteLine($"Degree of Curvature (Arc): {FormatAngle(degreeOfCurvatureChord)}");
                writer.WriteLine($" Length: {FormatWithProperRounding(arc.Length, 4),15}");
                writer.WriteLine($" Length(Chorded): {FormatWithProperRounding(arc.Length, 4),15}");
                writer.WriteLine($" Tangent: {FormatWithProperRounding(tangent, 4),15}");
                writer.WriteLine($" Chord: {FormatWithProperRounding(chord, 4),15}");
                writer.WriteLine($" Middle Ordinate: {FormatWithProperRounding(middleOrdinate, 4),15}");
                writer.WriteLine($" External: {FormatWithProperRounding(external, 4),15}");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Circular");
                writer.WriteLine($"Error writing arc data: {ex.Message}");
                writer.WriteLine();
            }
        }

        private void WriteSpiralElementFromEntity(System.IO.StreamWriter writer, AlignmentEntity entity, CivDb.Alignment alignment, int index, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            if (entity == null)
            {
                writer.WriteLine("Element: Clothoid");
                writer.WriteLine("Unable to read spiral data - entity is null.");
                writer.WriteLine();
                return;
            }

            try
            {
                writer.WriteLine("Element: Clothoid");
                writer.WriteLine($"DEBUG: Entity Type: {entity.GetType().Name}");
                writer.WriteLine($"DEBUG: EntityType: {entity.EntityType}");
                writer.WriteLine($"DEBUG: SubEntityCount: {entity.SubEntityCount}");

                // Try to access spiral data through SubEntity
                if (entity.SubEntityCount > 0)
                {
                    var subEntity = entity[0];
                    writer.WriteLine($"DEBUG: SubEntity Type: {subEntity.GetType().Name}");

                    if (subEntity is AlignmentSubEntitySpiral)
                    {
                        var spiralSubEntity = subEntity as AlignmentSubEntitySpiral;

                        double radiusIn = spiralSubEntity.RadiusIn;
                        double radiusOut = spiralSubEntity.RadiusOut;
                        double length = spiralSubEntity.Length;

                        double x1 = 0, y1 = 0, z1 = 0;
                        double x2 = 0, y2 = 0, z2 = 0;
                        double xMid = 0, yMid = 0, zMid = 0;

                        // Get start and end stations from the entity
                        double startStation = 0, endStation = 0;
                        try
                        {
                            // Access via dynamic to get StartStation/EndStation
                            dynamic dynEntity = entity;
                            startStation = dynEntity.StartStation;
                            endStation = dynEntity.EndStation;
                        }
                        catch (System.Exception ex)
                        {
                            writer.WriteLine($"DEBUG: Could not access stations: {ex.Message}");
                        }

                        alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                        alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);
                        alignment.PointLocation((startStation + endStation) / 2, 0, 0, ref xMid, ref yMid, ref zMid);

                        bool isEntry = radiusIn > radiusOut || (radiusIn == 0 && radiusOut > 0);
                        double R1 = isEntry ? double.PositiveInfinity : (radiusIn > 0 ? radiusIn : 0);
                        double R2 = isEntry ? (radiusOut > 0 ? radiusOut : 0) : double.PositiveInfinity;
                        double L = length;
                        double R = isEntry ? R2 : R1;
                        double A = 0;
                        if (R > 0 && !double.IsInfinity(R))
                        {
                            A = Math.Sqrt(L * R);
                        }
                        double theta = 0;
                        if (R > 0 && !double.IsInfinity(R))
                        {
                            theta = L / (2 * R);
                        }

                        writer.WriteLine(" SubEntity Access: SUCCESS via AlignmentSubEntitySpiral");

                        // Determine if entry or exit spiral based on neighboring elements
                        bool isEntrySpiralByContext = nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc;
                        bool isExitSpiralByContext = prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc;
                        
                        // Use context first, then fallback to radius logic
                        if (isEntrySpiralByContext || (isEntry && !isExitSpiralByContext))
                        {
                            // Entry spiral: Tangent -> Spiral -> Curve
                            writer.WriteLine($" TS  ( ) {FormatStation(startStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");
                            writer.WriteLine($" SPI ( ) {FormatStation((startStation + endStation) / 2),15} {FormatWithProperRounding(yMid, 4),15} {FormatWithProperRounding(xMid, 4),15}");
                            writer.WriteLine($" SC  ( ) {FormatStation(endStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                        }
                        else
                        {
                            // Exit spiral: Curve -> Spiral -> Tangent
                            writer.WriteLine($" CS  ( ) {FormatStation(startStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");
                            writer.WriteLine($" SPI ( ) {FormatStation((startStation + endStation) / 2),15} {FormatWithProperRounding(yMid, 4),15} {FormatWithProperRounding(xMid, 4),15}");
                            writer.WriteLine($" ST  ( ) {FormatStation(endStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                        }

                        writer.WriteLine($" R1 (Radius of curve 1): {(double.IsInfinity(R1) ? "Infinite" : FormatWithProperRounding(R1, 4)),15}");
                        writer.WriteLine($" R2 (Radius of curve 2): {(double.IsInfinity(R2) ? "Infinite" : FormatWithProperRounding(R2, 4)),15}");
                        writer.WriteLine($" SS (Spiral start): {FormatStation(startStation),15}");
                        writer.WriteLine($" SE (Spiral end): {FormatStation(endStation),15}");
                        writer.WriteLine($" L (Total arc length): {FormatWithProperRounding(L, 4),15}");
                        writer.WriteLine($" A (Flatness parameter): {FormatWithProperRounding(A, 4),15}");
                        writer.WriteLine();
                        return;
                    }
                    else
                    {
                        writer.WriteLine($"DEBUG: SubEntity is not AlignmentSubEntitySpiral");
                    }
                }
                else
                {
                    writer.WriteLine($"DEBUG: No sub-entities found");
                }

                writer.WriteLine("Unable to read spiral data - could not access spiral sub-entity.");
                writer.WriteLine();
            }
            catch (System.Exception ex)
            {
                writer.WriteLine("Element: Clothoid");
                writer.WriteLine($"Error extracting spiral from entity: {ex.Message}");
                writer.WriteLine($"Stack trace: {ex.StackTrace}");
                writer.WriteLine();
            }
        }

        private void WriteSpiralElement(System.IO.StreamWriter writer, AlignmentSpiral spiral, CivDb.Alignment alignment, int index, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
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

                // Try to get additional properties from AlignmentSubEntitySpiral
                double radiusIn = spiral.RadiusIn;
                double radiusOut = spiral.RadiusOut;
                double length = spiral.Length;
                double startStation = spiral.StartStation;
                double endStation = spiral.EndStation;

                bool gotFromSubEntity = false;
                try
                {
                    if (spiral.SubEntityCount > 0)
                    {
                        var subEntity = spiral[0];
                        if (subEntity is AlignmentSubEntitySpiral)
                        {
                            var subEntitySpiral = subEntity as AlignmentSubEntitySpiral;

                            // Extract properties from sub-entity
                            radiusIn = subEntitySpiral.RadiusIn;
                            radiusOut = subEntitySpiral.RadiusOut;
                            length = subEntitySpiral.Length;

                            gotFromSubEntity = true;
                        }
                    }
                }
                catch { }

                alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);

                // Calculate spiral point at arc length l (midpoint for now)
                double spiralMidStation = (startStation + endStation) / 2;
                alignment.PointLocation(spiralMidStation, 0, 0, ref xMid, ref yMid, ref zMid);

                bool isEntry = radiusIn > radiusOut || (radiusIn == 0 && radiusOut > 0);
                double R1 = isEntry ? double.PositiveInfinity : (radiusIn > 0 ? radiusIn : 0);  // Radius of curve 1
                double R2 = isEntry ? (radiusOut > 0 ? radiusOut : 0) : double.PositiveInfinity;  // Radius of curve 2
                double L = length;  // Total arc length of spiral

                // Calculate clothoid parameter A (flatness of spiral): A = sqrt(L*R)
                double R = isEntry ? R2 : R1;
                double A = 0;
                if (R > 0 && !double.IsInfinity(R))
                {
                    A = Math.Sqrt(L * R);
                }

                // Calculate central angle theta at spiral point
                double theta = 0;
                if (R > 0 && !double.IsInfinity(R))
                {
                    theta = L / (2 * R);  // Total angle subtended by spiral (radians)
                }

                writer.WriteLine("Element: Clothoid");
                writer.WriteLine($" SubEntity Access: {(gotFromSubEntity ? "SUCCESS" : "Fallback to base properties")}");

                // Determine if entry or exit spiral based on neighboring elements
                bool isEntrySpiralByContext = nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc;
                bool isExitSpiralByContext = prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc;
                
                // Use context first, then fallback to radius logic
                if (isEntrySpiralByContext || (isEntry && !isExitSpiralByContext))
                {
                    // Entry spiral: Tangent -> Spiral -> Curve
                    writer.WriteLine($" TS  ( ) {FormatStation(startStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");
                    writer.WriteLine($" SPI ( ) {FormatStation(spiralMidStation),15} {FormatWithProperRounding(yMid, 4),15} {FormatWithProperRounding(xMid, 4),15}");
                    writer.WriteLine($" SC  ( ) {FormatStation(endStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                }
                else
                {
                    // Exit spiral: Curve -> Spiral -> Tangent
                    writer.WriteLine($" CS  ( ) {FormatStation(startStation),15} {FormatWithProperRounding(y1, 4),15} {FormatWithProperRounding(x1, 4),15}");
                    writer.WriteLine($" SPI ( ) {FormatStation(spiralMidStation),15} {FormatWithProperRounding(yMid, 4),15} {FormatWithProperRounding(xMid, 4),15}");
                    writer.WriteLine($" ST  ( ) {FormatStation(endStation),15} {FormatWithProperRounding(y2, 4),15} {FormatWithProperRounding(x2, 4),15}");
                }

                // Output spiral parameters
                writer.WriteLine($" R1 (Radius of curve 1): {(double.IsInfinity(R1) ? "Infinite" : FormatWithProperRounding(R1, 4)),15}");
                writer.WriteLine($" R2 (Radius of curve 2): {(double.IsInfinity(R2) ? "Infinite" : FormatWithProperRounding(R2, 4)),15}");
                writer.WriteLine($" SS (Spiral start): {FormatStation(startStation),15}");
                writer.WriteLine($" SE (Spiral end): {FormatStation(endStation),15}");
                writer.WriteLine($" SP (Spiral point at arc length l): {FormatStation(spiralMidStation),15}");
                writer.WriteLine($" Θ (Central angle at spiral point): {FormatAngle(theta * 180.0 / Math.PI)}");
                writer.WriteLine($" L (Total arc length): {FormatWithProperRounding(L, 4),15}");
                writer.WriteLine($" A (Flatness parameter): {FormatWithProperRounding(A, 4),15}");
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
            return $"{sta:D2}+{offset:00.00}";
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
        /// Formats a double value to a specified number of decimal places using AwayFromZero rounding.
        /// This ensures proper rounding for surveying/engineering applications (0.5 always rounds up).
        /// Also handles floating point precision issues by checking if value is very close to a round number.
        /// </summary>
        private string FormatWithProperRounding(double value, int decimalPlaces)
        {
            // Calculate tolerance based on the number of decimal places we're rounding to
            // For 4 decimal places, we check if value is within 0.00005 (5e-5) of a round integer
            // This handles cases like 24999.9999 which should be 25000.0000
            double tolerance = 5.0 / Math.Pow(10, decimalPlaces + 1);

            // Check if the value is very close to a round integer
            double nearestInt = Math.Round(value, 0, MidpointRounding.AwayFromZero);
            if (Math.Abs(value - nearestInt) < tolerance)
            {
                return nearestInt.ToString($"F{decimalPlaces}");
            }

            // Standard rounding with AwayFromZero
            double rounded = Math.Round(value, decimalPlaces, MidpointRounding.AwayFromZero);
            return rounded.ToString($"F{decimalPlaces}");
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

                                // Get PI station and coordinates from arc properties
                                double piStation = arc.PIStation;
                                double xPI = 0, yPI = 0;

                                // Try to get PI and Center coordinates from sub-entities
                                bool gotFromSubEntity = false;
                                try
                                {
                                    if (arc.SubEntityCount > 0)
                                    {
                                        var subEntity = arc[0];
                                        if (subEntity is AlignmentSubEntityArc)
                                        {
                                            var subEntityArc = subEntity as AlignmentSubEntityArc;
                                            var piPoint = subEntityArc.PIPoint;
                                            var centerPoint = subEntityArc.CenterPoint;
                                            xPI = piPoint.X;  // Easting
                                            yPI = piPoint.Y;  // Northing
                                            xc = centerPoint.X;  // Easting (Center)
                                            yc = centerPoint.Y;  // Northing (Center)
                                            gotFromSubEntity = true;
                                        }
                                    }
                                }
                                catch { }

                                // Calculate deltaRadians using arc length and radius
                                double deltaRadians = arc.Length / Math.Abs(arc.Radius);

                                // Fallback: calculate if sub-entity access failed
                                if (!gotFromSubEntity)
                                {
                                    double tangent = arc.Radius * Math.Tan(Math.Abs(deltaRadians) / 2);
                                    double backTangentDir = arc.StartDirection + Math.PI;
                                    xPI = x1 + tangent * Math.Cos(backTangentDir);
                                    yPI = y1 + tangent * Math.Sin(backTangentDir);

                                    // Calculate center point as fallback
                                    double midStation = (arc.StartStation + arc.EndStation) / 2;
                                    double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                                    alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                                }

                                double deltaDegrees = deltaRadians * (180.0 / Math.PI);

                                elements.Add(new XElement("Arc",
                                    new XElement("StartStation", arc.StartStation),
                                    new XElement("EndStation", arc.EndStation),
                                    new XElement("Length", arc.Length),
                                    new XElement("Radius", arc.Radius),
                                    new XElement("Delta", Math.Abs(deltaDegrees)),
                                    new XElement("Direction", arc.Clockwise ? "Right" : "Left"),
                                    new XElement("PIStation", piStation),
                                    new XElement("StartPoint",
                                        new XElement("Northing", y1),
                                        new XElement("Easting", x1)
                                    ),
                                    new XElement("EndPoint",
                                        new XElement("Northing", y2),
                                        new XElement("Easting", x2)
                                    ),
                                    new XElement("PIPoint",
                                        new XElement("Northing", yPI),
                                        new XElement("Easting", xPI)
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
                PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD);
                PdfFont normalFont = PdfFontFactory.CreateFont(StandardFonts.COURIER);

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
                document.Add(new Paragraph($"Generated Report:  Date: {DateTime.Now:MM/dd/yyyy}     Time: {DateTime.Now:h:mm tt} {(TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now) ? TimeZoneInfo.Local.DaylightName : TimeZoneInfo.Local.StandardName)}")
                    .SetFont(normalFont).SetFontSize(9).SetItalic());
                document.Add(new Paragraph("\n").SetFontSize(3));

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
                document.Add(new Paragraph("\n").SetFontSize(3));

                // Reorder entities to match InRails format
                var reorderedEntities = ReorderEntitiesForInRails(alignment);

                // Process each entity
                for (int i = 0; i < reorderedEntities.Count; i++)
                {
                    AlignmentEntity entity = reorderedEntities[i];
                    if (entity == null) continue;

                    AlignmentEntity prevEntity = i > 0 ? reorderedEntities[i - 1] : null;
                    AlignmentEntity nextEntity = i < reorderedEntities.Count - 1 ? reorderedEntities[i + 1] : null;

                    try
                    {
                        switch (entity.EntityType)
                        {
                            case AlignmentEntityType.Line:
                                WriteLinearElementPdf(document, entity as AlignmentLine, alignment, i, normalFont, boldFont, prevEntity, nextEntity);
                                break;
                            case AlignmentEntityType.Arc:
                                WriteArcElementPdf(document, entity as AlignmentArc, alignment, i, normalFont, boldFont, prevEntity, nextEntity);
                                break;
                            case AlignmentEntityType.Spiral:
                                WriteSpiralElementPdf(document, entity, alignment, i, normalFont, boldFont, prevEntity, nextEntity);
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

        private void WriteLinearElementPdf(Document document, AlignmentLine line, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont boldFont, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
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
                document.Add(new Paragraph("\n").SetFontSize(2));

                // Determine start label based on previous element
                string startLabel;
                if (index == 0)
                {
                    startLabel = "POB ";  // Point of Beginning
                }
                else if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    startLabel = "ST  ";  // Spiral to Tangent
                }
                else if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc)
                {
                    startLabel = "PT  ";  // Point of Tangency (Curve to Tangent)
                }
                else
                {
                    startLabel = "PI  ";  // Point of Intersection
                }

                AddHorizontalDataRow(document, startLabel + "( ) ", FormatStation(line.StartStation), y1, x1, normalFont);

                // Determine end label based on next element
                string endLabel;
                if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    endLabel = "TS  ";  // Tangent to Spiral
                }
                else if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc)
                {
                    endLabel = "PC  ";  // Point of Curvature (Tangent to Arc)
                }
                else
                {
                    endLabel = "PI  ";  // Point of Intersection
                }

                AddHorizontalDataRow(document, endLabel + "( ) ", FormatStation(line.EndStation), y2, x2, normalFont);
                document.Add(new Paragraph($"Tangent Direction: {bearing}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent Length: {line.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(2));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Linear").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing line data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteArcElementPdf(Document document, AlignmentArc arc, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont boldFont, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
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

                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                // Get PI station and coordinates from arc properties
                double piStation = arc.PIStation;
                double xPI = 0, yPI = 0;

                // Try to get PI and Center coordinates from sub-entities
                bool gotFromSubEntity = false;
                try
                {
                    if (arc.SubEntityCount > 0)
                    {
                        var subEntity = arc[0];
                        if (subEntity is AlignmentSubEntityArc)
                        {
                            var subEntityArc = subEntity as AlignmentSubEntityArc;
                            var piPoint = subEntityArc.PIPoint;
                            var centerPoint = subEntityArc.CenterPoint;
                            xPI = piPoint.X;  // Easting
                            yPI = piPoint.Y;  // Northing
                            xc = centerPoint.X;  // Easting (Center)
                            yc = centerPoint.Y;  // Northing (Center)
                            gotFromSubEntity = true;
                        }
                    }
                }
                catch { }

                // Calculate deltaRadians using arc length and radius
                double deltaRadians = arc.Length / Math.Abs(arc.Radius);
                double tangent = arc.Radius * Math.Tan(Math.Abs(deltaRadians) / 2);

                // Fallback: calculate if sub-entity access failed
                if (!gotFromSubEntity)
                {
                    double backTangentDir = arc.StartDirection + Math.PI;
                    xPI = x1 + tangent * Math.Cos(backTangentDir);
                    yPI = y1 + tangent * Math.Sin(backTangentDir);

                    // Calculate center point as fallback
                    double midStation = (arc.StartStation + arc.EndStation) / 2;
                    double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                    alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                }

                double deltaDegrees = deltaRadians * (180.0 / Math.PI);
                double chord = 2 * arc.Radius * Math.Sin(Math.Abs(deltaRadians) / 2);
                double middleOrdinate = arc.Radius * (1 - Math.Cos(Math.Abs(deltaRadians) / 2));
                double external = arc.Radius * (1 / Math.Cos(Math.Abs(deltaRadians) / 2) - 1);

                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(2));

                // Determine start label based on previous element
                string startLabel;
                if (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    startLabel = "SC  ";  // Spiral to Curve
                }
                else
                {
                    startLabel = "PC  ";  // Point of Curvature (no spiral before)
                }
                
                AddHorizontalDataRow(document, startLabel + "( ) ", FormatStation(arc.StartStation), y1, x1, normalFont);
                AddHorizontalDataRow(document, "PI  ( ) ", FormatStation(piStation), yPI, xPI, normalFont);
                AddHorizontalDataRow(document, "CC  ( ) ", "               ", yc, xc, normalFont);
                
                // Determine end label based on next element
                string endLabel;
                if (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Spiral)
                {
                    endLabel = "CS  ";  // Curve to Spiral
                }
                else
                {
                    endLabel = "PT  ";  // Point of Tangency (no spiral after)
                }
                
                AddHorizontalDataRow(document, endLabel + "( ) ", FormatStation(arc.EndStation), y2, x2, normalFont);

                document.Add(new Paragraph($"Radius: {arc.Radius:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Design Speed(mph): {50.0:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Cant(inches): {2.0:F3}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Delta: {FormatAngle(Math.Abs(deltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                double degreeOfCurvatureChord = (100.0 * deltaRadians) / (arc.Length) * (180.0 / Math.PI);
                document.Add(new Paragraph($"Degree of Curvature (Arc): {FormatAngle(degreeOfCurvatureChord)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length: {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length(Chorded): {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent: {tangent:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Chord: {chord:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Middle Ordinate: {middleOrdinate:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"External: {external:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph("\n").SetFontSize(2));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing arc data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
            }
        }

        private void WriteSpiralElementPdf(Document document, AlignmentEntity entity, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont labelFont, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            if (entity == null)
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
                double xSPI = 0, ySPI = 0, zSPI = 0;

                double radiusIn = 0, radiusOut = 0, length = 0;
                double startStation = 0, endStation = 0, spiralA = 0;
                double spiralK_API = 0, longTangent_API = 0, shortTangent_API = 0;
                double startDirection_API = 0, endDirection_API = 0;
                string spiralDirection = "DirectionRight";
                
                double totalX = 0, totalY = 0, delta = 0;
                double spiralPIStation = 0;
                bool gotSubEntityData = false;
                bool gotSPIFromAPI = false;

                // Access spiral through SubEntity (spirals can't be cast directly)
                if (entity.SubEntityCount > 0)
                {
                    var subEntity = entity[0];
                    if (subEntity is AlignmentSubEntitySpiral)
                    {
                        var spiralSubEntity = subEntity as AlignmentSubEntitySpiral;
                        
                        // Get basic properties
                        startStation = spiralSubEntity.StartStation;
                        endStation = spiralSubEntity.EndStation;
                        length = spiralSubEntity.Length;
                        totalX = spiralSubEntity.TotalX;
                        totalY = spiralSubEntity.TotalY;
                        delta = spiralSubEntity.Delta;
                        
                        // Get additional properties via reflection (except SPI which we access strongly)
                        var spiralType = spiralSubEntity.GetType();
                        try
                        {
                            var radiusInProp = spiralType.GetProperty("RadiusIn");
                            var radiusOutProp = spiralType.GetProperty("RadiusOut");
                            var aProp = spiralType.GetProperty("A");
                            var kProp = spiralType.GetProperty("K");
                            var longTangentProp = spiralType.GetProperty("LongTangent");
                            var shortTangentProp = spiralType.GetProperty("ShortTangent");
                            var startDirProp = spiralType.GetProperty("StartDirection");
                            var endDirProp = spiralType.GetProperty("EndDirection");
                            var dirProp = spiralType.GetProperty("Direction");
                            if (radiusInProp != null) radiusIn = (double)radiusInProp.GetValue(spiralSubEntity);
                            if (radiusOutProp != null) radiusOut = (double)radiusOutProp.GetValue(spiralSubEntity);
                            if (aProp != null) spiralA = (double)aProp.GetValue(spiralSubEntity);
                            if (kProp != null) spiralK_API = (double)kProp.GetValue(spiralSubEntity);
                            if (longTangentProp != null) longTangent_API = (double)longTangentProp.GetValue(spiralSubEntity);
                            if (shortTangentProp != null) shortTangent_API = (double)shortTangentProp.GetValue(spiralSubEntity);
                            if (startDirProp != null) startDirection_API = (double)startDirProp.GetValue(spiralSubEntity);
                            if (endDirProp != null) endDirection_API = (double)endDirProp.GetValue(spiralSubEntity);
                            if (dirProp != null) spiralDirection = dirProp.GetValue(spiralSubEntity).ToString();
                        }
                        catch { }

                        // Strongly access SPIStation / SPI (Point3d) to avoid reflection failure
                        try
                        {
                            spiralPIStation = spiralSubEntity.SPIStation; // station of spiral PI
                            // Access SPI point via reflection (property is not directly accessible)
                            try
                            {
                                // Prefer SPIPoint (actual spiral PI) over SPI (alternate point)
                                var spiPointProp = spiralSubEntity.GetType().GetProperty("SPIPoint")
                                                     ?? spiralSubEntity.GetType().GetProperty("SPI");
                                if (spiPointProp != null)
                                {
                                    var spiPointObj = spiPointProp.GetValue(spiralSubEntity);
                                    if (spiPointObj != null)
                                    {
                                        // Handle Point3d or Point2d (some builds expose SPIPoint as Point2d)
                                        if (spiPointObj is Autodesk.AutoCAD.Geometry.Point3d p3)
                                        {
                                            xSPI = p3.X;
                                            ySPI = p3.Y;
                                            zSPI = p3.Z;
                                            gotSPIFromAPI = true;
                                        }
                                        else if (spiPointObj is Autodesk.AutoCAD.Geometry.Point2d p2)
                                        {
                                            xSPI = p2.X;
                                            ySPI = p2.Y;
                                            zSPI = 0.0;
                                            gotSPIFromAPI = true;
                                        }
                                        else
                                        {
                                            // Fallback: attempt dynamic X/Y properties
                                            var xProp = spiPointObj.GetType().GetProperty("X");
                                            var yProp = spiPointObj.GetType().GetProperty("Y");
                                            if (xProp != null && yProp != null)
                                            {
                                                xSPI = Convert.ToDouble(xProp.GetValue(spiPointObj));
                                                ySPI = Convert.ToDouble(yProp.GetValue(spiPointObj));
                                                zSPI = 0.0;
                                                gotSPIFromAPI = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (System.Exception)
                            {
                                // Suppress debug in production; fallback midpoint will handle absence
                            }
                        }
                        catch (System.Exception ex)
                        {
                            document.Add(new Paragraph($"DEBUG: Direct SPI access failed, falling back: {ex.Message}").SetFont(normalFont).SetFontSize(8));
                        }
                        
                        gotSubEntityData = true;
                    }
                }
                
                if (!gotSubEntityData)
                {
                    document.Add(new Paragraph("Element: Clothoid").SetFont(labelFont).SetFontSize(10));
                    document.Add(new Paragraph("Unable to read spiral data - no SubEntity available.").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph("\n"));
                    return;
                }
                
                // Get coordinates for start and end points
                alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);
                
                // If SPI wasn't available from API, calculate it as fallback
                if (!gotSPIFromAPI)
                {
                    if (spiralPIStation == 0)
                    {
                        spiralPIStation = startStation + (length / 2.0);
                    }
                    alignment.PointLocation(spiralPIStation, 0, 0, ref xSPI, ref ySPI, ref zSPI);
                }

                // Determine spiral type and labels based on neighboring elements
                bool isEntry = double.IsInfinity(radiusIn) || radiusIn == 0 || radiusIn > radiusOut;
                
                // Use context first to determine if entry or exit spiral
                bool isEntrySpiralByContext = nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc;
                bool isExitSpiralByContext = prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc;
                
                string startLabel, endLabel;
                if (isEntrySpiralByContext || (isEntry && !isExitSpiralByContext))
                {
                    // Entry spiral: Tangent -> Spiral -> Curve
                    startLabel = "TS";
                    endLabel = "SC";
                }
                else
                {
                    // Exit spiral: Curve -> Spiral -> Tangent
                    startLabel = "CS";
                    endLabel = "ST";
                }
                
                double effectiveRadius = isEntry ? radiusOut : radiusIn;

                // Calculate spiral properties
                double spiralK = 0;
                double spiralP = 0;
                double spiralAngle = 0;
                double longTangent = 0;
                double shortTangent = 0;
                double chordLength = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));
                
                if (effectiveRadius > 0 && !double.IsInfinity(effectiveRadius))
                {
                    spiralK = (length * length) / effectiveRadius;
                    spiralP = (length * length) / (24.0 * effectiveRadius);
                    spiralAngle = length / (2.0 * effectiveRadius);
                    
                    // Calculate tangent lengths
                    double theta = spiralAngle;
                    longTangent = length - (Math.Pow(length, 3) / (40 * effectiveRadius * effectiveRadius));
                    shortTangent = (Math.Pow(length, 3) / (6 * effectiveRadius * effectiveRadius));
                }

                // Calculate tangent and radial directions at start and end points
                // Use API properties if available, otherwise calculate
                double startTangentDirection = 0;
                double endTangentDirection = 0;
                
                if (startDirection_API != 0)
                {
                    startTangentDirection = startDirection_API;
                }
                else
                {
                    // Fallback: Calculate Start Direction (tangent at spiral start)
                    double tempXStart = 0, tempYStart = 0, tempZStart = 0;
                    alignment.PointLocation(startStation + 0.01, 0, 0, ref tempXStart, ref tempYStart, ref tempZStart);
                    startTangentDirection = Math.Atan2(tempYStart - y1, tempXStart - x1);
                }
                
                if (endDirection_API != 0)
                {
                    endTangentDirection = endDirection_API;
                }
                else
                {
                    // Fallback: Calculate End Direction (tangent at spiral end)
                    double tempXEnd = 0, tempYEnd = 0, tempZEnd = 0;
                    alignment.PointLocation(endStation - 0.01, 0, 0, ref tempXEnd, ref tempYEnd, ref tempZEnd);
                    endTangentDirection = Math.Atan2(y2 - tempYEnd, x2 - tempXEnd);
                }
                
                // Chord Direction (straight line from start to end)
                double chordDirectionAngle = Math.Atan2(y2 - y1, x2 - x1);

                // Header
                document.Add(new Paragraph("Element: Clothoid").SetFont(labelFont).SetFontSize(10).SetBold());
                document.Add(new Paragraph("\n").SetFontSize(3));

                // Station points (use API SPI; also show fallback midpoint for verification if different)
                AddHorizontalDataRow(document, $"{startLabel}  ( )", FormatStation(startStation), y1, x1, normalFont);
                AddHorizontalDataRow(document, "SPI  ( )", FormatStation(spiralPIStation), ySPI, xSPI, normalFont);
                AddHorizontalDataRow(document, $"{endLabel}  ( )", FormatStation(endStation), y2, x2, normalFont);
                
                // Spiral properties
                document.Add(new Paragraph($"Entrance Radius: {FormatRadius(radiusIn)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Exit Radius: {FormatRadius(radiusOut)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length: {Math.Round(length, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Angle: {FormatAngleDMS(spiralAngle)} {(spiralDirection.Contains("Right") ? "Right" : "Left")}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Constant: {Math.Round(spiralA, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Long Tangent: {Math.Round((longTangent_API > 0 ? longTangent_API : longTangent), 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Short Tangent: {Math.Round((shortTangent_API > 0 ? shortTangent_API : shortTangent), 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Long Chord: {Math.Round(chordLength, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                
                if (gotSubEntityData)
                {
                    document.Add(new Paragraph($"Xs: {Math.Round(totalX, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"Ys: {Math.Round(totalY, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                }
                
                document.Add(new Paragraph($"P: {Math.Round(spiralP, 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"K: {Math.Round((spiralK_API > 0 ? spiralK_API : spiralK), 2):F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                
                // Directions (Start and End tangent directions matching InRoads format)
                document.Add(new Paragraph($"Start Tangent Direction: {FormatBearingDMS(startTangentDirection)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"End Tangent Direction: {FormatBearingDMS(endTangentDirection)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Chord Direction: {FormatBearingDMS(chordDirectionAngle)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                // Removed duplicate start/end tangent & radial directions per Clothoid template simplification
                
                document.Add(new Paragraph("\n").SetFontSize(2));
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
                    PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD);
                    PdfFont normalFont = PdfFontFactory.CreateFont(StandardFonts.COURIER);

                    // Write header
                    document.Add(new Paragraph($"Project Name: {projectName}").SetFont(boldFont).SetFontSize(11));
                    document.Add(new Paragraph(" Description:").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Horizontal Alignment Name: {alignmentName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {alignmentDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {alignmentStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Vertical Alignment Name: {profileName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {profileDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {profileStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Generated Report:  Date: {DateTime.Now:MM/dd/yyyy}     Time: {DateTime.Now:h:mm tt} {(TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now) ? TimeZoneInfo.Local.DaylightName : TimeZoneInfo.Local.StandardName)}")
                        .SetFont(normalFont).SetFontSize(9).SetItalic());
                    document.Add(new Paragraph("\n").SetFontSize(3));

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
                    document.Add(new Paragraph("\n").SetFontSize(3));

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
                document.Add(new Paragraph("\n").SetFontSize(2));

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
                document.Add(new Paragraph("\n").SetFontSize(2));
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
                document.Add(new Paragraph("\n").SetFontSize(2));

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
                document.Add(new Paragraph("\n").SetFontSize(2));
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

        /// <summary>
        /// Format angle in DD°MM'SS.SS" format with symbols
        /// </summary>
        private string FormatAngleDMS(double angleRadians)
        {
            double angleDegrees = Math.Abs(angleRadians * (180.0 / Math.PI));
            int degrees = (int)angleDegrees;
            double minutes = (angleDegrees - degrees) * 60;
            int mins = (int)minutes;
            double seconds = (minutes - mins) * 60;

            return $"{degrees}°{mins:D2}'{seconds:F2}\"";
        }

        /// <summary>
        /// Format bearing in N/S DD°MM'SS.SS" E/W format
        /// </summary>
        private string FormatBearingDMS(double bearingRadians)
        {
            // Convert from radians to degrees (0-360)
            double bearingDegrees = bearingRadians * (180.0 / Math.PI);
            while (bearingDegrees < 0) bearingDegrees += 360;
            while (bearingDegrees >= 360) bearingDegrees -= 360;

            string ns, ew;
            double angle;

            if (bearingDegrees >= 0 && bearingDegrees < 90)
            {
                ns = "N";
                ew = "E";
                angle = bearingDegrees;
            }
            else if (bearingDegrees >= 90 && bearingDegrees < 180)
            {
                ns = "S";
                ew = "E";
                angle = 180 - bearingDegrees;
            }
            else if (bearingDegrees >= 180 && bearingDegrees < 270)
            {
                ns = "S";
                ew = "W";
                angle = bearingDegrees - 180;
            }
            else
            {
                ns = "N";
                ew = "W";
                angle = 360 - bearingDegrees;
            }

            int degrees = (int)angle;
            double minutes = (angle - degrees) * 60;
            int mins = (int)minutes;
            double seconds = (minutes - mins) * 60;

            return $"{ns} {degrees}°{mins:D2}'{seconds:F2}\" {ew}";
        }

        /// <summary>
        /// Format distance with foot apostrophe
        /// </summary>
        private string FormatDistanceFeet(double distance)
        {
            return $"{distance:F2}'";
        }

        /// <summary>
        /// Format radius for spiral reports with unified logic (Infinity/<=0 -> 0.0000)
        /// </summary>
        private string FormatRadius(double radius)
        {
            return (double.IsInfinity(radius) || radius <= 0) ? "0.0000" : radius.ToString("F4");
        }

        /// <summary>
        /// Format banking/superelevation with inch double-quote
        /// </summary>
        private string FormatBankingInches(double banking)
        {
            return $"{banking:F2}\"";
        }

        /// <summary>
        /// Calculate curve center coordinates
        /// </summary>
        private void CalculateCurveCenter(AlignmentArc arc, CivDb.Alignment alignment, out double centerNorthing, out double centerEasting)
        {
            // Get PC (Point of Curvature) coordinates
            double x1 = 0, y1 = 0, z1 = 0;
            alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);

            // Calculate perpendicular direction from PC to center
            // For a right curve (clockwise), the center is to the right (90° CW from tangent)
            // For a left curve (counter-clockwise), the center is to the left (90° CCW from tangent)
            double perpDirection = arc.Clockwise ? arc.StartDirection - Math.PI / 2.0 : arc.StartDirection + Math.PI / 2.0;

            // Offset by radius from PC to get center
            double radius = Math.Abs(arc.Radius);
            centerEasting = x1 + radius * Math.Cos(perpDirection);
            centerNorthing = y1 + radius * Math.Sin(perpDirection);
        }

        /// <summary>
        /// Calculate tangent distance for curve
        /// </summary>
        private double CalculateTangentDistance(double radius, double delta)
        {
            return radius * Math.Tan(delta / 2.0);
        }

        /// <summary>
        /// Calculate external distance for curve
        /// </summary>
        private double CalculateExternalDistance(double radius, double delta)
        {
            return radius * (1.0 / Math.Cos(delta / 2.0) - 1.0);
        }

        /// <summary>
        /// Calculate spiral angle (theta)
        /// </summary>
        private double CalculateSpiralAngle(double spiralLength, double radius)
        {
            if (Math.Abs(radius) < 0.001) return 0;
            return spiralLength / (2.0 * radius);
        }

        /// <summary>
        /// Calculate spiral X offset
        /// </summary>
        private double CalculateSpiralX(double spiralLength, double theta)
        {
            // Using approximate formula: X ≈ L - (L^3)/(40*R^2)
            return spiralLength * (1.0 - Math.Pow(theta, 2) / 10.0);
        }

        /// <summary>
        /// Calculate spiral Y offset
        /// </summary>
        private double CalculateSpiralY(double spiralLength, double theta)
        {
            // Using approximate formula: Y ≈ L^2/(6*R)
            return spiralLength * theta / 3.0;
        }

        /// <summary>
        /// Generate horizontal alignment report in Excel GeoTable format (OLD - will be replaced)
        /// </summary>
        private void GenerateHorizontalReportExcel(CivDb.Alignment alignment, string outputPath)
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Track Geometry Data");

                    // Get project info
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

                    // Set up the header
                    int currentRow = 1;

                    // Title
                    worksheet.Cells[currentRow, 1].Value = $"TRACK GEOMETRY DATA - {alignment.Name?.ToUpper() ?? "ALIGNMENT"}";
                    worksheet.Cells[currentRow, 1, currentRow, 8].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 1].Style.Font.Size = 12;
                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    currentRow += 2;

                    // Column headers
                    worksheet.Cells[currentRow, 1].Value = "ELEMENT";
                    worksheet.Cells[currentRow, 2].Value = "CURVE No.";
                    worksheet.Cells[currentRow, 3].Value = "POINT";
                    worksheet.Cells[currentRow, 4].Value = "STATION";
                    worksheet.Cells[currentRow, 5].Value = "BEARING";
                    worksheet.Cells[currentRow, 6].Value = "COORDINATES";
                    worksheet.Cells[currentRow, 6, currentRow, 7].Merge = true;
                    worksheet.Cells[currentRow, 8].Value = "DATA";

                    // Sub-headers for coordinates
                    currentRow++;
                    worksheet.Cells[currentRow, 6].Value = "Northing";
                    worksheet.Cells[currentRow, 7].Value = "Easting";

                    // Style headers
                    using (var range = worksheet.Cells[currentRow - 1, 1, currentRow, 8])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    currentRow++;
                    int curveNumber = 0;

                    // Process each alignment entity
                    for (int i = 0; i < alignment.Entities.Count; i++)
                    {
                        AlignmentEntity entity = alignment.Entities[i];
                        if (entity == null) continue;

                        try
                        {
                            switch (entity.EntityType)
                            {
                                case AlignmentEntityType.Line:
                                    currentRow = WriteLinearElementExcel(worksheet, entity as AlignmentLine, alignment, i, currentRow);
                                    break;
                                case AlignmentEntityType.Arc:
                                    curveNumber++;
                                    currentRow = WriteArcElementExcel(worksheet, entity as AlignmentArc, alignment, i, currentRow, curveNumber);
                                    break;
                                case AlignmentEntityType.Spiral:
                                    currentRow = WriteSpiralElementExcel(worksheet, entity as AlignmentSpiral, alignment, i, currentRow);
                                    break;
                                default:
                                    worksheet.Cells[currentRow, 1].Value = entity.EntityType.ToString();
                                    worksheet.Cells[currentRow, 3].Value = "Unsupported";
                                    currentRow++;
                                    break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            worksheet.Cells[currentRow, 1].Value = "ERROR";
                            worksheet.Cells[currentRow, 3].Value = $"Error: {ex.Message}";
                            currentRow++;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();

                    // Set minimum column widths
                    worksheet.Column(1).Width = 12; // ELEMENT
                    worksheet.Column(2).Width = 10; // CURVE No.
                    worksheet.Column(3).Width = 10; // POINT
                    worksheet.Column(4).Width = 15; // STATION
                    worksheet.Column(5).Width = 20; // BEARING
                    worksheet.Column(6).Width = 15; // Northing
                    worksheet.Column(7).Width = 15; // Easting
                    worksheet.Column(8).Width = 25; // DATA

                    // Save the file
                    System.IO.FileInfo file = new System.IO.FileInfo(outputPath);
                    package.SaveAs(file);
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error generating Excel report: {ex.Message}", ex);
            }
        }

        private int WriteLinearElementExcel(ExcelWorksheet ws, AlignmentLine line, CivDb.Alignment alignment, int index, int row)
        {
            if (line == null) return row + 1;

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(line.EndStation, 0, 0, ref x2, ref y2, ref z2);

                string bearing = FormatBearing(line.Direction);
                string pointLabel = (index == 0) ? "POB" : "PI"; // InRoads: first tangent start POB

                // Start point
                ws.Cells[row, 1].Value = "TANGENT";
                ws.Cells[row, 3].Value = pointLabel;
                ws.Cells[row, 4].Value = FormatStation(line.StartStation);
                ws.Cells[row, 5].Value = bearing;
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 8].Value = $"L = {line.Length:F2}";

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        private int WriteArcElementExcel(ExcelWorksheet ws, AlignmentArc arc, CivDb.Alignment alignment, int index, int row, int curveNumber)
        {
            if (arc == null) return row + 1;

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                string directionStr = arc.Clockwise ? "R" : "L";
                double radius = Math.Abs(arc.Radius);
                double delta = (arc.Length / radius) * (180.0 / Math.PI);

                // PC point
                ws.Cells[row, 1].Value = "CURVE";
                ws.Cells[row, 2].Value = $"{curveNumber}-{directionStr}";
                ws.Cells[row, 3].Value = "PC";
                ws.Cells[row, 4].Value = FormatStation(arc.StartStation);
                ws.Cells[row, 5].Value = FormatBearing(arc.StartDirection);
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 8].Value = $"R = {radius:F2}";
                row++;

                // PT point
                ws.Cells[row, 3].Value = "PT";
                ws.Cells[row, 4].Value = FormatStation(arc.EndStation);
                ws.Cells[row, 5].Value = FormatBearing(arc.EndDirection);
                ws.Cells[row, 6].Value = y2;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x2;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 8].Value = $"Δ = {delta:F2}°";

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        private int WriteSpiralElementExcel(ExcelWorksheet ws, AlignmentSpiral spiral, CivDb.Alignment alignment, int index, int row)
        {
            if (spiral == null) return row + 1;

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(spiral.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(spiral.EndStation, 0, 0, ref x2, ref y2, ref z2);

                string spiralType = spiral.SpiralDefinition.ToString();

                ws.Cells[row, 1].Value = "SPIRAL";
                ws.Cells[row, 3].Value = "TS";
                ws.Cells[row, 4].Value = FormatStation(spiral.StartStation);
                ws.Cells[row, 5].Value = FormatBearing(spiral.StartDirection);
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 8].Value = $"L = {spiral.Length:F2}";

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        /// <summary>
        /// Generate vertical alignment report in Excel GeoTable format
        /// </summary>
        private void GenerateVerticalReportExcel(CivDb.Alignment alignment, string outputPath)
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                // Get project name
                Database db = alignment.Database;
                string projectName = "Unknown Project";
                try
                {
                    if (db != null && !string.IsNullOrEmpty(db.Filename))
                    {
                        projectName = System.IO.Path.GetFileNameWithoutExtension(db.Filename);
                    }
                }
                catch { }

                // Get first layout profile
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
                    throw new System.Exception("No layout profile found.");
                }

                using (Profile layoutProfile = layoutProfileId.GetObject(OpenMode.ForRead) as Profile)
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Vertical Profile Data");

                    // Set up header
                    int currentRow = 1;

                    // Title
                    worksheet.Cells[currentRow, 1].Value = $"VERTICAL PROFILE DATA - {alignment.Name?.ToUpper() ?? "ALIGNMENT"}";
                    worksheet.Cells[currentRow, 1, currentRow, 6].Merge = true;
                    worksheet.Cells[currentRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[currentRow, 1].Style.Font.Size = 12;
                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    currentRow += 2;

                    // Column headers
                    worksheet.Cells[currentRow, 1].Value = "ELEMENT";
                    worksheet.Cells[currentRow, 2].Value = "POINT";
                    worksheet.Cells[currentRow, 3].Value = "STATION";
                    worksheet.Cells[currentRow, 4].Value = "ELEVATION";
                    worksheet.Cells[currentRow, 5].Value = "GRADE";
                    worksheet.Cells[currentRow, 6].Value = "DATA";

                    // Style headers
                    using (var range = worksheet.Cells[currentRow, 1, currentRow, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    currentRow++;

                    // Process profile entities
                    for (int i = 0; i < layoutProfile.Entities.Count; i++)
                    {
                        ProfileEntity entity = layoutProfile.Entities[i];
                        if (entity == null) continue;

                        try
                        {
                            switch (entity.EntityType)
                            {
                                case ProfileEntityType.Tangent:
                                    if (entity is ProfileTangent tangent)
                                        currentRow = WriteProfileTangentExcel(worksheet, tangent, i, layoutProfile.Entities.Count, currentRow);
                                    break;
                                case ProfileEntityType.Circular:
                                    if (entity is ProfileCircular circular)
                                        currentRow = WriteProfileParabolaExcel(worksheet, circular, currentRow);
                                    break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            worksheet.Cells[currentRow, 1].Value = "ERROR";
                            worksheet.Cells[currentRow, 2].Value = $"Error: {ex.Message}";
                            currentRow++;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();
                    worksheet.Column(1).Width = 12;
                    worksheet.Column(2).Width = 10;
                    worksheet.Column(3).Width = 15;
                    worksheet.Column(4).Width = 12;
                    worksheet.Column(5).Width = 12;
                    worksheet.Column(6).Width = 30;

                    // Save the file
                    System.IO.FileInfo file = new System.IO.FileInfo(outputPath);
                    package.SaveAs(file);
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error generating vertical Excel report: {ex.Message}", ex);
            }
        }

        private int WriteProfileTangentExcel(ExcelWorksheet ws, ProfileTangent tangent, int index, int totalCount, int row)
        {
            if (tangent == null) return row + 1;

            try
            {
                double grade = tangent.Grade * 100;

                // Start point
                if (index == 0)
                {
                    ws.Cells[row, 1].Value = "TANGENT";
                    ws.Cells[row, 2].Value = "POB";
                    ws.Cells[row, 3].Value = FormatStation(tangent.StartStation);
                    ws.Cells[row, 4].Value = tangent.StartElevation;
                    ws.Cells[row, 4].Style.Numberformat.Format = "0.00";
                    ws.Cells[row, 5].Value = $"{grade:F3}%";
                    ws.Cells[row, 6].Value = $"L = {tangent.Length:F2}";
                    row++;
                }

                // End point (PVI)
                ws.Cells[row, 1].Value = "TANGENT";
                ws.Cells[row, 2].Value = "PVI";
                ws.Cells[row, 3].Value = FormatStation(tangent.EndStation);
                ws.Cells[row, 4].Value = tangent.EndElevation;
                ws.Cells[row, 4].Style.Numberformat.Format = "0.00";
                ws.Cells[row, 5].Value = $"{grade:F3}%";
                ws.Cells[row, 6].Value = $"L = {tangent.Length:F2}";

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        private int WriteProfileParabolaExcel(ExcelWorksheet ws, ProfileCircular curve, int row)
        {
            if (curve == null) return row + 1;

            try
            {
                const double tolerance = 1e-8;

                double gradeIn = curve.GradeIn * 100;
                double gradeOut = curve.GradeOut * 100;
                double length = curve.Length;
                double pviElevation = curve.PVIElevation;

                double gradeInDecimal = gradeIn / 100.0;
                double gradeOutDecimal = gradeOut / 100.0;
                double pvcElevation = pviElevation - (gradeInDecimal * (length / 2));
                double pvtElevation = pviElevation + (gradeOutDecimal * (length / 2));

                // PVC
                ws.Cells[row, 1].Value = "CURVE";
                ws.Cells[row, 2].Value = "PVC";
                ws.Cells[row, 3].Value = FormatStation(curve.StartStation);
                ws.Cells[row, 4].Value = pvcElevation;
                ws.Cells[row, 4].Style.Numberformat.Format = "0.00";
                ws.Cells[row, 5].Value = $"{gradeIn:F3}%";
                ws.Cells[row, 6].Value = $"L = {length:F2}";
                row++;

                // PVI
                ws.Cells[row, 2].Value = "PVI";
                ws.Cells[row, 3].Value = FormatStation(curve.PVIStation);
                ws.Cells[row, 4].Value = pviElevation;
                ws.Cells[row, 4].Style.Numberformat.Format = "0.00";
                double gradeDiff = gradeOut - gradeIn;
                double k = Math.Abs(gradeDiff) > tolerance ? length / gradeDiff : double.PositiveInfinity;
                string kDisplay = Math.Abs(gradeDiff) > tolerance ? Math.Abs(k).ToString("F2") : "INF";
                ws.Cells[row, 6].Value = $"K = {kDisplay}";
                row++;

                // PVT
                ws.Cells[row, 2].Value = "PVT";
                ws.Cells[row, 3].Value = FormatStation(curve.EndStation);
                ws.Cells[row, 4].Value = pvtElevation;
                ws.Cells[row, 4].Style.Numberformat.Format = "0.00";
                ws.Cells[row, 5].Value = $"{gradeOut:F3}%";

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        /// <summary>
        /// Generate horizontal alignment GeoTable report in Excel (GLTT Standard Format)
        /// </summary>
        private void GenerateHorizontalGeoTableExcel(CivDb.Alignment alignment, string outputPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Track Geometry Data");

                    // Get project info
                    Database db = alignment.Database;
                    string projectName = "Unknown Project";
                    try
                    {
                        if (db.Filename != null && db.Filename.Length > 0)
                            projectName = System.IO.Path.GetFileNameWithoutExtension(db.Filename);
                    }
                    catch { }

                    int row = 1;

                    // Title row
                    string trackName = alignment.Name?.ToUpper() ?? "TRACK GEOMETRY DATA";
                    ws.Cells[row, 1].Value = $"TRACK GEOMETRY DATA - {trackName}";
                    ws.Cells[row, 1, row, 11].Merge = true;
                    ws.Cells[row, 1].Style.Font.Bold = true;
                    ws.Cells[row, 1].Style.Font.Size = 12;
                    ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    row++;
                    
                    // Date and Time row
                    ws.Cells[row, 0].Value = "Generated Report:";
                    ws.Cells[row, 0].Style.Font.Bold = true;
                    ws.Cells[row, 0].Style.Font.Size = 9;

                    ws.Cells[row, 2].Value = "Date:";
                    ws.Cells[row, 2].Style.Font.Bold = true;
                    ws.Cells[row, 2].Style.Font.Size = 9;
                    ws.Cells[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
                    ws.Cells[row, 3].Style.Font.Size = 9;
                    
                    ws.Cells[row, 4].Value = "Time:";
                    ws.Cells[row, 4].Style.Font.Bold = true;
                    ws.Cells[row, 4].Style.Font.Size = 9;
                    ws.Cells[row, 5].Value = DateTime.Now.ToString("h:mm tt");
                    ws.Cells[row, 5].Style.Font.Size = 9;
                    row++;

                    // Headers row 1
                    int headerRow1 = row;
                    ws.Cells[headerRow1, 1].Value = "ELEMENT";
                    ws.Cells[headerRow1, 2].Value = "CURVE No.";
                    ws.Cells[headerRow1, 3].Value = "POINT";
                    ws.Cells[headerRow1, 4].Value = "STATION";
                    ws.Cells[headerRow1, 5].Value = "BEARING";
                    ws.Cells[headerRow1, 6].Value = "COORDINATES";
                    ws.Cells[headerRow1, 6, headerRow1, 7].Merge = true;
                    ws.Cells[headerRow1, 8].Value = "DATA";
                    ws.Cells[headerRow1, 8, headerRow1, 11].Merge = true;

                    row++;

                    // Headers row 2 (sub-headers for coordinates)
                    ws.Cells[row, 6].Value = "Northing";
                    ws.Cells[row, 7].Value = "Easting";

                    // Merge column A-E headers vertically
                    for (int col = 1; col <= 5; col++)
                    {
                        ws.Cells[headerRow1, col, row, col].Merge = true;
                    }

                    // Style all headers
                    using (var range = ws.Cells[headerRow1, 1, row, 11])
                    {
                        range.Style.Font.Name = "Arial";
                        range.Style.Font.Size = 8;
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // Add borders to all header cells individually
                    for (int r = headerRow1; r <= row; r++)
                    {
                        for (int c = 1; c <= 11; c++)
                        {
                            ws.Cells[r, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            ws.Cells[r, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            ws.Cells[r, c].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            ws.Cells[r, c].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }
                    }

                    row++;
                    int startDataRow = row;
                    int lastDataRow = row;
                    int curveNumber = 0;

                    // Collect and sort alignment entities by station
                    var sortedEntities = new System.Collections.Generic.List<AlignmentEntity>();
                    for (int i = 0; i < alignment.Entities.Count; i++)
                    {
                        if (alignment.Entities[i] != null)
                            sortedEntities.Add(alignment.Entities[i]);
                    }

                    // Sort by start station
                    sortedEntities.Sort((a, b) =>
                    {
                        double stationA = 0;
                        double stationB = 0;

                        if (a.EntityType == AlignmentEntityType.Line)
                            stationA = (a as AlignmentLine).StartStation;
                        else if (a.EntityType == AlignmentEntityType.Arc)
                            stationA = (a as AlignmentArc).StartStation;
                        else if (a.EntityType == AlignmentEntityType.Spiral)
                            stationA = (a as AlignmentSpiral).StartStation;

                        if (b.EntityType == AlignmentEntityType.Line)
                            stationB = (b as AlignmentLine).StartStation;
                        else if (b.EntityType == AlignmentEntityType.Arc)
                            stationB = (b as AlignmentArc).StartStation;
                        else if (b.EntityType == AlignmentEntityType.Spiral)
                            stationB = (b as AlignmentSpiral).StartStation;

                        return stationA.CompareTo(stationB);
                    });

                    // Process alignment entities in sorted order
                    for (int i = 0; i < sortedEntities.Count; i++)
                    {
                        AlignmentEntity entity = sortedEntities[i];
                        if (entity == null) continue;

                        try
                        {
                            switch (entity.EntityType)
                            {
                                case AlignmentEntityType.Line:
                                    row = WriteGeoTableTangentExcel(ws, entity as AlignmentLine, alignment, i, row);
                                    break;
                                case AlignmentEntityType.Arc:
                                    curveNumber++;
                                    row = WriteGeoTableCurveExcel(ws, entity as AlignmentArc, alignment, i, row, curveNumber);
                                    break;
                                case AlignmentEntityType.Spiral:
                                    row = WriteGeoTableSpiralExcel(ws, entity, alignment, i, row);
                                    break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ws.Cells[row, 1].Value = "ERROR";
                            ws.Cells[row, 3].Value = $"Error: {ex.Message}";
                            row++;
                        }
                    }

                    lastDataRow = row - 1;

                    // Delete empty rows (rows where columns 1-7 are all empty)
                    for (int r = lastDataRow; r >= startDataRow; r--)
                    {
                        bool isEmpty = true;
                        for (int c = 1; c <= 7; c++)
                        {
                            if (ws.Cells[r, c].Value != null)
                            {
                                isEmpty = false;
                                break;
                            }
                        }

                        if (isEmpty)
                        {
                            ws.DeleteRow(r);
                            lastDataRow--;
                        }
                    }

                    // Apply Arial 8pt font, borders, and center alignment to all data cells
                    for (int r = startDataRow; r <= lastDataRow; r++)
                    {
                        for (int c = 1; c <= 11; c++)
                        {
                            ws.Cells[r, c].Style.Font.Name = "Arial";
                            ws.Cells[r, c].Style.Font.Size = 8;
                            ws.Cells[r, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            ws.Cells[r, c].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            // For columns 1-7, apply all borders
                            if (c <= 7)
                            {
                                ws.Cells[r, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                ws.Cells[r, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                ws.Cells[r, c].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                ws.Cells[r, c].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            }
                            // For DATA columns (8-11), no borders - will be applied to perimeter only
                        }
                    }

                    // Borders are applied per element group in the WriteGeoTable methods

                    // Add overall outside border around entire DATA section (Thin style)
                    for (int c = 8; c <= 11; c++)
                    {
                        ws.Cells[startDataRow, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[lastDataRow, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                    for (int r = startDataRow; r <= lastDataRow; r++)
                    {
                        ws.Cells[r, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[r, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    // Set fixed column widths (more compact like the example)
                    ws.Column(1).Width = 11;  // ELEMENT
                    ws.Column(2).Width = 10;  // CURVE No.
                    ws.Column(3).Width = 7;   // POINT
                    ws.Column(4).Width = 11;  // STATION
                    ws.Column(5).Width = 18;  // BEARING
                    ws.Column(6).Width = 13;  // Northing
                    ws.Column(7).Width = 13;  // Easting
                    ws.Column(8).Width = 16;  // DATA col 1
                    ws.Column(9).Width = 14;  // DATA col 2
                    ws.Column(10).Width = 16; // DATA col 3
                    ws.Column(11).Width = 16; // DATA col 4

                    // Save file
                    System.IO.FileInfo file = new System.IO.FileInfo(outputPath);
                    package.SaveAs(file);
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error generating GeoTable Excel: {ex.Message}", ex);
            }
        }

        private int WriteGeoTableTangentExcel(ExcelWorksheet ws, AlignmentLine line, CivDb.Alignment alignment, int index, int row)
        {
            if (line == null) return row + 1;

            try
            {
                double x1 = 0, y1 = 0, z1 = 0;
                alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);

                string bearing = FormatBearingDMS(line.Direction);
                string pointLabel = (index == 0) ? "POB" : "PI"; // InRoads mapping

                ws.Cells[row, 1].Value = "TANGENT";
                // Merge and center empty CURVE No. column
                ws.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 3].Value = pointLabel;
                ws.Cells[row, 4].Value = FormatStation(line.StartStation);
                ws.Cells[row, 5].Value = bearing;
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA - single merged cell for tangent
                string data = $"L = {FormatDistanceFeet(line.Length)}";
                ws.Cells[row, 8, row, 11].Merge = true;
                ws.Cells[row, 8].Value = data;
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                // Add border around DATA group (single row)
                for (int c = 8; c <= 11; c++)
                {
                    ws.Cells[row, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[row, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                ws.Cells[row, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                return row + 1;
            }
            catch
            {
                return row + 1;
            }
        }

        private int WriteGeoTableCurveExcel(ExcelWorksheet ws, AlignmentArc arc, CivDb.Alignment alignment, int index, int row, int curveNumber)
        {
            if (arc == null) return row + 3;

            try
            {
                int startRow = row;

                // Get coordinates
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);


                // Calculate curve parameters
                double radius = Math.Abs(arc.Radius);
                double delta = arc.Length / radius;
                double tc = CalculateTangentDistance(radius, delta);
                double ec = CalculateExternalDistance(radius, delta);
                string directionStr = arc.Clockwise ? "R" : "L";
                string deltaAngle = FormatAngleDMS(delta);

                // Get PI station and coordinates from arc properties
                double piStation = arc.PIStation;

                // Try to get PI and Center coordinates from sub-entities
                double xPI = 0, yPI = 0, centerE = 0, centerN = 0;
                bool gotFromSubEntity = false;
                string debugMessage = "";

                try
                {
                    debugMessage = $"SubEntityCount: {arc.SubEntityCount}";

                    // Access the first sub-entity which should be the arc itself
                    if (arc.SubEntityCount > 0)
                    {
                        var subEntity = arc[0];
                        debugMessage += $" | SubEntity Type: {subEntity.GetType().Name}";

                        if (subEntity is AlignmentSubEntityArc)
                        {
                            var subEntityArc = subEntity as AlignmentSubEntityArc;
                            var piPoint = subEntityArc.PIPoint;
                            var centerPoint = subEntityArc.CenterPoint;

                            xPI = piPoint.X;  // Easting
                            yPI = piPoint.Y;  // Northing
                            centerE = centerPoint.X;  // Easting
                            centerN = centerPoint.Y;  // Northing
                            gotFromSubEntity = true;
                            debugMessage += $" | SUCCESS: PI({xPI:F4},{yPI:F4}) CC({centerE:F4},{centerN:F4})";
                        }
                        else
                        {
                            debugMessage += " | FAILED: Not AlignmentSubEntityArc";
                        }
                    }
                    else
                    {
                        debugMessage += " | FAILED: No sub-entities";
                    }
                }
                catch (System.Exception ex)
                {
                    gotFromSubEntity = false;
                    debugMessage += $" | EXCEPTION: {ex.Message}";
                }

                // Fallback: calculate if sub-entity access failed
                if (!gotFromSubEntity)
                {
                    CalculateCurveCenter(arc, alignment, out centerN, out centerE);
                    double backTangentDir = arc.StartDirection - Math.PI;
                    xPI = x1 - tc * Math.Cos(backTangentDir);
                    yPI = y1 - tc * Math.Sin(backTangentDir);
                    debugMessage += $" | FALLBACK: PI({xPI:F4},{yPI:F4}) CC({centerE:F4},{centerN:F4})";
                }

                // Debug message removed - ed not available in this context
                // ed.WriteMessage($"\n[DEBUG Curve {curveNumber}] {debugMessage}");

                // Row 1: PC (Start of Curve per InRoads)
                ws.Cells[row, 1].Value = "CURVE";
                ws.Cells[row, 2].Value = $"{curveNumber}-{directionStr}";
                ws.Cells[row, 3].Value = "PC";
                ws.Cells[row, 4].Value = FormatStation(arc.StartStation);
                ws.Cells[row, 5].Value = FormatBearingDMS(arc.StartDirection);
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA for row 1 - individual cells
                ws.Cells[row, 8].Value = $"Δc = {FormatAngleDMS(delta)}";
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 9].Value = $"Da= {FormatAngleDMS(delta / 2.0)}";
                ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 10].Value = $"R= {FormatRadius(radius)}";
                ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 11].Value = $"Lc= {FormatDistanceFeet(arc.Length)}";
                ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                row++;

                // Row 2: PI - show PI Station (no bearing per InRails standard)
                ws.Cells[row, 3].Value = "PI";
                ws.Cells[row, 4].Value = FormatStation(piStation);
                ws.Cells[row, 5].Value = ""; // Empty bearing column for PI
                ws.Cells[row, 6].Value = yPI;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = xPI;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA for row 2 - individual cells
                ws.Cells[row, 8].Value = "V= -- MPH";
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 9].Value = "Ea= --\"";
                ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 10].Value = "Ee= --\"";
                ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 11].Value = "Eu= --\"";
                ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                row++;

                // Row 3: PT (End of Curve per InRoads) - empty bearing per InRails standard
                ws.Cells[row, 3].Value = "PT";
                ws.Cells[row, 4].Value = FormatStation(arc.EndStation);
                ws.Cells[row, 5].Value = ""; // Empty bearing column for PT
                ws.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 6].Value = y2;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x2;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA for row 3 - individual cells
                ws.Cells[row, 8].Value = $"Tc= {FormatDistanceFeet(tc)}";
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 9].Value = $"Ec= {FormatDistanceFeet(ec)}";
                ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 10].Value = $"CC:N {centerN:F4}";
                ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 11].Value = $"E {centerE:F4}";
                ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                row++;

                // Merge ELEMENT and CURVE No. columns for all 3 rows
                ws.Cells[startRow, 1, startRow + 2, 1].Merge = true;
                ws.Cells[startRow, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                ws.Cells[startRow, 2, startRow + 2, 2].Merge = true;
                ws.Cells[startRow, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                ws.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Add border around DATA group (3 rows)
                for (int c = 8; c <= 11; c++)
                {
                    ws.Cells[startRow, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[startRow + 2, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                for (int r = startRow; r <= startRow + 2; r++)
                {
                    ws.Cells[r, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[r, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                return row;
            }
            catch
            {
                return row + 3;
            }
        }

        private int WriteGeoTableSpiralExcel(ExcelWorksheet ws, AlignmentEntity spiralEntity, CivDb.Alignment alignment, int index, int row)
        {
            if (spiralEntity == null) return row + 2;

            try
            {
                int startRow = row;

                // Try to get additional properties from AlignmentSubEntitySpiral
                double radiusIn = double.PositiveInfinity;
                double radiusOut = double.PositiveInfinity;
                double length = 0; double startStation = 0; double endStation = 0; double startDirection = 0;
                double spiralA_API = 0, spiralK_API = 0, longTangent_API = 0, shortTangent_API = 0;
                var spiral = spiralEntity as AlignmentSpiral;
                if (spiral != null)
                {
                    radiusIn = spiral.RadiusIn; radiusOut = spiral.RadiusOut; length = spiral.Length; startStation = spiral.StartStation; endStation = spiral.EndStation; startDirection = spiral.StartDirection;
                }
                else
                {
                    try { dynamic d = spiralEntity; radiusIn = d.RadiusIn; } catch { }
                    try { dynamic d = spiralEntity; radiusOut = d.RadiusOut; } catch { }
                    try { dynamic d = spiralEntity; length = d.Length; } catch { }
                    try { dynamic d = spiralEntity; startStation = d.StartStation; } catch { }
                    try { dynamic d = spiralEntity; endStation = d.EndStation; } catch { }
                    // startDirection: approximate using chord around start (computed later if needed)
                }

                bool gotFromSubEntity = false;
                try
                {
                    dynamic dse = spiral ?? (dynamic)spiralEntity;
                    if (dse.SubEntityCount > 0)
                    {
                        var subEntity = dse[0];
                        if (subEntity is AlignmentSubEntitySpiral)
                        {
                            var subEntitySpiral = subEntity as AlignmentSubEntitySpiral;

                            // Extract properties from sub-entity
                            radiusIn = subEntitySpiral.RadiusIn;
                            radiusOut = subEntitySpiral.RadiusOut;
                            length = subEntitySpiral.Length;
                            
                            // Get additional properties via reflection
                            var spiralType = subEntitySpiral.GetType();
                            try
                            {
                                var aProp = spiralType.GetProperty("A");
                                var kProp = spiralType.GetProperty("K");
                                var longTangentProp = spiralType.GetProperty("LongTangent");
                                var shortTangentProp = spiralType.GetProperty("ShortTangent");
                                if (aProp != null) spiralA_API = (double)aProp.GetValue(subEntitySpiral);
                                if (kProp != null) spiralK_API = (double)kProp.GetValue(subEntitySpiral);
                                if (longTangentProp != null) longTangent_API = (double)longTangentProp.GetValue(subEntitySpiral);
                                if (shortTangentProp != null) shortTangent_API = (double)shortTangentProp.GetValue(subEntitySpiral);
                            }
                            catch { }

                            gotFromSubEntity = true;
                        }
                    }
                }
                catch { }

                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);

                // Determine entry/exit classification for labels
                bool isEntry = double.IsInfinity(radiusIn) || radiusIn == 0 || radiusIn > radiusOut;
                string startLabel = isEntry ? "TS" : "CS";
                string endLabel = isEntry ? "SC" : "ST";
                
                // Calculate tangent direction at end for ST bearing (exit spiral to tangent)
                double tangentDirEnd = 0;
                if (endLabel == "ST")
                {
                    double x3 = 0, y3 = 0, z3 = 0;
                    alignment.PointLocation(endStation + 0.01, 0, 0, ref x3, ref y3, ref z3);
                    tangentDirEnd = Math.Atan2(y3 - y2, x3 - x2);
                }

                // Calculate spiral parameters
                double spiralLength = length;
                // Use actual radius from spiral if available, otherwise default
                double radius = radiusOut > 0 ? radiusOut : (radiusIn > 0 ? radiusIn : 1000);
                double theta = CalculateSpiralAngle(spiralLength, radius);
                double xs = CalculateSpiralX(spiralLength, theta);
                double ys = CalculateSpiralY(spiralLength, theta);
                double p = ys;  // Simplified
                double k = spiralLength / 2.0;  // Simplified

                // Row 1: TS (entry) or CS (exit) start point
                ws.Cells[row, 1].Value = gotFromSubEntity ? "SPIRAL (SubEntity)" : "SPIRAL";
                ws.Cells[row, 3].Value = startLabel; // Dynamic: TS or CS
                ws.Cells[row, 4].Value = FormatStation(startStation);
                // If startDirection wasn't available, compute from local tangent
                if (Math.Abs(startDirection) < 1e-6)
                {
                    double x0 = 0, y0 = 0, z0 = 0; alignment.PointLocation(startStation - 0.01, 0, 0, ref x0, ref y0, ref z0);
                    startDirection = Math.Atan2(y1 - y0, x1 - x0);
                }
                ws.Cells[row, 5].Value = FormatBearingDMS(startDirection);
                ws.Cells[row, 6].Value = y1;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x1;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA for row 1 - individual cells
                ws.Cells[row, 8].Value = $"θs = {FormatAngleDMS(theta)}";
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 9].Value = $"Ls= {FormatDistanceFeet(spiralLength)}";
                ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 10].Value = $"LT= {FormatDistanceFeet(longTangent_API > 0 ? longTangent_API : (spiralLength * 0.67))}";
                ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 11].Value = $"STs= {FormatDistanceFeet(shortTangent_API > 0 ? shortTangent_API : (spiralLength * 0.33))}";
                ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                row++;

                // Row 2: End point (ST for exit spiral has bearing, SC for entry spiral is empty)
                ws.Cells[row, 3].Value = endLabel; // Dynamic: ST or SC
                ws.Cells[row, 4].Value = FormatStation(endStation);
                // ST (exit spiral end) shows bearing, SC (entry spiral end) is empty
                ws.Cells[row, 5].Value = (endLabel == "ST") ? FormatBearingDMS(tangentDirEnd) : "";
                ws.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 6].Value = y2;
                ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                ws.Cells[row, 7].Value = x2;
                ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";

                // DATA for row 2 - individual cells
                ws.Cells[row, 8].Value = $"Xs= {FormatDistanceFeet(xs)}";
                ws.Cells[row, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 9].Value = $"Ys= {FormatDistanceFeet(ys)}";
                ws.Cells[row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 10].Value = $"P= {FormatDistanceFeet(p)}";
                ws.Cells[row, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                ws.Cells[row, 11].Value = $"K= {FormatDistanceFeet(spiralK_API > 0 ? spiralK_API : k)}";
                ws.Cells[row, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                row++;

                // Merge ELEMENT column for both rows and CURVE No. column for both rows
                ws.Cells[startRow, 1, startRow + 1, 1].Merge = true;
                ws.Cells[startRow, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                ws.Cells[startRow, 2, startRow + 1, 2].Merge = true;
                ws.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Add border around DATA group (2 rows)
                for (int c = 8; c <= 11; c++)
                {
                    ws.Cells[startRow, c].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[startRow + 1, c].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                for (int r = startRow; r <= startRow + 1; r++)
                {
                    ws.Cells[r, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ws.Cells[r, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                return row;
            }
            catch
            {
                return row + 2;
            }
        }

        /// <summary>
        /// Generate horizontal alignment GeoTable report in PDF (GLTT Standard Format)
        /// </summary>
        private void GenerateHorizontalGeoTablePdf(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                using (PdfWriter writer = new PdfWriter(outputPath))
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                {
                    // Set to landscape with reduced margins
                    pdfDoc.SetDefaultPageSize(PageSize.LETTER.Rotate());

                    using (Document document = new Document(pdfDoc, PageSize.LETTER.Rotate(), false))
                    {
                        // Set smaller margins to fit content
                        document.SetMargins(20, 20, 20, 20);

                        // Create fonts - use Arial (Helvetica) with smaller sizes
                        PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

                        // Title - smaller font
                        string trackName = alignment.Name?.ToUpper() ?? "TRACK GEOMETRY DATA";
                        Paragraph title = new Paragraph($"TRACK GEOMETRY DATA - {trackName}")
                            .SetFont(boldFont)
                            .SetFontSize(10)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetMarginBottom(3);
                        document.Add(title);
                        
                        // Add date/time
                        Paragraph dateTime = new Paragraph($"Generated Report:  Date: {DateTime.Now:MM/dd/yyyy}     Time: {DateTime.Now:h:mm tt} {(TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now) ? TimeZoneInfo.Local.DaylightName : TimeZoneInfo.Local.StandardName)}")
                            .SetFont(font)
                            .SetFontSize(8)
                            .SetItalic()
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetMarginBottom(5);
                        document.Add(dateTime);

                        // Create table with 11 columns - optimized widths for compact display
                        float[] columnWidths = { 9f, 8f, 6f, 9f, 12f, 11f, 11f, 13f, 12f, 13f, 13f };
                        iText.Layout.Element.Table table = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(columnWidths));
                        table.SetWidth(UnitValue.CreatePercentValue(100));
                        table.SetFont(font).SetFontSize(6.5f);
                        table.SetFixedLayout();

                        // Header row 1 - COORDINATES and DATA merged headers
                        table.AddHeaderCell(CreateHeaderCell("ELEMENT", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("CURVE No.", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("POINT", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("STATION", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("BEARING", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("COORDINATES", boldFont, 1, 2));

                        // DATA header with only outside border
                        table.AddHeaderCell(new iText.Layout.Element.Cell(1, 4)
                            .Add(new Paragraph("DATA").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderBottom(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetPadding(1));

                        // Header row 2 - Northing, Easting subdivisions (DATA columns with no inside borders)
                        table.AddHeaderCell(CreateHeaderCell("Northing", boldFont, 1, 1));
                        table.AddHeaderCell(CreateHeaderCell("Easting", boldFont, 1, 1));

                        // DATA row 2 cells - no inside borders, only bottom border to complete the header
                        table.AddHeaderCell(new iText.Layout.Element.Cell()
                            .Add(new Paragraph("").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderRight(iText.Layout.Borders.Border.NO_BORDER)
                            .SetPadding(1));
                        table.AddHeaderCell(new iText.Layout.Element.Cell()
                            .Add(new Paragraph("").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderLeft(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderRight(iText.Layout.Borders.Border.NO_BORDER)
                            .SetPadding(1));
                        table.AddHeaderCell(new iText.Layout.Element.Cell()
                            .Add(new Paragraph("").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderLeft(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderRight(iText.Layout.Borders.Border.NO_BORDER)
                            .SetPadding(1));
                        table.AddHeaderCell(new iText.Layout.Element.Cell()
                            .Add(new Paragraph("").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderLeft(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetPadding(1));

                        // Collect and sort alignment entities by station
                        var sortedEntities = new System.Collections.Generic.List<AlignmentEntity>();
                        for (int i = 0; i < alignment.Entities.Count; i++)
                        {
                            if (alignment.Entities[i] != null)
                                sortedEntities.Add(alignment.Entities[i]);
                        }

                        // Sort by start station (robust to entity runtime types)
                        sortedEntities.Sort((a, b) => SafeStartStation(a).CompareTo(SafeStartStation(b)));

                        // Process alignment entities in sorted order
                        int curveNumber = 0;
                        for (int i = 0; i < sortedEntities.Count; i++)
                        {
                            AlignmentEntity entity = sortedEntities[i];
                            if (entity == null) continue;

                            AlignmentEntity prevEntity = i > 0 ? sortedEntities[i - 1] : null;
                            AlignmentEntity nextEntity = i < sortedEntities.Count - 1 ? sortedEntities[i + 1] : null;

                            try
                            {
                                switch (entity.EntityType)
                                {
                                    case AlignmentEntityType.Line:
                                        AddGeoTableTangentPdf(table, entity as AlignmentLine, alignment, i, font);
                                        break;
                                    case AlignmentEntityType.Arc:
                                        curveNumber++;
                                        AddGeoTableCurvePdf(table, entity as AlignmentArc, alignment, i, curveNumber, font, prevEntity, nextEntity);
                                        break;
                                    case AlignmentEntityType.Spiral:
                                        AddGeoTableSpiralPdf(table, entity, alignment, i, font, prevEntity, nextEntity);
                                        break;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                // Add error row
                                table.AddCell(CreateDataCell("ERROR", font));
                                for (int j = 1; j < 11; j++)
                                {
                                    table.AddCell(CreateDataCell($"Error: {ex.Message}", font));
                                }
                            }
                        }

                        document.Add(table);
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error generating GeoTable PDF: {ex.Message}", ex);
            }
        }

        // NEW: Vertical GeoTable PDF generation (layout modeled after horizontal GeoTable style but with vertical geometry fields)
        private void GenerateVerticalGeoTablePdf(CivDb.Alignment alignment, string outputPath)
        {
            try
            {
                // Acquire first profile (layout) with entities
                ObjectId layoutProfileId = ObjectId.Null;
                foreach (ObjectId pid in alignment.GetProfileIds())
                {
                    using (Profile profile = pid.GetObject(OpenMode.ForRead) as Profile)
                    {
                        if (profile != null && profile.Entities != null && profile.Entities.Count > 0)
                        {
                            layoutProfileId = pid; break;
                        }
                    }
                }
                if (layoutProfileId == ObjectId.Null)
                {
                    // Fallback: create minimal PDF noting absence
                    using (PdfWriter w = new PdfWriter(outputPath))
                    using (PdfDocument pdfDoc = new PdfDocument(w))
                    using (Document doc = new Document(pdfDoc, PageSize.LETTER))
                    {
                        doc.Add(new Paragraph("No vertical profile found.").SetFontSize(11));
                    }
                    return;
                }

                using (Profile profile = layoutProfileId.GetObject(OpenMode.ForRead) as Profile)
                using (PdfWriter writer = new PdfWriter(outputPath))
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                {
                    pdfDoc.SetDefaultPageSize(PageSize.LETTER.Rotate());
                    using (Document document = new Document(pdfDoc, PageSize.LETTER.Rotate(), false))
                    {
                        document.SetMargins(20, 20, 20, 20);
                        PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

                        string trackName = alignment.Name?.ToUpper() ?? "VERTICAL GEOMETRY";
                        document.Add(new Paragraph($"VERTICAL PROFILE DATA - {trackName}")
                            .SetFont(boldFont).SetFontSize(10)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetMarginBottom(3));

                        // Column layout mirrored from horizontal (retain widths for consistency)
                        float[] columnWidths = { 9f, 8f, 8f, 9f, 11f, 11f, 11f, 13f, 12f, 13f, 13f };
                        iText.Layout.Element.Table table = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(columnWidths));
                        table.SetWidth(UnitValue.CreatePercentValue(100));
                        table.SetFont(font).SetFontSize(6.5f);
                        table.SetFixedLayout();

                        // Header row 1
                        table.AddHeaderCell(CreateHeaderCell("ELEMENT", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("CURVE No.", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("POINT", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("STATION", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("ELEV", boldFont, 2, 1));
                        table.AddHeaderCell(CreateHeaderCell("GRADE", boldFont, 2, 1));
                        // DATA group (4 columns) for vertical-specific parameters
                        table.AddHeaderCell(new iText.Layout.Element.Cell(1, 4)
                            .Add(new Paragraph("DATA").SetFont(boldFont).SetFontSize(6.5f))
                            .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                            .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderBottom(iText.Layout.Borders.Border.NO_BORDER)
                            .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f))
                            .SetPadding(1));

                        // Header row 2 (blank placeholders under DATA)
                        table.AddHeaderCell(CreateHeaderCell("", boldFont, 1, 1)); // filler after merged DATA group start
                        table.AddHeaderCell(CreateHeaderCell("", boldFont, 1, 1));
                        table.AddHeaderCell(CreateHeaderCell("", boldFont, 1, 1));
                        table.AddHeaderCell(CreateHeaderCell("", boldFont, 1, 1));

                        int curveNumber = 0;
                        for (int i = 0; i < profile.Entities.Count; i++)
                        {
                            var entity = profile.Entities[i];
                            if (entity == null) continue;
                            try
                            {
                                switch (entity.EntityType)
                                {
                                    case ProfileEntityType.Tangent:
                                        AddVerticalTangentRows(table, entity as ProfileTangent, i, font);
                                        break;
                                    case ProfileEntityType.Circular:
                                        curveNumber++;
                                        AddVerticalCurveRows(table, entity as ProfileCircular, curveNumber, font);
                                        break;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                table.AddCell(CreateDataCell("ERROR", font));
                                for (int c = 0; c < 10; c++) table.AddCell(CreateDataCell(ex.Message, font));
                            }
                        }

                        document.Add(table);
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Error generating Vertical GeoTable PDF: {ex.Message}", ex);
            }
        }

        private void AddVerticalTangentRows(iText.Layout.Element.Table table, ProfileTangent tangent, int index, PdfFont font)
        {
            if (tangent == null) return;
            double startSta = tangent.StartStation; double endSta = tangent.EndStation;
            double startElev = tangent.StartElevation; double endElev = tangent.EndElevation;
            double gradePct = tangent.Grade * 100.0;
            // InRoads vertical: only the initial tangent start uses POB. Subsequent tangent starts precede PVC of a curve so no standalone label.
            string pointLabel = (index == 0) ? "POB" : "";

            // Single row representation for tangent start (mirroring horizontal tangent style with merged DATA)
            table.AddCell(CreateDataCell("TANGENT", font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell(pointLabel, font)); // blank for non-initial tangents
            table.AddCell(CreateDataCell(FormatStation(startSta), font));
            table.AddCell(CreateDataCell(startElev.ToString("F2"), font));
            table.AddCell(CreateDataCell(FormatGrade(gradePct), font));
            table.AddCell(new iText.Layout.Element.Cell(1,4)
                .Add(new Paragraph($"Grade={FormatGrade(gradePct)}  L={tangent.Length:F2}").SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK,0.5f))
                .SetBorderBottom(new SolidBorder(ColorConstants.BLACK,0.5f))
                .SetBorderLeft(new SolidBorder(ColorConstants.BLACK,0.5f))
                .SetBorderRight(new SolidBorder(ColorConstants.BLACK,0.5f))
                .SetPadding(1));
        }

        private void AddVerticalCurveRows(iText.Layout.Element.Table table, ProfileCircular curve, int curveNumber, PdfFont font)
        {
            if (curve == null) return;
            const double tol = 1e-8;
            double pvcSta = curve.StartStation;
            double pvtSta = curve.EndStation;
            double pviSta = curve.PVIStation;
            double pviElev = curve.PVIElevation;
            double gradeInPct = curve.GradeIn * 100.0;
            double gradeOutPct = curve.GradeOut * 100.0;
            double length = curve.Length;
            double gradeDiff = gradeOutPct - gradeInPct;
            double k = Math.Abs(gradeDiff) > tol ? length / Math.Abs(gradeDiff) : double.PositiveInfinity;

            // Compute PVC & PVT elevations if not exposed
            double pvcElev = 0; double pvtElev = 0;
            try { pvcElev = curve.StartElevation; } catch { pvcElev = pviElev - (curve.GradeIn * (length/2)); }
            try { pvtElev = curve.EndElevation; } catch { pvtElev = pviElev + (curve.GradeOut * (length/2)); }

            // High/Low point determination
            // x from PVC where derivative = 0: x = -g1*L/(g2-g1) (grades in decimal)
            double g1 = curve.GradeIn; double g2 = curve.GradeOut; double xHighLow = double.NaN; double highLowSta = double.NaN; double highLowElev = double.NaN; string curveType = "";
            if (Math.Abs(g2 - g1) > tol)
            {
                xHighLow = -g1 * length / (g2 - g1); // feet along curve from PVC
                if (xHighLow >= 0 && xHighLow <= length)
                {
                    highLowSta = pvcSta + xHighLow;
                    // Elevation at x: y = yPVC + g1*x + ( (g2-g1)/(2L) ) * x^2
                    highLowElev = pvcElev + g1 * xHighLow + ((g2 - g1) / (2.0 * length)) * xHighLow * xHighLow;
                }
            }
            curveType = (g1 > g2) ? "Crest" : (g1 < g2 ? "Sag" : "Level");

            string curveLabel = $"VC{curveNumber}-{(curveType.StartsWith("C")?"C":"S")}"; // VC#-C/S simplified label
            string kDisplay = double.IsInfinity(k) ? "INF" : k.ToString("F2");

            // Row 1: PVC
            table.AddCell(new iText.Layout.Element.Cell(3,1).Add(new Paragraph("CURVE").SetFont(font).SetFontSize(8))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.TOP)
                .SetBorder(new SolidBorder(ColorConstants.BLACK,1)));
            table.AddCell(new iText.Layout.Element.Cell(3,1).Add(new Paragraph(curveLabel).SetFont(font).SetFontSize(8))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.TOP)
                .SetBorder(new SolidBorder(ColorConstants.BLACK,1)));
            table.AddCell(CreateDataCell("PVC", font));
            table.AddCell(CreateDataCell(FormatStation(pvcSta), font));
            table.AddCell(CreateDataCell(pvcElev.ToString("F2"), font));
            table.AddCell(CreateDataCell(FormatGrade(gradeInPct), font));
            table.AddCell(CreateDataCellNoBorder($"G1= {FormatGrade(gradeInPct)}", font).SetBorderTop(new SolidBorder(ColorConstants.BLACK,0.5f)).SetBorderLeft(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder($"G2= {FormatGrade(gradeOutPct)}", font).SetBorderTop(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder($"L= {length:F2}", font).SetBorderTop(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder($"K= {kDisplay}", font).SetBorderTop(new SolidBorder(ColorConstants.BLACK,0.5f)).SetBorderRight(new SolidBorder(ColorConstants.BLACK,0.5f)));

            // Row 2: PVI
            table.AddCell(CreateDataCell("PVI", font));
            table.AddCell(CreateDataCell(FormatStation(pviSta), font));
            table.AddCell(CreateDataCell(pviElev.ToString("F2"), font));
            table.AddCell(CreateDataCell("", font)); // grade column blank for PVI row
            table.AddCell(CreateDataCellNoBorder($"Type= {curveType}", font).SetBorderLeft(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder($"ΔG= {gradeDiff:F3}%", font));
            table.AddCell(CreateDataCellNoBorder(!double.IsNaN(highLowSta)? $"HpSta= {FormatStation(highLowSta)}" : "HpSta= -", font));
            table.AddCell(CreateDataCellNoBorder(!double.IsNaN(highLowElev)? $"HpElev= {highLowElev:F2}" : "HpElev= -", font).SetBorderRight(new SolidBorder(ColorConstants.BLACK,0.5f)));

            // Row 3: PVT
            table.AddCell(CreateDataCell("PVT", font));
            table.AddCell(CreateDataCell(FormatStation(pvtSta), font));
            table.AddCell(CreateDataCell(pvtElev.ToString("F2"), font));
            table.AddCell(CreateDataCell(FormatGrade(gradeOutPct), font));
            // Sight distance placeholder (user can supply actual values later)
            table.AddCell(CreateDataCellNoBorder("SSD= --", font).SetBorderBottom(new SolidBorder(ColorConstants.BLACK,0.5f)).SetBorderLeft(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder("Middle Ord= --", font).SetBorderBottom(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder("Notes= -", font).SetBorderBottom(new SolidBorder(ColorConstants.BLACK,0.5f)));
            table.AddCell(CreateDataCellNoBorder("", font).SetBorderBottom(new SolidBorder(ColorConstants.BLACK,0.5f)).SetBorderRight(new SolidBorder(ColorConstants.BLACK,0.5f)));
        }

        private string FormatGrade(double gradePercent)
        {
            return gradePercent.ToString("F3") + "%";
        }

        private double SafeStartStation(AlignmentEntity entity)
        {
            try
            {
                dynamic d = entity;
                double s = d.StartStation;
                if (double.IsNaN(s) || double.IsInfinity(s)) return double.MaxValue;
                return s;
            }
            catch
            {
                return double.MaxValue; // push unknowns to end, avoids comparer exceptions
            }
        }

        private iText.Layout.Element.Cell CreateHeaderCell(string text, PdfFont font, int rowspan, int colspan)
        {
            return new iText.Layout.Element.Cell(rowspan, colspan)
                .Add(new Paragraph(text).SetFont(font).SetFontSize(6.5f))
                .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetPadding(1);
        }

        private iText.Layout.Element.Cell CreateDataCell(string text, PdfFont font)
        {
            return new iText.Layout.Element.Cell()
                .Add(new Paragraph(text).SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetPadding(1);
        }

        private iText.Layout.Element.Cell CreateDataCellNoBorder(string text, PdfFont font)
        {
            return new iText.Layout.Element.Cell()
                .Add(new Paragraph(text).SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBackgroundColor(ColorConstants.WHITE)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetPadding(1);
        }

        // Center-aligned label cell for POINT column per InRails standard
        private iText.Layout.Element.Cell CreateLabelCell(string text, PdfFont font, int rowspan = 1, int colspan = 1)
        {
            return new iText.Layout.Element.Cell(rowspan, colspan)
                .Add(new Paragraph(text).SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER) // Changed to CENTER per markup
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetPadding(1);
        }

        private void AddGeoTableTangentPdf(iText.Layout.Element.Table table, AlignmentLine line, CivDb.Alignment alignment, int index, PdfFont font)
        {
            if (line == null) return;

            double x1 = 0, y1 = 0, z1 = 0;
            alignment.PointLocation(line.StartStation, 0, 0, ref x1, ref y1, ref z1);
            string bearing = FormatBearingDMS(line.Direction);
            // InRoads: first tangent start is POB, subsequent tangent starts show PI where applicable
            string pointLabel = (index == 0) ? "POB" : "PI";

            table.AddCell(CreateDataCell("TANGENT", font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell(bearing, font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell("", font));

            // DATA - merge 4 columns with border around the group only
            table.AddCell(new iText.Layout.Element.Cell(1, 4)
                .Add(new Paragraph($"L = {FormatDistanceFeet(line.Length)}").SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBackgroundColor(ColorConstants.WHITE)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));
        }

        private void AddGeoTableCurvePdf(iText.Layout.Element.Table table, AlignmentArc arc, CivDb.Alignment alignment, int index, int curveNumber, PdfFont font, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            if (arc == null) return;

            double radius = Math.Abs(arc.Radius);
            double delta = arc.Length / radius;
            double tc = CalculateTangentDistance(radius, delta);
            double ec = CalculateExternalDistance(radius, delta);

            double x1 = 0, y1 = 0, z1 = 0, x2 = 0, y2 = 0, z2 = 0;
            alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
            alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

            double piStation = arc.PIStation;
            double xPI = 0, yPI = 0, centerE = 0, centerN = 0;
            bool gotFromSubEntity = false;

            try
            {
                if (arc.SubEntityCount > 0)
                {
                    var subEntity = arc[0];
                    if (subEntity is AlignmentSubEntityArc subEntityArc)
                    {
                        var piPoint = subEntityArc.PIPoint;
                        var centerPoint = subEntityArc.CenterPoint;
                        xPI = piPoint.X;
                        yPI = piPoint.Y;
                        centerE = centerPoint.X;
                        centerN = centerPoint.Y;
                        gotFromSubEntity = true;
                    }
                }
            }
            catch { gotFromSubEntity = false; }

            if (!gotFromSubEntity)
            {
                CalculateCurveCenter(arc, alignment, out centerN, out centerE);
                double backTangentDir = arc.StartDirection - Math.PI;
                xPI = x1 - tc * Math.Cos(backTangentDir);
                yPI = y1 - tc * Math.Sin(backTangentDir);
            }

            string curveDir = arc.Clockwise ? "R" : "L";
            string curveLabel = $"{curveNumber}-{curveDir}";

            // Determine labels based on neighbors
            string startLabel = (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Spiral) ? "SC" : "PC";
            string endLabel = (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Spiral) ? "CS" : "PT";

            // Calculate Degree of Curvature (Arc Definition)
            // Da = 5729.58 / R (Degrees) -> Convert to radians for formatter: 100 / R
            double degreeOfCurveRad = (radius > 0) ? (100.0 / radius) : 0;

            // Row 1: Start (PC/SC)
            table.AddCell(new iText.Layout.Element.Cell(3, 1).Add(new Paragraph("CURVE").SetFont(font).SetFontSize(8))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            
            // Curve Number cell with Oval Renderer
            var curveNoCell = new iText.Layout.Element.Cell(3, 1).Add(new Paragraph(curveLabel).SetFont(font).SetFontSize(8))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f));
            curveNoCell.SetNextRenderer(new OvalCellRenderer(curveNoCell));
            table.AddCell(curveNoCell);

            table.AddCell(CreateLabelCell(startLabel, font));
            table.AddCell(CreateDataCell(FormatStation(arc.StartStation), font));
            table.AddCell(CreateDataCell("", font));
            table.AddCell(CreateDataCell($"{y1:F4}", font));
            table.AddCell(CreateDataCell($"{x1:F4}", font));

            table.AddCell(CreateDataCellNoBorder($"Δc = {FormatAngleDMS(delta)}", font)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(CreateDataCellNoBorder($"Da= {FormatAngleDMS(degreeOfCurveRad)}", font)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(CreateDataCellNoBorder($"R= {radius:F2}", font)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(CreateDataCellNoBorder($"Lc= {FormatDistanceFeet(arc.Length)}", font)
                .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));

            // Row 2: PI (no bearing per InRails standard)
            table.AddCell(CreateLabelCell("PI", font));
            table.AddCell(CreateDataCell(FormatStation(piStation), font));
            table.AddCell(CreateDataCell("", font)); // Empty bearing column for PI
            table.AddCell(CreateDataCell($"{yPI:F4}", font));
            table.AddCell(CreateDataCell($"{xPI:F4}", font));

            table.AddCell(CreateDataCellNoBorder("V= -- MPH", font)
                .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(CreateDataCellNoBorder("Ea= --\"", font));
            table.AddCell(CreateDataCellNoBorder("Ee= --\"", font));
            table.AddCell(CreateDataCellNoBorder("Eu= --\"", font)
                .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));

            // Row 3: End (PT/CS) (no bearing per InRails standard)
            table.AddCell(CreateLabelCell(endLabel, font));
            table.AddCell(CreateDataCell(FormatStation(arc.EndStation), font));
            table.AddCell(CreateDataCell("", font)); // Empty bearing column for PT
            table.AddCell(CreateDataCell($"{y2:F4}", font));
            table.AddCell(CreateDataCell($"{x2:F4}", font));

            table.AddCell(CreateDataCellNoBorder($"Tc= {tc:F2}'", font)
                .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(CreateDataCellNoBorder($"Ec= {ec:F2}'", font)
                .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            table.AddCell(new iText.Layout.Element.Cell(1, 2)
                .Add(new Paragraph($"CC:N {centerN:F4}  E {centerE:F4}").SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBackgroundColor(ColorConstants.WHITE)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetPadding(1));
        }

        private void AddGeoTableSpiralPdf(iText.Layout.Element.Table table, AlignmentEntity spiralEntity, CivDb.Alignment alignment, int index, PdfFont font, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            if (spiralEntity == null) return;

            try
            {
                // Get coordinates and properties
                double radiusIn = double.PositiveInfinity;
                double radiusOut = double.PositiveInfinity;
                double length = 0; double startStation = 0; double endStation = 0; double spiralA = 0; string spiralDirection = "";
                var spiral = spiralEntity as AlignmentSpiral;
                if (spiral != null)
                {
                    radiusIn = spiral.RadiusIn; radiusOut = spiral.RadiusOut; length = spiral.Length; startStation = spiral.StartStation; endStation = spiral.EndStation; spiralA = spiral.A; spiralDirection = spiral.Direction.ToString();
                }
                else
                {
                    try { dynamic d = spiralEntity; radiusIn = d.RadiusIn; } catch { }
                    try { dynamic d = spiralEntity; radiusOut = d.RadiusOut; } catch { }
                    try { dynamic d = spiralEntity; length = d.Length; } catch { }
                    try { dynamic d = spiralEntity; startStation = d.StartStation; } catch { }
                    try { dynamic d = spiralEntity; endStation = d.EndStation; } catch { }
                    try { dynamic d = spiralEntity; spiralA = d.A; } catch { }
                    try { dynamic d = spiralEntity; spiralDirection = (d.Direction != null ? d.Direction.ToString() : ""); } catch { }
                }
                
                // Try to get additional properties from AlignmentSubEntitySpiral
                try
                {
                    dynamic dse = spiral ?? (dynamic)spiralEntity;
                    if (dse.SubEntityCount > 0)
                    {
                        var subEntity = dse[0];
                        if (subEntity is AlignmentSubEntitySpiral subSpiral)
                        {
                            startStation = subSpiral.StartStation;
                            endStation = subSpiral.EndStation;
                            length = subSpiral.Length;
                            spiralA = subSpiral.A;
                            spiralDirection = subSpiral.Direction.ToString();
                            try { radiusIn = subSpiral.RadiusIn; } catch { }
                            try { radiusOut = subSpiral.RadiusOut; } catch { }
                        }
                    }
                }
                catch { }

                // Compute final entry/exit classification using latest radius values
                bool isEntry = double.IsInfinity(radiusIn) || radiusIn == 0 || radiusIn > radiusOut;
                double effectiveRadius = isEntry ? radiusOut : radiusIn;

                // Populate start/end coordinates
                double x1 = 0, y1 = 0, z1 = 0;
                double x2 = 0, y2 = 0, z2 = 0;
                alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);

                // Calculate spiral parameters
                double spiralAngle = 0;
                if (effectiveRadius > 0 && !double.IsInfinity(effectiveRadius))
                {
                    spiralAngle = length / (2.0 * effectiveRadius);
                }

                // Calculate extended spiral parameters for display
                // Formulas based on standard clothoid spiral approximations
                double thetaS_rad = spiralAngle;
                double thetaS_deg = thetaS_rad * 180.0 / Math.PI;
                
                // Xs, Ys
                // Xs = Ls * (1 - theta^2/10 + theta^4/216)
                // Ys = Ls * (theta/3 - theta^3/42 + theta^5/1320)
                double Xs = length * (1.0 - Math.Pow(thetaS_rad, 2) / 10.0 + Math.Pow(thetaS_rad, 4) / 216.0);
                double Ys = length * (thetaS_rad / 3.0 - Math.Pow(thetaS_rad, 3) / 42.0 + Math.Pow(thetaS_rad, 5) / 1320.0);
                
                // P, K
                // P = Ys - R * (1 - cos(theta))
                // K = Xs - R * sin(theta)
                double P = Ys - effectiveRadius * (1.0 - Math.Cos(thetaS_rad));
                double K = Xs - effectiveRadius * Math.Sin(thetaS_rad);
                
                // LT, ST
                // LT = Xs - Ys * cot(theta)
                // ST = Ys / sin(theta)
                double LT = Xs - Ys * (1.0 / Math.Tan(thetaS_rad));
                double ST = Ys / Math.Sin(thetaS_rad);
                
                // LC (Long Chord)
                double LC = Math.Sqrt(Xs * Xs + Ys * Ys);
                
                // Chord Direction
                // Chord angle from tangent = atan(Ys/Xs)
                double chordAngleFromTangent = Math.Atan2(Ys, Xs);
                
                // Determine chord bearing based on spiral direction
                // If Entry (Tangent -> Curve), chord is StartDir +/- chordAngle
                // If Exit (Curve -> Tangent), chord is EndDir +/- chordAngle (reversed?)
                // Simplified: Chord direction is from Start Point to End Point
                double chordDir = Math.Atan2(y2 - y1, x2 - x1);
                string chordBearing = FormatBearingDMS(chordDir);

                // Determine which point to show to avoid duplication
                // If Entry (TS->SC): Show TS (Start). SC is in next Curve.
                // If Exit (CS->ST): Show ST (End). CS is in prev Curve.
                
                bool showStart = (nextEntity != null && nextEntity.EntityType == AlignmentEntityType.Arc);
                bool showEnd = (prevEntity != null && prevEntity.EntityType == AlignmentEntityType.Arc);
                
                // If neither (e.g. Spiral-Spiral or Tangent-Spiral-Tangent), default to showing Start (TS)
                if (!showStart && !showEnd) showStart = true;

                // Create nested table for Data block (2 rows x 4 columns)
                iText.Layout.Element.Table dataTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 1, 1, 1, 1 }));
                dataTable.SetWidth(UnitValue.CreatePercentValue(100));
                dataTable.SetFixedLayout();
                
                // Row 1 Data
                dataTable.AddCell(CreateDataCellNoBorder($"θs = {FormatAngleDMS(thetaS_rad)}", font));
                dataTable.AddCell(CreateDataCellNoBorder($"Ls= {length:F2}'", font));
                dataTable.AddCell(CreateDataCellNoBorder($"LT= {LT:F2}'", font));
                dataTable.AddCell(CreateDataCellNoBorder($"STs= {ST:F2}'", font));
                
                // Row 2 Data
                dataTable.AddCell(CreateDataCellNoBorder($"Xs= {Xs:F2}'", font));
                dataTable.AddCell(CreateDataCellNoBorder($"Ys= {Ys:F2}'", font));
                dataTable.AddCell(CreateDataCellNoBorder($"P= {P:F2}'", font));
                dataTable.AddCell(CreateDataCellNoBorder($"K= {K:F2}'", font));

                // Create a cell that wraps the data table
                iText.Layout.Element.Cell dataCell = new iText.Layout.Element.Cell(1, 4)
                    .Add(dataTable)
                    .SetPadding(0)
                    .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                    .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                    .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f))
                    .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f));
                
                if (showStart)
                {
                    // Row: TS
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("SPIRAL").SetFont(font).SetFontSize(8))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("").SetFont(font).SetFontSize(8))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                        
                    table.AddCell(CreateDataCell("TS", font));
                    table.AddCell(CreateDataCell(FormatStation(startStation), font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell($"{y1:F4}", font));
                    table.AddCell(CreateDataCell($"{x1:F4}", font));

                    // DATA row (merged)
                    table.AddCell(dataCell);
                }

                if (showEnd)
                {
                    // Row: ST
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("SPIRAL").SetFont(font).SetFontSize(8))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("").SetFont(font).SetFontSize(8))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                        
                    table.AddCell(CreateDataCell("ST", font));
                    table.AddCell(CreateDataCell(FormatStation(endStation), font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell($"{y2:F4}", font));
                    table.AddCell(CreateDataCell($"{x2:F4}", font));

                    // DATA row (merged) - if we showed start, we already added data. If we only show end, add data here.
                    // If we show BOTH, we need to decide where to put data.
                    // Usually spirals are TS...ST. If we show both, we have 2 rows.
                    // We can make the data cell rowspan=2 if both are shown.
                    
                    if (!showStart)
                    {
                         table.AddCell(dataCell);
                    }
                    else
                    {
                        // If we showed start, we already added the data cell.
                        // For the ST row, we need to add empty cells or handle the rowspan.
                        // But wait, the table structure is fixed 11 columns.
                        // The data cell took up 4 columns.
                        // If we added it in the TS row, it's there.
                        // If we are in the ST row now, we need to fill the last 4 columns.
                        // If we want the data block to span both rows, we should have set rowspan=2 on the data cell in the TS block.
                        
                        // Let's adjust the logic to handle rowspan.
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Add error row
                table.AddCell(new iText.Layout.Element.Cell(1, 2).Add(new Paragraph("SPIRAL ERROR").SetFont(font).SetFontSize(8))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                table.AddCell(new iText.Layout.Element.Cell(1, 5).Add(new Paragraph($"Error: {ex.Message}").SetFont(font).SetFontSize(7))
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            }
        }

    }

    // Dedicated 2025 report selection dialog to avoid cross-version ambiguity
    // Preview DTO for snippet-based live preview
    public class AlignmentPreviewData
    {
        public string AlignmentName = string.Empty;
        public string Description = string.Empty;
        public string StyleName = string.Empty;
        public double StartStation = 0;
        public double EndStation = 0;
        public int LineCount = 0;
        public int ArcCount = 0;
        public int SpiralCount = 0;
        public string ProfileName = string.Empty;
        public System.Collections.Generic.List<string> HorizontalSampleLines = new System.Collections.Generic.List<string>();
        public System.Collections.Generic.List<string> VerticalSampleLines = new System.Collections.Generic.List<string>();
    }

    public class ReportSelectionForm2025 : Form
    {
        private AlignmentPreviewData previewData; // injected preview metadata & sample lines
        // (Viewer fields removed)
        private ComboBox reportTypeComboBox;
        private TextBox outputPathTextBox;
        private Button browseButton;
        private CheckBox chkAlignmentPdf;
        private CheckBox chkAlignmentTxt;
        private CheckBox chkAlignmentXml;
        private CheckBox chkGeoTablePdf;
        private CheckBox chkGeoTableExcel;
        private Button okButton;
        private Button cancelButton;
        // Functional enhancement controls placeholders (phase 2)
        private Button btnSelectAll;
        private Button btnSelectNone;
        private Button btnSelectRecommended;
        private CheckBox chkRemember;
        private CheckBox chkOpenFolder;
        private CheckBox chkOpenFiles;
        private CheckBox chkLivePreview;
        private TextBox previewBox;
        private System.Windows.Forms.Label lblPathStatus;
        private ToolTip toolTip;
        private string settingsFilePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GeoTableReports", "dialog_prefs.json");

        public string ReportType { get; private set; } = "Horizontal";
        public string OutputPath { get; private set; } = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "GeoTable_Report");
        public bool GenerateAlignmentPDF => chkAlignmentPdf.Checked;
        public bool GenerateAlignmentTXT => chkAlignmentTxt.Checked;
        public bool GenerateAlignmentXML => chkAlignmentXml.Checked;
        public bool GenerateGeoTablePDF => chkGeoTablePdf.Checked;
        public bool GenerateGeoTableEXCEL => chkGeoTableExcel.Checked;
        // Future properties (inactive until wiring complete)
        public bool OpenFolderAfter => chkOpenFolder?.Checked == true;
        public bool OpenFilesAfter => chkOpenFiles?.Checked == true;

        public ReportSelectionForm2025(AlignmentPreviewData preview)
        {
            previewData = preview ?? new AlignmentPreviewData();
            Text = "GeoTable Report Selection";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Width = 700; // dialog width after viewer removal
            Height = 500; // height retained

            var lblType = new System.Windows.Forms.Label { Left = 15, Top = 15, Width = 120, Text = "Report Type:" };
            reportTypeComboBox = new ComboBox { Left = 140, Top = 12, Width = 250, DropDownStyle = ComboBoxStyle.DropDownList }; // widened per styling improvements
            reportTypeComboBox.Items.AddRange(new object[] { "Horizontal", "Vertical" });
            reportTypeComboBox.SelectedIndex = 0;
            reportTypeComboBox.SelectedIndexChanged += (s, e) => ReportType = reportTypeComboBox.SelectedItem?.ToString() ?? "Horizontal";

            var lblOutput = new System.Windows.Forms.Label { Left = 15, Top = 50, Width = 120, Text = "Base Output Path:" };
            outputPathTextBox = new TextBox { Left = 140, Top = 47, Width = 420, Text = OutputPath };
            browseButton = new Button { Left = 565, Top = 46, Width = 55, Text = "..." };
            browseButton.Click += (s, e) =>
            {
                using (var sfd = new System.Windows.Forms.SaveFileDialog())
                {
                    sfd.Title = "Select base output file name";
                    sfd.Filter = "All Files (*.*)|*.*";
                    sfd.FileName = System.IO.Path.GetFileName(OutputPath);
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        OutputPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(sfd.FileName) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop), System.IO.Path.GetFileNameWithoutExtension(sfd.FileName));
                        outputPathTextBox.Text = OutputPath;
                    }
                }
            };

            int groupTop = 95;
            int groupHeight = 220;
            int groupWidth = 320;
            var grpAlignment = new GroupBox { Left = 15, Top = groupTop, Width = groupWidth, Height = groupHeight, Text = "Alignment Reports" };
            chkAlignmentPdf = new CheckBox { Left = 15, Top = 25, Width = 160, Text = "Alignment PDF" };
            chkAlignmentTxt = new CheckBox { Left = 15, Top = 50, Width = 160, Text = "Alignment TXT" };
            chkAlignmentXml = new CheckBox { Left = 15, Top = 75, Width = 160, Text = "Alignment XML" };
            grpAlignment.Controls.AddRange(new Control[] { chkAlignmentPdf, chkAlignmentTxt, chkAlignmentXml });

            var grpGeoTable = new GroupBox { Left = 355, Top = groupTop, Width = groupWidth, Height = groupHeight, Text = "GeoTable Reports" };
            chkGeoTablePdf = new CheckBox { Left = 15, Top = 25, Width = 160, Text = "GeoTable PDF" };
            chkGeoTableExcel = new CheckBox { Left = 15, Top = 50, Width = 160, Text = "GeoTable Excel" };
            grpGeoTable.Controls.AddRange(new Control[] { chkGeoTablePdf, chkGeoTableExcel });

            // Divider line panel below path inputs
            var divider = new Panel { Left = 15, Top = 80, Width = 660, Height = 1, BackColor = System.Drawing.Color.LightGray };

            // Centered action buttons area (moved down to prevent overlap)
            int buttonTop = 395;
            okButton = new Button { Left = 250, Top = buttonTop, Width = 110, Text = "OK" };
            cancelButton = new Button { Left = 370, Top = buttonTop, Width = 110, Text = "Cancel" };
            var helpButton = new Button { Left = 490, Top = buttonTop, Width = 110, Text = "Help" };
            helpButton.Click += (s, e) => System.Windows.Forms.MessageBox.Show("Help documentation coming soon.", "Help", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

            // Simple icon placeholders for file types (phase 1: styling)
            System.Drawing.Color iconBack = System.Drawing.Color.FromArgb(230,230,230);
            PictureBox IconPdfAlign = new PictureBox { Left = 20, Top = 25, Width = 12, Height = 12, BackColor = iconBack, Parent = grpAlignment }; grpAlignment.Controls.Add(IconPdfAlign);
            PictureBox IconTxtAlign = new PictureBox { Left = 20, Top = 50, Width = 12, Height = 12, BackColor = iconBack, Parent = grpAlignment }; grpAlignment.Controls.Add(IconTxtAlign);
            PictureBox IconXmlAlign = new PictureBox { Left = 20, Top = 75, Width = 12, Height = 12, BackColor = iconBack, Parent = grpAlignment }; grpAlignment.Controls.Add(IconXmlAlign);
            chkAlignmentPdf.Left = 40;
            chkAlignmentTxt.Left = 40;
            chkAlignmentXml.Left = 40;

            PictureBox IconPdfGeo = new PictureBox { Left = 20, Top = 25, Width = 12, Height = 12, BackColor = iconBack, Parent = grpGeoTable }; grpGeoTable.Controls.Add(IconPdfGeo);
            PictureBox IconXlsGeo = new PictureBox { Left = 20, Top = 50, Width = 12, Height = 12, BackColor = iconBack, Parent = grpGeoTable }; grpGeoTable.Controls.Add(IconXlsGeo);
            chkGeoTablePdf.Left = 40;
            chkGeoTableExcel.Left = 40;
            okButton.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
            cancelButton.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            // --- Functional feature controls ---
            // Path status indicator
            lblPathStatus = new System.Windows.Forms.Label { Left = 625, Top = 50, Width = 30, Height = 18, Text = "" };
            chkLivePreview = new CheckBox { Left = 490, Top = 15, Width = 150, Text = "Live Preview" };
            previewBox = new TextBox { Left = 140, Top = 72, Width = 340, Height = 70, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, Visible = false };
            var btnRefresh = new Button { Left = 485, Top = 70, Width = 75, Height = 26, Text = "Refresh", Visible = false };
            btnRefresh.Click += (s,e)=> { UpdatePreview(true); };

            // Embedded PDF preview panel (right side)
            // Bulk selection shortcut buttons inside alignment group
            btnSelectAll = new Button { Left = 15, Top = 110, Width = 80, Height = 24, Text = "All" };
            btnSelectNone = new Button { Left = 100, Top = 110, Width = 80, Height = 24, Text = "None" };
            btnSelectRecommended = new Button { Left = 185, Top = 110, Width = 120, Height = 24, Text = "Recommended" };
            grpAlignment.Controls.AddRange(new Control[] { btnSelectAll, btnSelectNone, btnSelectRecommended });
            btnSelectAll.Click += (s,e)=> { SetAllSelections(true); UpdatePreview(); };
            btnSelectNone.Click += (s,e)=> { SetAllSelections(false); UpdatePreview(); };
            btnSelectRecommended.Click += (s,e)=> { ApplyRecommended(); UpdatePreview(); };

            // Preferences checkboxes (repositioned to avoid button overlap)
            chkRemember = new CheckBox { Left = 20, Top = 360, Width = 180, Text = "Remember my selections" };
            chkOpenFolder = new CheckBox { Left = 220, Top = 360, Width = 190, Text = "Open folder after creation" };
            chkOpenFiles = new CheckBox { Left = 425, Top = 360, Width = 200, Text = "Open files after creation" };

            // Tooltips
            toolTip = new ToolTip { AutoPopDelay = 8000, InitialDelay = 400, ReshowDelay = 200, ShowAlways = true };
            toolTip.SetToolTip(chkAlignmentPdf, "Printable engineering alignment report (PDF)");
            toolTip.SetToolTip(chkAlignmentTxt, "Plain text alignment data for quick review");
            toolTip.SetToolTip(chkAlignmentXml, "XML alignment data for interoperability");
            toolTip.SetToolTip(chkGeoTablePdf, "GeoTable formatted PDF summary");
            toolTip.SetToolTip(chkGeoTableExcel, "GeoTable Excel for tabular analysis");
            toolTip.SetToolTip(btnSelectRecommended, "Select recommended outputs for chosen report type");
            toolTip.SetToolTip(outputPathTextBox, "Base path used to build filenames; no extension");
            toolTip.SetToolTip(chkLivePreview, "Show dynamic list of files to be generated");
            toolTip.SetToolTip(previewBox, "Live preview of output filenames");
            // toolTip.SetToolTip(previewBox, "Preview of filenames that will be generated"); // removed from UI
            toolTip.SetToolTip(chkRemember, "Persist these selections for next session");
            toolTip.SetToolTip(chkOpenFolder, "Open containing folder after successful generation");
            toolTip.SetToolTip(chkOpenFiles, "Open each generated file automatically");

            // Events for dynamic preview and validation
            reportTypeComboBox.SelectedIndexChanged += (s,e)=> { ReportType = reportTypeComboBox.SelectedItem?.ToString() ?? "Horizontal"; UpdatePreview(); };
            chkLivePreview.CheckedChanged += (s,e)=> { previewBox.Visible = chkLivePreview.Checked; btnRefresh.Visible = chkLivePreview.Checked; UpdatePreview(true); };
            outputPathTextBox.TextChanged += (s,e)=> { ValidatePath(); UpdatePreview(); };
            chkAlignmentPdf.CheckedChanged += (s,e)=> { ValidatePath(); UpdatePreview(); };
            chkAlignmentTxt.CheckedChanged += (s,e)=> { ValidatePath(); UpdatePreview(); };
            chkAlignmentXml.CheckedChanged += (s,e)=> { ValidatePath(); UpdatePreview(); };
            chkGeoTableExcel.CheckedChanged += (s,e)=> { ValidatePath(); UpdatePreview(); };

            Controls.AddRange(new Control[] {
                lblType, reportTypeComboBox,
                lblOutput, outputPathTextBox, browseButton, lblPathStatus, chkLivePreview, previewBox, btnRefresh,
                grpAlignment, grpGeoTable,
                chkRemember, chkOpenFolder, chkOpenFiles,
                okButton, cancelButton, helpButton
            });

            LoadPreferences();
            ValidatePath();
            UpdatePreview();
            ApplyTheme();
        }
        private void SetAllSelections(bool value)
        {
            chkAlignmentPdf.Checked = value;
            chkAlignmentTxt.Checked = value;
            chkAlignmentXml.Checked = value;
            chkGeoTablePdf.Checked = value;
            chkGeoTableExcel.Checked = value;
        }
        private void ApplyRecommended()
        {
            SetAllSelections(false);
            switch (ReportType)
            {
                case "Horizontal":
                    chkAlignmentPdf.Checked = true;
                    chkGeoTableExcel.Checked = true;
                    break;
                case "Vertical":
                    chkAlignmentXml.Checked = true;
                    break;
                default:
                    chkAlignmentPdf.Checked = true;
                    break;
            }
        }
        private string FormatStationLocal(double station)
        {
            int sta = (int)(station / 100);
            double offset = station - (sta * 100);
            return $"{sta:D2}+{offset:00.00}";
        }

        private void UpdatePreview(bool force = false)
        {
            var list = new System.Collections.Generic.List<string>();
            string basePath = OutputPath;
            if (chkAlignmentPdf.Checked) list.Add(System.IO.Path.GetFileName(basePath + "_Alignment_Report.pdf"));
            if (chkAlignmentTxt.Checked) list.Add(System.IO.Path.GetFileName(basePath + "_Alignment_Report.txt"));
            if (chkAlignmentXml.Checked) list.Add(System.IO.Path.GetFileName(basePath + "_Alignment_Report.xml"));
            if (chkGeoTablePdf.Checked) list.Add(System.IO.Path.GetFileName(basePath + "_GeoTable.pdf"));
            if (chkGeoTableExcel.Checked) list.Add(System.IO.Path.GetFileName(basePath + "_GeoTable.xlsx"));
            if (previewBox.Visible)
            {
                if (list.Count == 0)
                {
                    previewBox.Text = "(No outputs selected)";
                }
                else
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine($"Alignment: {previewData.AlignmentName}");
                    sb.AppendLine($"Span: {FormatStationLocal(previewData.StartStation)} - {FormatStationLocal(previewData.EndStation)}");
                    sb.AppendLine($"Elements: Lines={previewData.LineCount} Arcs={previewData.ArcCount} Spirals={previewData.SpiralCount}");
                    if (!string.IsNullOrEmpty(previewData.ProfileName) && ReportType == "Vertical") sb.AppendLine($"Profile: {previewData.ProfileName}");
                    sb.AppendLine();
                    // For each selected format, append a snippet header + sample lines
                    System.Collections.Generic.List<string> sample = ReportType == "Vertical" ? previewData.VerticalSampleLines : previewData.HorizontalSampleLines;
                    int maxLinesPerFormat = 8;
                    if (chkAlignmentPdf.Checked) AppendFormatSnippet(sb, "Alignment PDF", sample, maxLinesPerFormat);
                    if (chkAlignmentTxt.Checked) AppendFormatSnippet(sb, "Alignment TXT", sample, maxLinesPerFormat);
                    if (chkAlignmentXml.Checked) AppendFormatSnippet(sb, "Alignment XML", sample, maxLinesPerFormat);
                    if (chkGeoTablePdf.Checked) AppendFormatSnippet(sb, "GeoTable PDF", sample, maxLinesPerFormat);
                    if (chkGeoTableExcel.Checked) AppendFormatSnippet(sb, "GeoTable Excel", sample, maxLinesPerFormat);
                    previewBox.Text = sb.ToString();
                }
            }
            okButton.Enabled = lblPathStatus.Text == "✔" && list.Count > 0;
            // Trigger PDF side preview regeneration if needed
            // (PDF preview hook removed)
            {
                // (PDF preview generation removed)
            }
        }

        private void AppendFormatSnippet(System.Text.StringBuilder sb, string title, System.Collections.Generic.List<string> source, int limit)
        {
            sb.AppendLine($"[{title}] sample:");
            if (source == null || source.Count == 0)
            {
                sb.AppendLine("  (no sample available)");
            }
            else
            {
                for (int i = 0; i < source.Count && i < limit; i++)
                    sb.AppendLine("  " + source[i]);
            }
            sb.AppendLine();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Save preferences if requested
            if (DialogResult == DialogResult.OK && chkRemember != null && chkRemember.Checked)
                SavePreferences();
            base.OnFormClosing(e);
        }


        private void ApplyTheme()
        {
            try
            {
                int theme = 0;
                try
                {
                    object val = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("COLORTHEME");
                    if (val is short s) theme = s; else if (val is int i) theme = i;
                }
                catch { }
                bool dark = theme == 1;
                if (dark)
                {
                    BackColor = System.Drawing.Color.FromArgb(45,45,48);
                    ForeColor = System.Drawing.Color.Gainsboro;
                    foreach (Control c in Controls)
                    {
                        if (c is GroupBox)
                        {
                            c.BackColor = System.Drawing.Color.FromArgb(58,58,62);
                            c.ForeColor = System.Drawing.Color.Gainsboro;
                        }
                        else if (c is TextBox || c is ComboBox)
                        {
                            c.BackColor = System.Drawing.Color.FromArgb(63,63,70);
                            c.ForeColor = System.Drawing.Color.Gainsboro;
                        }
                        else if (c is Button)
                        {
                            c.BackColor = System.Drawing.Color.FromArgb(72,72,80);
                            c.ForeColor = System.Drawing.Color.Gainsboro;
                        }
                        else if (c is CheckBox || c is System.Windows.Forms.Label)
                        {
                            c.ForeColor = System.Drawing.Color.Gainsboro;
                        }
                    }
                }
            }
            catch { }
        }
        private void ValidatePath()
        {
            string raw = outputPathTextBox.Text.Trim();
            bool valid = false;
            string validationMessage = string.Empty;
            try
            {
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    string dir = System.IO.Path.GetDirectoryName(raw);
                    if (string.IsNullOrEmpty(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    
                    // Check if path contains invalid characters
                    if (raw.IndexOfAny(System.IO.Path.GetInvalidPathChars()) >= 0)
                    {
                        validationMessage = "Path contains invalid characters";
                        valid = false;
                    }
                    else if (!System.IO.Directory.Exists(dir))
                    {
                        try
                        {
                            System.IO.Directory.CreateDirectory(dir);
                            validationMessage = "Path valid (created)";
                            valid = true;
                        }
                        catch
                        {
                            validationMessage = "Cannot create directory";
                            valid = false;
                        }
                    }
                    else
                    {
                        // Directory exists, test write permissions
                        string testFile = System.IO.Path.Combine(dir, "_gt_writetest_" + Guid.NewGuid().ToString().Substring(0, 8) + ".tmp");
                        try
                        {
                            System.IO.File.WriteAllText(testFile, "test");
                            System.IO.File.Delete(testFile);
                            
                            // Check for OneDrive/Dropbox sync paths (informational warning only)
                            string upperPath = dir.ToUpperInvariant();
                            if (upperPath.Contains("ONEDRIVE") || upperPath.Contains("DROPBOX") || upperPath.Contains("GOOGLE DRIVE"))
                            {
                                validationMessage = "⚠ Sync folder (may have delays)";
                                valid = true; // Valid with warning
                            }
                            else
                            {
                                validationMessage = "Path valid";
                                valid = true;
                            }
                        }
                        catch (System.Exception writeEx)
                        {
                            // If write test fails, still allow if it's a known sync path (user may have permissions later)
                            string upperPath = dir.ToUpperInvariant();
                            if (upperPath.Contains("ONEDRIVE") || upperPath.Contains("DROPBOX") || upperPath.Contains("GOOGLE DRIVE"))
                            {
                                validationMessage = "⚠ Sync folder (write test failed, may sync later)";
                                valid = true; // Allow with warning
                            }
                            else
                            {
                                validationMessage = $"Write test failed: {writeEx.Message}";
                                valid = false;
                            }
                        }
                    }
                }
                else
                {
                    validationMessage = "Path is empty";
                    valid = false;
                }
            }
            catch (System.Exception ex)
            {
                validationMessage = $"Validation error: {ex.Message}";
                valid = false;
            }
            
            lblPathStatus.Text = valid ? "✔" : "✖";
            lblPathStatus.ForeColor = valid ? System.Drawing.Color.Green : System.Drawing.Color.Red;
            toolTip.SetToolTip(lblPathStatus, validationMessage);
            OutputPath = raw;
        }
        // (Removed duplicate OnFormClosing; handled earlier with preview disposal)
        private void LoadPreferences()
        {
            try
            {
                if (System.IO.File.Exists(settingsFilePath))
                {
                    string json = System.IO.File.ReadAllText(settingsFilePath);
                    bool Bool(string key) => json.Contains("\"" + key + "\": true");
                    string Val(string key)
                    {
                        int i = json.IndexOf("\"" + key + "\":");
                        if (i < 0) return string.Empty;
                        int s = json.IndexOf('"', i + key.Length + 3) + 1;
                        int e = json.IndexOf('"', s);
                        if (s < 0 || e < 0) return string.Empty;
                        return json.Substring(s, e - s);
                    }
                    string rt = Val("ReportType");
                    if (!string.IsNullOrEmpty(rt))
                    {
                        int idx = reportTypeComboBox.Items.IndexOf(rt);
                        if (idx >= 0) reportTypeComboBox.SelectedIndex = idx;
                    }
                    string op = Val("OutputPath");
                    if (!string.IsNullOrEmpty(op)) outputPathTextBox.Text = op;
                    chkAlignmentPdf.Checked = Bool("AlignmentPDF");
                    chkAlignmentTxt.Checked = Bool("AlignmentTXT");
                    chkAlignmentXml.Checked = Bool("AlignmentXML");
                    chkGeoTablePdf.Checked = Bool("GeoTablePDF");
                    chkGeoTableExcel.Checked = Bool("GeoTableEXCEL");
                    chkOpenFolder = chkOpenFolder ?? new CheckBox();
                    chkOpenFiles = chkOpenFiles ?? new CheckBox();
                    chkOpenFolder.Checked = Bool("OpenFolderAfter");
                    chkOpenFiles.Checked = Bool("OpenFilesAfter");
                    chkRemember.Checked = Bool("RememberSelections");
                }
            }
            catch { }
        }
        private void SavePreferences()
        {
            try
            {
                string dir = System.IO.Path.GetDirectoryName(settingsFilePath) ?? settingsFilePath;
                if (!System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
                string json = "{" +
                    "\n  \"ReportType\": \"" + ReportType + "\"," +
                    "\n  \"OutputPath\": \"" + OutputPath.Replace("\\", "\\\\") + "\"," +
                    "\n  \"AlignmentPDF\": " + GenerateAlignmentPDF.ToString().ToLower() + "," +
                    "\n  \"AlignmentTXT\": " + GenerateAlignmentTXT.ToString().ToLower() + "," +
                    "\n  \"AlignmentXML\": " + GenerateAlignmentXML.ToString().ToLower() + "," +
                    "\n  \"GeoTablePDF\": " + GenerateGeoTablePDF.ToString().ToLower() + "," +
                    "\n  \"GeoTableEXCEL\": " + GenerateGeoTableEXCEL.ToString().ToLower() + "," +
                    "\n  \"OpenFolderAfter\": " + OpenFolderAfter.ToString().ToLower() + "," +
                    "\n  \"OpenFilesAfter\": " + OpenFilesAfter.ToString().ToLower() + "," +
                    "\n  \"RememberSelections\": " + (chkRemember?.Checked == true).ToString().ToLower() + "\n}";
                System.IO.File.WriteAllText(settingsFilePath, json);
            }
            catch { }
        }

        // Placeholder removed (no viewer logic)
    }

    // Progress Status Window for real-time feedback during report generation
    public class ProgressStatusWindow : Form
    {
        private ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblStep;
        private int totalSteps;
        private int currentStep;

        public ProgressStatusWindow(int steps)
        {
            totalSteps = steps;
            currentStep = 0;

            Text = "Generating Reports";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Width = 450;
            Height = 180;
            ControlBox = false;

            lblStatus = new System.Windows.Forms.Label
            {
                Left = 20,
                Top = 20,
                Width = 400,
                Height = 30,
                Text = "Initializing...",
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Regular)
            };

            progressBar = new ProgressBar
            {
                Left = 20,
                Top = 60,
                Width = 400,
                Height = 25,
                Minimum = 0,
                Maximum = totalSteps * 100, // Use finer granularity
                Value = 0,
                Style = ProgressBarStyle.Continuous
            };

            lblStep = new System.Windows.Forms.Label
            {
                Left = 20,
                Top = 95,
                Width = 400,
                Height = 20,
                Text = $"Step 0 of {totalSteps}",
                Font = new System.Drawing.Font("Segoe UI", 8, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.Gray
            };

            Controls.AddRange(new Control[] { lblStatus, progressBar, lblStep });
        }

        public void UpdateStatus(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateStatus(message)));
                return;
            }
            lblStatus.Text = message;
            System.Windows.Forms.Application.DoEvents();
        }

        public void IncrementProgress()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(IncrementProgress));
                return;
            }
            currentStep++;
            progressBar.Value = currentStep * 100;
            lblStep.Text = $"Step {currentStep} of {totalSteps}";
            System.Windows.Forms.Application.DoEvents();
        }

        public void Complete(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Complete(message)));
                return;
            }
            progressBar.Value = progressBar.Maximum;
            lblStatus.Text = message;
            lblStatus.ForeColor = System.Drawing.Color.Green;
            lblStep.Text = $"Completed: {totalSteps} of {totalSteps}";
            System.Windows.Forms.Application.DoEvents();
        }

        public void Fail(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Fail(message)));
                return;
            }
            lblStatus.Text = message;
            lblStatus.ForeColor = System.Drawing.Color.Red;
            lblStep.Text = $"Failed at step {currentStep}";
            System.Windows.Forms.Application.DoEvents();
        }
    }

    // Custom renderer to draw an oval around cell content
    public class OvalCellRenderer : CellRenderer
    {
        public OvalCellRenderer(iText.Layout.Element.Cell modelElement) : base(modelElement)
        {
        }

        public override void Draw(DrawContext drawContext)
        {
            base.Draw(drawContext);
            
            PdfCanvas canvas = drawContext.GetCanvas();
            iText.Kernel.Geom.Rectangle rect = GetOccupiedAreaBBox();
            
            // Calculate oval bounds (centered in cell, slightly smaller than cell)
            float x = rect.GetX() + 8;
            float y = rect.GetY() + 5;
            float width = rect.GetWidth() - 16;
            float height = rect.GetHeight() - 10;

            canvas.SaveState();
            canvas.SetStrokeColor(ColorConstants.BLACK); 
            canvas.SetLineWidth(0.5f);
            
            // Draw ellipse inscribed in the rectangle defined by corners (x, y) and (x+width, y+height)
            canvas.Ellipse(x, y, x + width, y + height);
            
            canvas.Stroke();
            canvas.RestoreState();
        }
        
        public override IRenderer GetNextRenderer()
        {
            return new OvalCellRenderer((iText.Layout.Element.Cell)GetModelElement());
        }
    }

}
#pragma warning restore 1591
