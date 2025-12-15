#pragma warning disable 1591
#pragma warning disable CA1416
/*
 * Civil 3D GeoTable Reports - .NET Add-in
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
    public interface IReportGenerator
    {
        void GenerateReport();
    }

    public class ReportCommands2025 : IExtensionApplication, IReportGenerator
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

        public void Initialize()
        {
            AcApp.Document doc = AcApp.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.Editor.WriteMessage("\nGeoTable Reports loaded. Type GEOTABLE_PANEL to open the panel or GEOTABLE for quick generation.");
            }
        }

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
                // Select alignment before showing dialog so preview has real data
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
                        int totalSteps = (generateAlignmentPDF ? 1 : 0) + (generateAlignmentXML ? 1 : 0) + (generateGeoTablePDF ? 1 : 0) + (generateGeoTableEXCEL ? 1 : 0);
                        ProgressStatusWindow progressWindow = null;
                        if (totalSteps > 0)
                        {
                            progressWindow = new ProgressStatusWindow(totalSteps);
                            progressWindow.Show();
                            System.Windows.Forms.Application.DoEvents();
                        }
                        try
                        {
                            ExecuteReportGeneration(alignment, baseOutputPath, reportType, generateAlignmentPDF, generateAlignmentXML, generateGeoTablePDF, generateGeoTableEXCEL, progressWindow, generatedFiles);

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

        private void ExecuteReportGeneration(
            CivDb.Alignment alignment,
            string baseOutputPath,
            string reportType,
            bool generateAlignmentPDF,
            bool generateAlignmentXML,
            bool generateGeoTablePDF,
            bool generateGeoTableEXCEL,
            ProgressStatusWindow progressWindow,
            System.Collections.Generic.List<string> generatedFiles)
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
            if (generateGeoTablePDF && reportType != "Vertical")
            {
                progressWindow?.UpdateStatus("Generating GeoTable PDF...");
                string geoPdf = baseOutputPath + "_GeoTable.pdf";
                GenerateHorizontalGeoTablePdf(alignment, geoPdf);
                generatedFiles.Add(geoPdf);
                progressWindow?.IncrementProgress();
            }
            if (generateGeoTableEXCEL && reportType != "Vertical")
            {
                progressWindow?.UpdateStatus("Generating GeoTable Excel...");
                string geoXlsx = baseOutputPath + "_GeoTable.xlsx";
                GenerateHorizontalGeoTableExcel(alignment, geoXlsx);
                generatedFiles.Add(geoXlsx);
                progressWindow?.IncrementProgress();
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

                    PopulateAlignmentMetadata(alignment, data);
                    PopulateHorizontalSamples(alignment, data);
                    PopulateVerticalSamples(alignment, data);

                    tr.Commit();
                }
            }
            catch { }
            return data;
        }

        private void PopulateAlignmentMetadata(CivDb.Alignment alignment, AlignmentPreviewData data)
        {
            data.AlignmentName = alignment.Name ?? "(Unnamed)";
            try { data.Description = alignment.Description ?? ""; } catch { }
            try { data.StyleName = alignment.StyleName ?? ""; } catch { }
            
            double startSta = 0; 
            double endSta = 0;
            try 
            { 
                if (alignment.Entities.Count > 0) 
                { 
                    startSta = (alignment.Entities[0] as dynamic).StartStation; 
                    var last = alignment.Entities[alignment.Entities.Count - 1]; 
                    endSta = (last as dynamic).EndStation; 
                } 
            } 
            catch { }
            data.StartStation = startSta; 
            data.EndStation = endSta;
        }

        private void PopulateHorizontalSamples(CivDb.Alignment alignment, AlignmentPreviewData data)
        {
            int lines = 0, arcs = 0, spirals = 0;
            for (int i = 0; i < alignment.Entities.Count; i++)
            {
                var e = alignment.Entities[i];
                if (e == null) continue;

                // Civil 3D may represent Spiral-Curve-Spiral as a single AlignmentSCS entity.
                // For reporting, treat it as 2 spirals + 1 arc so downstream formats remain consistent.
                if (TryGetSpiralCurveSpiral(e, out var scs))
                {
                    if (scs.SpiralIn != null) spirals++;
                    if (scs.Arc != null) arcs++;
                    if (scs.SpiralOut != null) spirals++;

                    if (data.HorizontalSampleLines.Count < 15)
                    {
                        foreach (var expanded in ExpandEntityForReports(e))
                        {
                            if (expanded == null) continue;
                            if (data.HorizontalSampleLines.Count >= 15) break;
                            try { AddHorizontalSampleLine(alignment, expanded, i, data); } catch { }
                        }
                    }
                    continue;
                }

                if (e.EntityType == AlignmentEntityType.Line) lines++;
                else if (e.EntityType == AlignmentEntityType.Arc) arcs++;
                else if (e.EntityType == AlignmentEntityType.Spiral) spirals++;

                if (data.HorizontalSampleLines.Count < 15)
                {
                    try { AddHorizontalSampleLine(alignment, e, i, data); } catch { }
                }
            }
            data.LineCount = lines; 
            data.ArcCount = arcs; 
            data.SpiralCount = spirals;
        }

        private void AddHorizontalSampleLine(CivDb.Alignment alignment, AlignmentEntity e, int index, AlignmentPreviewData data)
        {
            switch (e.EntityType)
            {
                case AlignmentEntityType.Line:
                    if (e is AlignmentLine l)
                    {
                        var (x1, y1, _) = GetPointAtStation(alignment, l.StartStation);
                        var (x2, y2, _) = GetPointAtStation(alignment, l.EndStation);
                        string tangentStartLabel = (index == 0) ? "POB" : "PI";
                        data.HorizontalSampleLines.Add($"{tangentStartLabel} {FormatStation(l.StartStation),15} {FormatRounded(y1, 4),15} {FormatRounded(x1, 4),15}");
                        data.HorizontalSampleLines.Add($"PI {FormatStation(l.EndStation),15} {FormatRounded(y2, 4),15} {FormatRounded(x2, 4),15}");
                    }
                    break;
                case AlignmentEntityType.Arc:
                    if (e is AlignmentArc a)
                    {
                        var (x1, y1, _) = GetPointAtStation(alignment, a.StartStation);
                        var (x2, y2, _) = GetPointAtStation(alignment, a.EndStation);
                        data.HorizontalSampleLines.Add($"PC {FormatStation(a.StartStation),15} {FormatRounded(y1, 4),15} {FormatRounded(x1, 4),15}");
                        data.HorizontalSampleLines.Add($"PT {FormatStation(a.EndStation),15} {FormatRounded(y2, 4),15} {FormatRounded(x2, 4),15}");
                    }
                    break;
                case AlignmentEntityType.Spiral:
                    try
                    {
                        dynamic dyn = e;
                        double ss = dyn.StartStation; 
                        double es = dyn.EndStation; 
                        var (xm, ym, _) = GetPointAtStation(alignment, (ss + es) / 2);
                        var (x1, y1, _) = GetPointAtStation(alignment, ss);
                        var (x2, y2, _) = GetPointAtStation(alignment, es);
                        data.HorizontalSampleLines.Add($"TS {FormatStation(ss),15} {FormatRounded(y1, 4),15} {FormatRounded(x1, 4),15}");
                        data.HorizontalSampleLines.Add($"SPI{FormatStation((ss + es) / 2),15} {FormatRounded(ym, 4),15} {FormatRounded(xm, 4),15}");
                        data.HorizontalSampleLines.Add($"SC {FormatStation(es),15} {FormatRounded(y2, 4),15} {FormatRounded(x2, 4),15}");
                    }
                    catch { }
                    break;
            }
        }

        private void PopulateVerticalSamples(CivDb.Alignment alignment, AlignmentPreviewData data)
        {
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
                                AddVerticalSampleLine(profile.Entities[i], data);
                            }
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        private void AddVerticalSampleLine(ProfileEntity pe, AlignmentPreviewData data)
        {
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
            else if (pe.EntityType == ProfileEntityType.ParabolaSymmetric && pe is ProfileParabolaSymmetric ps)
            {
                data.VerticalSampleLines.Add($"PVC {FormatStation(ps.StartStation),15} {ps.StartElevation,12:F2}");
                data.VerticalSampleLines.Add($"PVI {FormatStation(ps.PVIStation),15} {ps.PVIElevation,12:F2}");
                data.VerticalSampleLines.Add($"PVT {FormatStation(ps.EndStation),15} {ps.EndElevation,12:F2}");
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
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            // NORTHING with left-divider
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(FormatRounded(northing, 4)).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            // EASTING
            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(FormatRounded(easting, 4)).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            document.Add(dataTable);
        }

        /// <summary>
        /// Helper method to add a data row to vertical alignment table (2 columns)
        /// </summary>
        private void AddVerticalDataRow(Document document, string label, string station, double elevation, PdfFont font)
        {
            iText.Layout.Element.Table dataTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 75, 25 }));
            dataTable.SetWidth(UnitValue.CreatePercentValue(100));

            // Inner table to separate Label (Left) and Station (Right)
            iText.Layout.Element.Table innerTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 50, 50 }));
            innerTable.SetWidth(UnitValue.CreatePercentValue(100));

            innerTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(label.Trim()).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            innerTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(station).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(innerTable)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetPadding(0));

            dataTable.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph(FormatRounded(elevation, 2)).SetFont(font).SetFontSize(9))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT)
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f)));

            document.Add(dataTable);
        }

        /// <summary>
        /// Generate horizontal alignment report
        /// </summary>
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

        private System.Collections.Generic.List<ProfileEntity> GetSortedProfileEntities(Profile profile)
        {
            var entities = new System.Collections.Generic.List<ProfileEntity>();
            if (profile.Entities != null)
            {
                for (int i = 0; i < profile.Entities.Count; i++)
                {
                    if (profile.Entities[i] != null)
                        entities.Add(profile.Entities[i]);
                }

                entities.Sort((a, b) =>
                {
                    try
                    {
                        // Use dynamic to access StartStation if needed, or cast if types are known
                        // ProfileEntity usually has StartStation
                        double sa = 0;
                        double sb = 0;
                        
                        if (a is ProfileTangent t1) sa = t1.StartStation;
                        else if (a is ProfileCircular c1) sa = c1.StartStation;
                        else if (a is ProfileParabolaSymmetric p1) sa = p1.StartStation;
                        else 
                        {
                            // Try dynamic fallback
                            try { sa = (a as dynamic).StartStation; } catch { }
                        }

                        if (b is ProfileTangent t2) sb = t2.StartStation;
                        else if (b is ProfileCircular c2) sb = c2.StartStation;
                        else if (b is ProfileParabolaSymmetric p2) sb = p2.StartStation;
                        else
                        {
                            // Try dynamic fallback
                            try { sb = (b as dynamic).StartStation; } catch { }
                        }

                        return sa.CompareTo(sb);
                    }
                    catch
                    {
                        return 0;
                    }
                });
            }
            return entities;
        }

        private struct ArcData
        {
            public double X1, Y1, X2, Y2, XC, YC, XPI, YPI;
            public double PIStation;
            public double DeltaRadians, DeltaDegrees;
            public double Tangent, Chord, MiddleOrdinate, External;
            public double DegreeOfCurvature;
        }

        private ArcData GetArcData(CivDb.Alignment alignment, AlignmentArc arc)
        {
            var data = new ArcData();
            var (x1, y1, _) = GetPointAtStation(alignment, arc.StartStation);
            var (x2, y2, _) = GetPointAtStation(alignment, arc.EndStation);
            data.X1 = x1; data.Y1 = y1;
            data.X2 = x2; data.Y2 = y2;
            data.PIStation = arc.PIStation;

            bool gotFromSubEntity = false;
            try
            {
                if (arc.SubEntityCount > 0 && arc[0] is AlignmentSubEntityArc sub)
                {
                    data.XPI = sub.PIPoint.X;
                    data.YPI = sub.PIPoint.Y;
                    data.XC = sub.CenterPoint.X;
                    data.YC = sub.CenterPoint.Y;
                    gotFromSubEntity = true;
                }
            }
            catch { }

            data.DeltaRadians = arc.Length / Math.Abs(arc.Radius);
            data.Tangent = arc.Radius * Math.Tan(Math.Abs(data.DeltaRadians) / 2);

            if (!gotFromSubEntity)
            {
                double backTangentDir = arc.StartDirection + Math.PI;
                data.XPI = x1 + data.Tangent * Math.Cos(backTangentDir);
                data.YPI = y1 + data.Tangent * Math.Sin(backTangentDir);
                
                double midStation = (arc.StartStation + arc.EndStation) / 2;
                double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                var (xc, yc, _) = GetPointAtStation(alignment, midStation, offset);
                data.XC = xc; data.YC = yc;
            }

            data.DeltaDegrees = data.DeltaRadians * (180.0 / Math.PI);
            data.Chord = 2 * arc.Radius * Math.Sin(Math.Abs(data.DeltaRadians) / 2);
            data.MiddleOrdinate = arc.Radius * (1 - Math.Cos(Math.Abs(data.DeltaRadians) / 2));
            data.External = arc.Radius * (1 / Math.Cos(Math.Abs(data.DeltaRadians) / 2) - 1);
            data.DegreeOfCurvature = (100.0 * data.DeltaRadians) / (arc.Length) * (180.0 / Math.PI);

            return data;
        }

        private string FormatStation(double station)
        {
            int sta = (int)(station / 100);
            double offset = station - (sta * 100);
            // Use AwayFromZero rounding for surveying precision
            double roundedOffset = Math.Round(offset, 2, MidpointRounding.AwayFromZero);
            return $"{sta:D2}+{roundedOffset:00.00}";
        }

        private string GetTimeZoneAbbreviation(TimeZoneInfo timeZone)
        {
            if (timeZone == null) return "";
            bool isDst = timeZone.IsDaylightSavingTime(DateTime.Now);
            string name = isDst ? timeZone.DaylightName : timeZone.StandardName;

            var timeZones = new System.Collections.Generic.Dictionary<string, (string Std, string Dst)>
            {
                { "Pacific", ("PST", "PDT") },
                { "Mountain", ("MST", "MDT") },
                { "Central", ("CST", "CDT") },
                { "Eastern", ("EST", "EDT") },
                { "Atlantic", ("AST", "ADT") },
                { "Alaska", ("AKST", "AKDT") },
                { "Hawaii", ("HST", "HDT") }
            };

            foreach (var kvp in timeZones)
            {
                if (name.Contains(kvp.Key)) return isDst ? kvp.Value.Dst : kvp.Value.Std;
            }

            // Fallback: First letter of each word
            var words = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string abbreviation = "";
            foreach (var word in words)
            {
                if (word.Length > 0 && char.IsLetter(word[0]))
                    abbreviation += word[0];
            }
            return abbreviation.ToUpper();
        }

        private (int Deg, int Min, double Sec) ToDms(double degrees)
        {
            degrees = Math.Abs(degrees);
            int d = (int)degrees;
            double rem = (degrees - d) * 60;
            int m = (int)rem;
            double s = (rem - m) * 60;
            return (d, m, s);
        }

        private (string Quad1, string Quad2, double Angle) GetQuadrantAndAngle(double radians)
        {
            double degrees = radians * (180.0 / Math.PI);
            while (degrees < 0) degrees += 360;
            while (degrees >= 360) degrees -= 360;

            if (degrees < 90) return ("N", "E", degrees);
            if (degrees < 180) return ("S", "E", 180 - degrees);
            if (degrees < 270) return ("S", "W", degrees - 180);
            return ("N", "W", 360 - degrees);
        }

        private (double X, double Y, double Z) GetPointAtStation(CivDb.Alignment alignment, double station, double offset = 0)
        {
            double x = 0, y = 0, z = 0;
            alignment.PointLocation(station, offset, 0, ref x, ref y, ref z);
            return (x, y, z);
        }

        private (string StartLabel, string EndLabel) GetLinearLabels(int index, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            string startLabel = (index == 0) ? "POB" :
                                (prevEntity?.EntityType == AlignmentEntityType.Spiral) ? "ST" :
                                (prevEntity?.EntityType == AlignmentEntityType.Arc) ? "PT" : "PI";

            string endLabel = (nextEntity?.EntityType == AlignmentEntityType.Spiral) ? "TS" :
                              (nextEntity?.EntityType == AlignmentEntityType.Arc) ? "PC" : "PI";
            
            return (startLabel, endLabel);
        }

        private (string StartLabel, string EndLabel) GetArcLabels(AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            string startLabel = (prevEntity?.EntityType == AlignmentEntityType.Spiral) ? "SC" : "PC";
            string endLabel = (nextEntity?.EntityType == AlignmentEntityType.Spiral) ? "CS" : "PT";
            return (startLabel, endLabel);
        }

        private string FormatBearing(double radians)
        {
            var (ns, ew, angle) = GetQuadrantAndAngle(radians);
            return $"{ns} {FormatAngle(angle)} {ew}";
        }

        private string FormatAngle(double degrees)
        {
            var (d, m, s) = ToDms(degrees);
            return $"{d}^{m:D2}'{s:F4}\"";
        }

        /// <summary>
        /// Formats a double value to a specified number of decimal places using AwayFromZero rounding.
        /// This ensures proper rounding for surveying/engineering applications (0.5 always rounds up).
        /// Also handles floating point precision issues by checking if value is very close to a round number.
        /// </summary>
        private string FormatRounded(double value, int decimalPlaces)
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

        [Obsolete("Use FormatRounded instead")]
        private string FormatWithProperRounding(double value, int decimalPlaces) => FormatRounded(value, decimalPlaces);

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

            // Expand AlignmentSCS (spiral-curve-spiral) into its component entities so
            // existing XML consumers continue receiving <Arc> and <Spiral> nodes.
            var reportEntities = new System.Collections.Generic.List<AlignmentEntity>();
            for (int i = 0; i < alignment.Entities.Count; i++)
            {
                var entity = alignment.Entities[i];
                if (entity == null) continue;
                foreach (var expanded in ExpandEntityForReports(entity))
                {
                    if (expanded != null) reportEntities.Add(expanded);
                }
            }

            for (int i = 0; i < reportEntities.Count; i++)
            {
                AlignmentEntity entity = reportEntities[i];
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
                            try
                            {
                                dynamic d = entity;
                                double startStation = d.StartStation;
                                double endStation = d.EndStation;
                                double length = d.Length;
                                double radiusIn = 0;
                                double radiusOut = 0;
                                try { radiusIn = d.RadiusIn; } catch { }
                                try { radiusOut = d.RadiusOut; } catch { }

                                double x1 = 0, y1 = 0, z1 = 0;
                                double x2 = 0, y2 = 0, z2 = 0;
                                alignment.PointLocation(startStation, 0, 0, ref x1, ref y1, ref z1);
                                alignment.PointLocation(endStation, 0, 0, ref x2, ref y2, ref z2);

                                elements.Add(new XElement("Spiral",
                                    new XElement("StartStation", startStation),
                                    new XElement("EndStation", endStation),
                                    new XElement("Length", length),
                                    new XElement("RadiusIn", radiusIn),
                                    new XElement("RadiusOut", radiusOut),
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
                            catch
                            {
                                // Skip spirals that can't be processed
                            }
                            break;

                        case AlignmentEntityType.SpiralCurveSpiral:
                            try
                            {
                                var scsEntity = entity as AlignmentSCS;
                                if (scsEntity != null)
                                {
                                    // Add SpiralIn
                                    if (scsEntity.SpiralIn != null)
                                    {
                                        var spiralIn = scsEntity.SpiralIn;
                                        double x1 = 0, y1 = 0, z1 = 0;
                                        double x2 = 0, y2 = 0, z2 = 0;
                                        alignment.PointLocation(spiralIn.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                        alignment.PointLocation(spiralIn.EndStation, 0, 0, ref x2, ref y2, ref z2);
                                        elements.Add(new XElement("Spiral",
                                            new XElement("Type", "SpiralIn"),
                                            new XElement("StartStation", spiralIn.StartStation),
                                            new XElement("EndStation", spiralIn.EndStation),
                                            new XElement("Length", spiralIn.Length),
                                            new XElement("RadiusIn", spiralIn.RadiusIn),
                                            new XElement("RadiusOut", spiralIn.RadiusOut),
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

                                    // Add Arc
                                    if (scsEntity.Arc != null)
                                    {
                                        var scsArc = scsEntity.Arc;
                                        double xa1 = 0, ya1 = 0, za1 = 0;
                                        double xa2 = 0, ya2 = 0, za2 = 0;
                                        alignment.PointLocation(scsArc.StartStation, 0, 0, ref xa1, ref ya1, ref za1);
                                        alignment.PointLocation(scsArc.EndStation, 0, 0, ref xa2, ref ya2, ref za2);
                                        double deltaRadians = scsArc.Length / Math.Abs(scsArc.Radius);
                                        double deltaDegrees = deltaRadians * (180.0 / Math.PI);
                                        elements.Add(new XElement("Arc",
                                            new XElement("Type", "SCSArc"),
                                            new XElement("StartStation", scsArc.StartStation),
                                            new XElement("EndStation", scsArc.EndStation),
                                            new XElement("Length", scsArc.Length),
                                            new XElement("Radius", scsArc.Radius),
                                            new XElement("Delta", Math.Abs(deltaDegrees)),
                                            new XElement("Direction", scsArc.Clockwise ? "Right" : "Left"),
                                            new XElement("StartPoint",
                                                new XElement("Northing", ya1),
                                                new XElement("Easting", xa1)
                                            ),
                                            new XElement("EndPoint",
                                                new XElement("Northing", ya2),
                                                new XElement("Easting", xa2)
                                            )
                                        ));
                                    }

                                    // Add SpiralOut
                                    if (scsEntity.SpiralOut != null)
                                    {
                                        var spiralOut = scsEntity.SpiralOut;
                                        double x1 = 0, y1 = 0, z1 = 0;
                                        double x2 = 0, y2 = 0, z2 = 0;
                                        alignment.PointLocation(spiralOut.StartStation, 0, 0, ref x1, ref y1, ref z1);
                                        alignment.PointLocation(spiralOut.EndStation, 0, 0, ref x2, ref y2, ref z2);
                                        elements.Add(new XElement("Spiral",
                                            new XElement("Type", "SpiralOut"),
                                            new XElement("StartStation", spiralOut.StartStation),
                                            new XElement("EndStation", spiralOut.EndStation),
                                            new XElement("Length", spiralOut.Length),
                                            new XElement("RadiusIn", spiralOut.RadiusIn),
                                            new XElement("RadiusOut", spiralOut.RadiusOut),
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
                                }
                            }
                            catch
                            {
                                // Skip SCS elements that can't be processed
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
            var sortedEntities = GetSortedProfileEntities(profile);

            for (int i = 0; i < sortedEntities.Count; i++)
            {
                ProfileEntity entity = sortedEntities[i];
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

                        case ProfileEntityType.ParabolaSymmetric:
                            var parabola = entity as ProfileParabolaSymmetric;
                            if (parabola != null)
                            {
                                double gradeIn = parabola.GradeIn * 100;
                                double gradeOut = parabola.GradeOut * 100;
                                double r = (gradeOut - gradeIn) / parabola.Length;
                                double k = parabola.Length / (gradeOut - gradeIn);

                                elements.Add(new XElement("ParabolaSymmetric",
                                    new XElement("StartStation", parabola.StartStation),
                                    new XElement("EndStation", parabola.EndStation),
                                    new XElement("PVIStation", parabola.PVIStation),
                                    new XElement("PVIElevation", parabola.PVIElevation),
                                    new XElement("Length", parabola.Length),
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
                document.Add(new Paragraph($"Horizontal Alignment Name: {alignment.Name}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($" Description: {alignment.Description ?? ""}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($" Style: {alignment.StyleName ?? "Default"}")
                    .SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph($"Generated Report:  Date: {DateTime.Now:MM/dd/yyyy}     Time: {DateTime.Now:h:mm tt} {GetTimeZoneAbbreviation(TimeZoneInfo.Local)}")
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
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                    .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f));
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
                    .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                    .SetBorderLeft(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f));
                headerTable.AddCell(cell3);

                document.Add(headerTable);
                document.Add(new Paragraph("\n").SetFontSize(3));

                // Reorder entities to match InRails format
                var reorderedEntities = ReorderEntitiesForInRails(alignment);

                // Expand AlignmentSCS into SpiralIn/Arc/SpiralOut so all parts appear in the report.
                var reportEntities = new System.Collections.Generic.List<AlignmentEntity>();
                for (int i = 0; i < reorderedEntities.Count; i++)
                {
                    var e = reorderedEntities[i];
                    if (e == null) continue;
                    foreach (var expanded in ExpandEntityForReports(e))
                    {
                        if (expanded != null) reportEntities.Add(expanded);
                    }
                }

                // Process each entity
                for (int i = 0; i < reportEntities.Count; i++)
                {
                    AlignmentEntity entity = reportEntities[i];
                    if (entity == null) continue;

                    AlignmentEntity prevEntity = i > 0 ? reportEntities[i - 1] : null;
                    AlignmentEntity nextEntity = i < reportEntities.Count - 1 ? reportEntities[i + 1] : null;

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
                            case AlignmentEntityType.SpiralCurveSpiral:
                                WriteSCSElementPdf(document, entity as AlignmentSCS, alignment, i, normalFont, boldFont, prevEntity, nextEntity);
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
                document.Add(new Paragraph("Unable to read line data.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
                return;
            }

            try
            {
                var (x1, y1, _) = GetPointAtStation(alignment, line.StartStation);
                var (x2, y2, _) = GetPointAtStation(alignment, line.EndStation);
                var (startLabel, endLabel) = GetLinearLabels(index, prevEntity, nextEntity);

                document.Add(new Paragraph("Element: Linear").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(2));

                AddHorizontalDataRow(document, $"{startLabel}  ( ) ", FormatStation(line.StartStation), y1, x1, normalFont);
                AddHorizontalDataRow(document, $"{endLabel}  ( ) ", FormatStation(line.EndStation), y2, x2, normalFont);
                
                document.Add(new Paragraph($"Tangent Direction: {FormatBearing(line.Direction)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
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
                document.Add(new Paragraph("Unable to read arc data.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
                return;
            }

            try
            {
                var d = GetArcData(alignment, arc);
                var (startLabel, endLabel) = GetArcLabels(prevEntity, nextEntity);

                document.Add(new Paragraph("Element: Circular").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(2));

                AddHorizontalDataRow(document, $"{startLabel}  ( ) ", FormatStation(arc.StartStation), d.Y1, d.X1, normalFont);
                AddHorizontalDataRow(document, "PI  ( ) ", FormatStation(d.PIStation), d.YPI, d.XPI, normalFont);
                AddHorizontalDataRow(document, "CC  ( ) ", "               ", d.YC, d.XC, normalFont);
                AddHorizontalDataRow(document, $"{endLabel}  ( ) ", FormatStation(arc.EndStation), d.Y2, d.X2, normalFont);

                document.Add(new Paragraph($"Radius: {arc.Radius:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Design Speed(mph): {50.0:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Cant(inches): {2.0:F3}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Delta: {FormatAngle(Math.Abs(d.DeltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Degree of Curvature (Arc): {FormatAngle(d.DegreeOfCurvature)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length: {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Length(Chorded): {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent: {d.Tangent:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Chord: {d.Chord:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Middle Ordinate: {d.MiddleOrdinate:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"External: {d.External:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
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
                    // Use Atan2(dx, dy) for Azimuth (0=N, 90=E) to match FormatBearingDMS expectation
                    startTangentDirection = Math.Atan2(tempXStart - x1, tempYStart - y1);
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
                    // Use Atan2(dx, dy) for Azimuth
                    endTangentDirection = Math.Atan2(x2 - tempXEnd, y2 - tempYEnd);
                }
                
                // Chord Direction (straight line from start to end)
                // Use Atan2(dx, dy) for Azimuth (0=N, 90=E) to match FormatBearingDMS expectation
                double chordDirectionAngle = Math.Atan2(x2 - x1, y2 - y1);

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
        /// Write Spiral-Curve-Spiral (SCS) element to horizontal alignment PDF report.
        /// This handles the AlignmentSCS entity type which contains SpiralIn, Arc, and SpiralOut sub-entities.
        /// </summary>
        private void WriteSCSElementPdf(Document document, AlignmentSCS scs, CivDb.Alignment alignment, int index, PdfFont normalFont, PdfFont boldFont, AlignmentEntity prevEntity, AlignmentEntity nextEntity)
        {
            if (scs == null)
            {
                document.Add(new Paragraph("Element: Spiral-Curve-Spiral").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("Unable to read SCS data.").SetFont(normalFont).SetFontSize(10));
                document.Add(new Paragraph("\n"));
                return;
            }

            try
            {
                document.Add(new Paragraph("Element: Spiral-Curve-Spiral (Compound)").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph("\n").SetFontSize(2));

                // SpiralIn
                if (scs.SpiralIn != null)
                {
                    var spiralIn = scs.SpiralIn;
                    var (x1, y1, _) = GetPointAtStation(alignment, spiralIn.StartStation);
                    var (x2, y2, _) = GetPointAtStation(alignment, spiralIn.EndStation);

                    document.Add(new Paragraph("  Entry Spiral (Clothoid)").SetFont(boldFont).SetFontSize(9).SetMarginLeft(10));
                    AddHorizontalDataRow(document, "TS  ( )", FormatStation(spiralIn.StartStation), y1, x1, normalFont);
                    AddHorizontalDataRow(document, "SC  ( )", FormatStation(spiralIn.EndStation), y2, x2, normalFont);
                    document.Add(new Paragraph($"    Length: {spiralIn.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Radius In: {FormatRadius(spiralIn.RadiusIn)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Radius Out: {FormatRadius(spiralIn.RadiusOut)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    A: {spiralIn.A:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph("\n").SetFontSize(2));
                }

                // Arc
                if (scs.Arc != null)
                {
                    var arc = scs.Arc;
                    var (x1, y1, _) = GetPointAtStation(alignment, arc.StartStation);
                    var (x2, y2, _) = GetPointAtStation(alignment, arc.EndStation);
                    double deltaRadians = arc.Length / Math.Abs(arc.Radius);
                    double deltaDegrees = deltaRadians * (180.0 / Math.PI);

                    // Calculate curve center
                    double centerN = 0, centerE = 0;
                    try
                    {
                        var centerPoint = arc.CenterPoint;
                        centerE = centerPoint.X;
                        centerN = centerPoint.Y;
                    }
                    catch
                    {
                        // Fallback calculation if CenterPoint not available
                        double midStation = (arc.StartStation + arc.EndStation) / 2;
                        double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                        double xc = 0, yc = 0, zc = 0;
                        alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                        centerE = xc;
                        centerN = yc;
                    }

                    document.Add(new Paragraph("  Circular Arc").SetFont(boldFont).SetFontSize(9).SetMarginLeft(10));
                    AddHorizontalDataRow(document, "SC  ( )", FormatStation(arc.StartStation), y1, x1, normalFont);
                    AddHorizontalDataRow(document, "CC  ( )", "               ", centerN, centerE, normalFont);
                    AddHorizontalDataRow(document, "CS  ( )", FormatStation(arc.EndStation), y2, x2, normalFont);
                    document.Add(new Paragraph($"    Radius: {arc.Radius:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Delta: {FormatAngle(Math.Abs(deltaDegrees))} {(arc.Clockwise ? "Right" : "Left")}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Length: {arc.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph("\n").SetFontSize(2));
                }

                // SpiralOut
                if (scs.SpiralOut != null)
                {
                    var spiralOut = scs.SpiralOut;
                    var (x1, y1, _) = GetPointAtStation(alignment, spiralOut.StartStation);
                    var (x2, y2, _) = GetPointAtStation(alignment, spiralOut.EndStation);

                    document.Add(new Paragraph("  Exit Spiral (Clothoid)").SetFont(boldFont).SetFontSize(9).SetMarginLeft(10));
                    AddHorizontalDataRow(document, "CS  ( )", FormatStation(spiralOut.StartStation), y1, x1, normalFont);
                    AddHorizontalDataRow(document, "ST  ( )", FormatStation(spiralOut.EndStation), y2, x2, normalFont);
                    document.Add(new Paragraph($"    Length: {spiralOut.Length:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Radius In: {FormatRadius(spiralOut.RadiusIn)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    Radius Out: {FormatRadius(spiralOut.RadiusOut)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph($"    A: {spiralOut.A:F4}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                    document.Add(new Paragraph("\n").SetFontSize(2));
                }

                document.Add(new Paragraph("\n").SetFontSize(2));
            }
            catch (System.Exception ex)
            {
                document.Add(new Paragraph("Element: Spiral-Curve-Spiral").SetFont(boldFont).SetFontSize(10));
                document.Add(new Paragraph($"Error writing SCS data: {ex.Message}").SetFont(normalFont).SetFontSize(10));
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
                    document.Add(new Paragraph($"Horizontal Alignment Name: {alignmentName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {alignmentDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {alignmentStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Vertical Alignment Name: {profileName}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Description: {profileDescription}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($" Style: {profileStyle}").SetFont(normalFont).SetFontSize(10));
                    document.Add(new Paragraph($"Generated Report:  Date: {DateTime.Now:MM/dd/yyyy}  Time: {DateTime.Now:h:mm tt} {GetTimeZoneAbbreviation(TimeZoneInfo.Local)}")
                        .SetFont(normalFont).SetFontSize(9).SetItalic());
                    document.Add(new Paragraph("\n").SetFontSize(3));

                    // Add column headers
                    iText.Layout.Element.Table headerTable = new iText.Layout.Element.Table(UnitValue.CreatePercentArray(new float[] { 75, 25 }));
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
                    var sortedEntities = GetSortedProfileEntities(layoutProfile);
                    if (sortedEntities.Count > 0)
                    {
                        for (int i = 0; i < sortedEntities.Count; i++)
                        {
                            try
                            {
                                ProfileEntity entity = sortedEntities[i];
                                if (entity == null) continue;

                                switch (entity.EntityType)
                                {
                                    case ProfileEntityType.Tangent:
                                        if (entity is ProfileTangent tangent)
                                            WriteProfileTangentPdf(document, tangent, i, sortedEntities.Count, normalFont, boldFont);
                                        else
                                            WriteUnsupportedProfileEntityPdf(document, entity, normalFont);
                                        break;
                                    case ProfileEntityType.Circular:
                                        if (entity is ProfileCircular circular)
                                            WriteProfileParabolaPdf(document, circular, normalFont, boldFont);
                                        else
                                            WriteUnsupportedProfileEntityPdf(document, entity, normalFont);
                                        break;
                                    case ProfileEntityType.ParabolaSymmetric:
                                        if (entity is ProfileParabolaSymmetric parabola)
                                            WriteProfileParabolaSymmetricPdf(document, parabola, normalFont, boldFont);
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
                {
                    AddVerticalDataRow(document, "POB ", FormatStation(startStation), startElevation, normalFont);
                }
                else
                {
                    // After a parabola: show PVT at start (carrying over from parabola's end)
                    AddVerticalDataRow(document, "PVT ", FormatStation(startStation), startElevation, normalFont);
                }

                AddVerticalDataRow(document, "PVC ", FormatStation(endStation), endElevation, normalFont);
                document.Add(new Paragraph($"Tangent Grade: {grade * 100:F3}%").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                document.Add(new Paragraph($"Tangent Length: {FormatRounded(length, 2)}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
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
                document.Add(new Paragraph($"r = ( g2 - g1 ) / L: {r * 100:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                string kDisplay = Math.Abs(gradeDiff) > tolerance ? Math.Abs(k).ToString("F2") : "INF";
                document.Add(new Paragraph($"K = L / ( g2 - g1 ): {kDisplay}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
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

        private void WriteProfileParabolaSymmetricPdf(Document document, ProfileParabolaSymmetric curve, PdfFont normalFont, PdfFont labelFont)
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
                document.Add(new Paragraph($"r = ( g2 - g1 ) / L: {r * 100:F2}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
                string kDisplay = Math.Abs(gradeDiff) > tolerance ? Math.Abs(k).ToString("F2") : "INF";
                document.Add(new Paragraph($"K = L / ( g2 - g1 ): {kDisplay}").SetFont(normalFont).SetFontSize(9).SetMarginLeft(10));
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
            var (d, m, s) = ToDms(angleRadians * (180.0 / Math.PI));
            return $"{d}°{m:D2}'{s:F2}\"";
        }

        /// <summary>
        /// Format bearing in N/S DD°MM'SS.SS" E/W format
        /// </summary>
        private string FormatBearingDMS(double bearingRadians)
        {
            var (ns, ew, angle) = GetQuadrantAndAngle(bearingRadians);
            var (d, m, s) = ToDms(angle);
            return $"{ns} {d}°{m:D2}'{s:F2}\" {ew}";
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
        /// Calculate the PI (Point of Intersection) as the intersection of the two adjacent tangent lines.
        /// This finds the tangent lines before and after the curve (looking past any spirals) and computes their intersection.
        /// Used for GeoTable reports per InRoads convention where PI is always the tangent intersection.
        /// </summary>
        private (double PIEasting, double PINorthing) CalculateTangentIntersectionPI(
            AlignmentArc arc, 
            CivDb.Alignment alignment, 
            System.Collections.Generic.List<AlignmentEntity> sortedEntities, 
            int arcIndex)
        {
            double xPI = 0, yPI = 0;
            double tempZ = 0;
            const double stationTolerance = 0.1; // feet

            // Check if this arc has adjacent spirals (is part of a spiral-curve-spiral)
            // For spiraled curves, we need to find the LINE entities adjacent to the spirals
            AlignmentEntity entrySpiral = null;
            AlignmentEntity exitSpiral = null;
            
            if (sortedEntities != null && arcIndex >= 0 && arcIndex < sortedEntities.Count)
            {
                foreach (var entity in sortedEntities)
                {
                    if (entity == null) continue;
                    double entityEnd = SafeEndStation(entity);
                    double entityStart = SafeStartStation(entity);
                    
                    // Entry spiral ends at arc start
                    if (Math.Abs(entityEnd - arc.StartStation) < stationTolerance)
                    {
                        if (entity.EntityType == AlignmentEntityType.Spiral || entity is AlignmentSCS)
                            entrySpiral = entity;
                    }
                    // Exit spiral starts at arc end
                    if (Math.Abs(entityStart - arc.EndStation) < stationTolerance)
                    {
                        if (entity.EntityType == AlignmentEntityType.Spiral || entity is AlignmentSCS)
                            exitSpiral = entity;
                    }
                    if (entrySpiral != null && exitSpiral != null) break;
                }
            }

            // For spiraled curves: Find the tangent LINES adjacent to the spirals
            // PI = intersection of entry tangent line and exit tangent line
            if (entrySpiral != null || exitSpiral != null)
            {
                double entryTangentDir = 0, exitTangentDir = 0;
                double entryPointE = 0, entryPointN = 0;
                double exitPointE = 0, exitPointN = 0;
                bool foundEntryTangent = false, foundExitTangent = false;
                
                // Get the station where entry spiral starts (TS)
                double tsStation = entrySpiral != null ? SafeStartStation(entrySpiral) : arc.StartStation;
                // Get the station where exit spiral ends (ST)
                double stStation = exitSpiral != null ? SafeEndStation(exitSpiral) : arc.EndStation;
                
                // Find LINE entities adjacent to the spirals
                foreach (var entity in sortedEntities)
                {
                    if (entity == null) continue;
                    
                    // Find entry tangent LINE (ends at TS)
                    if (!foundEntryTangent && entrySpiral != null)
                    {
                        double entityEnd = SafeEndStation(entity);
                        if (Math.Abs(entityEnd - tsStation) < stationTolerance)
                        {
                            if (entity.EntityType == AlignmentEntityType.Line)
                            {
                                var line = entity as AlignmentLine;
                                if (line != null)
                                {
                                    alignment.PointLocation(tsStation, 0, 0, ref entryPointE, ref entryPointN, ref tempZ);
                                    entryTangentDir = line.Direction;
                                    foundEntryTangent = true;
                                }
                            }
                        }
                    }
                    
                    // Find exit tangent LINE (starts at ST)
                    if (!foundExitTangent && exitSpiral != null)
                    {
                        double entityStart = SafeStartStation(entity);
                        if (Math.Abs(entityStart - stStation) < stationTolerance)
                        {
                            if (entity.EntityType == AlignmentEntityType.Line)
                            {
                                var line = entity as AlignmentLine;
                                if (line != null)
                                {
                                    alignment.PointLocation(stStation, 0, 0, ref exitPointE, ref exitPointN, ref tempZ);
                                    exitTangentDir = line.Direction;
                                    foundExitTangent = true;
                                }
                            }
                        }
                    }
                    
                    if (foundEntryTangent && foundExitTangent) break;
                }
                
                // If no entry spiral, use arc start direction
                if (entrySpiral == null)
                {
                    alignment.PointLocation(arc.StartStation, 0, 0, ref entryPointE, ref entryPointN, ref tempZ);
                    entryTangentDir = arc.StartDirection;
                    foundEntryTangent = true;
                }
                
                // If no exit spiral, use arc end direction
                if (exitSpiral == null)
                {
                    alignment.PointLocation(arc.EndStation, 0, 0, ref exitPointE, ref exitPointN, ref tempZ);
                    exitTangentDir = arc.EndDirection;
                    foundExitTangent = true;
                }
                
                // Calculate intersection of the two tangent lines
                if (foundEntryTangent && foundExitTangent)
                {
                    double dx1 = Math.Cos(entryTangentDir);
                    double dy1 = Math.Sin(entryTangentDir);
                    double dx2 = Math.Cos(exitTangentDir);
                    double dy2 = Math.Sin(exitTangentDir);

                    double denom = dx1 * dy2 - dy1 * dx2;
                    if (Math.Abs(denom) > 1e-10)
                    {
                        double diffX = exitPointE - entryPointE;
                        double diffY = exitPointN - entryPointN;
                        double t = (diffX * dy2 - diffY * dx2) / denom;
                        xPI = entryPointE + t * dx1;
                        yPI = entryPointN + t * dy1;
                        return (xPI, yPI);
                    }
                }
            }

            // For simple curves (no spirals): Use adjacent tangent LINE directions
            if (sortedEntities != null && arcIndex >= 0 && arcIndex < sortedEntities.Count)
            {
                double entryDir = 0, exitDir = 0;
                double tsE = 0, tsN = 0, stE = 0, stN = 0;
                bool foundEntry = false, foundExit = false;

                foreach (var entity in sortedEntities)
                {
                    if (entity == null) continue;

                    // Check if this entity ends at our arc's start (entry element)
                    if (!foundEntry)
                    {
                        double entityEndStation = 0;
                        try { entityEndStation = SafeEndStation(entity); } catch { continue; }

                        if (Math.Abs(entityEndStation - arc.StartStation) < stationTolerance)
                        {
                            if (entity.EntityType == AlignmentEntityType.Line)
                            {
                                var line = entity as AlignmentLine;
                                if (line != null)
                                {
                                    alignment.PointLocation(line.EndStation, 0, 0, ref tsE, ref tsN, ref tempZ);
                                    entryDir = line.Direction;
                                    foundEntry = true;
                                }
                            }
                        }
                    }

                    // Check if this entity starts at our arc's end (exit element)
                    if (!foundExit)
                    {
                        double entityStartStation = 0;
                        try { entityStartStation = SafeStartStation(entity); } catch { continue; }

                        if (Math.Abs(entityStartStation - arc.EndStation) < stationTolerance)
                        {
                            if (entity.EntityType == AlignmentEntityType.Line)
                            {
                                var line = entity as AlignmentLine;
                                if (line != null)
                                {
                                    alignment.PointLocation(line.StartStation, 0, 0, ref stE, ref stN, ref tempZ);
                                    exitDir = line.Direction;
                                    foundExit = true;
                                }
                            }
                        }
                    }

                    if (foundEntry && foundExit) break;
                }

                if (foundEntry && foundExit)
                {
                    double dx1 = Math.Cos(entryDir);
                    double dy1 = Math.Sin(entryDir);
                    double dx2 = Math.Cos(exitDir);
                    double dy2 = Math.Sin(exitDir);
                    double denom = dx1 * dy2 - dy1 * dx2;

                    if (Math.Abs(denom) > 1e-10)
                    {
                        double diffX = stE - tsE;
                        double diffY = stN - tsN;
                        double t = (diffX * dy2 - diffY * dx2) / denom;
                        xPI = tsE + t * dx1;
                        yPI = tsN + t * dy1;
                        return (xPI, yPI);
                    }
                }
            }

            // Fallback: intersect tangents at the curve endpoints (PC/PT for simple curves).
            double pcX = 0, pcY = 0, pcZ = 0;
            double ptX = 0, ptY = 0, ptZ = 0;
            alignment.PointLocation(arc.StartStation, 0, 0, ref pcX, ref pcY, ref pcZ);
            alignment.PointLocation(arc.EndStation, 0, 0, ref ptX, ref ptY, ref ptZ);

            double startDir = arc.StartDirection;
            double endDir = arc.EndDirection;

            {
                double dx1 = Math.Cos(startDir);
                double dy1 = Math.Sin(startDir);
                double dx2 = Math.Cos(endDir);
                double dy2 = Math.Sin(endDir);

                double denom = dx1 * dy2 - dy1 * dx2;
                if (Math.Abs(denom) > 1e-10)
                {
                    double diffX = ptX - pcX;
                    double diffY = ptY - pcY;
                    double t = (diffX * dy2 - diffY * dx2) / denom;
                    xPI = pcX + t * dx1;
                    yPI = pcY + t * dy1;
                    return (xPI, yPI);
                }
            }

            // Fallback: Try to get PI from sub-entity
            try
            {
                if (arc.SubEntityCount > 0)
                {
                    var subEntity = arc[0];
                    if (subEntity is AlignmentSubEntityArc subEntityArc)
                    {
                        var piPoint = subEntityArc.PIPoint;
                        xPI = piPoint.X;
                        yPI = piPoint.Y;
                        return (xPI, yPI);
                    }
                }
            }
            catch { }

            // Final fallback: Calculate from arc geometry
            double x1 = 0, y1 = 0, z1 = 0;
            alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
            double radius = Math.Abs(arc.Radius);
            double delta = arc.Length / radius;
            double tc = CalculateTangentDistance(radius, delta);
            double backTangentDir = arc.StartDirection + Math.PI;
            xPI = x1 + tc * Math.Cos(backTangentDir);
            yPI = y1 + tc * Math.Sin(backTangentDir);

            return (xPI, yPI);
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
                    
                    // Date/time omitted for GeoTables per markup (kept for alignment reports)

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
                    sortedEntities.Sort((a, b) => SafeStartStation(a).CompareTo(SafeStartStation(b)));

                    // Expand AlignmentSCS into component entities for GeoTable output
                    var reportEntities = new System.Collections.Generic.List<AlignmentEntity>();
                    for (int i = 0; i < sortedEntities.Count; i++)
                    {
                        var e = sortedEntities[i];
                        if (e == null) continue;
                        foreach (var expanded in ExpandEntityForReports(e))
                        {
                            if (expanded != null) reportEntities.Add(expanded);
                        }
                    }

                    // Process alignment entities in sorted order
                    for (int i = 0; i < reportEntities.Count; i++)
                    {
                        AlignmentEntity entity = reportEntities[i];
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
                                    row = WriteGeoTableCurveExcel(ws, entity as AlignmentArc, alignment, i, row, curveNumber, reportEntities);
                                    break;
                                case AlignmentEntityType.Spiral:
                                    row = WriteGeoTableSpiralExcel(ws, entity, alignment, i, row);
                                    break;
                                case AlignmentEntityType.SpiralCurveSpiral:
                                    row = WriteGeoTableSCSExcel(ws, entity as AlignmentSCS, alignment, i, row, ref curveNumber, reportEntities);
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

        private int WriteGeoTableCurveExcel(ExcelWorksheet ws, AlignmentArc arc, CivDb.Alignment alignment, int index, int row, int curveNumber, System.Collections.Generic.List<AlignmentEntity> sortedEntities)
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

                // Get PI as intersection of adjacent tangent lines (per InRoads convention for GeoTable)
                // This ensures PI is consistent whether curve has spirals or not
                var (xPI, yPI) = CalculateTangentIntersectionPI(arc, alignment, sortedEntities, index);

                // Get Center coordinates from sub-entity or calculate
                double centerE = 0, centerN = 0;
                bool gotCenterFromSubEntity = false;

                try
                {
                    if (arc.SubEntityCount > 0)
                    {
                        var subEntity = arc[0];
                        if (subEntity is AlignmentSubEntityArc subEntityArc)
                        {
                            var centerPoint = subEntityArc.CenterPoint;
                            centerE = centerPoint.X;
                            centerN = centerPoint.Y;
                            gotCenterFromSubEntity = true;
                        }
                    }
                }
                catch { gotCenterFromSubEntity = false; }

                if (!gotCenterFromSubEntity)
                {
                    CalculateCurveCenter(arc, alignment, out centerN, out centerE);
                }

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

                // Row 2: PI (no station/bearing per markup)
                ws.Cells[row, 3].Value = "PI";
                ws.Cells[row, 4].Value = ""; // Station deleted per markup
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
        /// Write Spiral-Curve-Spiral entity to GeoTable Excel (GLTT Standard Format)
        /// </summary>
        private int WriteGeoTableSCSExcel(ExcelWorksheet ws, AlignmentSCS scs, CivDb.Alignment alignment, int index, int row, ref int curveNumber, System.Collections.Generic.List<AlignmentEntity> sortedEntities)
        {
            if (scs == null) return row + 6;

            try
            {
                // Entry Spiral
                if (scs.SpiralIn != null)
                {
                    var spiralIn = scs.SpiralIn;
                    int startRow = row;

                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(spiralIn.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(spiralIn.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double spiralAngle = spiralIn.Length / (2.0 * Math.Abs(spiralIn.RadiusOut));
                    double theta = spiralAngle;

                    ws.Cells[row, 1].Value = "SPIRAL";
                    ws.Cells[row, 3].Value = "TS";
                    ws.Cells[row, 4].Value = FormatStation(spiralIn.StartStation);
                    ws.Cells[row, 5].Value = FormatBearingDMS(spiralIn.StartDirection);
                    ws.Cells[row, 6].Value = y1;
                    ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 7].Value = x1;
                    ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 8].Value = $"θs = {FormatAngleDMS(theta)}";
                    ws.Cells[row, 9].Value = $"Ls= {FormatDistanceFeet(spiralIn.Length)}";
                    ws.Cells[row, 10].Value = $"A= {spiralIn.A:F2}";
                    ws.Cells[row, 11].Value = "";
                    row++;
                    // SC row is written as part of the CURVE section to avoid duplication
                }

                // Arc
                if (scs.Arc != null)
                {
                    var arc = scs.Arc;
                    curveNumber++;
                    int startRow = row;

                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double radius = Math.Abs(arc.Radius);
                    double delta = arc.Length / radius;
                    double tc = CalculateTangentDistance(radius, delta);
                    double ec = CalculateExternalDistance(radius, delta);
                    string directionStr = arc.Clockwise ? "R" : "L";

                    double centerE = 0, centerN = 0;
                    try
                    {
                        var centerPoint = arc.CenterPoint;
                        centerE = centerPoint.X;
                        centerN = centerPoint.Y;
                    }
                    catch
                    {
                        // Fallback: calculate center from midpoint offset
                        double midStation = (arc.StartStation + arc.EndStation) / 2;
                        double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                        double xc = 0, yc = 0, zc = 0;
                        alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                        centerE = xc;
                        centerN = yc;
                    }

                    ws.Cells[row, 1].Value = "CURVE";
                    ws.Cells[row, 2].Value = $"{curveNumber}-{directionStr}";
                    ws.Cells[row, 3].Value = "SC";
                    ws.Cells[row, 4].Value = FormatStation(arc.StartStation);
                    ws.Cells[row, 5].Value = "";
                    ws.Cells[row, 6].Value = y1;
                    ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 7].Value = x1;
                    ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 8].Value = $"Δc = {FormatAngleDMS(delta)}";
                    ws.Cells[row, 9].Value = $"R= {radius:F2}'";
                    ws.Cells[row, 10].Value = $"Lc= {FormatDistanceFeet(arc.Length)}";
                    ws.Cells[row, 11].Value = "";
                    row++;

                    // Get PI (Point of Intersection) - for SCS, this is the intersection of tangents at TS and ST
                    // (the "CurvesetPoint" in InRoads terminology)
                    double piN = 0, piE = 0;
                    bool gotPI = false;
                    
                    // For SCS: Use the tangent directions at TS (entry spiral start) and ST (exit spiral end)
                    // These are the directions of the tangent lines that define the curveset PI
                    if (scs.SpiralIn != null && scs.SpiralOut != null)
                    {
                        // Get TS point and direction (entry tangent)
                        double tsE = 0, tsN = 0, tsZ = 0;
                        alignment.PointLocation(scs.SpiralIn.StartStation, 0, 0, ref tsE, ref tsN, ref tsZ);
                        double dir1 = scs.SpiralIn.StartDirection;

                        // Get ST point and direction (exit tangent)
                        double stE = 0, stN = 0, stZ = 0;
                        alignment.PointLocation(scs.SpiralOut.EndStation, 0, 0, ref stE, ref stN, ref stZ);
                        double dir2 = scs.SpiralOut.EndDirection;

                        // DEBUG: Show message box with direction values
                        double dir1Deg = dir1 * 180.0 / Math.PI;
                        double dir2Deg = dir2 * 180.0 / Math.PI;
                        string debugMsg = $"=== SCS PI DEBUG (CURVE {curveNumber}) ===\n\n" +
                            $"TS Point: N {tsN:F4}, E {tsE:F4}\n" +
                            $"ST Point: N {stN:F4}, E {stE:F4}\n\n" +
                            $"SpiralIn.StartDirection: {dir1:F6} rad ({dir1Deg:F4}°)\n" +
                            $"SpiralOut.EndDirection: {dir2:F6} rad ({dir2Deg:F4}°)\n\n" +
                            $"Bearing at TS: {FormatBearingDMS(dir1)}\n" +
                            $"Bearing at ST: {FormatBearingDMS(dir2)}\n\n" +
                            $"Angle difference: {Math.Abs(dir2Deg - dir1Deg):F4}°";
                        System.Windows.Forms.MessageBox.Show(debugMsg, $"PI Debug - Curve {curveNumber}");

                        // Calculate intersection of the two tangent lines
                        double dx1 = Math.Cos(dir1);
                        double dy1 = Math.Sin(dir1);
                        double dx2 = Math.Cos(dir2);
                        double dy2 = Math.Sin(dir2);

                        double denom = dx1 * dy2 - dy1 * dx2;
                        
                        if (Math.Abs(denom) > 1e-10)
                        {
                            double diffX = stE - tsE;
                            double diffY = stN - tsN;
                            double t = (diffX * dy2 - diffY * dx2) / denom;
                            piE = tsE + t * dx1;
                            piN = tsN + t * dy1;
                            gotPI = true;
                            
                            // DEBUG: Show calculation result
                            string calcMsg = $"Calculation:\n" +
                                $"dx1={dx1:F6}, dy1={dy1:F6}\n" +
                                $"dx2={dx2:F6}, dy2={dy2:F6}\n" +
                                $"denom: {denom:F10}\n" +
                                $"diffX={diffX:F4}, diffY={diffY:F4}\n" +
                                $"t parameter: {t:F4}\n\n" +
                                $"Calculated PI: N {piN:F4}, E {piE:F4}";
                            System.Windows.Forms.MessageBox.Show(calcMsg, $"PI Calculation - Curve {curveNumber}");
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show($"WARNING: denom too small ({denom:F10}), lines nearly parallel!", "PI Warning");
                        }
                    }
                    
                    // Fallback: Use arc's PIPoint if tangent calculation failed
                    if (!gotPI)
                    {
                        try
                        {
                            var piPoint = arc.PIPoint;
                            piE = piPoint.X;
                            piN = piPoint.Y;
                        }
                        catch { }
                    }

                    ws.Cells[row, 3].Value = "PI";
                    ws.Cells[row, 4].Value = "";
                    ws.Cells[row, 5].Value = "";
                    ws.Cells[row, 6].Value = piN;
                    ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 7].Value = piE;
                    ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 8].Value = "";
                    ws.Cells[row, 9].Value = "";
                    ws.Cells[row, 10].Value = "";
                    ws.Cells[row, 11].Value = "";
                    row++;

                    ws.Cells[row, 3].Value = "CS";
                    ws.Cells[row, 4].Value = FormatStation(arc.EndStation);
                    ws.Cells[row, 5].Value = "";
                    ws.Cells[row, 6].Value = y2;
                    ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 7].Value = x2;
                    ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 8].Value = $"Tc= {FormatDistanceFeet(tc)}";
                    ws.Cells[row, 9].Value = $"Ec= {FormatDistanceFeet(ec)}";
                    ws.Cells[row, 10].Value = $"CC:N {centerN:F4}";
                    ws.Cells[row, 11].Value = $"E {centerE:F4}";
                    row++;

                    ws.Cells[startRow, 1, row - 1, 1].Merge = true;
                    ws.Cells[startRow, 2, row - 1, 2].Merge = true;
                }

                // Exit Spiral
                if (scs.SpiralOut != null)
                {
                    var spiralOut = scs.SpiralOut;
                    int startRow = row;

                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(spiralOut.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(spiralOut.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double spiralAngle = spiralOut.Length / (2.0 * Math.Abs(spiralOut.RadiusIn));
                    double theta = spiralAngle;

                    // CS row is written as part of the CURVE section to avoid duplication
                    ws.Cells[row, 1].Value = "SPIRAL";
                    ws.Cells[row, 3].Value = "ST";
                    ws.Cells[row, 4].Value = FormatStation(spiralOut.EndStation);
                    ws.Cells[row, 5].Value = FormatBearingDMS(spiralOut.EndDirection);
                    ws.Cells[row, 6].Value = y2;
                    ws.Cells[row, 6].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 7].Value = x2;
                    ws.Cells[row, 7].Style.Numberformat.Format = "0.0000";
                    ws.Cells[row, 8].Value = $"θs = {FormatAngleDMS(theta)}";
                    ws.Cells[row, 9].Value = $"Ls= {FormatDistanceFeet(spiralOut.Length)}";
                    ws.Cells[row, 10].Value = $"A= {spiralOut.A:F2}";
                    ws.Cells[row, 11].Value = "";
                    row++;
                }

                return row;
            }
            catch
            {
                return row + 6;
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
                            .SetMarginBottom(8);
                        document.Add(title);
                        
                        // Date/time omitted for GeoTables per markup (kept for alignment reports)

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

                        // Expand AlignmentSCS into component entities so CURVE/SPIRAL rows are produced
                        var reportEntities = new System.Collections.Generic.List<AlignmentEntity>();
                        for (int i = 0; i < sortedEntities.Count; i++)
                        {
                            var e = sortedEntities[i];
                            if (e == null) continue;
                            foreach (var expanded in ExpandEntityForReports(e))
                            {
                                if (expanded != null) reportEntities.Add(expanded);
                            }
                        }

                        // Process alignment entities in sorted order
                        int curveNumber = 0;
                        for (int i = 0; i < reportEntities.Count; i++)
                        {
                            AlignmentEntity entity = reportEntities[i];
                            if (entity == null) continue;

                            AlignmentEntity prevEntity = i > 0 ? reportEntities[i - 1] : null;
                            AlignmentEntity nextEntity = i < reportEntities.Count - 1 ? reportEntities[i + 1] : null;

                            try
                            {
                                switch (entity.EntityType)
                                {
                                    case AlignmentEntityType.Line:
                                        AddGeoTableTangentPdf(table, entity as AlignmentLine, alignment, i, font);
                                        break;
                                    case AlignmentEntityType.Arc:
                                        curveNumber++;
                                        AddGeoTableCurvePdf(table, entity as AlignmentArc, alignment, i, curveNumber, font, prevEntity, nextEntity, reportEntities);
                                        break;
                                    case AlignmentEntityType.Spiral:
                                        AddGeoTableSpiralPdf(table, entity, alignment, i, font, prevEntity, nextEntity);
                                        break;
                                    case AlignmentEntityType.SpiralCurveSpiral:
                                        AddGeoTableSCSPdf(table, entity as AlignmentSCS, alignment, i, ref curveNumber, font, reportEntities);
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

        private double SafeEndStation(AlignmentEntity entity)
        {
            try
            {
                dynamic d = entity;
                double s = d.EndStation;
                if (double.IsNaN(s) || double.IsInfinity(s)) return double.MinValue;
                return s;
            }
            catch
            {
                return double.MinValue;
            }
        }

        private sealed class SpiralCurveSpiral
        {
            public AlignmentEntity RawEntity { get; }
            public AlignmentEntity SpiralIn { get; }
            public AlignmentArc Arc { get; }
            public AlignmentEntity SpiralOut { get; }

            public double StartStation { get; }
            public double EndStation { get; }

            public double TotalLength
            {
                get
                {
                    double total = 0;
                    total += SafeLength(SpiralIn);
                    if (Arc != null) total += Arc.Length;
                    total += SafeLength(SpiralOut);
                    return total;
                }
            }

            public SpiralCurveSpiral(
                AlignmentEntity rawEntity,
                AlignmentEntity spiralIn,
                AlignmentArc arc,
                AlignmentEntity spiralOut,
                double startStation,
                double endStation)
            {
                RawEntity = rawEntity;
                SpiralIn = spiralIn;
                Arc = arc;
                SpiralOut = spiralOut;
                StartStation = startStation;
                EndStation = endStation;
            }
        }

        private static double SafeLength(AlignmentEntity entity)
        {
            if (entity == null) return 0;
            try
            {
                dynamic d = entity;
                double len = d.Length;
                if (double.IsNaN(len) || double.IsInfinity(len) || len < 0) return 0;
                return len;
            }
            catch
            {
                return 0;
            }
        }

        private bool TryGetSpiralCurveSpiral(AlignmentEntity entity, out SpiralCurveSpiral scs)
        {
            scs = null;
            if (entity == null) return false;

            // Civil 3D's API class is AlignmentSCS; use type check + dynamic property access for robustness.
            if (!(entity is AlignmentSCS) && !string.Equals(entity.GetType().Name, "AlignmentSCS", StringComparison.OrdinalIgnoreCase))
                return false;

            try
            {
                dynamic d = entity;
                object arcObj = null;
                object spiralInObj = null;
                object spiralOutObj = null;
                try { arcObj = d.Arc; } catch { }
                try { spiralInObj = d.SpiralIn; } catch { }
                try { spiralOutObj = d.SpiralOut; } catch { }

                var arc = arcObj as AlignmentArc;
                var spiralIn = spiralInObj as AlignmentEntity;
                var spiralOut = spiralOutObj as AlignmentEntity;

                double startStation = 0;
                double endStation = 0;
                try { startStation = d.StartStation; } catch { }
                try { endStation = d.EndStation; } catch { }

                // If start/end are not exposed at the parent, derive from components.
                if (startStation == 0 && spiralIn != null) startStation = SafeStartStation(spiralIn);
                if (startStation == 0 && arc != null) startStation = arc.StartStation;
                if (endStation == 0)
                {
                    try
                    {
                        if (spiralOut != null) endStation = (spiralOut as dynamic).EndStation;
                        else if (arc != null) endStation = arc.EndStation;
                    }
                    catch { }
                }

                scs = new SpiralCurveSpiral(entity, spiralIn, arc, spiralOut, startStation, endStation);
                return scs.Arc != null || scs.SpiralIn != null || scs.SpiralOut != null;
            }
            catch
            {
                scs = null;
                return false;
            }
        }

        private System.Collections.Generic.IEnumerable<AlignmentEntity> ExpandEntityForReports(AlignmentEntity entity)
        {
            if (entity == null) yield break;

            // Do NOT expand SCS entities - they need to be processed as compound entities
            // to correctly calculate PI from the entry and exit tangent directions
            if (TryGetSpiralCurveSpiral(entity, out var scs))
            {
                // Return the SCS as-is, don't break it apart
                yield return entity;
                yield break;
            }

            yield return entity;
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

        private void AddGeoTableCurvePdf(iText.Layout.Element.Table table, AlignmentArc arc, CivDb.Alignment alignment, int index, int curveNumber, PdfFont font, AlignmentEntity prevEntity, AlignmentEntity nextEntity, System.Collections.Generic.List<AlignmentEntity> sortedEntities)
        {
            if (arc == null) return;

            double radius = Math.Abs(arc.Radius);
            double delta = arc.Length / radius;
            double tc = CalculateTangentDistance(radius, delta);
            double ec = CalculateExternalDistance(radius, delta);

            double x1 = 0, y1 = 0, z1 = 0, x2 = 0, y2 = 0, z2 = 0;
            alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
            alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

            // Get PI as intersection of adjacent tangent lines (per InRoads convention for GeoTable)
            // This ensures PI is consistent whether curve has spirals or not
            var (xPI, yPI) = CalculateTangentIntersectionPI(arc, alignment, sortedEntities, index);

            // Get Center coordinates from sub-entity or calculate
            double centerE = 0, centerN = 0;
            bool gotCenterFromSubEntity = false;

            try
            {
                if (arc.SubEntityCount > 0)
                {
                    var subEntity = arc[0];
                    if (subEntity is AlignmentSubEntityArc subEntityArc)
                    {
                        var centerPoint = subEntityArc.CenterPoint;
                        centerE = centerPoint.X;
                        centerN = centerPoint.Y;
                        gotCenterFromSubEntity = true;
                    }
                }
            }
            catch { gotCenterFromSubEntity = false; }

            if (!gotCenterFromSubEntity)
            {
                CalculateCurveCenter(arc, alignment, out centerN, out centerE);
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
            table.AddCell(new iText.Layout.Element.Cell(3, 1).Add(new Paragraph("CURVE").SetFont(font).SetFontSize(6.5f))
                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            
            // Curve Number cell with Oval Renderer
            var curveNoCell = new iText.Layout.Element.Cell(3, 1).Add(new Paragraph(curveLabel).SetFont(font).SetFontSize(6.5f))
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

            // Row 2: PI (no station/bearing per markup)
            table.AddCell(CreateLabelCell("PI", font));
            table.AddCell(CreateDataCell("", font)); // Station deleted per markup
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
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("SPIRAL").SetFont(font).SetFontSize(6.5f))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("").SetFont(font).SetFontSize(6.5f))
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
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("SPIRAL").SetFont(font).SetFontSize(6.5f))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(new iText.Layout.Element.Cell().Add(new Paragraph("").SetFont(font).SetFontSize(6.5f))
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
                table.AddCell(new iText.Layout.Element.Cell(1, 2).Add(new Paragraph("SPIRAL ERROR").SetFont(font).SetFontSize(6.5f))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                table.AddCell(new iText.Layout.Element.Cell(1, 5).Add(new Paragraph($"Error: {ex.Message}").SetFont(font).SetFontSize(6.5f))
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
            }
        }

        /// <summary>
        /// Add Spiral-Curve-Spiral entity to GeoTable PDF (GLTT Standard Format)
        /// </summary>
        private void AddGeoTableSCSPdf(iText.Layout.Element.Table table, AlignmentSCS scs, CivDb.Alignment alignment, int index, ref int curveNumber, PdfFont font, System.Collections.Generic.List<AlignmentEntity> sortedEntities)
        {
            if (scs == null) return;

            try
            {
                // Entry Spiral
                if (scs.SpiralIn != null)
                {
                    var spiralIn = scs.SpiralIn;
                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(spiralIn.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(spiralIn.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double spiralAngle = spiralIn.Length / (2.0 * Math.Abs(spiralIn.RadiusOut));

                    // Row: TS
                    table.AddCell(CreateDataCell("SPIRAL", font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell("TS", font));
                    table.AddCell(CreateDataCell(FormatStation(spiralIn.StartStation), font));
                    table.AddCell(CreateDataCell(FormatBearingDMS(spiralIn.StartDirection), font));
                    table.AddCell(CreateDataCell($"{y1:F4}", font));
                    table.AddCell(CreateDataCell($"{x1:F4}", font));
                    table.AddCell(CreateDataCellNoBorder($"θs = {FormatAngleDMS(spiralAngle)}", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"Ls= {spiralIn.Length:F2}'", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"A= {spiralIn.A:F2}", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder("", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    // SC row is written as part of the CURVE section to avoid duplication
                }

                // Arc
                if (scs.Arc != null)
                {
                    var arc = scs.Arc;
                    curveNumber++;

                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(arc.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(arc.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double radius = Math.Abs(arc.Radius);
                    double delta = arc.Length / radius;
                    double tc = CalculateTangentDistance(radius, delta);
                    double ec = CalculateExternalDistance(radius, delta);
                    string curveDir = arc.Clockwise ? "R" : "L";
                    string curveLabel = $"{curveNumber}-{curveDir}";

                    double centerE = 0, centerN = 0;
                    try
                    {
                        var centerPoint = arc.CenterPoint;
                        centerE = centerPoint.X;
                        centerN = centerPoint.Y;
                    }
                    catch
                    {
                        // Fallback: calculate center from midpoint offset
                        double midStation = (arc.StartStation + arc.EndStation) / 2;
                        double offset = arc.Clockwise ? -arc.Radius : arc.Radius;
                        double xc = 0, yc = 0, zc = 0;
                        alignment.PointLocation(midStation, offset, 0, ref xc, ref yc, ref zc);
                        centerE = xc;
                        centerN = yc;
                    }

                    // Row: SC (start of curve from spiral)
                    table.AddCell(new iText.Layout.Element.Cell(3, 1).Add(new Paragraph("CURVE").SetFont(font).SetFontSize(6.5f))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    var curveNoCell = new iText.Layout.Element.Cell(3, 1).Add(new Paragraph(curveLabel).SetFont(font).SetFontSize(6.5f))
                        .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                        .SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE)
                        .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f));
                    curveNoCell.SetNextRenderer(new OvalCellRenderer(curveNoCell));
                    table.AddCell(curveNoCell);

                    table.AddCell(CreateDataCell("SC", font));
                    table.AddCell(CreateDataCell(FormatStation(arc.StartStation), font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell($"{y1:F4}", font));
                    table.AddCell(CreateDataCell($"{x1:F4}", font));
                    table.AddCell(CreateDataCellNoBorder($"Δc = {FormatAngleDMS(delta)}", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"R= {radius:F2}'", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"Lc= {arc.Length:F2}'", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder("", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));

                    // Get PI (Point of Intersection) - for SCS, this is the intersection of tangent directions
                    // at the TS (spiral start) and ST (spiral end) points
                    double piN = 0, piE = 0;
                    bool gotPI = false;
                    
                    // Calculate PI using intersection of tangent directions at TS and ST
                    // Get TS point and direction (entry tangent)
                    double tsE = 0, tsN = 0, tsZ = 0;
                    alignment.PointLocation(scs.SpiralIn.StartStation, 0, 0, ref tsE, ref tsN, ref tsZ);
                    double dir1 = scs.SpiralIn.StartDirection;
                    
                    // Get ST point and direction (exit tangent)
                    double stE = 0, stN = 0, stZ2 = 0;
                    alignment.PointLocation(scs.SpiralOut.EndStation, 0, 0, ref stE, ref stN, ref stZ2);
                    double dir2 = scs.SpiralOut.EndDirection;
                    
                    // Calculate intersection of the two tangent lines
                    double dx1 = Math.Cos(dir1);
                    double dy1 = Math.Sin(dir1);
                    double dx2 = Math.Cos(dir2);
                    double dy2 = Math.Sin(dir2);
                    
                    double denom = dx1 * dy2 - dy1 * dx2;
                    if (Math.Abs(denom) > 1e-10)
                    {
                        double diffX = stE - tsE;
                        double diffY = stN - tsN;
                        double t = (diffX * dy2 - diffY * dx2) / denom;
                        piE = tsE + t * dx1;
                        piN = tsN + t * dy1;
                        gotPI = true;
                    }
                    
                    // Fallback: Use arc's PIPoint if calculation failed
                    if (!gotPI)
                    {
                        try
                        {
                            var piPoint = arc.PIPoint;
                            piE = piPoint.X;
                            piN = piPoint.Y;
                        }
                        catch { }
                    }

                    table.AddCell(CreateDataCell("PI", font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell($"{piN:F4}", font));
                    table.AddCell(CreateDataCell($"{piE:F4}", font));
                    table.AddCell(CreateDataCellNoBorder("", font)
                        .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder("", font));
                    table.AddCell(CreateDataCellNoBorder("", font));
                    table.AddCell(CreateDataCellNoBorder("", font)
                        .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));

                    // Row: CS (end of curve to spiral)
                    table.AddCell(CreateDataCell("CS", font));
                    table.AddCell(CreateDataCell(FormatStation(arc.EndStation), font));
                    table.AddCell(CreateDataCell("", font));
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

                // Exit Spiral
                if (scs.SpiralOut != null)
                {
                    var spiralOut = scs.SpiralOut;
                    double x1 = 0, y1 = 0, z1 = 0;
                    double x2 = 0, y2 = 0, z2 = 0;
                    alignment.PointLocation(spiralOut.StartStation, 0, 0, ref x1, ref y1, ref z1);
                    alignment.PointLocation(spiralOut.EndStation, 0, 0, ref x2, ref y2, ref z2);

                    double spiralAngle = spiralOut.Length / (2.0 * Math.Abs(spiralOut.RadiusIn));

                    // CS row is written as part of the CURVE section to avoid duplication
                    // Row: ST only (with spiral parameters)
                    table.AddCell(CreateDataCell("SPIRAL", font));
                    table.AddCell(CreateDataCell("", font));
                    table.AddCell(CreateDataCell("ST", font));
                    table.AddCell(CreateDataCell(FormatStation(spiralOut.EndStation), font));
                    table.AddCell(CreateDataCell(FormatBearingDMS(spiralOut.EndDirection), font));
                    table.AddCell(CreateDataCell($"{y2:F4}", font));
                    table.AddCell(CreateDataCell($"{x2:F4}", font));
                    table.AddCell(CreateDataCellNoBorder($"θs = {FormatAngleDMS(spiralAngle)}", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderLeft(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"Ls= {spiralOut.Length:F2}'", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder($"A= {spiralOut.A:F2}", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                    table.AddCell(CreateDataCellNoBorder("", font)
                        .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderBottom(new SolidBorder(ColorConstants.BLACK, 0.5f))
                        .SetBorderRight(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                }
            }
            catch (System.Exception ex)
            {
                // Add error row
                table.AddCell(new iText.Layout.Element.Cell(1, 2).Add(new Paragraph("SCS ERROR").SetFont(font).SetFontSize(6.5f))
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f)));
                table.AddCell(new iText.Layout.Element.Cell(1, 9).Add(new Paragraph($"Error: {ex.Message}").SetFont(font).SetFontSize(6.5f))
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
                        string selectedPath = sfd.FileName;
                        string directory = System.IO.Path.GetDirectoryName(selectedPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(selectedPath);

                        // Strip known suffixes to avoid duplication
                        if (fileName.EndsWith("_Alignment_Report"))
                            fileName = fileName.Substring(0, fileName.Length - "_Alignment_Report".Length);
                        else if (fileName.EndsWith("_GeoTable"))
                            fileName = fileName.Substring(0, fileName.Length - "_GeoTable".Length);

                        OutputPath = System.IO.Path.Combine(directory, fileName);
                        outputPathTextBox.Text = OutputPath;
                    }
                }
            };

            int groupTop = 95;
            int groupHeight = 220;
            int groupWidth = 320;
            var grpAlignment = new GroupBox { Left = 15, Top = groupTop, Width = groupWidth, Height = groupHeight, Text = "Alignment Reports" };
            chkAlignmentPdf = new CheckBox { Left = 15, Top = 25, Width = 160, Text = "Alignment PDF" };

            chkAlignmentXml = new CheckBox { Left = 15, Top = 75, Width = 160, Text = "Alignment XML" };
            grpAlignment.Controls.AddRange(new Control[] { chkAlignmentPdf, chkAlignmentXml });

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

            PictureBox IconXmlAlign = new PictureBox { Left = 20, Top = 75, Width = 12, Height = 12, BackColor = iconBack, Parent = grpAlignment }; grpAlignment.Controls.Add(IconXmlAlign);
            chkAlignmentPdf.Left = 40;

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
