/*
 * Civil 3D GeoTable Reports - .NET Add-in Starter Template
 *
 * This is a basic template showing how a .NET add-in would work
 * Compile this with Visual Studio and Civil 3D API references
 */

using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

[assembly: CommandClass(typeof(GeoTableReports.ReportCommands))]

namespace GeoTableReports
{
    public class ReportCommands : IExtensionApplication
    {
        // Called when Civil 3D starts up
        public void Initialize()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
                                return;
                            }

                            // Show progress
                            ed.WriteMessage($"\nProcessing alignment: {alignment.Name}...");

                            if (reportType == "Vertical")
                            {
                                GenerateVerticalReport(alignment, outputPath);
                            }
                            else
                            {
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
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n✗ Error: {ex.Message}");
                MessageBox.Show(
                    $"Error generating report:\n\n{ex.Message}",
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
            Document doc = Application.DocumentManager.MdiActiveDocument;
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
        private void GenerateVerticalReport(Alignment alignment, string outputPath)
        {
            // Extract vertical alignment data
            var reportData = new System.Collections.Generic.Dictionary<string, object>();

            reportData["ProjectName"] = ""; // Get from drawing properties
            reportData["AlignmentName"] = alignment.Name;
            reportData["Description"] = alignment.Description;
            reportData["Length"] = alignment.Length;
            reportData["StartStation"] = alignment.StartingStation;
            reportData["EndStation"] = alignment.EndingStation;

            var elements = new System.Collections.Generic.List<object>();

            // Get profiles for this alignment
            foreach (ObjectId profileId in alignment.GetProfileIds())
            {
                using (Profile profile = profileId.GetObject(OpenMode.ForRead) as Profile)
                {
                    if (profile == null || profile.ProfileType != ProfileType.FGProfile)
                        continue;

                    // Extract profile entities
                    foreach (ProfileEntity entity in profile.Entities)
                    {
                        var element = new System.Collections.Generic.Dictionary<string, object>();

                        switch (entity.EntityType)
                        {
                            case ProfileEntityType.Tangent:
                                element["Type"] = "Linear";
                                var tangent = entity as ProfileTangent;
                                element["StartStation"] = tangent.StartStation;
                                element["EndStation"] = tangent.EndStation;
                                element["StartElevation"] = tangent.StartElevation;
                                element["EndElevation"] = tangent.EndElevation;
                                element["Grade"] = tangent.Grade;
                                element["Length"] = tangent.Length;
                                break;

                            case ProfileEntityType.Circular:
                                element["Type"] = "Parabola";
                                var curve = entity as ProfileCircular;
                                element["StartStation"] = curve.StartStation;
                                element["EndStation"] = curve.EndStation;
                                element["PVIStation"] = curve.PVIStation;
                                element["PVIElevation"] = curve.PVIElevation;
                                element["Length"] = curve.Length;
                                element["GradeIn"] = curve.GradeIn;
                                element["GradeOut"] = curve.GradeOut;
                                element["K"] = Math.Abs(curve.Length / (curve.GradeOut - curve.GradeIn));
                                break;
                        }

                        elements.Add(element);
                    }
                }
            }

            reportData["Elements"] = elements;

            // Generate PDF using report data
            // TODO: Implement PDF generation (use iTextSharp or similar)
            // For now, generate text file
            GenerateTextReport(reportData, outputPath.Replace(".pdf", ".txt"), "Vertical");
        }

        /// <summary>
        /// Generate horizontal alignment report
        /// </summary>
        private void GenerateHorizontalReport(Alignment alignment, string outputPath)
        {
            // Extract horizontal alignment data
            var reportData = new System.Collections.Generic.Dictionary<string, object>();

            reportData["ProjectName"] = "";
            reportData["AlignmentName"] = alignment.Name;
            reportData["Description"] = alignment.Description;

            var elements = new System.Collections.Generic.List<object>();

            // Extract alignment entities
            foreach (AlignmentEntity entity in alignment.Entities)
            {
                var element = new System.Collections.Generic.Dictionary<string, object>();
                element["Type"] = entity.EntityType.ToString();
                element["StartStation"] = entity.StartStation;
                element["EndStation"] = entity.EndStation;
                element["Length"] = entity.Length;

                // Get coordinates
                double x, y, z;
                alignment.PointLocation(entity.StartStation, 0, out x, out y, out z);
                element["StartX"] = x;
                element["StartY"] = y;

                alignment.PointLocation(entity.EndStation, 0, out x, out y, out z);
                element["EndX"] = x;
                element["EndY"] = y;

                // Type-specific properties
                switch (entity.EntityType)
                {
                    case AlignmentEntityType.Line:
                        var line = entity as AlignmentLine;
                        element["Direction"] = line.Direction;
                        break;

                    case AlignmentEntityType.Arc:
                        var arc = entity as AlignmentArc;
                        element["Radius"] = arc.Radius;
                        element["Clockwise"] = arc.Clockwise;
                        element["Delta"] = arc.Delta;
                        break;

                    case AlignmentEntityType.Spiral:
                        var spiral = entity as AlignmentSpiral;
                        element["RadiusIn"] = spiral.RadiusIn;
                        element["RadiusOut"] = spiral.RadiusOut;
                        element["SpiralType"] = spiral.SpiralDefinition.ToString();
                        break;
                }

                elements.Add(element);
            }

            reportData["Elements"] = elements;

            // Generate PDF using report data
            GenerateTextReport(reportData, outputPath.Replace(".pdf", ".txt"), "Horizontal");
        }

        /// <summary>
        /// Generate text report (fallback if PDF not implemented)
        /// </summary>
        private void GenerateTextReport(System.Collections.Generic.Dictionary<string, object> data, string path, string type)
        {
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(path))
            {
                writer.WriteLine($"Project Name: {data["ProjectName"]}");
                writer.WriteLine($"Alignment Name: {data["AlignmentName"]}");
                writer.WriteLine($"Description: {data["Description"]}");
                writer.WriteLine();

                if (type == "Vertical")
                {
                    writer.WriteLine($"{"",40} {"STATION",15} {"ELEVATION",15}");
                }
                else
                {
                    writer.WriteLine($"{"",40} {"STATION",15} {"NORTHING",15} {"EASTING",15}");
                }

                writer.WriteLine();

                var elements = data["Elements"] as System.Collections.Generic.List<object>;
                if (elements != null)
                {
                    foreach (var element in elements)
                    {
                        var dict = element as System.Collections.Generic.Dictionary<string, object>;
                        writer.WriteLine($"Element: {dict["Type"]}");

                        // Write element properties
                        foreach (var kvp in dict)
                        {
                            if (kvp.Key != "Type")
                            {
                                writer.WriteLine($"    {kvp.Key}: {kvp.Value}");
                            }
                        }

                        writer.WriteLine();
                    }
                }
            }
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

            var lblType = new Label { Text = "Report Type:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            this.Controls.Add(lblType);

            rbVertical = new RadioButton { Text = "Vertical Alignment", Location = new System.Drawing.Point(40, 45), Checked = true, AutoSize = true };
            rbHorizontal = new RadioButton { Text = "Horizontal Alignment", Location = new System.Drawing.Point(40, 70), AutoSize = true };
            this.Controls.Add(rbVertical);
            this.Controls.Add(rbHorizontal);

            var lblOutput = new Label { Text = "Output Path:", Location = new System.Drawing.Point(20, 100), AutoSize = true };
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

            var lblFolder = new Label { Text = "Output Folder:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            this.Controls.Add(lblFolder);

            txtOutputFolder = new TextBox { Location = new System.Drawing.Point(20, 40), Width = 350 };
            txtOutputFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            this.Controls.Add(txtOutputFolder);

            btnBrowse = new Button { Text = "Browse...", Location = new System.Drawing.Point(380, 38), Width = 80 };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            var lblTypes = new Label { Text = "Report Types:", Location = new System.Drawing.Point(20, 75), AutoSize = true };
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
