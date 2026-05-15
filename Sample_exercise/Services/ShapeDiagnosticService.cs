using Sample_exercise.Utilities;
using Sample_exercise.Helpers;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Sample_exercise.Services
{
    public class ShapeDiagnosticService
    {
        private readonly Logger _logger;

        public ShapeDiagnosticService()
        {
            _logger = new Logger(typeof(ShapeDiagnosticService));
        }

        public void ExportShapeDiagnostics(
            PowerPoint.Application application)
        {
            try
            {
                _logger.Info("Shape diagnostic export started.");

                if (!ShapeValidationHelper.HasPresentation(application))
                {
                    MessageBox.Show(
                        "No active presentation found.");

                    return;
                }

                PowerPoint.Presentation presentation =
                    application.ActivePresentation;

                if (!ShapeValidationHelper.HasSlides(presentation))
                {
                    MessageBox.Show(
                        "No slides found in presentation.");

                    return;
                }

                List<ShapeDiagnosticInfo> diagnostics =
                    ReadAllShapeDetails(presentation);

                string report = BuildDiagnosticReport(diagnostics);

                ShowDiagnosticReport(report);

                _logger.Info("Shape diagnostic export completed.");
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Error while exporting shape diagnostics.",
                    ex);

                MessageBox.Show(
                    "Failed to export shape diagnostics.");
            }
        }

        private List<ShapeDiagnosticInfo> ReadAllShapeDetails(
            PowerPoint.Presentation presentation)
        {
            var diagnostics = new List<ShapeDiagnosticInfo>();

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                _logger.Info(
                    $"Reading slide: {slide.SlideIndex}");

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    ShapeDiagnosticInfo info =
                        ReadShapeInformation(
                            slide.SlideIndex,
                            shape);

                    diagnostics.Add(info);
                }
            }

            return diagnostics;
        }

        private ShapeDiagnosticInfo ReadShapeInformation(
            int slideNumber,
            PowerPoint.Shape shape)
        {
            var diagnosticInfo = new ShapeDiagnosticInfo();

            diagnosticInfo.SlideNumber = slideNumber;

            diagnosticInfo.ShapeName = shape.Name;

            diagnosticInfo.ShapeType = shape.Type.ToString();

            diagnosticInfo.Left = shape.Left;

            diagnosticInfo.Top = shape.Top;

            diagnosticInfo.Width = shape.Width;

            diagnosticInfo.Height = shape.Height;

            diagnosticInfo.HasTextFrame =
                ShapeValidationHelper.HasTextFrame(shape);

            diagnosticInfo.IsGroupShape =
                ShapeValidationHelper.IsGroupShape(shape);

            diagnosticInfo.PlaceholderType =
                GetPlaceholderType(shape);

            return diagnosticInfo;
        }

        private string GetPlaceholderType(
            PowerPoint.Shape shape)
        {
            try
            {
                if (shape.PlaceholderFormat != null)
                {
                    return shape.PlaceholderFormat.Type.ToString();
                }
            }
            catch
            {
                return "Not a Placeholder";
            }

            return "Not a Placeholder";
        }

        private string BuildDiagnosticReport(
            List<ShapeDiagnosticInfo> diagnostics)
        {
            var reportBuilder = new StringBuilder();

            reportBuilder.AppendLine(
                "===== Shape Diagnostics =====");

            reportBuilder.AppendLine();

            foreach (ShapeDiagnosticInfo item in diagnostics)
            {
                reportBuilder.AppendLine(
                    $"Slide Number : {item.SlideNumber}");

                reportBuilder.AppendLine(
                    $"Shape Name : {item.ShapeName}");

                reportBuilder.AppendLine(
                    $"Shape Type : {item.ShapeType}");

                reportBuilder.AppendLine(
                    $"Left : {item.Left}");

                reportBuilder.AppendLine(
                    $"Top : {item.Top}");

                reportBuilder.AppendLine(
                    $"Width : {item.Width}");

                reportBuilder.AppendLine(
                    $"Height : {item.Height}");

                reportBuilder.AppendLine(
                    $"Has Text Frame : {item.HasTextFrame}");

                reportBuilder.AppendLine(
                    $"Is Group Shape : {item.IsGroupShape}");

                reportBuilder.AppendLine(
                    $"Placeholder Type : {item.PlaceholderType}");

                reportBuilder.AppendLine(
                    "--------------------------------");
            }

            return reportBuilder.ToString();
        }

        private void ShowDiagnosticReport(string report)
        {
            MessageBox.Show(
                report,
                "Shape Diagnostics");
        }
    }
}