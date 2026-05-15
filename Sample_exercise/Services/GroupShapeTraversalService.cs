using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Sample_exercise.Utilities;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sample_exercise.Services
{
    public class GroupShapeTraversalService
    {
        private readonly Logger _logger;

        public GroupShapeTraversalService()
        {
            _logger = new Logger(typeof(GroupShapeTraversalService));
        }

        public BooleanResult<string>GenerateGroupShapeReport()
        {
            try
            {
                _logger.Info("Starting group shape traversal.");

PowerPoint.Application application =
    Sample_exercise.Globals.ThisAddIn.Application;
                if (application == null)
                {
                    return BooleanResult<string>.FailResult("PowerPoint application not found.");
                }

                PowerPoint.Presentation presentation = application.ActivePresentation;

                if (presentation == null)
                {
                    return BooleanResult<string>.FailResult("No active presentation found.");
                }

                if (application.ActiveWindow == null)
                {
                    return BooleanResult<string>.FailResult("No active window found.");
                }

                PowerPoint.Slide slide = application.ActiveWindow.View.Slide;

                if (slide == null)
                {
                    return BooleanResult<string>.FailResult("No active slide found.");
                }

                List<GroupShapeReportInfo> reportItems = new List<GroupShapeReportInfo>();

                TraverseShapes(slide.Shapes, reportItems, 0);

                string report = CreateReport(reportItems);

                _logger.Info("Group shape traversal completed.");

                return BooleanResult<string>.SuccessResult(report);
            }
            catch (Exception ex)
            {
                _logger.Error("Failed to generate group shape report.", ex);

                return BooleanResult<string>.FailResult(ex.Message);
            }
        }

        private void TraverseShapes(
     dynamic shapes,
     List<GroupShapeReportInfo> reportItems,
     int depthLevel)
        {
            try
            {
                foreach (PowerPoint.Shape shape
                    in shapes)
                {
                    if (shape == null)
                    {
                        continue;
                    }

                    bool isGroupShape =
                        shape.Type ==
                        MsoShapeType.msoGroup;

                    _logger.Info(
                        $"Processing shape: {shape.Name}");

                    reportItems.Add(
                     new GroupShapeReportInfo
                     {
                         ShapeId = shape.Id,
                         ShapeName = shape.Name,
                         ShapeType = shape.Type.ToString(),
                         LeftPosition = shape.Left,
                         TopPosition = shape.Top,
                         Width = shape.Width,
                         Height = shape.Height,
                         Rotation = shape.Rotation,
                         DepthLevel = depthLevel,
                         IsGroupShape = isGroupShape
                     });

                    if (isGroupShape)
                    {
                        _logger.Info(
                            $"Entering group shape: {shape.Name}");

                        TraverseShapes(
                            shape.GroupItems,
                            reportItems,
                            depthLevel + 1);
                    }
                }
            }
            catch (Exception exception)
            {
                _logger.Error(
                    "Shape traversal failed.",
                    exception);
            }
        }

        private string CreateReport(
    List<GroupShapeReportInfo> reportItems)
        {
            StringBuilder reportBuilder =
                new StringBuilder();

            foreach (GroupShapeReportInfo item
                in reportItems)
            {
                string indentation =
                    new string(
                        ' ',
                        item.DepthLevel * 4);

                reportBuilder.AppendLine(
                    $"{indentation}Shape Id : {item.ShapeId}");

                reportBuilder.AppendLine(
                    $"{indentation}Shape Name : {item.ShapeName}");

                reportBuilder.AppendLine(
                    $"{indentation}Shape Type : {item.ShapeType}");

                reportBuilder.AppendLine(
                    $"{indentation}Left : {item.LeftPosition}");

                reportBuilder.AppendLine(
                    $"{indentation}Top : {item.TopPosition}");

                reportBuilder.AppendLine(
                    $"{indentation}Width : {item.Width}");

                reportBuilder.AppendLine(
                    $"{indentation}Height : {item.Height}");

                reportBuilder.AppendLine(
                    $"{indentation}Rotation : {item.Rotation}");

                reportBuilder.AppendLine(
                    $"{indentation}Is Group Shape : {item.IsGroupShape}");

                reportBuilder.AppendLine();
            }

            return reportBuilder.ToString();
        }
    }
}