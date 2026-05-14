using Microsoft.Office.Core;
using PPTSafeCommands.Utilities;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Drawing;

namespace PPTSafeCommands.Services
{
    /// <summary>
    /// Handles rectangle creation and formatting operations.
    /// </summary>
    public class RectangleShapeService
    {
        private readonly Logger _logger;

        public RectangleShapeService()
        {
            _logger = new Logger(typeof(RectangleShapeService));
        }

        /// <summary>
        /// Adds a formatted rectangle to the slide.
        /// </summary>
        public BooleanResult AddRectangle(PowerPoint.Slide slide)
        {
            try
            {
                _logger.Info(
                    "Starting rectangle creation process.");

                if (slide == null)
                {
                    _logger.Warn("Slide instance is null.");

                    return BooleanResult.FailResult(
                        "Slide instance is null.");
                }

                _logger.Info("Adding rectangle shape.");

                PowerPoint.Shape rectangleShape =
                    slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRectangle,
                        120,
                        120,
                        300,
                        120);

                _logger.Info(
                    "Rectangle shape added successfully.");

                ApplyFill(rectangleShape);

                ApplyBorder(rectangleShape);

                ApplyText(rectangleShape);

                _logger.Info(
                    "Rectangle formatting completed successfully.");

                return BooleanResult.SuccessResult(
                    "Rectangle added successfully.");
            }
            catch (Exception exception)
            {
                _logger.Error(
                    "Error while adding rectangle.",
                    exception);

                return BooleanResult.FailResult(
                    exception.Message);
            }
        }

        /// <summary>
        /// Applies fill formatting to the rectangle.
        /// </summary>
        private void ApplyFill(PowerPoint.Shape rectangleShape)
        {
            _logger.Info("Applying fill formatting.");

            rectangleShape.Fill.ForeColor.RGB =
                Color.LightBlue.ToArgb();

            rectangleShape.Fill.Transparency = 0.1f;

            _logger.Info("Fill formatting applied.");
        }

        /// <summary>
        /// Applies border formatting to the rectangle.
        /// </summary>
        private void ApplyBorder(PowerPoint.Shape rectangleShape)
        {
            _logger.Info("Applying border formatting.");

            rectangleShape.Line.ForeColor.RGB =
                Color.DarkBlue.ToArgb();

            rectangleShape.Line.Weight = 2;

            _logger.Info("Border formatting applied.");
        }

        /// <summary>
        /// Applies text formatting to the rectangle.
        /// </summary>
        private void ApplyText(PowerPoint.Shape rectangleShape)
        {
            _logger.Info("Applying rectangle text.");

            rectangleShape.TextFrame.TextRange.Text =
                "Exercise 2 Rectangle";

            rectangleShape.TextFrame.TextRange.Font.Size = 20;

            rectangleShape.TextFrame.TextRange.Font.Bold =
                MsoTriState.msoTrue;

            _logger.Info("Rectangle text applied.");
        }
    }
}