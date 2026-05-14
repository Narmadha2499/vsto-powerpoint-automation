using Sample_exercise.Utilities;
using Sample_exercise.Helpers;
using System;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sample_exercise.Services
{
    public class PresentationInfoService
    {
        private readonly Logger _logger =
            new Logger(typeof(PresentationInfoService));

        public string GetPresentationInformation()
        {
            _logger.Info("Started reading presentation information.");

            try
            {
                PowerPoint.Application application =
                    Globals.ThisAddIn.Application;

                ValidateApplication(application);

                PowerPoint.Presentation presentation =
                    application.ActivePresentation;

                ValidatePresentation(presentation);

                StringBuilder informationBuilder =
                    new StringBuilder();

                AppendPresentationInformation(
                    informationBuilder,
                    presentation);

                AppendSlideInformation(
                    informationBuilder,
                    application);

                AppendShapeInformation(
                    informationBuilder,
                    application);

                _logger.Info("Successfully retrieved presentation information.");

                return informationBuilder.ToString();
            }
            catch (Exception exception)
            {
                _logger.Error(
                    "Failed while retrieving presentation information.",
                    exception);

                throw;
            }
        }

        private void ValidateApplication(
            PowerPoint.Application application)
        {
            if (application == null)
            {
                throw new InvalidOperationException(
                    "PowerPoint application instance is not available.");
            }
        }

        private void ValidatePresentation(
            PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                throw new InvalidOperationException(
                    "No active presentation found.");
            }
        }

        private void AppendPresentationInformation(
            StringBuilder informationBuilder,
            PowerPoint.Presentation presentation)
        {
            informationBuilder.AppendLine(
                "Presentation Name: " +
                presentation.Name);

            informationBuilder.AppendLine();

            informationBuilder.AppendLine(
                "Total Slides: " +
                presentation.Slides.Count);

            informationBuilder.AppendLine();
        }

        private void AppendSlideInformation(
            StringBuilder informationBuilder,
            PowerPoint.Application application)
        {
            PowerPoint.Slide activeSlide =
                PowerPointContextHelper.GetActiveSlide(application);

            if (activeSlide == null)
            {
                informationBuilder.AppendLine(
                    "No active slide available.");

                informationBuilder.AppendLine();

                return;
            }

            informationBuilder.AppendLine(
                "Active Slide Number: " +
                activeSlide.SlideIndex);

            informationBuilder.AppendLine();
        }

        private void AppendShapeInformation(
            StringBuilder informationBuilder,
            PowerPoint.Application application)
        {
            PowerPoint.Selection selection =
                PowerPointContextHelper.GetSelection(application);

            if (selection == null)
            {
                informationBuilder.AppendLine(
                    "Selection is not available.");

                return;
            }

            if (selection.Type !=
                PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                informationBuilder.AppendLine(
                    "No shapes selected.");

                return;
            }

            PowerPoint.ShapeRange shapeRange =
                selection.ShapeRange;

            informationBuilder.AppendLine(
                "Selected Shape Count: " +
                shapeRange.Count);

            informationBuilder.AppendLine();

            informationBuilder.AppendLine(
                "Shape Details:");

            foreach (PowerPoint.Shape shape in shapeRange)
            {
                AppendShapeDetails(
                    informationBuilder,
                    shape);
            }
        }

        private void AppendShapeDetails(
            StringBuilder informationBuilder,
            PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            informationBuilder.AppendLine(
                shape.Name +
                " - " +
                shape.Type);
        }
    }
}