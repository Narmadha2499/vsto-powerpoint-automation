using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Sample_exercise.Helpers;
using Sample_exercise.TaskPaneViews;
using Sample_exercise.Utilities;
using Sample_exercise.Services;
using System;

namespace Sample_exercise.Services
{
    public class SlideInfoTaskPaneService
    {
        private readonly PowerPoint.Application application;

        private CustomTaskPane customTaskPane;

        private SlideInfoTaskPaneControl slideInfoTaskPaneControl;

        public SlideInfoTaskPaneService(
            PowerPoint.Application application)
        {
            this.application = application;
        }

        public void ShowTaskPane()
        {
            try
            {
                if (customTaskPane == null)
                {
                    CreateTaskPane();
                }

                customTaskPane.Visible = true;

                RefreshSlideInformation();
            }
            catch (Exception exception)
            {
                Logger.LogError(
                    exception);
            }
        }

        private void CreateTaskPane()
        {
            slideInfoTaskPaneControl =
                new SlideInfoTaskPaneControl();

            slideInfoTaskPaneControl.RefreshButtonClicked +=
                SlideInfoTaskPaneControl_RefreshButtonClicked;

            slideInfoTaskPaneControl.CloseButtonClicked +=
                SlideInfoTaskPaneControl_CloseButtonClicked;

            customTaskPane =
                Globals.ThisAddIn.CustomTaskPanes.Add(
                    slideInfoTaskPaneControl,
                    "Slide Information");

            customTaskPane.Width = 300;
        }

        private void SlideInfoTaskPaneControl_RefreshButtonClicked(
            object sender,
            EventArgs e)
        {
            RefreshSlideInformation();
        }

        private void SlideInfoTaskPaneControl_CloseButtonClicked(
            object sender,
            EventArgs e)
        {
            customTaskPane.Visible = false;
        }

        public void RefreshSlideInformation()
        {
            try
            {
                PowerPoint.Presentation activePresentation =
                    application.ActivePresentation;

                BooleanResult presentationValidationResult =
                    TaskPaneValidationHelper.ValidatePresentation(
                        activePresentation);

                if (!presentationValidationResult.IsValid)
                {
                    UpdateError(
                        presentationValidationResult.Message);

                    return;
                }

                PowerPoint.Slide currentSlide =
                    application.ActiveWindow.View.Slide;

                BooleanResult slideValidationResult =
                    TaskPaneValidationHelper.ValidateSlide(
                        currentSlide);

                if (!slideValidationResult.IsValid)
                {
                    UpdateError(
                        slideValidationResult.Message);

                    return;
                }

                SlideInfoReportInfo slideInfoReportInfo =
                    new SlideInfoReportInfo
                    {
                        SlideName = currentSlide.Name,
                        SlideIndex = currentSlide.SlideIndex,
                        ShapeCount = currentSlide.Shapes.Count,
                        ErrorMessage = string.Empty
                    };

                slideInfoTaskPaneControl.UpdateSlideInformation(
                    slideInfoReportInfo);
            }
            catch (Exception exception)
            {
                Logger.LogError(
                    exception);

                UpdateError(
                    "Unable to retrieve current slide information.");
            }
        }

        private void UpdateError(
            string errorMessage)
        {
            SlideInfoReportInfo slideInfoReportInfo =
                new SlideInfoReportInfo
                {
                    SlideName = string.Empty,
                    SlideIndex = 0,
                    ShapeCount = 0,
                    ErrorMessage = errorMessage
                };

            slideInfoTaskPaneControl.UpdateSlideInformation(
                slideInfoReportInfo);
        }
    }
}