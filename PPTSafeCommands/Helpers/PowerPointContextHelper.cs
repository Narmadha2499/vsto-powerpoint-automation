using Microsoft.Office.Interop.PowerPoint;
using PPTSafeCommands.Utilities;
using System;

namespace PPTSafeCommands.Helpers
{
    /// <summary>
    /// Provides validation helpers for PowerPoint context objects.
    /// </summary>
    public static class PowerPointContextHelper
    {
        private static readonly Logger Logger =
            new Logger(typeof(PowerPointContextHelper));

        /// <summary>
        /// Validates the active presentation.
        /// </summary>
        public static BooleanResult<Presentation> GetActivePresentation(
            Application application)
        {
            try
            {
                Logger.Info("Validating active presentation.");

                if (application == null)
                {
                    Logger.Warn("PowerPoint application instance is null.");

                    return BooleanResult<Presentation>.FailResult(
                        "Application instance is null.");
                }

                Presentation activePresentation =
                    application.ActivePresentation;

                if (activePresentation == null)
                {
                    Logger.Warn("No active presentation found.");

                    return BooleanResult<Presentation>.FailResult(
                        "No active presentation available.");
                }

                Logger.Info(
                    "Active presentation validated successfully.");

                return BooleanResult<Presentation>.SuccessResult(
                    activePresentation);
            }
            catch (Exception exception)
            {
                Logger.Error(
                    "Error while validating active presentation.",
                    exception);

                return BooleanResult<Presentation>.FailResult(
                    exception.Message);
            }
        }

        /// <summary>
        /// Validates the active slide.
        /// </summary>
        public static BooleanResult<Slide> GetActiveSlide(
            Application application)
        {
            try
            {
                Logger.Info("Validating active slide.");

                if (application?.ActiveWindow?.View?.Slide == null)
                {
                    Logger.Warn("No active slide found.");

                    return BooleanResult<Slide>.FailResult(
                        "No active slide available.");
                }

                Slide activeSlide =
                    application.ActiveWindow.View.Slide;

                Logger.Info(
                    "Active slide validated successfully.");

                return BooleanResult<Slide>.SuccessResult(
                    activeSlide);
            }
            catch (Exception exception)
            {
                Logger.Error(
                    "Error while validating active slide.",
                    exception);

                return BooleanResult<Slide>.FailResult(
                    exception.Message);
            }
        }
    }
}