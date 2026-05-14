using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sample_exercise.Helpers
{
    public static class PowerPointContextHelper
    {
        public static PowerPoint.Slide GetActiveSlide(
            PowerPoint.Application application)
        {
            try
            {
                if (application?.ActiveWindow?.View == null)
                {
                    return null;
                }

                return application.ActiveWindow.View.Slide;
            }
            catch
            {
                return null;
            }
        }

        public static PowerPoint.Selection GetSelection(
            PowerPoint.Application application)
        {
            try
            {
                return application?.ActiveWindow?.Selection;
            }
            catch
            {
                return null;
            }
        }
    }
}