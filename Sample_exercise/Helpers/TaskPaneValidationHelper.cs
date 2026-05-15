using Microsoft.Office.Interop.PowerPoint;
using Sample_exercise.Utilities;

namespace Sample_exercise.Helpers
{
    public static class TaskPaneValidationHelper
    {
        public static BooleanResult ValidatePresentation(
            Presentation presentation)
        {
            if (presentation == null)
            {
                return new BooleanResult(
                    false,
                    "No active presentation available.");
            }

            return new BooleanResult(
                true,
                string.Empty);
        }

        public static BooleanResult ValidateSlide(
            Slide slide)
        {
            if (slide == null)
            {
                return new BooleanResult(
                    false,
                    "No active slide available.");
            }

            return new BooleanResult(
                true,
                string.Empty);
        }
    }
}