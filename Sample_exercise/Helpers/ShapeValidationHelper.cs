using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sample_exercise.Helpers
{
    public static class ShapeValidationHelper
    {
        public static bool HasPresentation(PowerPoint.Application application)
        {
            return application?.ActivePresentation != null;
        }

        public static bool HasSlides(PowerPoint.Presentation presentation)
        {
            return presentation != null &&
                   presentation.Slides != null &&
                   presentation.Slides.Count > 0;
        }

        public static bool HasTextFrame(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }

            return shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        public static bool IsGroupShape(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }

            return shape.Type ==
                   Microsoft.Office.Core.MsoShapeType.msoGroup;
        }
    }
}