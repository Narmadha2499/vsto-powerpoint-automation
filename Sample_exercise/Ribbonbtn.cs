//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Windows.Forms;
//using Microsoft.Office.Tools.Ribbon;
//using Office = Microsoft.Office.Core;
//using PowerPoint = Microsoft.Office.Interop.PowerPoint;

//namespace Sample_exercise
//{
//    public partial class Ribbonbtn
//    {
//        private void Ribbonbtn_Load(object sender, RibbonUIEventArgs e)
//        {

//        }
//        private void btnPresentation_Click(object sender, RibbonControlEventArgs e)
//        {
//            PowerPoint.Application app =
//                Globals.ThisAddIn.Application;

//            PowerPoint.Presentation presentation =
//                app.ActivePresentation;

//            PowerPoint.Slide slide =
//                app.ActiveWindow.View.Slide;

//            PowerPoint.Selection selection =
//                app.ActiveWindow.Selection;

//            StringBuilder info = new StringBuilder();

//            // Presentation Info
//            info.AppendLine("Presentation Name: "
//                + presentation.Name);

//            info.AppendLine();

//            info.AppendLine("Total Slides: "
//                + presentation.Slides.Count);

//            info.AppendLine();

//            // Active Slide
//            info.AppendLine("Active Slide Number: "
//                + slide.SlideIndex);

//            info.AppendLine();

//            // Selected Shapes
//            if (selection.Type ==
//                PowerPoint.PpSelectionType.ppSelectionShapes)
//            {
//                PowerPoint.ShapeRange shapeRange =
//                    selection.ShapeRange;

//                info.AppendLine("Selected Shape Count: "
//                    + shapeRange.Count);

//                info.AppendLine();

//                info.AppendLine("Shape Details:");

//                foreach (PowerPoint.Shape shape in shapeRange)
//                {
//                    info.AppendLine(
//                        shape.Name
//                        + " - "
//                        + shape.Type.ToString());
//                }
//            }
//            else
//            {
//                info.AppendLine("No Shapes Selected");
//            }

//            MessageBox.Show(
//                info.ToString(),
//                "Presentation Info");
//        }
//    }
//}

using Microsoft.Office.Tools.Ribbon;
using Sample_exercise.Services;
using System;
using System.Windows.Forms;

namespace Sample_exercise
{
    public partial class Ribbonbtn
    {
        private readonly PresentationInfoService _presentationInfoService =
            new PresentationInfoService();

        private void Ribbonbtn_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnPresentation_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            try
            {
                string presentationInformation =
                    _presentationInfoService.GetPresentationInformation();

                MessageBox.Show(
                    presentationInformation,
                    "Presentation Info",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    exception.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}