using Microsoft.Office.Tools.Ribbon;
using Sample_exercise.Services;
using Sample_exercise.Utilities;
using System;
using System.Windows.Forms;

namespace Sample_exercise
{
    public partial class Ribbonbtn
    {
        private readonly PresentationInfoService _presentationInfoService =
            new PresentationInfoService();

        private void Ribbonbtn_Load(
            object sender,
            RibbonUIEventArgs e)
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

        private void btnShapeDiagnostics_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            try
            {
                var shapeDiagnosticService =
                    new ShapeDiagnosticService();

                shapeDiagnosticService.ExportShapeDiagnostics(
                    Globals.ThisAddIn.Application);
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

        private void btnGroupShapeTraversal_Click(
    object sender,
    RibbonControlEventArgs e)
        {
            try
            {
                GroupShapeTraversalService service =
                    new GroupShapeTraversalService();

                BooleanResult<string> result =
                    service.GenerateGroupShapeReport();

                if (!result.Success)
                {
                    MessageBox.Show(
                        result.Message,
                        "Group Shape Report");

                    return;
                }

                MessageBox.Show(
                    result.Result,
                    "Group Shape Report");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Error");
            }
        }
    }
}