using Microsoft.Office.Tools.Ribbon;
using PPTSafeCommands.Helpers;
using PPTSafeCommands.Services;
using PPTSafeCommands.Utilities;
using System.Windows.Forms;

namespace PPTSafeCommands.Ribbon
{
    public partial class Ribbonbtn
    {
        private readonly Logger _logger =
            new Logger(typeof(Ribbonbtn));

        private void Ribbonbtn_Load(
            object sender,
            RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// Handles Add Rectangle button click.
        /// </summary>
        private void btnAddRectangle_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            _logger.Info(
                "Add Rectangle button clicked.");

            var presentationResult =
                PowerPointContextHelper.GetActivePresentation(
                    Globals.ThisAddIn.Application);

            if (!presentationResult.Success)
            {
                _logger.Warn(
                    presentationResult.Message);

                MessageBox.Show(
                    presentationResult.Message,
                    "Validation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            var slideResult =
                PowerPointContextHelper.GetActiveSlide(
                    Globals.ThisAddIn.Application);

            if (!slideResult.Success)
            {
                _logger.Warn(
                    slideResult.Message);

                MessageBox.Show(
                    slideResult.Message,
                    "Validation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            var rectangleShapeService =
                new RectangleShapeService();

            var result =
                rectangleShapeService.AddRectangle(
                    slideResult.Result);

            if (!result.Success)
            {
                _logger.Warn(result.Message);

                MessageBox.Show(
                    result.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            _logger.Info(
                "Rectangle command completed successfully.");

            MessageBox.Show(
                result.Message,
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}