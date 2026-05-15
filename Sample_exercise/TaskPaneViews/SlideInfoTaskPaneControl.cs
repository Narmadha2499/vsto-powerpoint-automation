using System;
using System.Windows.Forms;
using Sample_exercise.Utilities;

namespace Sample_exercise.TaskPaneViews
{
    public partial class SlideInfoTaskPaneControl : UserControl
    {
        public event EventHandler RefreshButtonClicked;

        public event EventHandler CloseButtonClicked;

        public SlideInfoTaskPaneControl()
        {
            InitializeComponent();
        }

        public void UpdateSlideInformation(
            SlideInfoReportInfo slideInfoReportInfo)
        {
            if (slideInfoReportInfo == null)
            {
                return;
            }

            lblSlideNameValue.Text =
                slideInfoReportInfo.SlideName;

            lblSlideIndexValue.Text =
                slideInfoReportInfo.SlideIndex.ToString();

            lblShapeCountValue.Text =
                slideInfoReportInfo.ShapeCount.ToString();

            txtErrorMessage.Text =
                slideInfoReportInfo.ErrorMessage;
        }

        private void btnRefresh_Click(
            object sender,
            EventArgs e)
        {
            RefreshButtonClicked?.Invoke(
                this,
                EventArgs.Empty);
        }

        private void btnClose_Click(
            object sender,
            EventArgs e)
        {
            CloseButtonClicked?.Invoke(
                this,
                EventArgs.Empty);
        }
    }
}