namespace Sample_exercise
{
    partial class Ribbonbtn : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbonbtn()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">
        /// true if managed resources should be disposed;
        /// otherwise, false.
        /// </param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support.
        /// Do not modify the contents of this method
        /// with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnPresentation = this.Factory.CreateRibbonButton();
            this.btnShapeDiagnostics = this.Factory.CreateRibbonButton();

            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();

            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType =
                Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;

            this.tab1.Groups.Add(this.group1);

            this.tab1.Label = "TabAddIns";

            this.tab1.Name = "tab1";

            // 
            // group1
            // 
            this.group1.Items.Add(this.btnPresentation);

            this.group1.Items.Add(this.btnShapeDiagnostics);

            this.group1.Label = "group1";

            this.group1.Name = "group1";

            // 
            // btnPresentation
            // 
            this.btnPresentation.Label = "Get Presentation Info";

            this.btnPresentation.Name = "btnPresentation";

            this.btnPresentation.Click +=
                new Microsoft.Office.Tools.Ribbon
                .RibbonControlEventHandler(
                    this.btnPresentation_Click);

            // 
            // btnShapeDiagnostics
            // 
            this.btnShapeDiagnostics.Label =
                "Shape Diagnostics";

            this.btnShapeDiagnostics.Name =
                "btnShapeDiagnostics";

            this.btnShapeDiagnostics.Click +=
                new Microsoft.Office.Tools.Ribbon
                .RibbonControlEventHandler(
                    this.btnShapeDiagnostics_Click);

            // 
            // Ribbonbtn
            // 
            this.Name = "Ribbonbtn";

            this.RibbonType =
                "Microsoft.PowerPoint.Presentation";

            this.Tabs.Add(this.tab1);

            this.Load +=
                new Microsoft.Office.Tools.Ribbon
                .RibbonUIEventHandler(
                    this.Ribbonbtn_Load);

            this.tab1.ResumeLayout(false);

            this.tab1.PerformLayout();

            this.group1.ResumeLayout(false);

            this.group1.PerformLayout();

            this.ResumeLayout(false);
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPresentation;

        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShapeDiagnostics;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbonbtn Ribbonbtn
        {
            get
            {
                return this.GetRibbon<Ribbonbtn>();
            }
        }
    }
}