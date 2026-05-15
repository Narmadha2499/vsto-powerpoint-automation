namespace Sample_exercise.TaskPaneViews
{
    partial class SlideInfoTaskPaneControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
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
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblSlideNameValue = new System.Windows.Forms.Label();
            this.lblSlideIndexValue = new System.Windows.Forms.Label();
            this.lblShapeCountValue = new System.Windows.Forms.Label();
            this.txtErrorMessage = new System.Windows.Forms.TextBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblSlideNameValue
            // 
            this.lblSlideNameValue.AutoSize = true;
            this.lblSlideNameValue.Location = new System.Drawing.Point(20, 20);
            this.lblSlideNameValue.Name = "lblSlideNameValue";
            this.lblSlideNameValue.Size = new System.Drawing.Size(0, 13);
            this.lblSlideNameValue.TabIndex = 0;
            // 
            // lblSlideIndexValue
            // 
            this.lblSlideIndexValue.AutoSize = true;
            this.lblSlideIndexValue.Location = new System.Drawing.Point(20, 50);
            this.lblSlideIndexValue.Name = "lblSlideIndexValue";
            this.lblSlideIndexValue.Size = new System.Drawing.Size(0, 13);
            this.lblSlideIndexValue.TabIndex = 1;
            // 
            // lblShapeCountValue
            // 
            this.lblShapeCountValue.AutoSize = true;
            this.lblShapeCountValue.Location = new System.Drawing.Point(20, 80);
            this.lblShapeCountValue.Name = "lblShapeCountValue";
            this.lblShapeCountValue.Size = new System.Drawing.Size(0, 13);
            this.lblShapeCountValue.TabIndex = 2;
            // 
            // txtErrorMessage
            // 
            this.txtErrorMessage.Location = new System.Drawing.Point(20, 110);
            this.txtErrorMessage.Multiline = true;
            this.txtErrorMessage.Name = "txtErrorMessage";
            this.txtErrorMessage.ReadOnly = true;
            this.txtErrorMessage.Size = new System.Drawing.Size(200, 60);
            this.txtErrorMessage.TabIndex = 3;
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(20, 180);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 23);
            this.btnRefresh.TabIndex = 4;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(100, 180);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // SlideInfoTaskPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.txtErrorMessage);
            this.Controls.Add(this.lblShapeCountValue);
            this.Controls.Add(this.lblSlideIndexValue);
            this.Controls.Add(this.lblSlideNameValue);
            this.Name = "SlideInfoTaskPaneControl";
            this.Size = new System.Drawing.Size(250, 250);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSlideNameValue;
        private System.Windows.Forms.Label lblSlideIndexValue;
        private System.Windows.Forms.Label lblShapeCountValue;
        private System.Windows.Forms.TextBox txtErrorMessage;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnClose;
    }
}
