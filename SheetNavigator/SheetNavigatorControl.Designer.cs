namespace SheetNavigator
{
    partial class SheetNavigatorControl
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
            this.WorksheetList = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // WorksheetList
            // 
            this.WorksheetList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.WorksheetList.FormattingEnabled = true;
            this.WorksheetList.HorizontalScrollbar = true;
            this.WorksheetList.Location = new System.Drawing.Point(0, 0);
            this.WorksheetList.Name = "WorksheetList";
            this.WorksheetList.Size = new System.Drawing.Size(150, 150);
            this.WorksheetList.TabIndex = 0;
            // 
            // SheetNavigatorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.WorksheetList);
            this.Name = "SheetNavigatorControl";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox WorksheetList;
    }
}
