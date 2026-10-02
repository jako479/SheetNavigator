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
            this.worksheetList = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            //
            // worksheetList
            //
            this.worksheetList.AccessibleName = "Worksheets";
            this.worksheetList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.worksheetList.FormattingEnabled = true;
            this.worksheetList.HorizontalScrollbar = true;
            this.worksheetList.Location = new System.Drawing.Point(0, 0);
            this.worksheetList.Name = "worksheetList";
            this.worksheetList.Size = new System.Drawing.Size(150, 150);
            this.worksheetList.TabIndex = 0;
            // 
            // SheetNavigatorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.worksheetList);
            this.Name = "SheetNavigatorControl";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox worksheetList;
    }
}
