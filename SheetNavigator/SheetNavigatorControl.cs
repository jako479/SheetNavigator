using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetNavigator
{
    /// <summary>
    /// The pane's content: a list of the workbook's visible sheets. Clicking a name activates
    /// that sheet in this pane's own window, so several windows on one workbook stay independent.
    /// </summary>
    public partial class SheetNavigatorControl : UserControl
    {
        /// <summary>
        /// Excel's generic "can't do that now" error, raised for example while a sheet tab name is being typed.
        /// </summary>
        private const int ExcelBusyHResult = unchecked((int)0x800A03EC);

        /// <summary>
        /// The Excel window this pane belongs to. Each window has its own active sheet.
        /// </summary>
        internal Excel.Window Window { get; set; }

        /// <summary>
        /// The workbook shown in that window. Falls back to the active workbook.
        /// </summary>
        internal Excel.Workbook Workbook { get; set; }

        private Excel.Workbook TargetWorkbook
        {
            get { return Workbook ?? Globals.ThisAddIn.Application.ActiveWorkbook; }
        }

        /// <summary>
        /// Wires the list's events; the layout itself comes from the designer.
        /// </summary>
        public SheetNavigatorControl()
        {
            InitializeComponent();

            // The jump waits for the mouse release, so a drag across the list jumps once, where it ends;
            // keys are swallowed so the highlight can only move by mouse
            this.WorksheetList.MouseUp += new MouseEventHandler(WorksheetList_MouseUp);
            this.WorksheetList.KeyDown += new KeyEventHandler(WorksheetList_KeyDown);

            // Excel raises no event for a sheet rename or reorder, so check as the pointer arrives
            this.WorksheetList.MouseEnter += new EventHandler(WorksheetList_MouseEnter);
        }

        /// <summary>
        /// Brings the list in line with the workbook: rebuilds it only if the visible sheet names
        /// changed (added, removed, renamed, reordered, hidden or unhidden), then highlights the active sheet.
        /// </summary>
        public void RefreshWorksheets(Excel.Workbook activeWorkbook)
        {
            if (IsDisposed || activeWorkbook == null) return;

            List<string> names = VisibleSheetNames(activeWorkbook);
            if (!SameAsList(names))
            {
                this.WorksheetList.BeginUpdate();
                try
                {
                    this.WorksheetList.Items.Clear();
                    foreach (string name in names) this.WorksheetList.Items.Add(name);
                }
                finally
                {
                    this.WorksheetList.EndUpdate();
                }
            }

            HighlightActiveSheet();
        }

        /// <summary>
        /// Moves the highlight to this window's active sheet. A chart sheet is not in the list, so
        /// while one is active nothing is highlighted.
        /// </summary>
        public void HighlightActiveSheet()
        {
            if (IsDisposed) return;

            try
            {
                object active = Window != null ? Window.ActiveSheet : TargetWorkbook?.ActiveSheet;
                if (active is Excel.Worksheet currentSheet)
                {
                    this.WorksheetList.SelectedItem = currentSheet.Name;
                }
                else
                {
                    // Leaving the old sheet highlighted would make the list disagree with the window
                    this.WorksheetList.SelectedIndex = -1;
                }
            }
            catch { /* Leave the highlight alone */ }
        }

        /// <summary>
        /// The names of the workbook's visible worksheets, in tab order. Chart sheets are not
        /// worksheets, so they are left out.
        /// </summary>
        private static List<string> VisibleSheetNames(Excel.Workbook workbook)
        {
            List<string> names = new List<string>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) names.Add(worksheet.Name);
            }
            return names;
        }

        /// <summary>
        /// True when the list already shows exactly these names in this order.
        /// </summary>
        private bool SameAsList(List<string> names)
        {
            if (this.WorksheetList.Items.Count != names.Count) return false;
            for (int i = 0; i < names.Count; i++)
            {
                if (!string.Equals(this.WorksheetList.Items[i] as string, names[i], StringComparison.Ordinal)) return false;
            }
            return true;
        }

        /// <summary>
        /// Refreshes if Excel will answer, otherwise just re-highlights the active sheet.
        /// </summary>
        internal void RefreshQuietly()
        {
            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null) return;

                if (IsExcelEditing(Globals.ThisAddIn.Application))
                {
                    HighlightActiveSheet();
                }
                else
                {
                    RefreshWorksheets(workbook);
                }
            }
            catch { /* Keep the current list */ }
        }

        /// <summary>
        /// True while Excel is busy or mid-edit (typing in a cell or a sheet tab name).
        /// Most Ribbon commands are disabled then, which is the only reliable signal.
        /// </summary>
        private static bool IsExcelEditing(Excel.Application app)
        {
            if (app.Ready == false || app.Interactive == false) return true;
            return app.CommandBars.GetEnabledMso("FileNewDefault") == false;
        }

        /// <summary>
        /// The worksheet with this name, or null if it has been renamed or removed since the list was filled.
        /// </summary>
        private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase)) return worksheet;
            }
            return null;
        }

        /// <summary>
        /// Catches a rename or reorder as the pointer arrives, before the user can click a stale name.
        /// </summary>
        private void WorksheetList_MouseEnter(object sender, EventArgs e)
        {
            RefreshQuietly();
        }

        /// <summary>
        /// Blocks keyboard navigation in the list; Excel's Ctrl+PgUp/PgDn already covers that.
        /// </summary>
        private void WorksheetList_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        /// <summary>
        /// Jumps to the name under the pointer when the left button is released, so a drag across the
        /// list jumps once, where it ends. A release off the names only puts the highlight back.
        /// </summary>
        private void WorksheetList_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || IsDisposed) return;

            int index = this.WorksheetList.IndexFromPoint(e.Location);
            if (index == ListBox.NoMatches)
            {
                HighlightActiveSheet();
                return;
            }

            JumpTo(this.WorksheetList.Items[index] as string);
        }

        /// <summary>
        /// Activates the named sheet in this pane's window. If the jump cannot happen,
        /// the list is refreshed and the highlight returns to the sheet the window is still on.
        /// </summary>
        private void JumpTo(string sheetName)
        {
            if (sheetName == null) return;

            Excel.Application app = Globals.ThisAddIn.Application;
            bool jumped = false;
            bool? previousScreenUpdating = null;

            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null || IsExcelEditing(app)) return;

                // The sheet may have been renamed or removed since the list was filled
                Excel.Worksheet targetSheet = FindSheet(workbook, sheetName);
                if (targetSheet == null) return;

                previousScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                // Worksheet.Activate acts on the workbook's active window, so make it this one first
                Window?.Activate();
                targetSheet.Activate();
                jumped = true;
            }
            catch (COMException ex) when (ex.HResult == ExcelBusyHResult)
            {
                // Excel refused because it is mid-edit; nothing to report
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Jump failed: " + ex);
                MessageBox.Show($"Could not jump to sheet: {ex.Message}", "Navigation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Restore rather than force on: a running macro may have turned screen updating off itself
                if (previousScreenUpdating.HasValue)
                {
                    try { app.ScreenUpdating = previousScreenUpdating.Value; }
                    catch { }
                }

                if (!jumped) RefreshQuietly();
            }
        }
    }
}
