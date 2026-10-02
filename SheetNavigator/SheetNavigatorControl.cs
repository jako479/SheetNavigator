using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetNavigator
{
    /// <summary>
    /// The pane's content: a list of the workbook's visible sheets. Enter or a click on a name
    /// activates that sheet in this pane's own window, so several windows on one workbook stay independent.
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

        /// <summary>
        /// The workbook the list reads: the pane's own, or the active workbook if none was set.
        /// </summary>
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

            // Keys move the highlight natively; only Enter or a click activates, and either hands focus back to the window
            this.worksheetList.PreviewKeyDown += new PreviewKeyDownEventHandler(WorksheetList_PreviewKeyDown);
            this.worksheetList.KeyDown += new KeyEventHandler(WorksheetList_KeyDown);
            this.worksheetList.MouseClick += new MouseEventHandler(WorksheetList_MouseClick);

            // A double-click's second release raises only this event, so it takes the same path to hand focus back
            this.worksheetList.MouseDoubleClick += new MouseEventHandler(WorksheetList_MouseClick);

            // Once the keyboard leaves the list, the highlight follows the active sheet again
            this.worksheetList.LostFocus += new EventHandler(WorksheetList_LostFocus);

            // Excel raises no event for a sheet rename or reorder, so check as the pointer arrives
            this.worksheetList.MouseEnter += new EventHandler(WorksheetList_MouseEnter);
        }

        /// <summary>
        /// Brings the list in line with the workbook: rebuilds it only if the visible sheet names changed, then highlights the active sheet.
        /// </summary>
        public void RefreshWorksheets(Excel.Workbook activeWorkbook)
        {
            if (IsDisposed || activeWorkbook == null) return;

            // Comparing names catches sheets added, removed, renamed, reordered, hidden or unhidden
            List<string> names = VisibleSheetNames(activeWorkbook);
            if (SameAsList(names))
            {
                HighlightActiveSheet();
                return;
            }

            // A rebuild drops the highlight; a keyboard user gets it back on the same name if it survived
            string highlighted = IsKeyboardNavigating ? this.worksheetList.SelectedItem as string : null;

            this.worksheetList.BeginUpdate();
            try
            {
                this.worksheetList.Items.Clear();
                foreach (string name in names) this.worksheetList.Items.Add(name);
            }
            finally
            {
                this.worksheetList.EndUpdate();
            }

            if (highlighted != null && names.Contains(highlighted))
            {
                this.worksheetList.SelectedItem = highlighted;
            }
            else
            {
                HighlightActiveSheet(force: true);
            }
        }

        /// <summary>
        /// Moves the highlight to this window's active sheet, unless the keyboard is using the list:
        /// the highlight is the user's cursor then. While a chart sheet is active nothing is highlighted.
        /// </summary>
        public void HighlightActiveSheet()
        {
            HighlightActiveSheet(force: false);
        }

        /// <summary>
        /// The move itself. A rebuilt list has no highlight, so a rebuild forces one even while
        /// the keyboard is using the list.
        /// </summary>
        private void HighlightActiveSheet(bool force)
        {
            if (IsDisposed) return;
            if (!force && IsKeyboardNavigating) return;

            try
            {
                object active = Window != null ? Window.ActiveSheet : TargetWorkbook?.ActiveSheet;
                if (active is Excel.Worksheet currentSheet)
                {
                    this.worksheetList.SelectedItem = currentSheet.Name;
                }
                else
                {
                    // Leaving the old sheet highlighted would make the list disagree with the window
                    this.worksheetList.SelectedIndex = -1;
                }
            }
            catch (Exception ex) { Diagnostics.Write("Highlight failed: " + ex); }
        }

        /// <summary>
        /// True while the list has keyboard focus, when the highlight belongs to the user rather than to Excel.
        /// </summary>
        private bool IsKeyboardNavigating
        {
            get { return this.worksheetList.Focused; }
        }

        /// <summary>
        /// The names of the workbook's visible worksheets, in tab order.
        /// </summary>
        private static List<string> VisibleSheetNames(Excel.Workbook workbook)
        {
            // Chart sheets are not in Worksheets, so they are left out
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
            if (this.worksheetList.Items.Count != names.Count) return false;
            for (int i = 0; i < names.Count; i++)
            {
                if (!string.Equals(this.worksheetList.Items[i] as string, names[i], StringComparison.Ordinal)) return false;
            }
            return true;
        }

        /// <summary>
        /// Refreshes if Excel will answer and no macro has screen updating off, otherwise just re-highlights the active sheet.
        /// </summary>
        internal void RefreshQuietly()
        {
            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null) return;

                // Walking every sheet during a macro would slow it down; the list catches up once the macro is done
                Excel.Application app = Globals.ThisAddIn.Application;
                if (IsExcelEditing(app) || app.ScreenUpdating == false)
                {
                    HighlightActiveSheet();
                }
                else
                {
                    RefreshWorksheets(workbook);
                }
            }
            catch (Exception ex) { Diagnostics.Write("Quiet refresh failed: " + ex); }
        }

        /// <summary>
        /// True while Excel is busy or mid-edit (typing in a cell or a sheet tab name).
        /// </summary>
        private static bool IsExcelEditing(Excel.Application app)
        {
            if (app.Ready == false || app.Interactive == false) return true;

            // Most Ribbon commands are disabled mid-edit, which is the only reliable signal
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
        /// Puts the highlight back on the active sheet once the keyboard has left the list.
        /// </summary>
        private void WorksheetList_LostFocus(object sender, EventArgs e)
        {
            HighlightActiveSheet(force: true);
        }

        /// <summary>
        /// Enter is a dialog key the container would otherwise keep; the list wants it.
        /// </summary>
        private void WorksheetList_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Return) e.IsInputKey = true;
        }

        /// <summary>
        /// Enter activates the highlighted sheet. Every other key keeps the list's native
        /// behavior: arrows, Home/End, PgUp/PgDn and typed letters move the highlight.
        /// </summary>
        private void WorksheetList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Return) return;

            // Swallowed so Enter does not beep
            e.Handled = true;
            e.SuppressKeyPress = true;
            ActivateHighlightedSheet();
            ActivateWindow();
        }

        /// <summary>
        /// A left click on a name activates it, even one that was already highlighted; a click on
        /// the blank space below the names activates nothing. Either way focus goes back to the window.
        /// </summary>
        private void WorksheetList_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || IsDisposed) return;

            // IndexFromPoint names the nearest item for any point inside the list, blank space below
            // the names included, so the item's own rectangle decides whether a name was hit.
            int index = this.worksheetList.IndexFromPoint(e.Location);
            bool onName = index != ListBox.NoMatches && this.worksheetList.GetItemRectangle(index).Contains(e.Location);

            if (onName)
            {
                this.worksheetList.SelectedIndex = index;
                ActivateHighlightedSheet();
            }

            ActivateWindow();
        }

        /// <summary>
        /// Hands focus back to the pane's window, so Excel's own keys (Ctrl+PgUp/PgDn included) work right after a click or Enter.
        /// </summary>
        private void ActivateWindow()
        {
            try
            {
                Window?.Activate();
            }
            catch (Exception ex) { Diagnostics.Write("Window activation failed: " + ex); }
        }

        /// <summary>
        /// Activates the highlighted sheet in this pane's window. If the jump cannot happen,
        /// the list is refreshed so a renamed or removed sheet drops out.
        /// </summary>
        private void ActivateHighlightedSheet()
        {
            if (IsDisposed) return;

            Excel.Application app = Globals.ThisAddIn.Application;
            bool jumped = false;
            bool? previousScreenUpdating = null;
            string failure = null;

            try
            {
                string sheetName = this.worksheetList.SelectedItem as string;
                if (sheetName == null) return;

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
                // Excel refused because it is mid-edit; expected, so nothing is logged or shown
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Jump failed: " + ex);
                failure = ex.Message;
            }
            finally
            {
                // Restore rather than force on: a running macro may have turned screen updating off itself
                if (previousScreenUpdating.HasValue)
                {
                    try { app.ScreenUpdating = previousScreenUpdating.Value; }
                    catch (Exception ex) { Diagnostics.Write("ScreenUpdating restore failed: " + ex); }
                }

                if (!jumped)
                {
                    // Enter or a click means the user is done navigating, so the highlight returns to the
                    // active sheet even while the list still has focus
                    RefreshQuietly();
                    HighlightActiveSheet(force: true);
                }
            }

            // Only after ScreenUpdating is back on, so Excel repaints while the box is up
            if (failure != null)
            {
                MessageBox.Show(this, $"Could not jump to sheet: {failure}", "Navigation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
