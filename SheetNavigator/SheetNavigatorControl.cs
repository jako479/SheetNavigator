using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetNavigator
{
    /// <summary>
    /// The pane's content: a list of the workbook's visible sheets. Selecting a name activates that sheet.
    /// </summary>
    public partial class SheetNavigatorControl : UserControl
    {
        /// <summary>
        /// Excel's generic "can't do that now" error, raised for example while a sheet tab name is being typed.
        /// </summary>
        private const int ExcelBusyHResult = unchecked((int)0x800A03EC);

        /// <summary>
        /// True while this control moves the highlight itself, so only the user's selections trigger a jump.
        /// </summary>
        private bool isUpdatingSelection;

        /// <summary>
        /// The workbook shown in the window this pane belongs to. Falls back to the active workbook.
        /// </summary>
        internal Excel.Workbook Workbook { get; set; }

        private Excel.Workbook TargetWorkbook
        {
            get { return Workbook ?? Globals.ThisAddIn.Application.ActiveWorkbook; }
        }

        public SheetNavigatorControl()
        {
            InitializeComponent();

            // A single click (or arrow key) selects and jumps
            this.WorksheetList.SelectedIndexChanged += new EventHandler(WorksheetList_SelectedIndexChanged);

            // Excel raises no event for a sheet rename or reorder, so refresh as the pointer arrives
            this.WorksheetList.MouseEnter += new EventHandler(WorksheetList_MouseEnter);
        }

        /// <summary>
        /// Rebuilds the list from the workbook's visible sheets and highlights the active one.
        /// </summary>
        public void RefreshWorksheets(Excel.Workbook activeWorkbook)
        {
            bool wasUpdating = isUpdatingSelection;
            isUpdatingSelection = true;
            this.WorksheetList.BeginUpdate();
            try
            {
                this.WorksheetList.Items.Clear();

                if (activeWorkbook != null)
                {
                    foreach (Excel.Worksheet ws in activeWorkbook.Worksheets)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            this.WorksheetList.Items.Add(ws.Name);
                        }
                    }

                    HighlightActiveSheet(activeWorkbook);
                }
            }
            finally
            {
                this.WorksheetList.EndUpdate();
                isUpdatingSelection = wasUpdating;
            }
        }

        /// <summary>
        /// Moves the highlight to the workbook's active sheet without triggering a jump.
        /// </summary>
        private void HighlightActiveSheet(Excel.Workbook workbook)
        {
            bool wasUpdating = isUpdatingSelection;
            isUpdatingSelection = true;
            try
            {
                if (workbook?.ActiveSheet is Excel.Worksheet currentSheet)
                {
                    this.WorksheetList.SelectedItem = currentSheet.Name;
                }
            }
            catch { /* Leave the highlight alone */ }
            finally
            {
                isUpdatingSelection = wasUpdating;
            }
        }

        /// <summary>
        /// Rebuilds the list if Excel will answer, otherwise just re-highlights the active sheet.
        /// </summary>
        private void RefreshQuietly()
        {
            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null) return;

                if (IsExcelEditing(Globals.ThisAddIn.Application))
                {
                    HighlightActiveSheet(workbook);
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

        private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
            }
            return null;
        }

        private void WorksheetList_MouseEnter(object sender, EventArgs e)
        {
            RefreshQuietly();
        }

        /// <summary>
        /// Activates the sheet the user selected. If the jump cannot happen, the list is refreshed
        /// and the highlight returns to the sheet Excel is still on.
        /// </summary>
        private void WorksheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingSelection || this.WorksheetList.SelectedItem == null) return;

            string selectedSheetName = this.WorksheetList.SelectedItem.ToString();
            Excel.Application app = Globals.ThisAddIn.Application;
            bool jumped = false;
            bool? previousScreenUpdating = null;

            try
            {
                Excel.Workbook workbook = TargetWorkbook;
                if (workbook == null || IsExcelEditing(app)) return;

                // The sheet may have been renamed or removed since the list was filled
                Excel.Worksheet targetSheet = FindSheet(workbook, selectedSheetName);
                if (targetSheet == null) return;

                previousScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                targetSheet.Activate();
                jumped = true;
            }
            catch (COMException ex) when (ex.HResult == ExcelBusyHResult)
            {
                // Excel refused because it is mid-edit; nothing to report
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not jump to sheet: {ex.Message}", "Navigation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Put Excel's setting back exactly as found
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
