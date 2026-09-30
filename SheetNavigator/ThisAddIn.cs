using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SheetNavigator
{
    /// <summary>
    /// Excel add-in that shows a "Worksheets" task pane listing the sheets of a workbook.
    /// Each Excel window gets its own pane. Which workbooks show the pane, and at what width,
    /// is remembered per user in the add-in's settings.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Saved widths are clamped to this range on both save and load.</summary>
        private const int MinPaneWidth = 100;
        private const int MaxPaneWidth = 800;

        private int? cachedDefaultWidth = null;

        /// <summary>One pane per Excel window, keyed by window handle.</summary>
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();
        private Ribbon ribbon;

        /// <summary>Delays the width save until the user has stopped dragging the pane border.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };
        private PaneEntry pendingResizeEntry;

        /// <summary>
        /// True while the add-in sets a pane's <c>Visible</c> itself, so
        /// <see cref="MyCustomTaskPane_VisibleChanged"/> only reacts to user actions.
        /// </summary>
        private bool isProgrammaticVisibilityChange = false;

        /// <summary>
        /// True when the active window's pane is shown.
        /// </summary>
        public bool IsSidebarVisible
        {
            get
            {
                try
                {
                    PaneEntry entry = FindPane(this.Application.ActiveWindow);
                    return entry != null && entry.Pane.Visible;
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// Saved "path|width" entries, one per tracked workbook. Created on first use.
        /// </summary>
        private StringCollection TrackedFiles
        {
            get
            {
                if (Properties.Settings.Default.TrackedFiles == null)
                {
                    Properties.Settings.Default.TrackedFiles = new StringCollection();
                }
                return Properties.Settings.Default.TrackedFiles;
            }
        }

        /// <summary>
        /// Supplies the Ribbon customization (the Worksheets toggle button on the View tab).
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon();
            return ribbon;
        }

        /// <summary>
        /// Wires up Excel events. Panes are created per window on demand.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                UpgradeSettingsIfNeeded();

                resizeSaveTimer.Tick += new EventHandler(ResizeSaveTimer_Tick);

                this.Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
                this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
                this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
                this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);

                InitialCheckOnLoad();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"SheetNavigator failed to initialize components.\n\nError Details: {ex.Message}",
                                "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Carries settings forward after an Office update. The settings folder is named after
        /// the Excel build, so a new build would otherwise start empty.
        /// </summary>
        private void UpgradeSettingsIfNeeded()
        {
            try
            {
                if (!Properties.Settings.Default.UpgradeRequired) return;

                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeRequired = false;
                Properties.Settings.Default.Save();
            }
            catch { /* Nothing to carry over */ }
        }

        /// <summary>
        /// Default pane width, scaled from 150 on a 1920-wide screen and kept within 120-400.
        /// </summary>
        private int EnsureResponsiveDefaultWidth()
        {
            if (cachedDefaultWidth.HasValue) return cachedDefaultWidth.Value;

            int defaultWidth = 150;

            try
            {
                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                double targetScalePercentage = defaultWidth / 1920.0;
                int calculatedWidth = (int)Math.Round(screenWidth * targetScalePercentage);

                cachedDefaultWidth = Math.Max(120, Math.Min(400, calculatedWidth));
            }
            catch
            {
                cachedDefaultWidth = defaultWidth;
            }

            return cachedDefaultWidth.Value;
        }

        /// <summary>
        /// Restores the pane for the workbook that is open when Excel starts.
        /// </summary>
        private void InitialCheckOnLoad()
        {
            try
            {
                RestoreSidebar(this.Application.ActiveWorkbook, this.Application.ActiveWindow);
            }
            catch { /* Excel may not be ready yet */ }
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            PruneDeadPanes();
            RestoreSidebar(workbook, window);
            ribbon?.RefreshToggleState();
        }

        /// <summary>
        /// Flags the workbook's panes as closing. Excel asks about unsaved changes after this event,
        /// so the close may still be cancelled; <see cref="PruneDeadPanes"/> removes the panes once
        /// the window is really gone.
        /// </summary>
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    PaneEntry entry = FindPane(window);
                    if (entry != null) entry.IsClosing = true;
                }
            }
            catch { /* Windows may already be gone */ }
        }

        /// <summary>
        /// Records a visible pane once a new workbook has been saved and has a path.
        /// </summary>
        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            if (!success) return;

            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    PaneEntry entry = FindPane(window);
                    if (entry != null && entry.Pane.Visible)
                    {
                        SaveTrackedWidth(workbook, entry.Pane);
                        break;
                    }
                }
            }
            catch { /* Nothing to record */ }
        }

        /// <summary>
        /// Keeps the list's highlight on the active sheet.
        /// </summary>
        private void Application_SheetActivate(object sheet)
        {
            try
            {
                PaneEntry entry = FindPane(this.Application.ActiveWindow);
                if (entry == null) return;

                // Working in the window again means any pending close was cancelled
                entry.IsClosing = false;

                if (entry.Pane.Visible && this.Application.ActiveWorkbook != null)
                {
                    entry.Control.RefreshWorksheets(this.Application.ActiveWorkbook);
                }
            }
            catch { /* Chart sheets and templates can refuse the refresh */ }
        }

        /// <summary>
        /// Shows or hides the active window's pane as a user action. Called by the Ribbon button.
        /// Tracking is updated by <see cref="MyCustomTaskPane_VisibleChanged"/>.
        /// </summary>
        public void ToggleSidebar()
        {
            try
            {
                PruneDeadPanes();

                Excel.Window window = this.Application.ActiveWindow;
                if (window == null || this.Application.ActiveWorkbook == null) return;

                PaneEntry entry = GetOrCreatePane(window);
                entry.Pane.Visible = !entry.Pane.Visible;
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// The first time a window is seen, shows its pane at the saved width if the workbook is tracked.
        /// After that the pane keeps whatever state the user left it in. Never writes settings.
        /// </summary>
        private void RestoreSidebar(Excel.Workbook workbook, Excel.Window window)
        {
            if (workbook == null || window == null) return;

            try
            {
                if (FindPane(window) != null) return;
                if (!HasPath(workbook)) return;
                if (!TryGetTrackedWidth(workbook.FullName, out int savedWidth)) return;

                PaneEntry entry = GetOrCreatePane(window);
                entry.Pane.Width = savedWidth;
                entry.Control.RefreshWorksheets(workbook);
                SetPaneVisibleSilently(entry.Pane, true);
            }
            catch { /* Leave the pane hidden */ }
        }

        /// <summary>
        /// Syncs the Ribbon button, then records user actions: showing the pane tracks the workbook
        /// at the current width, hiding it untracks the workbook. Changes made by the add-in itself are ignored.
        /// </summary>
        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            ribbon?.RefreshToggleState();

            if (isProgrammaticVisibilityChange || !(sender is CustomTaskPane pane)) return;

            PaneEntry entry = FindEntry(pane);
            if (entry == null || entry.IsClosing) return;

            try
            {
                Excel.Workbook workbook = WorkbookOf(pane);
                if (workbook == null) return;

                if (pane.Visible)
                {
                    SaveTrackedWidth(workbook, pane);
                    entry.Control.RefreshWorksheets(workbook);
                }
                else
                {
                    Untrack(workbook);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not update tracking preferences.\n\nDetails: {ex.Message}",
                                "Save Failure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Restarts the save delay on every resize, so a drag is written once when it ends.
        /// </summary>
        private void SheetNavigatorUI_Resize(object sender, EventArgs e)
        {
            PaneEntry entry = FindEntry(sender as SheetNavigatorControl);
            if (entry == null || entry.IsClosing) return;

            try
            {
                if (!entry.Pane.Visible) return;
            }
            catch { return; }

            pendingResizeEntry = entry;
            resizeSaveTimer.Stop();
            resizeSaveTimer.Start();
        }

        /// <summary>
        /// Saves the resized pane's width once resizing has settled.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            resizeSaveTimer.Stop();

            PaneEntry entry = pendingResizeEntry;
            pendingResizeEntry = null;
            if (entry == null || entry.IsClosing) return;

            try
            {
                if (!entry.Pane.Visible) return;

                Excel.Workbook workbook = WorkbookOf(entry.Pane);
                if (workbook != null) SaveTrackedWidth(workbook, entry.Pane);
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// The workbook shown in the window a pane belongs to.
        /// </summary>
        private Excel.Workbook WorkbookOf(CustomTaskPane pane)
        {
            Excel.Window window = pane.Window as Excel.Window;
            return (window?.Parent as Excel.Workbook) ?? this.Application.ActiveWorkbook;
        }

        /// <summary>
        /// The pane already created for a window, or null.
        /// </summary>
        private PaneEntry FindPane(Excel.Window window)
        {
            if (window == null) return null;
            return panes.TryGetValue(window.Hwnd, out PaneEntry entry) ? entry : null;
        }

        private PaneEntry FindEntry(CustomTaskPane pane)
        {
            if (pane == null) return null;
            foreach (PaneEntry entry in panes.Values)
            {
                if (ReferenceEquals(entry.Pane, pane)) return entry;
            }
            return null;
        }

        private PaneEntry FindEntry(SheetNavigatorControl control)
        {
            if (control == null) return null;
            foreach (PaneEntry entry in panes.Values)
            {
                if (ReferenceEquals(entry.Control, control)) return entry;
            }
            return null;
        }

        /// <summary>
        /// The pane for a window, created and wired on first use.
        /// </summary>
        private PaneEntry GetOrCreatePane(Excel.Window window)
        {
            int hwnd = window.Hwnd;
            if (panes.TryGetValue(hwnd, out PaneEntry existing)) return existing;

            SheetNavigatorControl control = new SheetNavigatorControl
            {
                Workbook = window.Parent as Excel.Workbook
            };

            CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Worksheets", window);
            try
            {
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                pane.Width = EnsureResponsiveDefaultWidth();

                // Width is meaningless when docked top or bottom, so keep the pane on a side. Excel's
                // "NoHorizontal" is the restriction compatible with a left-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { pane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch { }

                pane.VisibleChanged += new EventHandler(MyCustomTaskPane_VisibleChanged);
                control.Resize += new EventHandler(SheetNavigatorUI_Resize);
            }
            catch
            {
                // Don't leave a half-configured pane behind
                try { this.CustomTaskPanes.Remove(pane); } catch { }
                throw;
            }

            PaneEntry entry = new PaneEntry(hwnd, pane, control);
            panes[hwnd] = entry;
            return entry;
        }

        /// <summary>
        /// Removes panes whose window no longer exists and clears the closing flag on windows that survived a cancelled close.
        /// </summary>
        private void PruneDeadPanes()
        {
            HashSet<int> liveHwnds = new HashSet<int>();
            try
            {
                foreach (Excel.Window window in this.Application.Windows)
                {
                    liveHwnds.Add(window.Hwnd);
                }
            }
            catch { return; /* Try again on the next event */ }

            List<int> deadHwnds = new List<int>();
            foreach (int hwnd in panes.Keys)
            {
                if (!liveHwnds.Contains(hwnd)) deadHwnds.Add(hwnd);
            }

            foreach (int hwnd in deadHwnds) RemovePane(hwnd);

            foreach (PaneEntry entry in panes.Values) entry.IsClosing = false;
        }

        /// <summary>
        /// Drops a window's pane. Handlers are unhooked first so the removal is not seen as a user hide.
        /// </summary>
        private void RemovePane(int hwnd)
        {
            if (!panes.TryGetValue(hwnd, out PaneEntry entry)) return;
            panes.Remove(hwnd);
            if (ReferenceEquals(pendingResizeEntry, entry)) pendingResizeEntry = null;

            try { entry.Pane.VisibleChanged -= MyCustomTaskPane_VisibleChanged; } catch { }
            try { entry.Control.Resize -= SheetNavigatorUI_Resize; } catch { }
            try { this.CustomTaskPanes.Remove(entry.Pane); } catch { /* Already disposed with its window */ }
        }

        /// <summary>
        /// Changes a pane's visibility without treating it as a user action.
        /// </summary>
        private void SetPaneVisibleSilently(CustomTaskPane pane, bool visible)
        {
            if (pane == null) return;

            isProgrammaticVisibilityChange = true;
            try
            {
                pane.Visible = visible;
            }
            catch { /* Window may be closing */ }
            finally
            {
                isProgrammaticVisibilityChange = false;
            }
        }

        /// <summary>
        /// False for a workbook that has never been saved.
        /// </summary>
        private static bool HasPath(Excel.Workbook workbook)
        {
            return workbook != null && !string.IsNullOrEmpty(workbook.Path);
        }

        private static int ClampPaneWidth(int width)
        {
            return Math.Max(MinPaneWidth, Math.Min(MaxPaneWidth, width));
        }

        /// <summary>
        /// Records a pane's width for its workbook, replacing any older entry.
        /// Skips unsaved workbooks and widths already on file.
        /// </summary>
        private void SaveTrackedWidth(Excel.Workbook workbook, CustomTaskPane pane)
        {
            if (!HasPath(workbook)) return;

            string path = workbook.FullName;
            int width = ClampPaneWidth(pane.Width);
            if (TryGetTrackedWidth(path, out int savedWidth) && savedWidth == width) return;

            RemoveTrackedEntries(path);
            TrackedFiles.Add($"{path}|{width}");
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Forgets a workbook so its pane stays hidden next time.
        /// </summary>
        private void Untrack(Excel.Workbook workbook)
        {
            if (!HasPath(workbook)) return;

            if (RemoveTrackedEntries(workbook.FullName) > 0)
            {
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Looks up the saved width for a workbook path. Returns false if the path is not tracked.
        /// </summary>
        private bool TryGetTrackedWidth(string path, out int width)
        {
            width = EnsureResponsiveDefaultWidth();

            foreach (string entry in TrackedFiles)
            {
                if (!IsEntryForPath(entry, path)) continue;

                int separator = entry.LastIndexOf('|');
                if (separator > 0 && int.TryParse(entry.Substring(separator + 1), out int parsedWidth))
                {
                    width = ClampPaneWidth(parsedWidth);
                }
                return true;
            }

            return false;
        }

        /// <summary>
        /// Removes every entry for a workbook path and returns how many were removed.
        /// </summary>
        private int RemoveTrackedEntries(string path)
        {
            int removed = 0;
            StringCollection tracked = TrackedFiles;
            for (int i = tracked.Count - 1; i >= 0; i--)
            {
                if (IsEntryForPath(tracked[i], path))
                {
                    tracked.RemoveAt(i);
                    removed++;
                }
            }
            return removed;
        }

        private static bool IsEntryForPath(string entry, string path)
        {
            return entry != null && entry.StartsWith(path + "|", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Unhooks Excel events. The VSTO runtime has already disposed the panes by now.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            resizeSaveTimer.Stop();
            resizeSaveTimer.Dispose();

            this.Application.WindowActivate -= Application_WindowActivate;
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
            this.Application.SheetActivate -= Application_SheetActivate;

            panes.Clear();
        }

        #region VSTO Generated Code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        /// <summary>
        /// A task pane and its list control for one Excel window.
        /// </summary>
        private sealed class PaneEntry
        {
            public int Hwnd { get; }
            public CustomTaskPane Pane { get; }
            public SheetNavigatorControl Control { get; }

            /// <summary>
            /// Set while the workbook is closing, so a hide raised by the teardown is not treated as a user action.
            /// </summary>
            public bool IsClosing { get; set; }

            public PaneEntry(int hwnd, CustomTaskPane pane, SheetNavigatorControl control)
            {
                Hwnd = hwnd;
                Pane = pane;
                Control = control;
            }
        }
    }
}
