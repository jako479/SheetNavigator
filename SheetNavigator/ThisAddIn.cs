using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
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
        /// <summary>Pane width in points for a workbook with no saved width. Points already scale with DPI.</summary>
        private const int DefaultPaneWidth = 150;

        /// <summary>
        /// Ceiling for saved widths, on both save and load, to reject a garbage value in the settings file.
        /// There is no floor: Excel enforces its own minimum whenever a width is set.
        /// </summary>
        private const int MaxPaneWidth = 400;

        /// <summary>One pane per Excel window, keyed by window handle.</summary>
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();
        private Ribbon ribbon;

        /// <summary>Delays the width save until the user has stopped dragging the pane border.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };
        private PaneEntry pendingResizeEntry;

        /// <summary>
        /// Excel raises no event when a sheet tab is dragged to a new position, so visible panes
        /// re-check the sheet list once a second. Only a changed list triggers a rebuild.
        /// </summary>
        private readonly Timer refreshTimer = new Timer { Interval = 1000 };

        /// <summary>
        /// True while the add-in sets a pane's <c>Visible</c> itself, so
        /// <see cref="MyCustomTaskPane_VisibleChanged"/> only reacts to user actions.
        /// </summary>
        private bool isProgrammaticVisibilityChange = false;

        /// <summary>
        /// True when the pane of the given window (or the active window) is shown.
        /// </summary>
        public bool IsSidebarVisibleIn(Excel.Window window)
        {
            try
            {
                PaneEntry entry = FindPane(window ?? this.Application.ActiveWindow);
                return entry != null && entry.Pane.Visible;
            }
            catch { return false; }
        }

        /// <summary>
        /// Saved "path|width" entries, one per tracked workbook. Created on first use.
        /// </summary>
        private StringCollection TrackedFiles
        {
            get
            {
                return ReadSetting(() =>
                {
                    if (Properties.Settings.Default.TrackedFiles == null)
                    {
                        Properties.Settings.Default.TrackedFiles = new StringCollection();
                    }
                    return Properties.Settings.Default.TrackedFiles;
                });
            }
        }

        /// <summary>
        /// Reads a setting, resetting the settings file first if .NET reports it unreadable.
        /// </summary>
        private static T ReadSetting<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (ConfigurationErrorsException ex)
            {
                RecoverSettings(ex);
                return read();
            }
        }

        /// <summary>
        /// Saves the settings, resetting the settings file first if .NET reports it unreadable.
        /// </summary>
        private static void SaveSettings()
        {
            try
            {
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationErrorsException ex)
            {
                RecoverSettings(ex);
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// A user.config left truncated by a crash makes every settings call throw until it is deleted.
        /// Deletes it and reloads the defaults; the tracked list starts over.
        /// </summary>
        private static void RecoverSettings(ConfigurationErrorsException ex)
        {
            string file = ex.Filename ?? (ex.InnerException as ConfigurationErrorsException)?.Filename;
            Diagnostics.Write("Settings file unreadable, resetting it: " + (file ?? "(unknown path)") + " | " + ex.Message);

            try
            {
                if (!string.IsNullOrEmpty(file) && File.Exists(file)) File.Delete(file);
            }
            catch { }

            Properties.Settings.Default.Reload();
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
                Diagnostics.HookUnhandledExceptions();
                Diagnostics.Write("Startup");

                UpgradeSettingsIfNeeded();

                resizeSaveTimer.Tick += new EventHandler(ResizeSaveTimer_Tick);
                refreshTimer.Tick += new EventHandler(RefreshTimer_Tick);
                refreshTimer.Start();

                this.Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
                this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
                this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
                this.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Application_SheetActivate);

                InitialCheckOnLoad();
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Startup failed: " + ex);
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
                if (!ReadSetting(() => Properties.Settings.Default.UpgradeRequired)) return;

                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeRequired = false;
                SaveSettings();
                Diagnostics.Write("Settings upgraded from a previous Excel build");
            }
            catch { /* Nothing to carry over */ }
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
            try
            {
                PruneDeadPanes();
                RestoreSidebar(workbook, window);
                ribbon?.RefreshToggleState();
            }
            catch (Exception ex) { Diagnostics.Write("WindowActivate failed: " + ex); }
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
        /// Keeps every pane of the active workbook highlighting its own window's active sheet.
        /// The active window's pane also picks up any sheet list changes.
        /// </summary>
        private void Application_SheetActivate(object sheet)
        {
            try
            {
                Excel.Workbook activeWorkbook = this.Application.ActiveWorkbook;
                PaneEntry activeEntry = FindPane(this.Application.ActiveWindow);
                if (activeWorkbook == null) return;

                foreach (PaneEntry entry in panes.Values)
                {
                    // Working in the window again means any pending close was cancelled
                    if (ReferenceEquals(entry, activeEntry)) entry.IsClosing = false;

                    if (!IsAlive(entry) || !entry.Pane.Visible || !SameWorkbook(entry.Control.Workbook, activeWorkbook)) continue;

                    if (ReferenceEquals(entry, activeEntry))
                    {
                        entry.Control.RefreshWorksheets(activeWorkbook);
                    }
                    else
                    {
                        entry.Control.HighlightActiveSheet();
                    }
                }
            }
            catch (Exception ex) { Diagnostics.Write("SheetActivate failed: " + ex); }
        }

        /// <summary>
        /// Shows or hides a window's pane as a user action. Called by the Ribbon button with the
        /// state the button now shows. Tracking is updated by <see cref="MyCustomTaskPane_VisibleChanged"/>.
        /// </summary>
        public void SetSidebarVisible(Excel.Window window, bool visible)
        {
            try
            {
                PruneDeadPanes();

                window = window ?? this.Application.ActiveWindow;
                if (window == null || this.Application.ActiveWorkbook == null) return;

                PaneEntry entry = GetOrCreatePane(window);
                if (entry.Pane.Visible != visible) entry.Pane.Visible = visible;
            }
            catch (Exception ex) { Diagnostics.Write("SetSidebarVisible failed: " + ex); }
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
                Diagnostics.Write($"Restored pane for {workbook.Name} (window {entry.Hwnd}, width {savedWidth})");
            }
            catch (Exception ex) { Diagnostics.Write("RestoreSidebar failed: " + ex); }
        }

        /// <summary>
        /// Syncs the Ribbon button, then records user actions: showing the pane tracks the workbook
        /// at the current width, hiding it untracks the workbook. Changes made by the add-in itself are ignored.
        /// </summary>
        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                ribbon?.RefreshToggleState();

                if (isProgrammaticVisibilityChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || entry.IsClosing) return;

                Excel.Workbook workbook = entry.Control.Workbook ?? this.Application.ActiveWorkbook;
                if (workbook == null) return;

                if (pane.Visible)
                {
                    Diagnostics.Write($"User showed pane for {workbook.Name} (window {entry.Hwnd})");
                    SaveTrackedWidth(workbook, pane);
                    entry.Control.RefreshWorksheets(workbook);
                }
                else if (!IsWindowOpen(entry.Hwnd))
                {
                    // A closing window reports its pane as hidden; that is not the user's doing
                    Diagnostics.Write($"Ignored hide from closing window {entry.Hwnd} ({workbook.Name})");
                }
                else
                {
                    Diagnostics.Write($"User hid pane for {workbook.Name} (window {entry.Hwnd})");
                    Untrack(workbook);
                }
            }
            catch (Exception ex)
            {
                Diagnostics.Write("VisibleChanged failed: " + ex);
                MessageBox.Show($"Could not update tracking preferences.\n\nDetails: {ex.Message}",
                                "Save Failure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Restarts the save delay on every resize, so a drag is written once when it ends.
        /// </summary>
        private void SheetNavigatorUI_Resize(object sender, EventArgs e)
        {
            try
            {
                PaneEntry entry = FindEntry(sender as SheetNavigatorControl);
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible) return;

                pendingResizeEntry = entry;
                resizeSaveTimer.Stop();
                resizeSaveTimer.Start();
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Saves the resized pane's width once resizing has settled.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();

                PaneEntry entry = pendingResizeEntry;
                pendingResizeEntry = null;
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible) return;

                Excel.Workbook workbook = entry.Control.Workbook;
                if (workbook != null) SaveTrackedWidth(workbook, entry.Pane);
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Once a second, lets every visible pane catch sheet reorders that raise no Excel event.
        /// </summary>
        private void RefreshTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (panes.Count == 0 || this.Application.Ready == false) return;

                foreach (PaneEntry entry in panes.Values)
                {
                    if (entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible) continue;
                    entry.Control.RefreshQuietly();
                }
            }
            catch { /* Excel is busy; try again next tick */ }
        }

        /// <summary>
        /// True while Excel still lists a window with this handle.
        /// </summary>
        private bool IsWindowOpen(int hwnd)
        {
            try
            {
                foreach (Excel.Window window in this.Application.Windows)
                {
                    if (window.Hwnd == hwnd) return true;
                }
            }
            catch { }
            return false;
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
        /// False once Excel or VSTO has disposed the pane, for example because its window closed.
        /// </summary>
        private static bool IsAlive(PaneEntry entry)
        {
            try
            {
                bool unused = entry.Pane.Visible;
                return !entry.Control.IsDisposed;
            }
            catch { return false; }
        }

        private static bool SameWorkbook(Excel.Workbook a, Excel.Workbook b)
        {
            try
            {
                return a != null && b != null && string.Equals(a.FullName, b.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
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
                Window = window,
                Workbook = window.Parent as Excel.Workbook
            };

            CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Worksheets", window);
            try
            {
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                pane.Width = DefaultPaneWidth;

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
            Diagnostics.Write($"Created pane for {control.Workbook?.Name} (window {hwnd})");
            return entry;
        }

        /// <summary>
        /// Removes panes whose window no longer exists or that Excel has already disposed,
        /// and clears the closing flag on windows that survived a cancelled close.
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
            foreach (PaneEntry entry in panes.Values)
            {
                if (!liveHwnds.Contains(entry.Hwnd) || !IsAlive(entry)) deadHwnds.Add(entry.Hwnd);
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
            Diagnostics.Write($"Removed pane for window {hwnd}");
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
            if (width <= 0) return DefaultPaneWidth;
            return Math.Min(MaxPaneWidth, width);
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
            SaveSettings();
            Diagnostics.Write($"Tracked {path} at width {width}");
        }

        /// <summary>
        /// Forgets a workbook so its pane stays hidden next time.
        /// </summary>
        private void Untrack(Excel.Workbook workbook)
        {
            if (!HasPath(workbook)) return;

            if (RemoveTrackedEntries(workbook.FullName) > 0)
            {
                SaveSettings();
                Diagnostics.Write($"Untracked {workbook.FullName}");
            }
        }

        /// <summary>
        /// Looks up the saved width for a workbook path. Returns false if the path is not tracked.
        /// </summary>
        private bool TryGetTrackedWidth(string path, out int width)
        {
            width = DefaultPaneWidth;

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
            refreshTimer.Stop();
            refreshTimer.Dispose();

            this.Application.WindowActivate -= Application_WindowActivate;
            this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
            this.Application.SheetActivate -= Application_SheetActivate;

            panes.Clear();
            Diagnostics.Write("Shutdown");
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
