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
    /// Each Excel window gets its own pane. Which workbooks show the pane, and on which side and
    /// at what width, is remembered per user in the add-in's settings.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Docked side used when the saved default side cannot be read.</summary>
        private const Office.MsoCTPDockPosition FallbackDockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

        /// <summary>Pane width in points used when the saved default width cannot be read. Points already scale with DPI.</summary>
        private const int FallbackWidth = 150;

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
        /// True while the add-in sets a pane's <c>Visible</c>, <c>DockPosition</c> or <c>Width</c>
        /// itself, so the pane's change handlers only react to user actions.
        /// </summary>
        private bool isProgrammaticPaneChange = false;

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
        /// Saved "path|side|width" entries, one per tracked workbook. Created on first use.
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
        /// The side the user last docked a pane on. New panes, and tracked workbooks whose saved
        /// side cannot be read, open on this side.
        /// </summary>
        private Office.MsoCTPDockPosition DefaultDockPosition
        {
            get
            {
                return SavedDockPosition(ReadSetting(() => Properties.Settings.Default.DefaultDockPosition), FallbackDockPosition);
            }
        }

        /// <summary>
        /// The width the user last left a docked pane at. New panes, and tracked workbooks whose
        /// saved width cannot be read, open at this width.
        /// </summary>
        private int DefaultWidth
        {
            get
            {
                return ClampPaneWidth(ReadSetting(() => Properties.Settings.Default.DefaultWidth), FallbackWidth);
            }
        }

        /// <summary>
        /// Makes a docked pane's side and width the defaults for new panes, if they are not already.
        /// </summary>
        private void SaveDefaults(Office.MsoCTPDockPosition side, int width)
        {
            if (DefaultDockPosition == side && DefaultWidth == width) return;

            Properties.Settings.Default.DefaultDockPosition = DockPositionName(side);
            Properties.Settings.Default.DefaultWidth = width;
            SaveSettings();
            Diagnostics.Write($"Defaults are now {DockPositionName(side)}, width {width}");
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
                        SaveTrackedPane(workbook, entry.Pane);
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
        /// The first time a window is seen, shows its pane on the saved side at the saved width if the
        /// workbook is tracked. After that the pane keeps whatever state the user left it in. Never writes settings.
        /// </summary>
        private void RestoreSidebar(Excel.Workbook workbook, Excel.Window window)
        {
            if (workbook == null || window == null) return;

            try
            {
                if (FindPane(window) != null) return;
                if (!HasPath(workbook)) return;
                if (!TryGetTrackedPane(workbook.FullName, out Office.MsoCTPDockPosition savedSide, out int savedWidth)) return;

                PaneEntry entry = GetOrCreatePane(window);
                isProgrammaticPaneChange = true;
                try
                {
                    // Side before width: the pane API expects the dock position to be set first
                    entry.Pane.DockPosition = savedSide;
                    entry.Pane.Width = savedWidth;
                }
                finally
                {
                    isProgrammaticPaneChange = false;
                }
                entry.Control.RefreshWorksheets(workbook);
                SetPaneVisibleSilently(entry.Pane, true);
                Diagnostics.Write($"Restored pane for {workbook.Name} (window {entry.Hwnd}, {DockPositionName(savedSide)}, width {savedWidth})");
            }
            catch (Exception ex) { Diagnostics.Write("RestoreSidebar failed: " + ex); }
        }

        /// <summary>
        /// Syncs the Ribbon button, then records user actions: showing the pane tracks the workbook
        /// at the current side and width, hiding it untracks the workbook. Changes made by the add-in itself are ignored.
        /// </summary>
        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                ribbon?.RefreshToggleState();

                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || entry.IsClosing) return;

                Excel.Workbook workbook = entry.Control.Workbook ?? this.Application.ActiveWorkbook;
                if (workbook == null) return;

                if (pane.Visible)
                {
                    Diagnostics.Write($"User showed pane for {workbook.Name} (window {entry.Hwnd})");
                    SaveTrackedPane(workbook, pane);
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
        /// A floating pane is never saved.
        /// </summary>
        private void SheetNavigatorUI_Resize(object sender, EventArgs e)
        {
            try
            {
                PaneEntry entry = FindEntry(sender as SheetNavigatorControl);
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible || !IsDockedOnSide(entry.Pane)) return;

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
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible || !IsDockedOnSide(entry.Pane)) return;

                Excel.Workbook workbook = entry.Control.Workbook;
                if (workbook != null) SaveTrackedPane(workbook, entry.Pane);
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Saves the pane's new side when the user docks it left or right. Floating is not saved,
        /// so the pane comes back docked on its last saved side.
        /// </summary>
        private void MyCustomTaskPane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !pane.Visible || !IsDockedOnSide(pane)) return;

                Excel.Workbook workbook = entry.Control.Workbook;
                if (workbook != null) SaveTrackedPane(workbook, pane);
            }
            catch (Exception ex) { Diagnostics.Write("DockPositionChanged failed: " + ex); }
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
                pane.DockPosition = DefaultDockPosition;
                pane.Width = DefaultWidth;

                // Width is meaningless when docked top or bottom, so keep the pane on a side. Excel's
                // "NoHorizontal" is the restriction compatible with a side-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { pane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch { }

                pane.VisibleChanged += new EventHandler(MyCustomTaskPane_VisibleChanged);
                pane.DockPositionChanged += new EventHandler(MyCustomTaskPane_DockPositionChanged);
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
            try { entry.Pane.DockPositionChanged -= MyCustomTaskPane_DockPositionChanged; } catch { }
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

            isProgrammaticPaneChange = true;
            try
            {
                pane.Visible = visible;
            }
            catch { /* Window may be closing */ }
            finally
            {
                isProgrammaticPaneChange = false;
            }
        }

        /// <summary>
        /// False for a workbook that has never been saved.
        /// </summary>
        private static bool HasPath(Excel.Workbook workbook)
        {
            return workbook != null && !string.IsNullOrEmpty(workbook.Path);
        }

        /// <summary>
        /// Keeps a width within the saved range; a width of zero or less (a garbage value) is the fallback.
        /// </summary>
        private static int ClampPaneWidth(int width, int fallback)
        {
            if (width <= 0) return fallback;
            return Math.Min(MaxPaneWidth, width);
        }

        /// <summary>
        /// True while the pane is docked left or right, the only states that are saved.
        /// </summary>
        private static bool IsDockedOnSide(CustomTaskPane pane)
        {
            Office.MsoCTPDockPosition position = pane.DockPosition;
            return position == Office.MsoCTPDockPosition.msoCTPDockPositionLeft
                || position == Office.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        /// <summary>
        /// The side as written in settings: "Left" or "Right".
        /// </summary>
        private static string DockPositionName(Office.MsoCTPDockPosition side)
        {
            return side == Office.MsoCTPDockPosition.msoCTPDockPositionRight ? "Right" : "Left";
        }

        /// <summary>
        /// Reads a saved side; anything but "Left" or "Right" (a garbage value) is the fallback.
        /// </summary>
        private static Office.MsoCTPDockPosition SavedDockPosition(string value, Office.MsoCTPDockPosition fallback)
        {
            if (string.Equals(value, "Right", StringComparison.OrdinalIgnoreCase))
            {
                return Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
            if (string.Equals(value, "Left", StringComparison.OrdinalIgnoreCase))
            {
                return Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            }
            return fallback;
        }

        /// <summary>
        /// Records a pane's side and width for its workbook, replacing any older entry, and makes a
        /// docked pane's side and width the defaults for new panes. A floating pane keeps what is already
        /// saved (or the defaults), so only docked state is ever written. Skips unsaved workbooks and entries already on file.
        /// </summary>
        private void SaveTrackedPane(Excel.Workbook workbook, CustomTaskPane pane)
        {
            if (!HasPath(workbook)) return;

            bool docked = IsDockedOnSide(pane);
            if (docked) SaveDefaults(pane.DockPosition, ClampPaneWidth(pane.Width, DefaultWidth));

            string path = workbook.FullName;
            bool tracked = TryGetTrackedPane(path, out Office.MsoCTPDockPosition savedSide, out int savedWidth);

            Office.MsoCTPDockPosition side = docked ? pane.DockPosition : savedSide;
            int width = docked ? ClampPaneWidth(pane.Width, DefaultWidth) : savedWidth;
            if (tracked && savedSide == side && savedWidth == width) return;

            RemoveTrackedEntries(path);
            TrackedFiles.Add($"{path}|{DockPositionName(side)}|{width}");
            SaveSettings();
            Diagnostics.Write($"Tracked {path} at {DockPositionName(side)}, width {width}");
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
        /// Looks up the saved side and width for a workbook path. Returns false if the path is not
        /// tracked. A value that cannot be read falls back to its default.
        /// </summary>
        private bool TryGetTrackedPane(string path, out Office.MsoCTPDockPosition side, out int width)
        {
            side = DefaultDockPosition;
            width = DefaultWidth;

            foreach (string entry in TrackedFiles)
            {
                if (!IsEntryForPath(entry, path)) continue;

                string[] values = entry.Substring(path.Length + 1).Split('|');
                if (values.Length > 0) side = SavedDockPosition(values[0], side);
                if (values.Length > 1 && int.TryParse(values[1], out int parsedWidth)) width = ClampPaneWidth(parsedWidth, width);
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
