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
    /// Each Excel window gets its own pane. Each workbook file keeps an entry in the user's settings
    /// saying whether its pane is shown, where it is docked, how wide it is and, while floating, how
    /// tall. See DESIGN.md.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>Docked side used when the saved default side cannot be read.</summary>
        private const Office.MsoCTPDockPosition FallbackDockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

        /// <summary>Pane width in points used when the saved default width cannot be read. Points already scale with DPI.</summary>
        private const int FallbackWidth = 150;

        /// <summary>
        /// Floating pane height in points used when the saved default height cannot be read. Only a
        /// floating pane has a height of its own; a docked pane is stretched to the window.
        /// </summary>
        private const int FallbackHeight = 400;

        /// <summary>
        /// Ceiling for saved widths, on both save and load, to reject a garbage value in the settings file.
        /// There is no floor: Excel enforces its own minimum whenever a width is set.
        /// </summary>
        private const int MaxPaneWidth = 400;

        /// <summary>One pane per Excel window, keyed by window handle.</summary>
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();
        private Ribbon ribbon;

        /// <summary>Delays the save until the user has stopped dragging the pane border.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };
        private PaneEntry pendingResizeEntry;

        /// <summary>
        /// Excel raises no event when a sheet tab is dragged to a new position, so visible panes
        /// re-check the sheet list once a second. Only a changed list triggers a rebuild.
        /// </summary>
        private readonly Timer refreshTimer = new Timer { Interval = 1000 };

        /// <summary>
        /// True while the add-in shows or hides a pane itself, so the pane's change handlers only
        /// react to user actions.
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
        /// Saved "path|dock|width|height|visible" entries, one per workbook file. Created on first use.
        /// </summary>
        private StringCollection FileEntries
        {
            get
            {
                return ReadSetting(() =>
                {
                    if (Properties.Settings.Default.FileEntries == null)
                    {
                        Properties.Settings.Default.FileEntries = new StringCollection();
                    }
                    return Properties.Settings.Default.FileEntries;
                });
            }
        }

        /// <summary>
        /// The dock position the user last used, floating included. New panes, and workbooks whose
        /// entry's dock position cannot be read, open here.
        /// </summary>
        private Office.MsoCTPDockPosition DefaultDockPosition
        {
            get
            {
                return SavedDockPosition(ReadSetting(() => Properties.Settings.Default.DefaultDockPosition), FallbackDockPosition);
            }
        }

        /// <summary>
        /// The width the user last used. New panes, and workbooks whose entry's width cannot be
        /// read, open at this width.
        /// </summary>
        private int DefaultWidth
        {
            get
            {
                return ClampPaneWidth(ReadSetting(() => Properties.Settings.Default.DefaultWidth), FallbackWidth);
            }
        }

        /// <summary>
        /// The floating height the user last used. New panes that start floating, and workbooks whose
        /// entry's height cannot be read, open at this height.
        /// </summary>
        private int DefaultHeight
        {
            get
            {
                return PositiveHeight(ReadSetting(() => Properties.Settings.Default.DefaultHeight), FallbackHeight);
            }
        }

        /// <summary>
        /// Makes a pane's dock position, width and floating height the defaults for new panes.
        /// </summary>
        private void SaveDefaults(Office.MsoCTPDockPosition dock, int width, int height)
        {
            Properties.Settings.Default.DefaultDockPosition = DockPositionName(dock);
            Properties.Settings.Default.DefaultWidth = width;
            Properties.Settings.Default.DefaultHeight = height;
            SaveSettings();
            Diagnostics.Write($"Defaults are now {DockPositionName(dock)}, width {width}, height {height}");
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
        /// Deletes it and reloads the defaults; the file entries start over.
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
        /// After a save under a new path (first save or Save As), writes the workbook's pane state as that
        /// path's entry, replacing any entry the path already had; a shown pane also sets the defaults.
        /// A save under the same path writes nothing.
        /// </summary>
        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            if (!success) return;

            try
            {
                if (!HasPath(workbook)) return;
                string path = workbook.FullName;

                // Prefer a shown pane as the source; with several windows the shown one is what the user sees
                PaneEntry source = null;
                foreach (Excel.Window window in workbook.Windows)
                {
                    PaneEntry entry = FindPane(window);
                    if (entry == null || string.Equals(entry.RecordedPath, path, StringComparison.OrdinalIgnoreCase)) continue;

                    entry.RecordedPath = path;
                    if (source == null || (!source.Pane.Visible && entry.Pane.Visible)) source = entry;
                }
                if (source == null) return;

                if (source.Pane.Visible)
                {
                    WriteEntryFromPane(path, source);
                    SaveDefaults(source.RecordedDock, source.RecordedWidth, source.RecordedHeight);
                }
                else
                {
                    WriteEntry(path, source.Pane.DockPosition, PaneWidth(source.Pane), PaneHeight(source.Pane, source.RecordedHeight), false);
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
        /// state the button now shows. The file's entry is updated by <see cref="MyCustomTaskPane_VisibleChanged"/>.
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
        /// The first time a window is seen, shows its pane if the workbook's entry says visible; the
        /// pane is created from the entry's dock position and size. After that the pane keeps whatever
        /// state the user left it in. Never writes settings.
        /// </summary>
        private void RestoreSidebar(Excel.Workbook workbook, Excel.Window window)
        {
            if (workbook == null || window == null) return;

            try
            {
                if (FindPane(window) != null) return;
                if (!HasPath(workbook)) return;
                if (!TryGetEntry(workbook.FullName, out _, out _, out _, out bool visible) || !visible) return;

                PaneEntry entry = GetOrCreatePane(window);
                entry.Control.RefreshWorksheets(workbook);
                SetPaneVisibleSilently(entry.Pane, true);
                Diagnostics.Write($"Restored pane for {workbook.Name} (window {entry.Hwnd}, {DockPositionName(entry.RecordedDock)}, width {entry.RecordedWidth}, height {entry.RecordedHeight})");
            }
            catch (Exception ex) { Diagnostics.Write("RestoreSidebar failed: " + ex); }
        }

        /// <summary>
        /// Syncs the Ribbon button, then records user actions: showing the pane marks the file's entry
        /// visible, adding one from the pane if the file has none; hiding it marks the entry hidden.
        /// A workbook without a path has no entry. Changes made by the add-in itself are ignored.
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

                string path = HasPath(workbook) ? workbook.FullName : null;

                if (pane.Visible)
                {
                    Diagnostics.Write($"User showed pane for {workbook.Name} (window {entry.Hwnd})");
                    if (path != null)
                    {
                        if (HasEntry(path)) SetEntryVisible(path, true);
                        else WriteEntryFromPane(path, entry);
                    }
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
                    if (path != null && HasEntry(path)) SetEntryVisible(path, false);
                }
            }
            catch (Exception ex)
            {
                Diagnostics.Write("VisibleChanged failed: " + ex);
                MessageBox.Show($"Could not save the pane settings.\n\nDetails: {ex.Message}",
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
        /// Records the resized pane once resizing has settled.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();

                PaneEntry entry = pendingResizeEntry;
                pendingResizeEntry = null;
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible) return;

                RecordPaneChange(entry);
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// Records the pane when the user docks it left or right or floats it. A pane that just floated
        /// is given the recorded height first, since only a floating pane has a height of its own; Excel
        /// rejects property sets inside this handler, so that is queued to run right after it returns.
        /// Floating fires as the pane detaches; Excel may change the width at the same time.
        /// </summary>
        private void MyCustomTaskPane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || entry.IsClosing || !IsAlive(entry) || !pane.Visible) return;

                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    entry.Control.BeginInvoke(new Action(() => ApplyFloatingHeight(entry)));
                    return;
                }

                RecordPaneChange(entry);
            }
            catch (Exception ex) { Diagnostics.Write("DockPositionChanged failed: " + ex); }
        }

        /// <summary>
        /// Runs right after the float event's handler returns: gives the floating pane its recorded
        /// height, then records the pane.
        /// </summary>
        private void ApplyFloatingHeight(PaneEntry entry)
        {
            try
            {
                if (entry.IsClosing || !IsAlive(entry) || !entry.Pane.Visible) return;
                if (entry.Pane.DockPosition != Office.MsoCTPDockPosition.msoCTPDockPositionFloating) return;

                entry.Pane.Height = entry.RecordedHeight;
                RecordPaneChange(entry);
            }
            catch (Exception ex) { Diagnostics.Write("Floating height failed: " + ex); }
        }

        /// <summary>
        /// After the user docks, floats or resizes a pane: writes its dock position, width and, while
        /// floating, height to the file's entry if the file has one, and to the defaults. A layout event
        /// that changed nothing since the last record is ignored, and so is a pane docked top or bottom.
        /// A docked pane's height is Excel's, so the recorded height stays.
        /// </summary>
        private void RecordPaneChange(PaneEntry entry)
        {
            Office.MsoCTPDockPosition dock = entry.Pane.DockPosition;
            int width = PaneWidth(entry.Pane);
            int height = PaneHeight(entry.Pane, entry.RecordedHeight);
            if (dock == entry.RecordedDock && width == entry.RecordedWidth && height == entry.RecordedHeight) return;

            // Top and bottom are only reachable if Excel rejected the dock restriction; they have no saved form
            if (dock == Office.MsoCTPDockPosition.msoCTPDockPositionTop || dock == Office.MsoCTPDockPosition.msoCTPDockPositionBottom) return;

            entry.RecordedDock = dock;
            entry.RecordedWidth = width;
            entry.RecordedHeight = height;

            Excel.Workbook workbook = entry.Control.Workbook;
            if (HasPath(workbook) && HasEntry(workbook.FullName)) WriteEntry(workbook.FullName, dock, width, height, true);
            SaveDefaults(dock, width, height);
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

            // A file with an entry gets its pane where the entry says; any other pane starts at the defaults
            Office.MsoCTPDockPosition dock = DefaultDockPosition;
            int width = DefaultWidth;
            int height = DefaultHeight;
            if (HasPath(control.Workbook)) TryGetEntry(control.Workbook.FullName, out dock, out width, out height, out _);

            CustomTaskPane pane = this.CustomTaskPanes.Add(control, "Worksheets", window);
            try
            {
                // Dock position before size: the pane API expects that order. Only a floating pane takes a height.
                pane.DockPosition = dock;
                pane.Width = width;
                if (dock == Office.MsoCTPDockPosition.msoCTPDockPositionFloating) pane.Height = height;

                // Width is meaningless when docked top or bottom, so keep the pane on a side or floating. Excel's
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

            PaneEntry entry = new PaneEntry(hwnd, pane, control)
            {
                RecordedPath = HasPath(control.Workbook) ? control.Workbook.FullName : null,
                RecordedDock = pane.DockPosition,
                RecordedWidth = PaneWidth(pane),
                RecordedHeight = PaneHeight(pane, height)
            };
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
        /// The pane's width, within the saved range.
        /// </summary>
        private int PaneWidth(CustomTaskPane pane)
        {
            return ClampPaneWidth(pane.Width, DefaultWidth);
        }

        /// <summary>
        /// A height of zero or less (a garbage value) is the fallback. There is no ceiling: Excel keeps
        /// a floating pane on screen.
        /// </summary>
        private static int PositiveHeight(int height, int fallback)
        {
            return height <= 0 ? fallback : height;
        }

        /// <summary>
        /// The pane's height while floating; a docked pane's height is Excel's, so the fallback stands.
        /// </summary>
        private static int PaneHeight(CustomTaskPane pane, int fallback)
        {
            return pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PositiveHeight(pane.Height, fallback)
                : fallback;
        }

        /// <summary>
        /// The dock position as written in settings: "Left", "Right" or "Floating".
        /// </summary>
        private static string DockPositionName(Office.MsoCTPDockPosition dock)
        {
            switch (dock)
            {
                case Office.MsoCTPDockPosition.msoCTPDockPositionRight: return "Right";
                case Office.MsoCTPDockPosition.msoCTPDockPositionFloating: return "Floating";
                default: return "Left";
            }
        }

        /// <summary>
        /// Reads a saved dock position; anything but "Left", "Right" or "Floating" (a garbage value) is the fallback.
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
            if (string.Equals(value, "Floating", StringComparison.OrdinalIgnoreCase))
            {
                return Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            }
            return fallback;
        }

        /// <summary>
        /// True if the file has an entry, whatever it says.
        /// </summary>
        private bool HasEntry(string path)
        {
            foreach (string entry in FileEntries)
            {
                if (IsEntryForPath(entry, path)) return true;
            }
            return false;
        }

        /// <summary>
        /// Reads a file's entry. Returns false if the file has none. A value that cannot be read falls
        /// back to its default: the default dock position, width and height, hidden.
        /// </summary>
        private bool TryGetEntry(string path, out Office.MsoCTPDockPosition dock, out int width, out int height, out bool visible)
        {
            dock = DefaultDockPosition;
            width = DefaultWidth;
            height = DefaultHeight;
            visible = false;

            foreach (string entry in FileEntries)
            {
                if (!IsEntryForPath(entry, path)) continue;

                string[] values = entry.Substring(path.Length + 1).Split('|');
                if (values.Length > 0) dock = SavedDockPosition(values[0], dock);
                if (values.Length > 1 && int.TryParse(values[1], out int parsedWidth)) width = ClampPaneWidth(parsedWidth, width);
                if (values.Length > 2 && int.TryParse(values[2], out int parsedHeight)) height = PositiveHeight(parsedHeight, height);
                if (values.Length > 3 && bool.TryParse(values[3], out bool parsedVisible)) visible = parsedVisible;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Writes a file's entry, replacing any older one. Entries are never removed.
        /// </summary>
        private void WriteEntry(string path, Office.MsoCTPDockPosition dock, int width, int height, bool visible)
        {
            RemoveEntries(path);
            FileEntries.Add($"{path}|{DockPositionName(dock)}|{width}|{height}|{visible}");
            SaveSettings();
            Diagnostics.Write($"Entry for {path}: {DockPositionName(dock)}, width {width}, height {height}, visible {visible}");
        }

        /// <summary>
        /// Writes the pane's dock position, width and floating height as the file's entry, marked visible,
        /// and remembers them on the pane so a later layout event that changed nothing is not written again.
        /// </summary>
        private void WriteEntryFromPane(string path, PaneEntry entry)
        {
            entry.RecordedPath = path;
            entry.RecordedDock = entry.Pane.DockPosition;
            entry.RecordedWidth = PaneWidth(entry.Pane);
            entry.RecordedHeight = PaneHeight(entry.Pane, entry.RecordedHeight);
            WriteEntry(path, entry.RecordedDock, entry.RecordedWidth, entry.RecordedHeight, true);
        }

        /// <summary>
        /// Marks a file's entry visible or hidden; its dock position, width and height stay as they are.
        /// </summary>
        private void SetEntryVisible(string path, bool visible)
        {
            if (!TryGetEntry(path, out Office.MsoCTPDockPosition dock, out int width, out int height, out _)) return;
            WriteEntry(path, dock, width, height, visible);
        }

        /// <summary>
        /// Removes every entry for a workbook path and returns how many were removed.
        /// </summary>
        private int RemoveEntries(string path)
        {
            int removed = 0;
            StringCollection entries = FileEntries;
            for (int i = entries.Count - 1; i >= 0; i--)
            {
                if (IsEntryForPath(entries[i], path))
                {
                    entries.RemoveAt(i);
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

            /// <summary>
            /// The workbook path this pane's entry was last written under (the workbook's path at
            /// creation, null for an unsaved workbook). A save under another path is a first save or Save As.
            /// </summary>
            public string RecordedPath { get; set; }

            /// <summary>
            /// The dock position last recorded for this pane (set at creation, updated on every
            /// write), so a layout event that changed nothing is not written again.
            /// </summary>
            public Office.MsoCTPDockPosition RecordedDock { get; set; }

            /// <summary>
            /// The width last recorded for this pane; see <see cref="RecordedDock"/>.
            /// </summary>
            public int RecordedWidth { get; set; }

            /// <summary>
            /// The floating height last recorded for this pane; see <see cref="RecordedDock"/>. A docked
            /// pane's height is Excel's, so this is always the last floating height.
            /// </summary>
            public int RecordedHeight { get; set; }

            public PaneEntry(int hwnd, CustomTaskPane pane, SheetNavigatorControl control)
            {
                Hwnd = hwnd;
                Pane = pane;
                Control = control;
            }
        }
    }
}
