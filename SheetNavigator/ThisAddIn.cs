using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Xml;
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

        /// <summary>
        /// How many half-second checks a pane change waits while its workbook's close is pending. A
        /// window still open after that many survived a cancelled close, so the change is the user's.
        /// </summary>
        private const int ClosingRecheckLimit = 4;

        /// <summary>One pane per Excel window, keyed by window handle.</summary>
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();

        /// <summary>
        /// Each window's workbook path as last seen, keyed by window handle. A file whose entry says
        /// hidden never gets a pane, so after a Save As this is the only way to find the entry to copy.
        /// </summary>
        private readonly Dictionary<int, string> lastKnownPaths = new Dictionary<int, string>();

        private Ribbon ribbon;

        /// <summary>Delays recording a resize, dock or float until the user has stopped dragging.</summary>
        private readonly Timer resizeSaveTimer = new Timer { Interval = 500 };
        private readonly List<PaneEntry> pendingResizeEntries = new List<PaneEntry>();

        /// <summary>
        /// Delays recording a hide until a closing window has had time to go: such a window reports
        /// its pane hidden too, sometimes while Excel still lists the window.
        /// </summary>
        private readonly Timer hideSaveTimer = new Timer { Interval = 500 };
        private readonly List<PaneEntry> pendingHideEntries = new List<PaneEntry>();

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
        public bool IsPaneVisibleIn(Excel.Window window)
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
        /// Changes the settings in memory only; the caller saves once it has written everything.
        /// </summary>
        private void WriteDefaults(Office.MsoCTPDockPosition dock, int width, int height)
        {
            string dockName = DockPositionName(dock);
            WriteSetting(() =>
            {
                Properties.Settings.Default.DefaultDockPosition = dockName;
                Properties.Settings.Default.DefaultWidth = width;
                Properties.Settings.Default.DefaultHeight = height;
            });
        }

        /// <summary>
        /// Reads a setting, resetting the settings file first if it is corrupt.
        /// </summary>
        private static T ReadSetting<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!TryRecoverSettings(ex)) throw;
                return read();
            }
        }

        /// <summary>
        /// Changes a setting in memory, resetting the settings file first if it is corrupt.
        /// </summary>
        private static void WriteSetting(Action write)
        {
            try
            {
                write();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!TryRecoverSettings(ex)) throw;
                write();
            }
        }

        /// <summary>
        /// Saves every settings change made so far, once per user action; a corrupt settings file is reset first.
        /// </summary>
        private static void SaveSettings()
        {
            try
            {
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationErrorsException ex)
            {
                if (!TryRecoverSettings(ex)) throw;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Deletes a corrupt user.config (a crash can leave it truncated) and reloads the defaults, so
        /// the file entries start over. Returns false, keeping the file, for any other failure.
        /// </summary>
        private static bool TryRecoverSettings(ConfigurationErrorsException ex)
        {
            string file = ex.Filename ?? (ex.InnerException as ConfigurationErrorsException)?.Filename;

            // A locked or unreadable file is not corrupt; deleting it would throw the user's settings away
            if (!IsCorruptUserConfig(ex, file))
            {
                Diagnostics.Write("Settings unavailable, file kept: " + (file ?? "(unknown path)") + " | " + ex);
                return false;
            }

            Diagnostics.Write("Settings file corrupt, resetting it: " + file + " | " + ex.Message);
            try
            {
                File.Delete(file);
            }
            catch (Exception deleteException)
            {
                Diagnostics.Write("Could not delete the settings file: " + deleteException);
                return false;
            }

            Properties.Settings.Default.Reload();
            return true;
        }

        /// <summary>
        /// True only for an existing file under the user's profile that .NET could not parse as XML.
        /// </summary>
        private static bool IsCorruptUserConfig(Exception ex, string file)
        {
            if (string.IsNullOrEmpty(file) || !File.Exists(file) || !IsUnderUserProfile(file)) return false;

            for (Exception inner = ex; inner != null; inner = inner.InnerException)
            {
                if (inner is XmlException) return true;
            }
            return false;
        }

        /// <summary>
        /// True if the file lives under the user's local or roaming application data, where user.config lives.
        /// </summary>
        private static bool IsUnderUserProfile(string file)
        {
            string fullPath = Path.GetFullPath(file);
            foreach (Environment.SpecialFolder folder in new[] { Environment.SpecialFolder.LocalApplicationData, Environment.SpecialFolder.ApplicationData })
            {
                string root = Environment.GetFolderPath(folder);
                if (string.IsNullOrEmpty(root)) continue;

                root = root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
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

                UpgradeSettingsIfNeeded();

                resizeSaveTimer.Tick += new EventHandler(ResizeSaveTimer_Tick);
                hideSaveTimer.Tick += new EventHandler(HideSaveTimer_Tick);
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
                MessageBox.Show(ExcelOwner(), $"SheetNavigator failed to initialize components.\n\nError Details: {ex.Message}",
                                "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Carries settings forward after an Office update, since the settings folder is named after the Excel build.
        /// </summary>
        private void UpgradeSettingsIfNeeded()
        {
            try
            {
                if (!ReadSetting(() => Properties.Settings.Default.UpgradeRequired)) return;

                Properties.Settings.Default.Upgrade();
                WriteSetting(() => Properties.Settings.Default.UpgradeRequired = false);
                SaveSettings();
            }
            catch (Exception ex)
            {
                // The flag stays set, so the next start tries again; without the log that retry would be invisible
                Diagnostics.Write("Settings upgrade failed: " + ex);
            }
        }

        /// <summary>
        /// Restores the pane for the workbook that is open when Excel starts.
        /// </summary>
        private void InitialCheckOnLoad()
        {
            try
            {
                RememberPath(this.Application.ActiveWindow, this.Application.ActiveWorkbook);
                RestorePane(this.Application.ActiveWorkbook, this.Application.ActiveWindow);
            }
            catch { /* Excel may not be ready yet */ }
        }

        /// <summary>
        /// Drops panes whose window is gone, restores the pane the first time a window is seen and syncs the Ribbon button.
        /// </summary>
        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                PruneDeadPanes();
                RememberPath(window, workbook);

                // Working in the window again means any pending close was cancelled
                PaneEntry entry = FindPane(window);
                if (entry != null) entry.IsClosing = false;

                RestorePane(workbook, window);
                ribbon?.RefreshToggleState();
            }
            catch (Exception ex) { Diagnostics.Write("WindowActivate failed: " + ex); }
        }

        /// <summary>
        /// Records anything the user changed just before the close, then flags the workbook's panes as
        /// closing so the hide raised by the teardown is not taken for the user's.
        /// </summary>
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                // A hide or resize still waiting for its delay happened before the close began, so it is the user's
                FlushPendingSaves();

                // Excel asks about unsaved changes after this event, so the close may still be cancelled; the
                // flag is cleared when the user works in the window again or when a change outlives the re-checks
                foreach (Excel.Window window in workbook.Windows)
                {
                    PaneEntry entry = FindPane(window);
                    if (entry == null) continue;

                    entry.IsClosing = true;
                    entry.ClosingRechecks = 0;
                }
            }
            catch (Exception ex) { Diagnostics.Write("BeforeClose failed: " + ex); }
        }

        /// <summary>
        /// After a first save or Save As, writes the pane's state as the new path's entry (a workbook
        /// with no pane passes its hidden entry on); a save under the same path writes nothing.
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
                string previousPath = null;
                foreach (Excel.Window window in workbook.Windows)
                {
                    int hwnd = window.Hwnd;
                    if (lastKnownPaths.TryGetValue(hwnd, out string known) && !string.Equals(known, path, StringComparison.OrdinalIgnoreCase)) previousPath = known;
                    lastKnownPaths[hwnd] = path;

                    PaneEntry entry = FindPane(window);
                    if (entry == null || string.Equals(entry.RecordedPath, path, StringComparison.OrdinalIgnoreCase)) continue;

                    entry.RecordedPath = path;
                    if (source == null || (!source.Pane.Visible && entry.Pane.Visible)) source = entry;
                }

                // A shown pane is the source and also sets the defaults; a hidden one only gets an entry
                if (source != null)
                {
                    if (source.Pane.Visible)
                    {
                        WriteEntryFromPane(path, source);
                        WriteDefaults(source.RecordedDock, source.RecordedWidth, source.RecordedHeight);
                    }
                    else
                    {
                        WriteEntry(path, source.Pane.DockPosition, PaneWidth(source.Pane), PaneHeight(source.Pane, source.RecordedHeight), false);
                    }
                    SaveSettings();
                }
                else if (previousPath != null && TryGetEntry(previousPath, out Office.MsoCTPDockPosition dock, out int width, out int height, out _))
                {
                    WriteEntry(path, dock, width, height, false);
                    SaveSettings();
                }
            }
            catch (Exception ex) { Diagnostics.Write("AfterSave failed: " + ex); }
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

                    // A macro with screen updating off may activate every sheet in turn; the list catches up on the next tick
                    if (ReferenceEquals(entry, activeEntry) && this.Application.ScreenUpdating)
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
        /// Shows or hides a window's pane as a user action (the Ribbon button, with the state it now
        /// shows). The file's entry is updated by <see cref="Pane_VisibleChanged"/>.
        /// </summary>
        public void SetPaneVisible(Excel.Window window, bool visible)
        {
            try
            {
                PruneDeadPanes();

                window = window ?? this.Application.ActiveWindow;
                if (window == null || this.Application.ActiveWorkbook == null) return;

                PaneEntry entry = GetOrCreatePane(window);

                // The user is working in this window, so any pending close was cancelled
                entry.IsClosing = false;

                if (entry.Pane.Visible != visible) entry.Pane.Visible = visible;
            }
            catch (Exception ex) { Diagnostics.Write("SetPaneVisible failed: " + ex); }
        }

        /// <summary>
        /// The first time a window is seen, shows its pane from the entry's dock position and size if
        /// the workbook's entry says visible. Never writes settings.
        /// </summary>
        private void RestorePane(Excel.Workbook workbook, Excel.Window window)
        {
            if (workbook == null || window == null) return;

            try
            {
                // After the first time the pane keeps whatever state the user left it in
                if (FindPane(window) != null) return;
                if (!HasPath(workbook)) return;
                if (!TryGetEntry(workbook.FullName, out _, out _, out _, out bool visible) || !visible) return;

                PaneEntry entry = GetOrCreatePane(window);
                entry.Control.RefreshWorksheets(workbook);
                SetPaneVisibleSilently(entry.Pane, true);
            }
            catch (Exception ex) { Diagnostics.Write("RestorePane failed: " + ex); }
        }

        /// <summary>
        /// Syncs the Ribbon button, then records a user's show at once and queues a user's hide;
        /// changes made by the add-in itself are ignored.
        /// </summary>
        private void Pane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                ribbon?.RefreshToggleState();

                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null) return;

                if (pane.Visible)
                {
                    // Only the user shows a pane, so any pending close was cancelled
                    entry.IsClosing = false;

                    Excel.Workbook workbook = entry.Control.Workbook ?? this.Application.ActiveWorkbook;
                    if (workbook == null) return;

                    // A workbook without a path has no entry until it is saved
                    if (HasPath(workbook)) RecordShow(workbook.FullName, entry);
                    entry.Control.RefreshWorksheets(workbook);
                }
                else
                {
                    // A closing window reports its pane as hidden too, sometimes while Excel still lists
                    // the window, so whether this was the user's doing is decided once the delay has passed
                    QueueHideSave(entry);
                }
            }
            catch (Exception ex) { Diagnostics.Write("VisibleChanged failed: " + ex); }
        }

        /// <summary>
        /// Marks the file's entry visible, adding one from the pane if the file has none; a failed save is reported.
        /// </summary>
        private void RecordShow(string path, PaneEntry entry)
        {
            try
            {
                if (HasEntry(path)) SetEntryVisible(path, true);
                else WriteEntryFromPane(path, entry);
                SaveSettings();
            }
            catch (Exception ex) { ReportSaveFailure("Show", ex); }
        }

        /// <summary>
        /// Queues a hide to be recorded once the delay has passed, restarting the delay.
        /// </summary>
        private void QueueHideSave(PaneEntry entry)
        {
            if (!pendingHideEntries.Contains(entry)) pendingHideEntries.Add(entry);
            hideSaveTimer.Stop();
            hideSaveTimer.Start();
        }

        /// <summary>
        /// Queues a resize, dock or float to be recorded once the delay has passed, restarting the delay.
        /// </summary>
        private void QueueResizeSave(PaneEntry entry)
        {
            if (!pendingResizeEntries.Contains(entry)) pendingResizeEntries.Add(entry);
            resizeSaveTimer.Stop();
            resizeSaveTimer.Start();
        }

        /// <summary>
        /// Records every queued hide and layout change now, for a close or shutdown that would otherwise drop them.
        /// </summary>
        private void FlushPendingSaves()
        {
            ResizeSaveTimer_Tick(null, null);
            HideSaveTimer_Tick(null, null);
        }

        /// <summary>
        /// True while a queued change must keep waiting because its workbook's close is still pending;
        /// a window still open after the last re-check survived a cancelled close.
        /// </summary>
        private static bool StillClosing(PaneEntry entry)
        {
            if (!entry.IsClosing) return false;
            if (++entry.ClosingRechecks < ClosingRecheckLimit) return true;

            entry.IsClosing = false;
            return false;
        }

        /// <summary>
        /// Queues the resize so a drag is written once, when it ends.
        /// </summary>
        private void Control_Resize(object sender, EventArgs e)
        {
            try
            {
                PaneEntry entry = FindEntry(sender as SheetNavigatorControl);
                if (entry == null || !IsAlive(entry) || !entry.Pane.Visible) return;

                QueueResizeSave(entry);
            }
            catch (Exception ex) { Diagnostics.Write("Resize failed: " + ex); }
        }

        /// <summary>
        /// Records the layout changes that have settled; a pane whose window is gone was changed by the teardown.
        /// </summary>
        private void ResizeSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                resizeSaveTimer.Stop();

                List<PaneEntry> entries = new List<PaneEntry>(pendingResizeEntries);
                pendingResizeEntries.Clear();

                foreach (PaneEntry entry in entries)
                {
                    if (!IsAlive(entry) || !IsWindowOpen(entry.Hwnd) || !entry.Pane.Visible) continue;
                    if (StillClosing(entry))
                    {
                        QueueResizeSave(entry);
                        continue;
                    }

                    RecordPaneChange(entry);
                }
            }
            catch (Exception ex) { ReportSaveFailure("Resize", ex); }
        }

        /// <summary>
        /// Records the hides that have settled; a pane whose window is gone was hidden by the teardown,
        /// and one shown again meanwhile is left alone.
        /// </summary>
        private void HideSaveTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                hideSaveTimer.Stop();

                List<PaneEntry> entries = new List<PaneEntry>(pendingHideEntries);
                pendingHideEntries.Clear();

                bool changed = false;
                foreach (PaneEntry entry in entries)
                {
                    if (!IsAlive(entry) || !IsWindowOpen(entry.Hwnd)) continue;
                    if (StillClosing(entry))
                    {
                        QueueHideSave(entry);
                        continue;
                    }
                    if (entry.Pane.Visible) continue;

                    Excel.Workbook workbook = entry.Control.Workbook ?? this.Application.ActiveWorkbook;
                    if (workbook == null) continue;

                    // A file without an entry has nothing to mark hidden
                    if (!HasPath(workbook) || !HasEntry(workbook.FullName)) continue;

                    SetEntryVisible(workbook.FullName, false);
                    changed = true;
                }

                if (changed) SaveSettings();
            }
            catch (Exception ex) { ReportSaveFailure("Hide", ex); }
        }

        /// <summary>
        /// Queues the dock or float to be recorded; a pane that just floated is given the saved height first.
        /// </summary>
        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || !IsAlive(entry) || !pane.Visible) return;

                // Excel rejects property sets inside this handler, so the height is applied right after it returns.
                // Floating fires as the pane detaches and Excel may change the width at the same time, which
                // is why every dock change is recorded through the resize delay rather than on the spot.
                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    entry.Control.BeginInvoke(new Action(() => ApplyFloatingHeight(entry)));
                    return;
                }

                QueueResizeSave(entry);
            }
            catch (Exception ex) { Diagnostics.Write("DockPositionChanged failed: " + ex); }
        }

        /// <summary>
        /// Runs right after the float event's handler returns: gives the floating pane the file's saved
        /// height (the default if the file has none), then queues the pane to be recorded.
        /// </summary>
        private void ApplyFloatingHeight(PaneEntry entry)
        {
            try
            {
                if (!IsAlive(entry) || !entry.Pane.Visible) return;
                if (entry.Pane.DockPosition != Office.MsoCTPDockPosition.msoCTPDockPositionFloating) return;

                // The saved height rather than this pane's own, so two windows on one workbook agree
                Excel.Workbook workbook = entry.Control.Workbook;
                int height = DefaultHeight;
                if (HasPath(workbook)) TryGetEntry(workbook.FullName, out _, out _, out height, out _);

                entry.Pane.Height = height;
                QueueResizeSave(entry);
            }
            catch (Exception ex) { Diagnostics.Write("Floating height failed: " + ex); }
        }

        /// <summary>
        /// Writes a pane's dock position, width and floating height to the file's entry if it has one
        /// and to the defaults, saved once; a change that left the pane as last recorded writes nothing.
        /// </summary>
        private void RecordPaneChange(PaneEntry entry)
        {
            // A docked pane's height is Excel's, so the recorded floating height stands while docked
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
            WriteDefaults(dock, width, height);
            SaveSettings();
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
                    if (!IsAlive(entry) || !entry.Pane.Visible) continue;
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
        /// Notes the window's current workbook path; see <see cref="lastKnownPaths"/>.
        /// </summary>
        private void RememberPath(Excel.Window window, Excel.Workbook workbook)
        {
            try
            {
                if (window != null && HasPath(workbook)) lastKnownPaths[window.Hwnd] = workbook.FullName;
            }
            catch { /* Excel is busy */ }
        }

        /// <summary>
        /// The pane already created for a window, or null.
        /// </summary>
        private PaneEntry FindPane(Excel.Window window)
        {
            if (window == null) return null;
            return panes.TryGetValue(window.Hwnd, out PaneEntry entry) ? entry : null;
        }

        /// <summary>
        /// The entry that owns this task pane, or null.
        /// </summary>
        private PaneEntry FindEntry(CustomTaskPane pane)
        {
            if (pane == null) return null;
            foreach (PaneEntry entry in panes.Values)
            {
                if (ReferenceEquals(entry.Pane, pane)) return entry;
            }
            return null;
        }

        /// <summary>
        /// The entry that owns this list control, or null.
        /// </summary>
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

        /// <summary>
        /// True when both are the same workbook; a COM failure counts as different.
        /// </summary>
        private static bool SameWorkbook(Excel.Workbook first, Excel.Workbook second)
        {
            try
            {
                return first != null && second != null && string.Equals(first.FullName, second.FullName, StringComparison.OrdinalIgnoreCase);
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
                // A value Excel rejects must not stop the pane from being created, or the file could never get one.
                ApplySavedOrDefault("dock position", () => pane.DockPosition = dock, () => pane.DockPosition = DefaultDockPosition);
                ApplySavedOrDefault("width", () => pane.Width = width, () => pane.Width = DefaultWidth);
                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating)
                {
                    ApplySavedOrDefault("height", () => pane.Height = height, () => pane.Height = DefaultHeight);
                }

                // Width is meaningless when docked top or bottom, so keep the pane on a side or floating. Excel's
                // "NoHorizontal" is the restriction compatible with a side-docked pane, despite its name.
                // Optional: a rejected restriction must never stop the pane from working.
                try { pane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal; }
                catch { }

                pane.VisibleChanged += new EventHandler(Pane_VisibleChanged);
                pane.DockPositionChanged += new EventHandler(Pane_DockPositionChanged);
                control.Resize += new EventHandler(Control_Resize);
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
            return entry;
        }

        /// <summary>
        /// Applies a saved pane value; if Excel rejects it the default is tried, and if that is rejected
        /// too the pane keeps Excel's own value.
        /// </summary>
        private static void ApplySavedOrDefault(string property, Action applySaved, Action applyDefault)
        {
            try
            {
                applySaved();
                return;
            }
            catch (Exception ex) { Diagnostics.Write($"Saved {property} rejected, trying the default: {ex.Message}"); }

            try
            {
                applyDefault();
            }
            catch (Exception ex) { Diagnostics.Write($"Default {property} rejected, keeping Excel's: {ex.Message}"); }
        }

        /// <summary>
        /// Removes panes whose window is gone or that Excel disposed, and forgets the paths of windows that are gone.
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

            foreach (int hwnd in new List<int>(lastKnownPaths.Keys))
            {
                if (!liveHwnds.Contains(hwnd)) lastKnownPaths.Remove(hwnd);
            }
        }

        /// <summary>
        /// Drops a window's pane. Handlers are unhooked first so the removal is not seen as a user hide.
        /// </summary>
        private void RemovePane(int hwnd)
        {
            if (!panes.TryGetValue(hwnd, out PaneEntry entry)) return;
            panes.Remove(hwnd);
            pendingResizeEntries.Remove(entry);
            pendingHideEntries.Remove(entry);

            try { entry.Pane.VisibleChanged -= Pane_VisibleChanged; } catch { }
            try { entry.Pane.DockPositionChanged -= Pane_DockPositionChanged; } catch { }
            try { entry.Control.Resize -= Control_Resize; } catch { }
            try { this.CustomTaskPanes.Remove(entry.Pane); } catch { /* Already disposed with its window */ }
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
        /// A height of zero or less (a garbage value) is the fallback; there is no ceiling, since Excel keeps a floating pane on screen.
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
        /// The dock position as written in settings: "Left", "Right" or "Floating"; top and bottom throw.
        /// </summary>
        private static string DockPositionName(Office.MsoCTPDockPosition dock)
        {
            // Top and bottom have no saved form and are only reachable if Excel rejected the dock restriction
            switch (dock)
            {
                case Office.MsoCTPDockPosition.msoCTPDockPositionLeft: return "Left";
                case Office.MsoCTPDockPosition.msoCTPDockPositionRight: return "Right";
                case Office.MsoCTPDockPosition.msoCTPDockPositionFloating: return "Floating";
                default: throw new ArgumentOutOfRangeException(nameof(dock), dock, "Only a side-docked or floating pane has a saved form");
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
        /// Reads a file's entry, or returns false if the file has none; a value that cannot be read falls back to its default.
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
        /// Writes a file's entry, replacing any older one, in memory only; the caller saves once it has written everything.
        /// </summary>
        private void WriteEntry(string path, Office.MsoCTPDockPosition dock, int width, int height, bool visible)
        {
            // The name can throw, so take it before the old entry is removed or that entry would be lost
            string dockName = DockPositionName(dock);
            RemoveEntries(path);
            FileEntries.Add($"{path}|{dockName}|{width}|{height}|{visible}");
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
        /// Removes every entry for a workbook path.
        /// </summary>
        private void RemoveEntries(string path)
        {
            StringCollection entries = FileEntries;
            for (int i = entries.Count - 1; i >= 0; i--)
            {
                if (IsEntryForPath(entries[i], path)) entries.RemoveAt(i);
            }
        }

        /// <summary>
        /// True if the saved entry belongs to this workbook path.
        /// </summary>
        private static bool IsEntryForPath(string entry, string path)
        {
            return entry != null && entry.StartsWith(path + "|", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Logs a failed settings save and tells the user, since the pane state they just changed was not kept.
        /// </summary>
        private void ReportSaveFailure(string action, Exception ex)
        {
            Diagnostics.Write(action + " save failed: " + ex);
            MessageBox.Show(ExcelOwner(), $"Could not save the pane settings.\n\nDetails: {ex.Message}",
                            "Save Failure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Excel's main window as a message box owner, so the box stays in front of Excel; null if Excel will not say.
        /// </summary>
        private IWin32Window ExcelOwner()
        {
            try
            {
                return new WindowHandle(new IntPtr(this.Application.Hwnd));
            }
            catch { return null; }
        }

        /// <summary>
        /// Records pending changes, stops the timers, unhooks the Excel and pane events and drops the
        /// panes; each step is guarded so one failure cannot skip the rest.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // A change still waiting for its delay would otherwise be lost with the timers
            try { FlushPendingSaves(); } catch { }

            try { resizeSaveTimer.Stop(); resizeSaveTimer.Dispose(); } catch { }
            try { hideSaveTimer.Stop(); hideSaveTimer.Dispose(); } catch { }
            try { refreshTimer.Stop(); refreshTimer.Dispose(); } catch { }

            try { this.Application.WindowActivate -= Application_WindowActivate; } catch { }
            try { this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose; } catch { }
            try { this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave; } catch { }
            try { this.Application.SheetActivate -= Application_SheetActivate; } catch { }

            foreach (int hwnd in new List<int>(panes.Keys)) RemovePane(hwnd);
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
            /// How many times a queued change has waited on <see cref="IsClosing"/>; see <see cref="ClosingRecheckLimit"/>.
            /// </summary>
            public int ClosingRechecks { get; set; }

            /// <summary>
            /// The workbook path this pane's entry was last written under, null for an unsaved workbook;
            /// a save under another path is a first save or Save As.
            /// </summary>
            public string RecordedPath { get; set; }

            /// <summary>
            /// The dock position last recorded for this pane, so a layout event that changed nothing is not written again.
            /// </summary>
            public Office.MsoCTPDockPosition RecordedDock { get; set; }

            /// <summary>
            /// The width last recorded for this pane; see <see cref="RecordedDock"/>.
            /// </summary>
            public int RecordedWidth { get; set; }

            /// <summary>
            /// The floating height last recorded for this pane; a docked pane's height is Excel's, so
            /// this is always the last floating height.
            /// </summary>
            public int RecordedHeight { get; set; }

            /// <summary>
            /// Pairs a window handle with its pane and list control; the recorded values are set by the caller.
            /// </summary>
            public PaneEntry(int hwnd, CustomTaskPane pane, SheetNavigatorControl control)
            {
                Hwnd = hwnd;
                Pane = pane;
                Control = control;
            }
        }

        /// <summary>
        /// Wraps a native window handle for use as a message box owner.
        /// </summary>
        private sealed class WindowHandle : IWin32Window
        {
            public IntPtr Handle { get; }

            /// <summary>
            /// Keeps the handle; nothing is created or owned.
            /// </summary>
            public WindowHandle(IntPtr handle)
            {
                Handle = handle;
            }
        }
    }
}
