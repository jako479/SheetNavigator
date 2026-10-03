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

        /// <summary>The settings file .NET keeps per user: the only file settings recovery may delete.</summary>
        private const string UserConfigFileName = "user.config";

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
            catch (Exception ex)
            {
                Diagnostics.Write("Pane visibility check failed: " + ex);
                return false;
            }
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
                return PaneSizeRules.ClampWidth(ReadSetting(() => Properties.Settings.Default.DefaultWidth), FallbackWidth);
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
                return PaneSizeRules.ClampHeight(ReadSetting(() => Properties.Settings.Default.DefaultHeight), FallbackHeight);
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
        /// Reads a setting. If .NET reports the settings file corrupt, the file is reset and the read retried.
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
                Properties.Settings.Default.Reload();
                return read();
            }
        }

        /// <summary>
        /// Changes a setting in memory. If .NET reports the settings file corrupt, the file is reset and the write retried.
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
                Properties.Settings.Default.Reload();
                write();
            }
        }

        /// <summary>
        /// Saves every settings change made so far, once per user action. If .NET reports the settings
        /// file corrupt, the file is reset and the save retried; the values being saved stay in memory, so none are lost.
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
        /// Deletes a corrupt user.config (a crash can leave it truncated, and then every settings call
        /// throws until it is gone). True if the file was deleted; the caller reloads if it needs the defaults.
        /// </summary>
        private static bool TryRecoverSettings(ConfigurationErrorsException ex)
        {
            // Only the add-in's own user.config inside the user's profile, and only on a parse error;
            // anything else (another file, a passing "file in use") is logged and left alone
            string file = ConfigFileNamedBy(ex);
            if (!IsOwnUserConfig(file) || !IsParseError(ex))
            {
                Diagnostics.Write("Settings unavailable, file kept: " + (file ?? "(unknown path)") + " | " + ex);
                return false;
            }

            Diagnostics.Write("Settings file corrupt, deleting it: " + file + " | " + ex.Message);
            try
            {
                File.Delete(file);
                return true;
            }
            catch (Exception deleteException)
            {
                Diagnostics.Write("Settings file delete failed: " + deleteException);
                return false;
            }
        }

        /// <summary>The first file name in the exception chain, or null if .NET named none.</summary>
        private static string ConfigFileNamedBy(ConfigurationErrorsException ex)
        {
            for (Exception current = ex; current != null; current = current.InnerException)
            {
                string file = (current as ConfigurationException)?.Filename;
                if (!string.IsNullOrEmpty(file)) return file;
            }
            return null;
        }

        /// <summary>
        /// True only for a file named user.config inside the user's local or roaming application data,
        /// which is where .NET keeps this add-in's settings and never where a workbook lives.
        /// </summary>
        private static bool IsOwnUserConfig(string file)
        {
            try
            {
                if (string.IsNullOrEmpty(file)) return false;
                if (!string.Equals(Path.GetFileName(file), UserConfigFileName, StringComparison.OrdinalIgnoreCase)) return false;

                string fullPath = Path.GetFullPath(file);
                return IsInside(fullPath, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
                    || IsInside(fullPath, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Settings path check failed: " + ex);
                return false;
            }
        }

        /// <summary>True if the path lies inside the folder; the folder itself does not count.</summary>
        private static bool IsInside(string fullPath, string folder)
        {
            if (string.IsNullOrEmpty(folder)) return false;
            string prefix = Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// True if the failure is a parse error in the file's contents: an XML error, or a configuration
        /// error that names a line. A locked or missing file reports neither.
        /// </summary>
        private static bool IsParseError(ConfigurationErrorsException ex)
        {
            for (Exception current = ex; current != null; current = current.InnerException)
            {
                if (current is XmlException) return true;
                if (current is ConfigurationException configurationError && configurationError.Line > 0) return true;
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
                MessageBox.Show(ExcelOwner(), $"Sheet Navigator failed to initialize components.\n\nError Details: {ex.Message}",
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
            catch (Exception ex) { Diagnostics.Write("Initial check failed: " + ex); }
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
            catch (Exception ex) { Diagnostics.Write("WorkbookBeforeClose failed: " + ex); }
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
            catch (Exception ex) { Diagnostics.Write("WorkbookAfterSave failed: " + ex); }
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

                    if (!IsShown(entry) || !SameWorkbook(entry.Control.Workbook, activeWorkbook)) continue;

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
                try
                {
                    // The quiet fill tolerates a busy Excel; the refresh timer fills the list a moment later
                    entry.Control.RefreshQuietly();
                    SetPaneVisibleSilently(entry.Pane, true);
                }
                catch
                {
                    // A pane that could not be shown is dropped, so the next activation tries again instead
                    // of finding a hidden pane and leaving it
                    RemovePane(entry.Hwnd);
                    throw;
                }
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
        /// Marks the file's entry visible, adding one from the pane if the file has none; a failed save is logged.
        /// </summary>
        private void RecordShow(string path, PaneEntry entry)
        {
            try
            {
                if (HasEntry(path)) SetEntryVisible(path, true);
                else WriteEntryFromPane(path, entry);
                SaveSettings();
            }
            catch (Exception ex) { Diagnostics.Write("Show save failed: " + ex); }
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
                if (entry == null || !IsShown(entry)) return;

                QueueResizeSave(entry);
            }
            catch (Exception ex) { Diagnostics.Write("Resize failed: " + ex); }
        }

        /// <summary>
        /// Records the layout changes that have settled, after first giving a pane that just floated its
        /// saved height; a pane whose window is gone was changed by the teardown.
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
                    if (!IsShown(entry) || !IsWindowOpen(entry.Hwnd)) continue;
                    if (StillClosing(entry))
                    {
                        QueueResizeSave(entry);
                        continue;
                    }

                    if (entry.IsFloatingHeightPending) ApplyFloatingHeight(entry);
                    RecordPaneChange(entry);
                }
            }
            catch (Exception ex) { Diagnostics.Write("Resize save failed: " + ex); }
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
            catch (Exception ex) { Diagnostics.Write("Hide save failed: " + ex); }
        }

        /// <summary>
        /// Queues the dock or float to be recorded; a pane that just floated is marked so the resize save
        /// gives it the saved height first.
        /// </summary>
        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                if (isProgrammaticPaneChange || !(sender is CustomTaskPane pane)) return;

                PaneEntry entry = FindEntry(pane);
                if (entry == null || !IsShown(entry)) return;

                // Excel rejects property sets inside this handler and may still resize the pane as the drag
                // ends, so the saved height is applied, and the pane recorded, by the resize save. Docking goes
                // through the same delay, since Excel may change the width at the same time.
                if (pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating) entry.IsFloatingHeightPending = true;

                QueueResizeSave(entry);
            }
            catch (Exception ex) { Diagnostics.Write("DockPositionChanged failed: " + ex); }
        }

        /// <summary>
        /// Gives a pane that just floated the file's saved height (the default if the file has none), so
        /// two windows on one workbook agree. A rejected height is logged and the pane keeps Excel's, which
        /// is then what gets recorded; the dock position and width are recorded either way.
        /// </summary>
        private void ApplyFloatingHeight(PaneEntry entry)
        {
            entry.IsFloatingHeightPending = false;
            if (entry.Pane.DockPosition != Office.MsoCTPDockPosition.msoCTPDockPositionFloating) return;

            try
            {
                Excel.Workbook workbook = entry.Control.Workbook;
                int height = DefaultHeight;
                if (HasPath(workbook)) TryGetEntry(workbook.FullName, out _, out _, out height, out _);

                entry.Pane.Height = height;
            }
            catch (Exception ex) { Diagnostics.Write("Floating height failed: " + ex); }
        }

        /// <summary>
        /// Writes a pane's dock position, width and floating height to the file's entry if it has one
        /// and to the defaults, saved once; a change that left the pane as last recorded writes nothing.
        /// The values are recorded only once saved, so a failed save is retried by the next layout event.
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

            Excel.Workbook workbook = entry.Control.Workbook;
            if (HasPath(workbook) && HasEntry(workbook.FullName)) WriteEntry(workbook.FullName, dock, width, height, true);
            WriteDefaults(dock, width, height);
            SaveSettings();

            entry.RecordedDock = dock;
            entry.RecordedWidth = width;
            entry.RecordedHeight = height;
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
                    if (!IsShown(entry)) continue;
                    entry.Control.RefreshQuietly();
                }
            }
            catch (Exception ex) { Diagnostics.Write("Refresh tick failed: " + ex); }
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
            catch (Exception ex) { Diagnostics.Write("Window list unreadable: " + ex); }
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
            catch (Exception ex) { Diagnostics.Write("Path lookup failed: " + ex); }
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
            catch (Exception ex)
            {
                Diagnostics.Write("Pane is gone: " + ex);
                return false;
            }
        }

        /// <summary>
        /// True while the pane exists and is shown: the only time a layout event or refresh concerns it.
        /// </summary>
        private static bool IsShown(PaneEntry entry)
        {
            return IsAlive(entry) && entry.Pane.Visible;
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
            catch (Exception ex)
            {
                Diagnostics.Write("Workbook comparison failed: " + ex);
                return false;
            }
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
                catch (Exception ex) { Diagnostics.Write("Dock restriction rejected: " + ex); }

                pane.VisibleChanged += new EventHandler(Pane_VisibleChanged);
                pane.DockPositionChanged += new EventHandler(Pane_DockPositionChanged);
                control.Resize += new EventHandler(Control_Resize);
            }
            catch
            {
                // Don't leave a half-configured pane behind
                try { this.CustomTaskPanes.Remove(pane); } catch (Exception ex) { Diagnostics.Write("Pane removal after failed creation failed: " + ex); }
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
            catch (Exception ex)
            {
                // Nothing is pruned on a guess; the next event tries again
                Diagnostics.Write("Window list unreadable, pruning skipped: " + ex);
                return;
            }

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

            try { entry.Pane.VisibleChanged -= Pane_VisibleChanged; } catch (Exception ex) { Diagnostics.Write("Visible unhook failed: " + ex); }
            try { entry.Pane.DockPositionChanged -= Pane_DockPositionChanged; } catch (Exception ex) { Diagnostics.Write("Dock unhook failed: " + ex); }
            try { entry.Control.Resize -= Control_Resize; } catch (Exception ex) { Diagnostics.Write("Resize unhook failed: " + ex); }
            try { this.CustomTaskPanes.Remove(entry.Pane); } catch (Exception ex) { Diagnostics.Write("Pane removal failed (already disposed with its window?): " + ex); }
        }

        /// <summary>
        /// Changes a pane's visibility without treating it as a user action. A failure propagates, so
        /// the caller can drop the pane.
        /// </summary>
        private void SetPaneVisibleSilently(CustomTaskPane pane, bool visible)
        {
            if (pane == null) return;

            isProgrammaticPaneChange = true;
            try
            {
                pane.Visible = visible;
            }
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
        /// The pane's width, within the saved range.
        /// </summary>
        private int PaneWidth(CustomTaskPane pane)
        {
            return PaneSizeRules.ClampWidth(pane.Width, DefaultWidth);
        }

        /// <summary>
        /// The pane's height while floating, within the saved range; a docked pane's height is Excel's, so the fallback stands.
        /// </summary>
        private static int PaneHeight(CustomTaskPane pane, int fallback)
        {
            return pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                ? PaneSizeRules.ClampHeight(pane.Height, fallback)
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
                if (values.Length > 1 && int.TryParse(values[1], out int parsedWidth)) width = PaneSizeRules.ClampWidth(parsedWidth, width);
                if (values.Length > 2 && int.TryParse(values[2], out int parsedHeight)) height = PaneSizeRules.ClampHeight(parsedHeight, height);
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
        /// Excel's main window as a message box owner, so the box stays in front of Excel; null if Excel will not say.
        /// </summary>
        private IWin32Window ExcelOwner()
        {
            try
            {
                return new WindowHandle(new IntPtr(this.Application.Hwnd));
            }
            catch (Exception ex)
            {
                Diagnostics.Write("Excel window lookup failed: " + ex);
                return null;
            }
        }

        /// <summary>
        /// Records pending changes, stops the timers, unhooks the Excel and pane events and drops the
        /// panes, each step on its own so one failure cannot skip the rest.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // A change still waiting for its delay would otherwise be lost with the timers
            RunLogged("Pending saves", FlushPendingSaves);

            RunLogged("Resize timer stop", () => { resizeSaveTimer.Stop(); resizeSaveTimer.Dispose(); });
            RunLogged("Hide timer stop", () => { hideSaveTimer.Stop(); hideSaveTimer.Dispose(); });
            RunLogged("Refresh timer stop", () => { refreshTimer.Stop(); refreshTimer.Dispose(); });

            RunLogged("WindowActivate unhook", () => this.Application.WindowActivate -= Application_WindowActivate);
            RunLogged("WorkbookBeforeClose unhook", () => this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose);
            RunLogged("WorkbookAfterSave unhook", () => this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave);
            RunLogged("SheetActivate unhook", () => this.Application.SheetActivate -= Application_SheetActivate);

            RunLogged("Pane removal", () => { foreach (int hwnd in new List<int>(panes.Keys)) RemovePane(hwnd); });

            // A repeat count still pending would otherwise be lost with Excel
            Diagnostics.Flush();
        }

        /// <summary>
        /// Runs one shutdown step, logging a failure instead of letting it stop the steps after it.
        /// </summary>
        private static void RunLogged(string step, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex) { Diagnostics.Write(step + " failed: " + ex); }
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
            /// Set when the pane floats. The next resize save applies the saved height before recording,
            /// because Excel rejects a property set inside the dock event and may resize the pane again as the drag ends.
            /// </summary>
            public bool IsFloatingHeightPending { get; set; }

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
