# Keyboard-Accessible Worksheet List Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The Worksheets pane list stops swallowing keys: arrows and typed letters move the highlight, Enter or a click activates the sheet, and focus then returns to the worksheet.

**Architecture:** One WinForms UserControl (`SheetNavigatorControl`) hosts a `ListBox`. The change replaces its key-swallowing and mouse-release handlers with the handlers ContractsFileNavigator uses (PreviewKeyDown, KeyDown, MouseClick, LostFocus), makes the highlight focus-aware, and renames the jump method. The existing window re-activation after every click stays and also runs after Enter.

**Tech Stack:** C# 7.3, .NET Framework 4.7.2, VSTO Excel add-in, WinForms, MSBuild from Visual Studio 2022+ (Release configuration only).

**Spec:** `docs/superpowers/specs/2026-10-02-keyboard-accessible-list-design.md`

## Global Constraints

- Release configuration only; never Debug, never F5.
- Build commands run from `SheetNavigator\` (the project folder inside the worktree), one command per call, never chained.
- .NET naming: PascalCase types, methods and constants; camelCase parameters, locals and private fields; no `Wb`/`Sh` abbreviations.
- Constants first in a class, nested classes last; every method has a one- or two-line `<summary>`; comments say why, not what.
- `.cs` files keep their UTF-8 BOM and CRLF endings; every other text file is UTF-8 with CRLF.
- Anything that changes Excel state temporarily is restored in a `finally`; COM calls that can fail while Excel is busy are wrapped in try/catch.
- Commit messages are one line, no body, never mentioning any AI tool.
- No test project exists; verification is the compile command, then the human's Ctrl+F5 checks in Excel.
- Space is not an activator; only Enter and a left click activate.

## Review Focus

There are no automated tests, so each line below is pinned by a manual check in Task 3's hand-off list.

1. A click on the blank space below the last name must activate nothing; `IndexFromPoint` names the nearest item there, so only the item-rectangle check prevents a jump to the last sheet.
2. Enter while Excel is mid-edit (typing in a cell or a tab name): no jump, no error box, focus handed back to the window.
3. Enter on a name whose sheet was renamed or removed since the list was filled: no jump, no error box, the list refreshes and the highlight returns to the active sheet.
4. Keyboard focus in the list while the once-a-second refresh runs: the highlight must stay where the user put it, not snap back to the active sheet.
5. Two windows on one workbook: Enter in one window's pane activates the sheet in that window only.

---

### Task 1: Keyboard and click handling in the control

**Files:**
- Modify: `SheetNavigator/SheetNavigatorControl.cs` (whole file replaced below)
- Modify: `SheetNavigator/SheetNavigatorControl.Designer.cs:36` (one line added)

**Interfaces:**
- Consumes: `Diagnostics.Write(string)` (existing), `Globals.ThisAddIn.Application` (existing).
- Produces: `public void RefreshWorksheets(Excel.Workbook)` and `public void HighlightActiveSheet()` keep their signatures; `ThisAddIn.cs` calls them at lines 463, 467, 513, 544 and `RefreshQuietly()` at 778, all unchanged.

- [ ] **Step 1: Replace the control file**

Write `SheetNavigator/SheetNavigatorControl.cs` with exactly this content, keeping the UTF-8 BOM and CRLF endings:

```csharp
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
            this.WorksheetList.PreviewKeyDown += new PreviewKeyDownEventHandler(WorksheetList_PreviewKeyDown);
            this.WorksheetList.KeyDown += new KeyEventHandler(WorksheetList_KeyDown);
            this.WorksheetList.MouseClick += new MouseEventHandler(WorksheetList_MouseClick);

            // Once the keyboard leaves the list, the highlight follows the active sheet again
            this.WorksheetList.LostFocus += new EventHandler(WorksheetList_LostFocus);

            // Excel raises no event for a sheet rename or reorder, so check as the pointer arrives
            this.WorksheetList.MouseEnter += new EventHandler(WorksheetList_MouseEnter);
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
            string highlighted = IsKeyboardNavigating ? this.WorksheetList.SelectedItem as string : null;

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

            if (highlighted != null && names.Contains(highlighted))
            {
                this.WorksheetList.SelectedItem = highlighted;
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
        /// True while the list has keyboard focus, when the highlight belongs to the user rather than to Excel.
        /// </summary>
        private bool IsKeyboardNavigating
        {
            get { return this.WorksheetList.Focused; }
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
            if (this.WorksheetList.Items.Count != names.Count) return false;
            for (int i = 0; i < names.Count; i++)
            {
                if (!string.Equals(this.WorksheetList.Items[i] as string, names[i], StringComparison.Ordinal)) return false;
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
            catch { /* Keep the current list */ }
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
            int index = this.WorksheetList.IndexFromPoint(e.Location);
            bool onName = index != ListBox.NoMatches && this.WorksheetList.GetItemRectangle(index).Contains(e.Location);

            if (onName)
            {
                this.WorksheetList.SelectedIndex = index;
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
            catch { /* Excel is busy; the next click or Enter tries again */ }
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
                string sheetName = this.WorksheetList.SelectedItem as string;
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
                // Excel refused because it is mid-edit; expected, so the user is not told
                Diagnostics.Write("Jump refused, Excel is busy: " + ex.Message);
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

                if (!jumped) RefreshQuietly();
            }

            // Only after ScreenUpdating is back on, so Excel repaints while the box is up
            if (failure != null)
            {
                MessageBox.Show(this, $"Could not jump to sheet: {failure}", "Navigation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
```

- [ ] **Step 2: Name the list for screen readers**

In `SheetNavigator/SheetNavigatorControl.Designer.cs`, inside `InitializeComponent()`, add one line directly above `this.WorksheetList.Dock = System.Windows.Forms.DockStyle.Fill;`:

```csharp
            this.WorksheetList.AccessibleName = "Worksheets";
```

- [ ] **Step 3: Check encoding and endings**

Run from the worktree root:

```bash
file SheetNavigator/SheetNavigatorControl.cs SheetNavigator/SheetNavigatorControl.Designer.cs
```

Expected: both lines say `UTF-8 (with BOM) text, with CRLF line terminators`. If either lost its BOM or has LF endings, fix it before compiling.

- [ ] **Step 4: Compile**

Run from `SheetNavigator\` (the project folder inside the worktree):

```powershell
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" SheetNavigator.csproj -t:Compile -p:Configuration=Release
```

Expected: `Build succeeded.` with 0 errors. Any error naming `WorksheetList_MouseUp` or `JumpTo` means an old reference survived.

- [ ] **Step 5: Commit**

```bash
git add SheetNavigator/SheetNavigatorControl.cs SheetNavigator/SheetNavigatorControl.Designer.cs
git commit -m "Make the sheet list keyboard accessible: arrows and letters move, Enter or click activates"
```

---

### Task 2: Docs

**Files:**
- Modify: `DESIGN.md:172-184`
- Modify: `STATUS.md:3,7,10,23`
- Modify: `TODO.md:1`
- Modify: `CHANGELOG.md:8`
- Modify: `README.md:3,19`
- Modify: `release/README.txt:5-6`

**Interfaces:** none; prose only. Every file keeps CRLF endings.

- [ ] **Step 1: DESIGN.md**

Under "**When it refreshes**", change the last bullet:

```markdown
- Activating a name could not jump (the sheet was renamed or removed).
```

Change the end of the paragraph that follows it so it reads:

```markdown
A refresh rebuilds the list only when the visible names changed, so it
catches sheets that were added, deleted, renamed, moved, hidden or unhidden,
even though Excel raises no event for a rename or a move. The highlight then
moves to the window's active sheet, unless the keyboard is using the list.
```

Replace the whole "**Clicking a name**" paragraph (six lines, ending "Ctrl+PgUp/PgDn already covers it.") with:

```markdown
**Keyboard and mouse**

- Arrows, Home, End, PgUp, PgDn and typed letters move the highlight without
  switching sheets, as in any Windows list.
- Enter or a click on a name activates that sheet in the pane's own window.
  A click on the blank space below the names activates nothing.
- After Enter or a click, focus goes back to the pane's window, whether or
  not the jump happened, so Excel's keys work right away.
- While the list has keyboard focus the highlight is the user's cursor:
  refreshes leave it where it is, and a rebuild puts it back on the same name
  if that sheet still exists. When focus leaves the list, the highlight
  returns to the active sheet.
- The pane handles no other keys.
- Screen readers announce the list as "Worksheets".
```

- [ ] **Step 2: STATUS.md**

Replace line 3 (`[Where things stand now]`) with:

```markdown
The keyboard-accessible list compiles and is not yet tried in Excel. Next: Ctrl+F5 and run the keyboard check in TODO.md.
```

Insert as the first bullet under `## Decisions`:

```markdown
- The list follows the Windows list pattern, as WCAG 2.1.1 requires keyboard operation: arrows and typed letters move the highlight, Enter or a click activates, and focus then goes back to the pane's window so Excel's keys work right away, as after a click on Excel's own sheet tabs. While the list has keyboard focus, refreshes leave the highlight alone. Screen readers get the name Worksheets.
```

Delete these two bullets:

```markdown
- Every mouse release on the list activates the pane's window, so Excel's own Ctrl+PgUp/PgDn work right after a click; the list itself still swallows keys.
```

```markdown
- A sheet is activated when the mouse button is released on its name, so dragging across the list jumps once; keyboard navigation in the pane is blocked, since Ctrl+PgUp/PgDn already covers it.
```

- [ ] **Step 3: TODO.md**

Add under `## TODO` (the file currently has only the heading):

```markdown

- [ ] Check in Excel: F6 reaches the pane; arrows and letters move the highlight without switching sheets; Enter and a click switch and leave focus on the grid; a click on blank space switches nothing; the highlight stays while the list has focus and snaps back when focus leaves; Narrator announces "Worksheets".
```

- [ ] **Step 4: CHANGELOG.md**

Replace line 8 with:

```markdown
- Worksheets pane on the View tab lists the active workbook's visible sheets; click a name, or move to it with the arrow keys or its first letter and press Enter, to jump.
```

- [ ] **Step 5: README.md**

Line 3: change `Click a name to jump to that sheet.` to `Click a name, or arrow to it and press Enter, to jump to that sheet.`

Line 19: replace with:

```markdown
- Click a sheet in the pane to activate it, or move to it with the arrow keys or its first letter and press Enter.
```

- [ ] **Step 6: release/README.txt**

Replace lines 5-6 with:

```text
Sheet Navigator adds a Worksheets pane to Excel that lists the
sheets in the active workbook. Click a name to jump to it, or
pick it with the arrow keys and press Enter.
```

- [ ] **Step 7: Check endings**

```bash
file DESIGN.md STATUS.md TODO.md CHANGELOG.md README.md release/README.txt
```

Expected: every line ends `with CRLF line terminators`.

- [ ] **Step 8: Commit**

```bash
git add DESIGN.md STATUS.md TODO.md CHANGELOG.md README.md release/README.txt
git commit -m "Docs for the keyboard-accessible list"
```

---

### Task 3: Full build for Excel and hand-off

**Files:**
- Copy: `SheetNavigator_TemporaryKey.pfx` from the main checkout's `SheetNavigator\` folder into the worktree's `SheetNavigator\` folder (gitignored; the worktree has none).

**Interfaces:** none.

- [ ] **Step 1: Copy the signing key**

```powershell
Copy-Item "C:\Users\Brian\Projects\Internal Projects\SheetNavigator\SheetNavigator\SheetNavigator_TemporaryKey.pfx" "C:\Users\Brian\Projects\Internal Projects\SheetNavigator\.claude\worktrees\keyboard-list\SheetNavigator\SheetNavigator_TemporaryKey.pfx"
```

- [ ] **Step 2: Full build in the worktree**

This registers the add-in from the worktree's `bin\Release`, which is what Excel will load. Run from `SheetNavigator\` inside the worktree:

```powershell
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" SheetNavigator.csproj -p:Configuration=Release
```

Expected: `Build succeeded.` with 0 errors.

- [ ] **Step 3: Confirm Excel will load the worktree build**

```powershell
(Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\SheetNavigator").Manifest
```

Expected: a path under `...\.claude\worktrees\keyboard-list\SheetNavigator\bin\Release\`.

- [ ] **Step 4: Hand off the Excel checks**

Tell the human to restart Excel, open a workbook with several sheets, and check:

1. F6 (repeat as needed) reaches the pane; Shift+F6 leaves it.
2. Up/Down, Home/End, PgUp/PgDn and a sheet's first letter move the highlight; the active sheet does not change.
3. Enter on a highlighted name switches to that sheet and focus lands on the grid (arrows then move the cell cursor).
4. A click on a name switches to it and focus lands on the grid.
5. A click on the blank space below the last name switches nothing and focus lands on the grid (Review Focus 1).
6. Start typing in a cell, F6 to the pane, press Enter: no switch, no error box, Esc then works on the cell (Review Focus 2).
7. F6 to the pane, highlight a sheet, with the mouse rename or delete that sheet on the tab bar, F6 back and press Enter on the stale name: no switch, no error box, the list refreshes and the highlight returns to the active sheet (Review Focus 3).
8. F6 to the pane, move the highlight off the active sheet, wait three seconds: the highlight stays put (Review Focus 4).
9. View, New Window; in the second window's pane press F6, arrow, Enter: only that window switches sheets (Review Focus 5).
10. Narrator (Win+Ctrl+Enter) announces "Worksheets" when the list gets focus.

- [ ] **Step 5: After the merge**

Before the worktree is removed, a full build must run in the main checkout's `SheetNavigator\` folder, or Excel reports a failed install for the deleted path. The hand-off message says so.
