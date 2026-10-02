# Keyboard-accessible worksheet list

2026-10-02

## Goal

The Worksheets pane list stops swallowing keys and works like any Windows
list: arrows and typed letters move the highlight, Enter activates. A click
still activates. Either way focus ends on the worksheet, like Excel's own
sheet tabs.

The design and code follow ContractsFileNavigator commit f31f30c ("Make the
sheet list keyboard accessible"), with the differences listed under
"Differences from ContractsFileNavigator".

## Behavior

**Keys**

- Arrows, Home, End, PgUp, PgDn and typed letters move the highlight without
  switching sheets.
- Enter activates the highlighted sheet. Enter is marked as an input key in
  PreviewKeyDown so the pane container does not take it, and swallowed in
  KeyDown so it does not beep.
- No other key is handled. Space is not an activator.

**Mouse**

- A left click on a name highlights and activates it, even one that was
  already highlighted. MouseClick replaces MouseUp.
- A click on the blank space below the names activates nothing.
  IndexFromPoint names the nearest item for any point inside the list, so
  the item's own rectangle decides whether a name was hit.

**Focus**

- After a click or Enter the pane hands focus back to its Excel window (the
  existing ActivateWindow), whether or not the jump happened, so Excel's
  keys work right away.
- When focus leaves the list, the highlight returns to the window's active
  sheet.

**Highlight while the list has focus**

- While the list has keyboard focus the highlight is the user's cursor:
  refreshes leave it alone.
- A rebuild of the list puts the highlight back on the same name if that
  sheet still exists, otherwise on the active sheet.

**Screen readers**

- The list's accessible name is "Worksheets".

## Code changes

All in `SheetNavigator/SheetNavigatorControl.cs` unless noted.

- Constructor hooks PreviewKeyDown, KeyDown, MouseClick, MouseDoubleClick
  (same handler as MouseClick, since a double-click's second release raises
  only that event) and LostFocus, keeps MouseEnter, drops MouseUp.
- `WorksheetList_PreviewKeyDown`: Enter is an input key.
- `WorksheetList_KeyDown`: Enter only; swallow it, activate the highlighted
  sheet, hand focus back to the window.
- `WorksheetList_MouseClick`: left button only; index from point plus the
  item-rectangle check; on a name, select it and activate; hand focus back
  to the window either way.
- `WorksheetList_LostFocus`: force the highlight to the active sheet.
- `HighlightActiveSheet()` stays public and calls a private
  `HighlightActiveSheet(bool force)` that returns early while the list has
  focus unless forced. Its silent catch stays: it runs once a second and
  the logger does not collapse repeats.
- `IsKeyboardNavigating`: true while the list has focus.
- `RefreshWorksheets`: unchanged names means highlight and return;
  otherwise remember the highlighted name while keyboard navigating,
  rebuild, then restore that name or force the highlight.
- `ActivateHighlightedSheet()` replaces `JumpTo(string)`: acts on the
  selected item with the same Excel calls, and takes the ContractsFileNavigator
  error flow minus its busy log line: log a failed ScreenUpdating restore,
  show the error box after the finally block so Excel repaints under it,
  with the pane as the box's owner, and when the jump did not happen force
  the highlight back to the active sheet after the refresh.
- `ActivateWindow()` unchanged.
- `SheetNavigatorControl.Designer.cs`: `WorksheetList.AccessibleName` is
  "Worksheets".
- Class summary comment updated.

## Differences from ContractsFileNavigator

- No Space.
- The item-rectangle check on a click. ContractsFileNavigator trusts
  IndexFromPoint alone, so there a click on blank space activates the last
  sheet.
- Focus is handed back to the window after a click or Enter.
- `HighlightActiveSheet` keeps its silent catch.
- The error box keeps the pane as its owner.
- A double-click's second release takes the click path, so it hands focus
  back too.
- No log line for a busy refusal: this project's log holds errors only.
- A refused activation forces the highlight back to the active sheet.

## Docs

- DESIGN.md: the "Clicking a name" paragraph becomes a "Keyboard and mouse"
  list; the refresh trigger reads "Activating a name could not jump".
- STATUS.md: the top line says the change is untested in Excel; the two
  "keyboard navigation is blocked" decisions become one keyboard-accessible
  list decision.
- TODO.md: one Excel check item.
- CHANGELOG.md (1.0.0 entry), README.md, release/README.txt: "click, or
  arrow to it and press Enter".

## Testing

- Compile with the MSBuild command in AGENTS.md.
- In Excel after Ctrl+F5: F6 reaches the pane; arrows and letters move the
  highlight without switching; Enter switches and focus lands on the grid;
  a click switches and focus lands on the grid; a click on blank space
  switches nothing and focus lands on the grid; the highlight stays put
  while the list has focus and snaps back when focus leaves; Narrator
  announces "Worksheets".

## Out of scope

- Space as an activator.
- Collapsing repeated log lines.
- ContractsFileNavigator's blank-space click bug.
