# Design

How the Worksheets pane behaves, and what the add-in does on each Excel event.

## Saved state

Everything lives in the user's settings, never in the workbook.

- **Entry per file**: `path|dock|width|height|visible`, for example
  `C:\Books\Sales.xlsx|Right|180|400|True`.
  - `dock` is `Left`, `Right` or `Floating` (Excel's own names).
  - `width` is in points, capped at 400.
  - `height` is in points, the pane's height while floating. A docked pane
    is stretched to the window, so docking leaves it alone.
  - `visible` is `True` or `False`.
  - Entries are never removed; hiding the pane marks the entry hidden, so
    the file keeps its dock, width and height.
- **DefaultDockPosition**: the dock the user last used, `Floating` included.
  New panes open here. Starts at `Right`.
- **DefaultWidth**: the width the user last used. New panes open at it.
  Starts at 150.
- **DefaultHeight**: the floating height the user last used. New panes that
  start floating open at it. Starts at 400.
- A value that cannot be read falls back to its default: dock to
  DefaultDockPosition, width to DefaultWidth, height to DefaultHeight,
  visible to `False`. A settings file whose XML is corrupt is deleted and
  starts over; a file that is merely locked or unavailable is kept and the
  failure is reported.
- A saved dock, width or height that Excel rejects when the pane is
  created falls back to the default; if Excel rejects that too, the pane
  keeps Excel's own value.
- After an Office update the settings are carried over from the previous
  Excel build's folder.

## Panes

- One pane per Excel window, created the first time it is needed and kept
  (hidden or shown) until the window closes.
- One entry per file. Two windows on the same workbook share the entry; the
  last action wins.
- A floating pane's screen position is Excel's; the add-in only knows it is
  floating. A `Floating` entry comes back floating at its saved width and
  height, wherever Excel puts it.
- The Ribbon button on the View tab shows the state of the active window's pane.

## Pane events

A brand-new workbook has no path and cannot have an entry until it is saved.

**Pane hidden** (Ribbon button or the pane's X)

1. File with an entry: mark it hidden.
2. File without an entry: nothing is written.

The hide is recorded half a second later, and only if the window is still
open by then: a closing window reports its pane hidden too, sometimes while
Excel still lists the window, and that is not the user's doing. While the
workbook's close is pending (Excel is asking about unsaved changes) the hide
is re-checked every half second for up to two seconds; a window still open
by then survived a cancelled close, so the hide is recorded. A hide still
waiting when the close begins is recorded at once, since it came first.

**Pane shown** (Ribbon button)

1. File with an entry, window has no pane yet: create the pane with the
   entry's dock and width, and its height if that dock is floating; mark
   the entry visible.
2. File with an entry, window already has a pane: show it as it is (no
   move, no resize); mark the entry visible.
3. Brand-new workbook: create the pane with the default dock and width, and
   the default height if that dock is floating; nothing is written (an
   entry is added when the workbook is saved).
4. Existing file without an entry: create the pane as in case 3; add an
   entry, visible.

**Pane docked or floated** (written once it settles, half a second later)

A pane that just floated is given the saved height first. Then:

1. File with an entry: write the pane's dock and width, and its height if
   it is floating, to the entry, and to the defaults.
2. File without an entry: write the same to the defaults.

Width is written too because Excel may change it when docking or floating.
Docking never touches the saved height.

**Pane resized** (written once the drag settles, floating or docked)

Same as docked or floated: the pane's dock position and width, and its
height if it is floating, go to the entry if the file has one, and to the
defaults.

A dock, float or resize that arrives while the workbook's close is pending
waits like a hide does (see "Pane hidden"), and one still waiting when a
close or Excel shutdown begins is written at once.

A dock, float or resize event that left the pane exactly as last recorded
(for example the layout event Excel raises right after a pane is shown)
writes nothing. A pane docked top or bottom, only possible if Excel rejected
the add-in's dock restriction, is never recorded: a layout event ignores it,
and any other attempt to save it is reported as a failure. A pane reporting a width or
height of zero or less keeps the saved value. A docked pane's height is
Excel's and is never written, so the saved height is always the last
floating height.

**Workbook saved**

1. Same path as before: nothing.
2. New path (first save, or Save As): write the pane's dock, width, height
   and shown/hidden state as that path's entry, replacing any entry the
   path already had; if the pane is shown, also write its dock, width and
   height to the defaults. With several windows, a shown pane is the
   source.

A workbook with no pane (its entry said hidden, so none was ever created)
passes the old path's entry on to the new path, still hidden.

After a Save As the old path's entry is left alone.

## When settings are applied and saved

**Applied to the pane**: dock position and width when a pane is created
("Pane shown" cases 1, 3 and 4, and restore below), plus height if it
starts floating; and height again whenever the pane floats. Nothing else is
ever written from the settings back to a pane.

**Saved from the pane**: dock position and width on every dock, float or
resize, and height too while the pane is floating. Docking never touches
the saved height. Hiding the pane only marks the entry hidden, and a
closing workbook saves nothing of its own teardown. Each user action saves
the settings file once, however many values it changed.

## Workbook and window events

- **Excel starts**: the active workbook's pane is restored if its entry is
  visible.
- **Workbook opened / window activated**: panes whose window is gone are
  dropped; the first time a window is seen, its pane is restored from the
  entry if the entry is visible; the Ribbon button is refreshed. Activating
  a window also brings back the pane after a cancelled close.
- **Workbook or window deactivated**: nothing.
- **Workbook closing**: any hide or layout change still waiting for its
  delay is written first, then its panes are flagged as closing. Excel
  reports the pane as hidden while the window closes; that is ignored,
  nothing is written, so the entry keeps saying visible and the pane comes
  back next time. Closing one window of a workbook that has several raises
  no closing event, so there the half-second delay before a hide is
  recorded is what tells the teardown from the user. The pane objects are
  dropped once the window is really gone; working in a window again, or a
  change that outlives the two-second re-check, clears its closing flag.
- **Workbook saved**: see "Pane events".
- **Sheet activated**: every pane of the active workbook re-highlights its
  own window's active sheet; the active window's pane also refreshes its
  list, unless a macro has screen updating off.
- **Excel closes**: any hide or layout change still waiting for its delay
  is written, then timers stop and events unhook; nothing else is written.

## Worksheet list

The list shows the workbook's visible worksheets, in tab order. Hidden and
very hidden sheets are not listed, and neither are chart sheets; while a
chart sheet is active, nothing is highlighted.

**When it refreshes**

- A sheet is activated.
- Once a second, for every shown pane. If Excel is mid-edit (typing in a
  cell or a tab name), or a macro has screen updating off, only the
  highlight is updated.
- The pointer enters the list.
- The pane is shown or restored.
- A click on a sheet name could not jump (the sheet was renamed or removed).

A refresh rebuilds the list only when the visible names changed, so it
catches sheets that were added, deleted, renamed, moved, hidden or unhidden,
even though Excel raises no event for a rename or a move. The highlight then
moves to the window's active sheet.

**Clicking a name** activates that sheet in the pane's own window. The jump
happens when the mouse button is released, so dragging across names jumps
once, where the drag ends. A release on the blank space below the names
only puts the highlight back. Every release hands focus back to the pane's
window, so Excel's keys work right after a click. Keyboard navigation in
the list is blocked; Ctrl+PgUp/PgDn already covers it.

## Diagnostics

Errors and unhandled exceptions go to `%TEMP%\SheetNavigator.log`, newest
256 KB kept once it passes 1 MB. Nothing else is logged.
