# Design

How the Worksheets pane behaves, and what the add-in does on each Excel event.

## Saved state

Everything lives in the user's settings, never in the workbook.

- **Entry per file**: `path|dock|width|visible`, for example
  `C:\Books\Sales.xlsx|Right|180|True`.
  - `dock` is `Left`, `Right` or `Floating` (Excel's own names).
  - `width` is in points, capped at 400.
  - `visible` is `True` or `False`.
  - Entries are never removed; hiding the pane marks the entry hidden, so
    the file keeps its dock and width.
- **DefaultDockPosition**: the dock the user last used, `Floating` included.
  New panes open here. Starts at `Right`.
- **DefaultWidth**: the width the user last used. New panes open at it.
  Starts at 150.
- A value that cannot be read falls back to its default: dock to
  DefaultDockPosition, width to DefaultWidth, visible to `False`. A settings
  file .NET cannot read at all is deleted and starts over.
- After an Office update the settings are carried over from the previous
  Excel build's folder.

## Panes

- One pane per Excel window, created the first time it is needed and kept
  (hidden or shown) until the window closes.
- One entry per file. Two windows on the same workbook share the entry; the
  last action wins.
- A floating pane's screen position is Excel's; the add-in only knows it is
  floating. A `Floating` entry comes back floating at its saved width,
  wherever Excel puts it.
- The Ribbon button on the View tab shows the state of the active window's pane.

## Pane events

A brand-new workbook has no path and cannot have an entry until it is saved.

**Pane hidden** (Ribbon button or the pane's X)

1. File with an entry: mark it hidden.
2. File without an entry: nothing is written.

**Pane shown** (Ribbon button)

1. File with an entry, window has no pane yet: create the pane with the
   entry's dock and width; mark the entry visible.
2. File with an entry, window already has a pane: show it as it is (no
   move, no resize); mark the entry visible.
3. Brand-new workbook: create the pane with the default dock and width;
   nothing is written (an entry is added when the workbook is saved).
4. Existing file without an entry: create the pane with the default dock
   and width; add an entry, visible.

**Pane docked or floated**

1. File with an entry: write the pane's dock and width to the entry, and to
   the defaults.
2. File without an entry: write the pane's dock and width to the defaults.

Width is written too because Excel may change it when docking or floating.

**Pane resized** (written once the drag settles, floating or docked)

Same as docked or floated: the pane's dock position and width go to the
entry if the file has one, and to the defaults.

A dock, float or resize event that left the pane exactly as last recorded
(for example the layout event Excel raises right after a pane is shown)
writes nothing. A pane docked top or bottom, only possible if Excel rejected
the add-in's dock restriction, is not recorded.

**Workbook saved**

1. Same path as before: nothing.
2. New path (first save, or Save As): write the pane's dock, width and
   shown/hidden state as that path's entry, replacing any entry the path
   already had; if the pane is shown, also write its dock and width to the
   defaults. With several windows, a shown pane is the source.

After a Save As the old path's entry is left alone.

Nothing is ever written from the entry back to a pane except in "Pane shown"
case 1 and on restore (below).

## Workbook and window events

- **Excel starts**: the active workbook's pane is restored if its entry is
  visible.
- **Workbook opened / window activated**: panes whose window is gone are
  dropped; the first time a window is seen, its pane is restored from the
  entry if the entry is visible; the Ribbon button is refreshed. Activating
  a window also brings back the pane after a cancelled close.
- **Workbook or window deactivated**: nothing.
- **Workbook closing**: its panes are flagged as closing. Excel reports the
  pane as hidden while the window closes; that is ignored, nothing is
  written, so the entry keeps saying visible and the pane comes back next
  time. The pane objects are dropped once the window is really gone.
- **Workbook saved**: see "Pane events".
- **Sheet activated**: every pane of the active workbook re-highlights its
  own window's active sheet; the active window's pane also refreshes its
  list.
- **Excel closes**: timers stop, events unhook, nothing is written.

## Worksheet list

The list shows the workbook's visible sheets, in tab order. Hidden and very
hidden sheets are not listed.

**When it refreshes**

- A sheet is activated.
- Once a second, for every shown pane. If Excel is mid-edit (typing in a
  cell or a tab name) only the highlight is updated.
- The pointer enters the list.
- The pane is shown or restored.
- A click on a sheet name could not jump (the sheet was renamed or removed).

A refresh rebuilds the list only when the visible names changed, so it
catches sheets that were added, deleted, renamed, moved, hidden or unhidden,
even though Excel raises no event for a rename or a move. The highlight then
moves to the window's active sheet.

**Clicking a name** activates that sheet in the pane's own window. Keyboard
navigation in the list is blocked; Ctrl+PgUp/PgDn already covers it.
