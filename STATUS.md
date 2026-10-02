# Status

[Where things stand now]

## Decisions

- A change queued while a workbook's close is pending is re-checked every half second for up to two seconds; a window still open by then survived a cancelled close, so the change is the user's and is recorded. A hide or resize still queued when a close or Excel shutdown begins is recorded first, since it happened before the close.
- Only a settings file .NET cannot parse as XML, under the user's profile, is deleted and reset; a locked or otherwise unavailable file is kept and the failure reported, so a transient error cannot wipe the saved entries.
- A saved dock position, width or height that Excel rejects falls back to the default, and a rejected default leaves Excel's own, so a bad entry can never stop a file from getting a pane.
- Every mouse release on the list activates the pane's window, so Excel's own Ctrl+PgUp/PgDn work right after a click; the list itself still swallows keys.
- The list is not rebuilt while a macro has screen updating off; it catches up on the first tick after the macro restores it.
- The log holds errors only; per-event lines were dropped before release.
- A dock or float is recorded through the same half-second delay as a resize, since Excel may still be changing the width when the dock event fires.
- A hide is recorded half a second after the event and only if the window is still open then: a closing window reports its pane hidden too, sometimes while Excel still lists the window, and closing one window of several raises no closing event at all.
- Each user action saves the settings file once; entries and defaults are written in memory first.
- A pane docked top or bottom has no saved form: a layout event ignores it and any other save of it throws, so Excel ignoring the dock restriction shows up as a failure instead of being saved as Left.
- Excel rejects any pane property set inside a pane event handler ("cannot be set during the object's event handler"), so the floating height is applied by code queued with BeginInvoke to run right after the dock event returns.
- Save As writes the pane's state as the new path's entry, replacing any entry that path had; a workbook with no pane passes its hidden entry on. The old path's entry is left alone.
- A floating pane is remembered as floating, at its width and height; its screen position is not saved because the pane API exposes none, so Excel places it. The file's saved height is applied whenever the pane floats, so two windows on one workbook agree, and height is saved only while the pane is floating, since a docked pane takes the window's height.
- A file's entry is never removed; hiding the pane marks it hidden so the position and size survive. Pane behavior per event is in DESIGN.md.
- No automated tests: the testable logic is private to ThisAddIn, which only exists inside Excel, and extracting it is not worth it for an add-in this size.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- A sheet is activated when the mouse button is released on its name, so dragging across the list jumps once; keyboard navigation in the pane is blocked, since Ctrl+PgUp/PgDn already covers it.
- The signing key stays out of the repo; contributors create their own test certificate.
- No keyboard hotkey: Application.OnKey cannot reach VSTO code without a VBA bridge.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
