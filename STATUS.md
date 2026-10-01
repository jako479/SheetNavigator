# Status

[Where things stand now]

## Decisions

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
