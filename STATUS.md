# Status

[Where things stand now]

## Decisions

- No automated tests: the testable logic is private to ThisAddIn, which only exists inside Excel, and extracting it is not worth it for an add-in this size.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- A single click activates a sheet; keyboard navigation in the pane is blocked, since Ctrl+PgUp/PgDn already covers it.
- The signing key stays out of the repo; contributors create their own test certificate.
- No keyboard hotkey: Application.OnKey cannot reach VSTO code without a VBA bridge.
- State is not carried across Save As; not worth the complexity yet.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
