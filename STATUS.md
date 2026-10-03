# Status

The code review fixes, the keyboard-accessible list and the changes matched to Contracts File Navigator on 2026-10-02 compile, and the size rules pass their tests; none are yet tried in Excel. Next: Ctrl+F5 and run the checks in TODO.md.

## Decisions

- A failed settings save is logged, never shown: a layout save is minor, and a disk or permission problem would otherwise pop a dialog after every drag. A layout change is recorded only once its save succeeds, so the next layout event retries it.
- Automated tests cover only the pure size rules (PaneSizeRules, compiled into SheetNavigator.Tests as a linked file); everything else needs Excel and is verified by hand.
- Floating height is capped at 1200 points on save and load, like the 400-point width cap: Excel puts no maximum on a floating pane, so a garbage value could push its bottom edge off the screen.
- Every caught error is logged; identical consecutive messages are written once, followed by "Last message repeated N times" when a different message arrives or Excel closes, so a failure repeating every second fills two lines. Only two failures stay unlogged: a failure of the log itself, and Excel's expected refusal when Enter or a click lands mid-edit.
- A pane whose first fill or show fails on restore is dropped and logged, so the next activation tries again instead of finding a hidden pane and leaving it.
- ClickOnce's update check is off: there is no install URL for it to check, and a new version is installed by running its own setup.
- The list follows the Windows list pattern, as WCAG 2.1.1 requires keyboard operation: arrows and typed letters move the highlight, Enter or a click activates, and focus then goes back to the pane's window so Excel's keys work right away, as after a click on Excel's own sheet tabs. While the list has keyboard focus, refreshes leave the highlight alone. Screen readers get the name Worksheets.
- A change queued while a workbook's close is pending is re-checked every half second for up to two seconds; a window still open by then survived a cancelled close, so the change is the user's and is recorded. A hide or resize still queued when a close or Excel shutdown begins is recorded first, since it happened before the close.
- Only a file named user.config under the user's profile that .NET reports as unparsable is deleted and reset; a locked or otherwise unavailable file is kept and the failure logged, so a transient error cannot wipe the saved entries. A reset during a save keeps the values being saved.
- A saved dock position, width or height that Excel rejects falls back to the default, and a rejected default leaves Excel's own, so a bad entry can never stop a file from getting a pane.
- The list is not rebuilt while a macro has screen updating off; it catches up on the first tick after the macro restores it.
- A dock or float is recorded through the same half-second delay as a resize, since Excel may still be changing the width when the dock event fires.
- A hide is recorded half a second after the event and only if the window is still open then: a closing window reports its pane hidden too, sometimes while Excel still lists the window, and closing one window of several raises no closing event at all.
- Each user action saves the settings file once; entries and defaults are written in memory first.
- A pane docked top or bottom has no saved form: a layout event ignores it and any other save of it throws, so Excel ignoring the dock restriction shows up in the log as a failure instead of being saved as Left.
- Excel rejects any pane property set inside a pane event handler ("cannot be set during the object's event handler") and may still resize the pane as the drag that floated it ends, so the saved height is applied by the resize save half a second later, in its own try, and the pane is recorded after it either way.
- Save As writes the pane's state as the new path's entry, replacing any entry that path had; a workbook with no pane passes its hidden entry on. The old path's entry is left alone.
- A floating pane is remembered as floating, at its width and height; its screen position is not saved because the pane API exposes none, so Excel places it. The file's saved height is applied whenever the pane floats, so two windows on one workbook agree, and height is saved only while the pane is floating, since a docked pane takes the window's height.
- A file's entry is never removed; hiding the pane marks it hidden so the position and size survive. Pane behavior per event is in DESIGN.md.
- State lives in the user's settings, never in the workbook, so opening the pane never dirties a file.
- The signing key stays out of the repo; contributors create their own test certificate.
- No keyboard hotkey: Application.OnKey cannot reach VSTO code without a VBA bridge.
- The release ZIP is made by hand; zipping is not a VSTO publish convention.
