# Changelog

What changed for users, grouped by release version. Not a commit log.

## 1.0.0 - 2026-09-30

- A settings save that fails is logged instead of shown in a dialog, and a layout change whose save failed is saved again on the next resize, dock or float.
- The startup error box names the add-in Sheet Navigator.
- Floating height is capped at 1200 points, like the 400-point width cap.
- Every caught error is logged; a repeating one fills two lines.
- Initial release.
- Worksheets pane on the View tab lists the active workbook's visible sheets; click a name, or move to it with the arrow keys or its first letter and press Enter, to jump.
- Pane open state, position (left, right or floating) and size are remembered per workbook file; hiding the pane keeps its position and size.
- A new pane opens where the pane was last used, at the size last used; right and 150 points wide to start, 400 points tall when floating.
- Works in every workbook window.
