# Sheet Navigator

An Excel add-in that adds a **Worksheets** pane listing the sheets in the active workbook. Click a name to jump to that sheet. Excel remembers, per file, whether the pane was open, where it was and how big it was.

## Requirements

- Windows with desktop Microsoft Excel 2013 or later.
- .NET Framework 4.7.2 and the Visual Studio 2010 Tools for Office Runtime. The installer adds both if they are missing.

## Install

1. Download the latest release ZIP and extract it. Don't run the installer from inside the ZIP window.
2. Close Excel and run `setup.exe`. Click **Install** when Office asks.
3. Open a workbook. The button is on the **View** tab, in the **Navigate** group.

## Use

- **View, Navigate, Worksheets** shows or hides the pane. Press Alt to see the KeyTips.
- Click a sheet in the pane to activate it.
- The pane's open state, position (left, right or floating) and size are remembered for each workbook file, and kept while the pane is hidden. This is stored in your user profile; nothing is written to the workbook. See DESIGN.md for the details.
- A new pane opens where you last had one, at the size you last used (right and 150 points wide at first; a floating pane starts 400 points tall).

## Troubleshooting

The add-in writes a short log of pane events and any unexpected error to `%TEMP%\SheetNavigator.log`.

## Uninstall

Windows Settings, Apps, **Sheet Navigator**, Uninstall.

## Build

- Visual Studio 2022 or later with the **Office/SharePoint development** workload.
- Open `SheetNavigator.slnx` and build the Release configuration. Ctrl+F5 starts Excel with the add-in loaded.
- ClickOnce signing needs a certificate, which is not in the repo. Before the first build, open Project Properties, **Signing**, and click **Create Test Certificate** (or choose your own).

## Release

1. In Visual Studio, right-click the project and choose **Publish**. Output goes to `SheetNavigator\publish\`; `release\README.txt` and `LICENSE` are copied there automatically.
2. Zip the contents of `publish\` and attach it to a GitHub release.

## License

MIT. See [LICENSE](LICENSE).
