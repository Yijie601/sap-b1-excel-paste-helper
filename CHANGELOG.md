# Changelog

## 0.1.0-beta.7

- Fixed `Test Calibration` and `Capture` closing and disposing the calibration window while it was temporarily hidden.
- Calibration now safely hides and returns as a modeless owned window.
- All six SAP positions must be explicitly captured before calibration can be tested, saved, or used by the global hotkey.
- Legacy built-in coordinates are treated as uncalibrated so a Supplier value cannot be entered into the wrong SAP field.
- Added a regression case for the COL33 invoice to verify the item paste starts at SAP Code, not Supplier Name.

## 0.1.0-beta.6

- Fixed calibration capture incorrectly reporting that SAP was inactive before Windows completed its focus change.
- Calibration now identifies SAP and the A/P Invoice from the window underneath the actual click position.
- Added support for both `AP Invoice` and SAP's standard `A/P Invoice` window titles.
- Added clearer errors when a captured click is outside SAP or outside the A/P Invoice window.

## 0.1.0-beta.5

- Removed Supplier Mapping from the workflow and interface.
- The exact Supplier Name from Excel is now pasted directly into the calibrated SAP Supplier field.
- Missing mapping files no longer block clipboard readiness or hotkey execution.
- Existing local `supplier_mapping.csv` files are preserved during updates but are no longer read.

## 0.1.0-beta.4

- Added a persistent custom global hotkey picker, available from the main status panel and tray menu.
- Function keys can be used alone; other shortcuts require Ctrl or Alt to avoid blocking normal typing.
- New shortcuts are checked with Windows before saving. If unavailable, the previous shortcut is restored.
- Confirmed the calibration workflow hides its main window while capturing a click in SAP, stores a window-relative coordinate, and restores the calibration window afterward.

## 0.1.0-beta.3

- Added automatic GitHub Release checks at startup, limited to once every 12 hours.
- Added a visible update prompt with release notes, Download & Install, Later, and download progress.
- Added a manual Check for Updates action in the main window and tray menu.
- Added prerelease-aware semantic version comparison, HTTPS host validation, and mandatory SHA-256 verification before launching the visible installer.
- Updates are never installed silently or forced while SAP automation is running.

## 0.1.0-beta.2

- Replaced the Inno Setup package with an NSIS installer whose license explicitly permits commercial use.
- Preserved per-user installation, optional desktop/startup shortcuts, in-place updates, and user configuration across uninstall/reinstall.

## 0.1.0-beta.1

- Added automatic clipboard monitoring and B:N invoice pre-validation.
- Added Supplier Mapping editor with CSV import/export.
- Added click-to-capture relative SAP field calibration and movement-only calibration test.
- Added global F8 hotkey and tray operation.
- Added direct Win32 input for SAP header fields and one-shot E:N item paste.
- Added clipboard restoration, safe-stop validation, notifications, duration logging, and a strict prohibition on Add/Update automation.
- Added a per-user Inno Setup installer and GitHub Actions release pipeline.
