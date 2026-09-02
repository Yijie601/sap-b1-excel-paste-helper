# Changelog

## 0.1.0-beta.12

- Replaced the single automatic run with a five-press F8 sequence: Supplier, Posting Date, Supplier Ref., Remarks, then Items.
- Each F8 press performs exactly one calibrated click and one paste, so the user decides when SAP is ready for the next field.
- The fifth F8 press clicks First Item No. and pastes every copied Excel E:N row as one 10-column tab-delimited block.
- Copying a valid B:N selection resets the sequence to Supplier; a failed step remains selected for a safe retry, and a completed sequence cannot accidentally cycle back into Supplier.
- Added visible next-step notifications and per-step log entries for company-PC troubleshooting.

## 0.1.0-beta.11

- Removed simulated Tab keys from the SAP paste flow because the target SAP client can reject synthetic Tab input.
- Restored Supplier Ref. as a calibrated absolute desktop position.
- Corrected the coordinate-click order to Supplier, Posting Date, Supplier Ref., Remarks, and First Item No.; the helper waits for SAP processing after Posting Date before continuing to Supplier Ref.
- Calibration now requires five click targets; beta 8 configurations that still contain Supplier Ref. remain compatible, while beta 9/10 configurations request only the missing Supplier Ref. position.

## 0.1.0-beta.10

- Changed each Supplier navigation Tab into a human-like key press with separate key-down, 110 ms hold, key-up, and 220 ms release delays so SAP Business One cannot miss an instantaneous key tap.
- Added a 500 ms settling delay after pasting Supplier before sending the first Tab.
- The main status now explicitly shows `Pressing Tab 1 of 2` and `Pressing Tab 2 of 2` while the sequence runs.

## 0.1.0-beta.9

- Corrected the SAP header sequence: paste Supplier, press Tab once to commit, wait for SAP, press Tab again, and paste Supplier Ref. into the focused field.
- Posting Date is now the only date field entered; SAP is allowed to update Document Date automatically.
- Reduced calibration from six click positions to four: Supplier, Posting Date, Remarks, and First Item No.
- Existing beta 8 absolute coordinates for those four fields remain compatible.

## 0.1.0-beta.8

- Switched calibration and runtime input to direct absolute desktop coordinates, removing SAP process, window-title, child-window, and focused-control detection from the paste path.
- Header values now use clipboard paste exactly once per target: the first Excel B:D row supplies Supplier, both dates, Supplier Ref., and Remarks.
- All selected Excel E:N rows are still rebuilt as one 10-column tab-delimited matrix and pasted once at First Item No.
- Subsequent item rows may leave B:D blank; nonblank B:D values are still checked to prevent mixed invoices.
- Added 50-row regression coverage and validation that Supplier Name never enters the item matrix.
- Absolute-coordinate v2 calibration invalidates older relative coordinates and requires a one-time recapture of all six positions.

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
