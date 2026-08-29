# Changelog

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
