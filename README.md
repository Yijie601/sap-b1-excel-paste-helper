# SAP B1 Excel Helper

A lightweight Windows tray utility that validates copied Excel AP Invoice rows and pastes them into SAP Business One AP Invoice using calibrated field positions.

## Daily workflow

1. In Excel, select one invoice across exactly columns **B:N**, without the Excel header.
2. Press `Ctrl+C`. The helper validates and prepares the invoice immediately.
3. Confirm the helper status is **Ready**.
4. Switch to SAP Business One and make sure **AP Invoice** is active.
5. Press the configured global hotkey (`F8` by default), or map a Logitech side button to that shortcut.

Click the **Hotkey** link in the main status panel, or choose **Hotkey Settings** from the tray menu, to change the shortcut. Function keys F1–F24 can be used alone; other keys require Ctrl or Alt. The setting is saved under `%LOCALAPPDATA%\SapB1ExcelHelper\Config` and remains after updates.
6. Review the completed invoice and click **Add** yourself.

The helper never clicks SAP **Add** or **Update**.

## First-time setup

No supplier mapping is required. The Supplier Name copied from Excel is pasted directly into the calibrated SAP Supplier field.

### SAP calibration

Open **Calibration**. For each field:

1. Click **Capture**.
2. Click the corresponding real field in SAP.
3. Repeat for Supplier, Supplier Ref., Posting Date, Document Date, Remarks, and the first Item No. cell.
4. Use **Test Calibration** to watch the mouse move through the saved positions without clicking or entering data.
5. Click **Save**.

Coordinates are stored relative to the AP Invoice window, so moving the SAP window does not invalidate calibration. A resolution or SAP layout change may require recalibration.

## Clipboard validation

The helper expects 13 tab-separated columns corresponding to Excel B:N. It validates:

- exactly 13 columns on every row;
- data rows only, without the Excel header;
- one Supplier Name, Document Date, and Document Number across all rows;
- dates in `dd-MM-yyyy`, `dd/MM/yyyy`, `dd.MM.yyyy`, or `yyyy-MM-dd` format.

The Excel Supplier Name is pasted directly into SAP. SAP dates are prepared as `dd.MM.yy`. Empty Department and UoM cells remain in their original column positions. The E:N item matrix is pasted as one multi-row operation.

## Data locations

User-editable data and logs are stored under:

```text
%LOCALAPPDATA%\SapB1ExcelHelper\
├── Config\calibration.json
├── Config\hotkey.json
└── Logs\
```

These files are kept when the application is updated.

## Install and update

Download the latest `SapB1ExcelHelper-Setup-...-win-x64.exe` from the GitHub Releases page. The installer is per-user and normally does not require Administrator permission. If SAP Business One runs as Administrator, the helper must run with the same privilege level for Windows input automation to work.

Installing a newer version over the existing version updates the application while preserving calibration, hotkey settings, and logs under `%LOCALAPPDATA%`. Older `supplier_mapping.csv` files are left untouched but are no longer used.

The helper checks this repository for a newer GitHub Release at startup, at most once every 12 hours. When an update is available it shows a visible prompt with **Download & Install** and **Later** options. It never installs silently. Choosing Download & Install downloads the Windows installer, verifies its GitHub-provided SHA-256 digest, and opens the normal setup wizard for the user to complete. A manual **Check for Updates** action is available in the main window and tray menu.

## Building locally

Requirements:

- Windows x64
- .NET 8 SDK
- NSIS 3.12 for the installer

```powershell
dotnet run --project .\tests\SapB1ExcelHelper.SmokeTests\SapB1ExcelHelper.SmokeTests.csproj -c Release
dotnet publish .\SapB1ExcelHelper\SapB1ExcelHelper.csproj -c Release -r win-x64 --self-contained true -o .\artifacts\publish
& "${env:ProgramFiles(x86)}\NSIS\makensis.exe" /DAPP_VERSION=0.1.0-beta.5 .\installer\SapB1ExcelHelper.nsi
```

## Publishing a new version

Update the version in `SapB1ExcelHelper.csproj`, commit it, and push a version tag:

```powershell
git tag v0.1.1
git push origin main --tags
```

GitHub Actions builds the self-contained executable, Windows installer, portable ZIP, and GitHub Release automatically.

## Current limitations

- Windows x64 and SAP Business One only.
- The exact Excel Supplier Name must be accepted by the target SAP Supplier field.
- The first public build should be treated as a beta until calibrated and verified against the target company's SAP installation.
- SAP UI layouts that move individual fields require recalibration.
