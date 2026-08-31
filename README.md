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
3. Repeat for Supplier, Supplier Ref., Posting Date, Remarks, and the first Item No. cell.
4. Use **Test Calibration** to watch the mouse move through the saved positions without clicking or entering data.
5. Click **Save**.

Coordinates are stored as direct Windows desktop positions. Keep SAP maximized or in the same screen position used during calibration. Recalibrate after moving SAP, changing display scaling, resolution, monitor arrangement, or the SAP layout.

All five positions must show a green check before **Test Calibration**, **Save**, or the global hotkey can run. Beta 8 and newer use absolute coordinates; older relative coordinates are intentionally rejected. A beta 9/10 installation normally needs to capture only the newly restored Supplier Ref. position.

During capture, the helper records the exact desktop point you click. During F8 execution it moves to those points directly; it does not depend on the SAP process name, A/P Invoice title, internal window class, or input-control detection.

## Clipboard validation

The helper expects 13 tab-separated columns corresponding to Excel B:N. It validates:

- exactly 13 columns on every row;
- data rows only, without the Excel header;
- Supplier Name, Document Date, and Document Number from the first row only; later B:D cells may repeat the same values or remain blank;
- a nonblank SAP Code / Item No. in column E for every selected row;
- dates in `dd-MM-yyyy`, `dd/MM/yyyy`, `dd.MM.yyyy`, or `yyyy-MM-dd` format.

The Excel Supplier Name is pasted once into Supplier. The helper then clicks Posting Date and pastes the first-row date once as `dd.MM.yy`, waits for SAP to finish Supplier/date processing, clicks Supplier Ref. and pastes the first-row Document Number, then clicks Remarks and pastes the same Document Number. No Tab key is sent, and SAP updates Document Date automatically. Empty item cells remain in their original positions. Whether there are 2 rows or dozens, all E:N rows are pasted once as one multi-row matrix beginning with column E SAP Code.

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
& "${env:ProgramFiles(x86)}\NSIS\makensis.exe" /DAPP_VERSION=0.1.0-beta.11 .\installer\SapB1ExcelHelper.nsi
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
- SAP must remain in the same desktop position used for calibration; moved fields or display-layout changes require recalibration.
