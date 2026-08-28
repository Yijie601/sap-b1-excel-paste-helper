# SAP B1 Excel Helper

A lightweight Windows tray utility that validates copied Excel AP Invoice rows and pastes them into SAP Business One AP Invoice using calibrated field positions.

## Daily workflow

1. In Excel, select one invoice across exactly columns **B:N**, without the Excel header.
2. Press `Ctrl+C`. The helper validates and prepares the invoice immediately.
3. Confirm the helper status is **Ready**.
4. Switch to SAP Business One and make sure **AP Invoice** is active.
5. Press `F8` (or map a Logitech side button to `F8`).
6. Review the completed invoice and click **Add** yourself.

The helper never clicks SAP **Add** or **Update**.

## First-time setup

### Supplier mappings

Open **Supplier Mapping**, then add each Excel Supplier Name and its SAP BP/Supplier Code. Matching trims spaces and ignores letter case; it never guesses or fuzzy-matches suppliers.

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
- Supplier Mapping existence;
- dates in `dd-MM-yyyy`, `dd/MM/yyyy`, `dd.MM.yyyy`, or `yyyy-MM-dd` format.

SAP dates are prepared as `dd.MM.yy`. Empty Department and UoM cells remain in their original column positions. The E:N item matrix is pasted as one multi-row operation.

## Data locations

User-editable data and logs are stored under:

```text
%LOCALAPPDATA%\SapB1ExcelHelper\
├── Config\calibration.json
├── Config\supplier_mapping.csv
└── Logs\
```

These files are kept when the application is updated.

## Install and update

Download the latest `SapB1ExcelHelper-Setup-...-win-x64.exe` from the GitHub Releases page. The installer is per-user and normally does not require Administrator permission. If SAP Business One runs as Administrator, the helper must run with the same privilege level for Windows input automation to work.

Installing a newer version over the existing version updates the application while preserving mappings and calibration under `%LOCALAPPDATA%`.

## Building locally

Requirements:

- Windows x64
- .NET 8 SDK
- Inno Setup 6 for the installer

```powershell
dotnet run --project .\tests\SapB1ExcelHelper.SmokeTests\SapB1ExcelHelper.SmokeTests.csproj -c Release
dotnet publish .\SapB1ExcelHelper\SapB1ExcelHelper.csproj -c Release -r win-x64 --self-contained true -o .\artifacts\publish
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" .\installer\SapB1ExcelHelper.iss
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
- The first public build should be treated as a beta until calibrated and verified against the target company's SAP installation.
- SAP UI layouts that move individual fields require recalibration.

