# SAP B1 Excel Helper

A lightweight Windows tray utility that validates copied Excel AP Invoice rows and pastes them into SAP Business One AP Invoice using calibrated field positions.

## Daily workflow

1. In Excel, select one invoice across exactly columns **B:N**, without the Excel header.
2. Press `Ctrl+C`. The helper validates and prepares the invoice immediately.
3. Confirm the helper status is **Ready**.
4. Switch to SAP Business One and make sure **AP Invoice** is active.
5. Press the configured global hotkey (`F8` by default) once. The helper runs these five paste actions automatically, with a 0.8-second guard between header pastes:
   1. Supplier
   2. Posting Date
   3. Supplier Ref.
   4. Remarks
   5. First Item No. — paste the entire copied E:N item block

6. After the sequence completes, review the invoice and click **Add** yourself.
7. Copy Excel B:N again before starting another run. Repeated F8 cannot replay the same prepared copy after a completed or interrupted attempt.

Click the **Hotkey** link in the main status panel, or choose **Hotkey Settings** from the tray menu, to change the shortcut. Function keys F1–F24 can be used alone; other keys require Ctrl or Alt. The setting is saved under `%LOCALAPPDATA%\SapB1ExcelHelper\Config` and remains after updates.
Release the shortcut keys after pressing them. The helper waits up to three seconds for Ctrl/Alt/Shift/Windows keys to be released before starting, and stops if you hold a modifier during a paste.

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

During capture, the helper records the exact desktop point you click. During F8 execution it moves to those points directly; it does not depend on the SAP process name, A/P Invoice title, internal window class, or input-control detection. A lightweight foreground/root-window check stops the run if a point is covered by another top-level window, focus leaves the target, or the target moves during execution. This check is not SAP or field recognition: you must still activate the correct AP Invoice and verify the coordinates before running.

## Clipboard validation

The helper expects 13 tab-separated columns corresponding to Excel B:N. It validates:

- exactly 13 columns on every row;
- data rows only, without the Excel header;
- Supplier Name, Document Date, and Document Number from the first row only; later B:D cells may repeat the same values or remain blank;
- a nonblank SAP Code / Item No. in column E for every selected row;
- dates in `dd-MM-yyyy`, `dd/MM/yyyy`, `dd.MM.yyyy`, or `yyyy-MM-dd` format.
- Excel quoted cells and escaped quotes; embedded tabs, line breaks, and NUL characters inside cells are rejected to prevent item-column misalignment.

The helper keeps the validated copy in memory and performs five actions from one F8 press. Actions 1–4 use the first row only: Excel B goes to Supplier, C goes to Posting Date as `dd.MM.yy`, and D goes to both Supplier Ref. and Remarks. No Tab key is sent, and SAP updates Document Date automatically. After each header paste the helper waits 0.8 seconds asynchronously, then moves to the next calibrated position. Action 5 clicks First Item No. and pastes all copied E:N rows once as one 10-column matrix. Empty item cells remain in their original positions, whether the copy contains 2 rows or dozens.

The helper keeps the parsed invoice in memory, prepares each of the five paste payloads once, and restores the copied B:N as plain text once at the end. This reduces a successful run from ten clipboard writes to six. It uses Unicode text directly, without reloading or restoring Excel's heavier OLE clipboard formats. Excel delayed text rendering is read in the background; clipboard contention retries yield asynchronously for up to eight seconds rather than sleeping on the UI thread.

The final Items payload remains available for a bounded row-count guard (0.8–5 seconds) before restoration. These waits do not prove that SAP accepted the input. Review every completed or partial invoice; never rerun over already populated Items without checking them. If you copy something new during a run, the helper stops at the next check and preserves your new clipboard content.

Choose **Stop Paste** from the tray menu to cancel before the next field. A paste already sent cannot be undone; its short payload guard finishes before the helper restores the clipboard. Opening Calibration or the main window is blocked while automation is active so the helper does not take focus from SAP.

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
& "${env:ProgramFiles(x86)}\NSIS\makensis.exe" /DAPP_VERSION=0.1.0-beta.15 .\installer\SapB1ExcelHelper.nsi
```

Optional native clipboard integration verification:

```powershell
dotnet run --project .\tests\SapB1ExcelHelper.SmokeTests\SapB1ExcelHelper.SmokeTests.csproj -c Release -- --native-clipboard-integration
```

This launches a disposable child process in a new noninteractive window station so the real user's clipboard is never accessed. Some Windows accounts refuse this isolation; the explicit integration run then fails before clipboard access. Ordinary smoke tests print a clear skip notice and cover the parser, workflow, retry/cancellation and safety rules with simulated I/O instead. Do not treat a skipped native test or passing simulations as real SAP acceptance verification.
The Windows release CI opts into this integration test, so a release build must pass native clipboard verification as well as workflow simulations.

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
- Coordinate-based input cannot prove the focused SAP field or read back the invoice. Same-window dialogs and slow SAP validation can still require intervention. Automated regression tests are not a substitute for testing on the target company PC.
- The self-contained package includes .NET so company PCs do not need to install a runtime. ReadyToRun remains enabled for startup compatibility; this revision reduces runtime work, not the bundled runtime size.
