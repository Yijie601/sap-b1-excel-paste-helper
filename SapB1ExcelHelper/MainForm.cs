using System.Diagnostics;
using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class MainForm : Form
{
    private const int HotkeyId = 0xB101;

    private readonly ExcelClipboardParser _parser = new();
    private readonly CalibrationService _calibrationService = new();
    private readonly HotkeySettingsService _hotkeySettingsService = new();
    private readonly UpdateService _updateService = new();
    private readonly UpdateStateService _updateStateService = new();
    private readonly SapAutomationService _automationService;
    private readonly NotifyIcon _trayIcon;
    private readonly Label _subtitleLabel;
    private readonly Label _statusLabel;
    private readonly Label _invoiceLabel;
    private readonly Label _modeLabel;
    private readonly LinkLabel _hotkeyLabel;
    private readonly Button _runButton;
    private readonly Button _updateButton;
    private readonly ToolStripMenuItem _trayRunItem;

    private HotkeyDefinition _hotkey;
    private InvoiceClipboardData? _preparedInvoice;
    private CalibrationForm? _calibrationForm;
    private string _lastValidationError = "Copy Excel columns B:N first.";
    private uint _lastClipboardSequence;
    private int _clipboardValidationRevision;
    private CancellationTokenSource? _clipboardValidationCancellation;
    private CancellationTokenSource? _runCancellation;
    private bool _needsFreshCopy;
    private bool _automationRunning;
    private bool _allowExit;
    private bool _hotkeyRegistered;
    private bool _updateCheckRunning;
    private bool _updatePromptOpen;

    public MainForm()
    {
        _automationService = new SapAutomationService(() => Handle);
        _hotkey = _hotkeySettingsService.Load();

        Text = "SAP B1 Excel Helper";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(560, 390);
        Size = new Size(640, 455);
        Font = new Font("Segoe UI", 10F);
        BackColor = Color.FromArgb(247, 248, 250);
        Icon = SystemIcons.Application;

        var header = new Label
        {
            Text = "SAP B1 Excel Helper",
            Font = new Font("Segoe UI Semibold", 20F),
            AutoSize = true,
            Location = new Point(28, 22)
        };

        _subtitleLabel = new Label
        {
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoEllipsis = true,
            Location = new Point(31, 68),
            Size = new Size(560, 24),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var statusPanel = new Panel
        {
            BackColor = Color.White,
            Location = new Point(30, 106),
            Size = new Size(562, 145),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        _statusLabel = new Label
        {
            Text = "● Waiting for Excel data",
            Font = new Font("Segoe UI Semibold", 13F),
            ForeColor = Color.FromArgb(171, 110, 0),
            AutoSize = true,
            Location = new Point(20, 17)
        };
        _invoiceLabel = new Label
        {
            Text = "Copy exactly 13 columns: B:N",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoEllipsis = true,
            Location = new Point(22, 54),
            Size = new Size(515, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _modeLabel = new Label
        {
            Text = "Mode: One F8 • 0.8s paste interval",
            AutoSize = true,
            ForeColor = Color.FromArgb(85, 91, 101),
            Location = new Point(22, 92)
        };
        _hotkeyLabel = new LinkLabel
        {
            AutoEllipsis = true,
            Location = new Point(300, 88),
            Size = new Size(240, 28),
            TextAlign = ContentAlignment.MiddleRight,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _hotkeyLabel.LinkClicked += (_, _) => OpenHotkeySettings();
        statusPanel.Controls.AddRange(new Control[]
        {
            _statusLabel, _invoiceLabel, _modeLabel, _hotkeyLabel
        });

        _runButton = CreateButton(string.Empty, 30, 278, 170);
        _runButton.BackColor = Color.FromArgb(31, 106, 68);
        _runButton.ForeColor = Color.White;
        _runButton.FlatAppearance.BorderSize = 0;
        _runButton.Click += async (_, _) => await RunAutomationAsync();

        var hotkeyButton = CreateButton("Hotkey Settings", 210, 278, 180);
        hotkeyButton.Click += (_, _) => OpenHotkeySettings();

        var calibrationButton = CreateButton("Calibration", 400, 278, 192);
        calibrationButton.Click += (_, _) => OpenCalibration();

        var logsButton = CreateButton("Open Logs", 30, 332, 170);
        logsButton.Click += (_, _) => OpenFolder(AppPaths.LogsDirectory);

        _updateButton = CreateButton("Check for Updates", 210, 332, 180);
        _updateButton.Click += async (_, _) => await CheckForUpdatesAsync(userInitiated: true);

        var minimizeButton = CreateButton("Minimize to Tray", 400, 332, 192);
        minimizeButton.Click += (_, _) => MinimizeToTray();

        Controls.AddRange(new Control[]
        {
            header, _subtitleLabel, statusPanel, _runButton,
            hotkeyButton, calibrationButton, logsButton, _updateButton, minimizeButton
        });

        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Open", null, (_, _) => RestoreFromTray());
        _trayRunItem = new ToolStripMenuItem(string.Empty, null, async (_, _) => await RunAutomationAsync());
        trayMenu.Items.Add(_trayRunItem);
        trayMenu.Items.Add("Calibration", null, (_, _) => OpenCalibration());
        trayMenu.Items.Add("Hotkey Settings...", null, (_, _) => OpenHotkeySettings());
        trayMenu.Items.Add("Check for Updates", null, async (_, _) => await CheckForUpdatesAsync(userInitiated: true));
        trayMenu.Items.Add("Open Log", null, (_, _) => OpenFolder(AppPaths.LogsDirectory));
        trayMenu.Items.Add("Stop Paste", null, (_, _) => _runCancellation?.Cancel());
        trayMenu.Items.Add("Open Data Folder", null, (_, _) => OpenFolder(AppPaths.DataDirectory));
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("Exit", null, (_, _) => ExitApplication());

        _trayIcon = new NotifyIcon
        {
            Text = "SAP B1 Excel Helper",
            Icon = SystemIcons.Application,
            ContextMenuStrip = trayMenu,
            Visible = true
        };
        _trayIcon.DoubleClick += (_, _) => RestoreFromTray();
        RefreshHotkeyText();

        Resize += (_, _) =>
        {
            if (WindowState == FormWindowState.Minimized)
            {
                MinimizeToTray();
            }
        };
        FormClosing += OnFormClosing;
        Shown += async (_, _) =>
        {
            await Task.Delay(100);
            await ValidateCurrentClipboardAsync();
            await CheckForUpdatesAsync(userInitiated: false);
        };
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        _hotkeyRegistered = RegisterHotkey(_hotkey);
        _ = NativeMethods.AddClipboardFormatListener(Handle);

        if (!_hotkeyRegistered)
        {
            BeginInvoke(() => ShowHotkeyUnavailable(_hotkey));
        }
    }

    protected override void OnHandleDestroyed(EventArgs e)
    {
        if (_hotkeyRegistered)
        {
            _ = NativeMethods.UnregisterHotKey(Handle, HotkeyId);
            _hotkeyRegistered = false;
        }

        _ = NativeMethods.RemoveClipboardFormatListener(Handle);
        base.OnHandleDestroyed(e);
    }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == NativeMethods.WmHotkey && message.WParam.ToInt32() == HotkeyId)
        {
            _ = RunAutomationAsync();
        }
        else if (message.Msg == NativeMethods.WmClipboardUpdate && !_automationRunning &&
                 ClipboardService.Sequence != _lastClipboardSequence)
        {
            _ = ValidateClipboardAfterDelayAsync();
        }

        base.WndProc(ref message);
    }

    private static Button CreateButton(string text, int x, int y, int width) => new()
    {
        Text = text,
        Location = new Point(x, y),
        Size = new Size(width, 42),
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Cursor = Cursors.Hand
    };

    private async Task ValidateClipboardAfterDelayAsync()
    {
        _clipboardValidationCancellation?.Cancel();
        var cancellation = new CancellationTokenSource();
        _clipboardValidationCancellation = cancellation;
        var revision = ++_clipboardValidationRevision;
        try
        {
            await Task.Delay(60, cancellation.Token);
            if (_automationRunning || IsDisposed || revision != _clipboardValidationRevision)
            {
                return;
            }

            await ValidateCurrentClipboardAsync(cancellation.Token, revision: revision);
        }
        catch (OperationCanceledException)
        {
        }
        finally
        {
            if (ReferenceEquals(_clipboardValidationCancellation, cancellation))
            {
                _clipboardValidationCancellation = null;
            }

            cancellation.Dispose();
        }
    }

    private async Task<bool> ValidateCurrentClipboardAsync(
        CancellationToken cancellationToken = default,
        bool duringRun = false,
        int? revision = null)
    {
        bool IsCurrentRequest() => !IsDisposed && !Disposing && !cancellationToken.IsCancellationRequested &&
            (duringRun || !_automationRunning) &&
            (revision is null || revision == _clipboardValidationRevision);

        try
        {
            var clipboard = await ClipboardService.ReadTextAsync(cancellationToken);
            if (!IsCurrentRequest())
            {
                return false;
            }

            if (clipboard.Sequence != ClipboardService.Sequence)
            {
                if (!duringRun)
                {
                    _ = ValidateClipboardAfterDelayAsync();
                }
                return false;
            }

            if (clipboard.Text is null)
            {
                SetNotReady("Clipboard does not contain Excel text.");
                return false;
            }

            var invoice = _parser.Parse(clipboard.Text);
            _preparedInvoice = invoice;
            _lastClipboardSequence = clipboard.Sequence;
            _needsFreshCopy = false;
            _lastValidationError = string.Empty;
            SetReady(invoice);
            return true;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (ClipboardValidationException exception)
        {
            if (IsCurrentRequest())
            {
                SetNotReady(exception.Message);
            }
            return false;
        }
        catch (Exception exception)
        {
            if (IsCurrentRequest())
            {
                SetNotReady($"Unable to read clipboard: {exception.Message}");
            }
            return false;
        }
    }

    private async Task RunAutomationAsync()
    {
        if (_automationRunning || _updatePromptOpen)
        {
            return;
        }

        if (_calibrationForm is { IsDisposed: false })
        {
            _calibrationForm.Show();
            _calibrationForm.Activate();
            ShowError("Finish and save Calibration before running the SAP paste.");
            return;
        }

        _clipboardValidationCancellation?.Cancel();
        ++_clipboardValidationRevision;
        _automationRunning = true;
        _runButton.Enabled = false;
        _trayRunItem.Enabled = false;
        using var cancellation = new CancellationTokenSource();
        _runCancellation = cancellation;
        InvoiceClipboardData? invoice = null;
        var runStarted = false;
        var started = Stopwatch.StartNew();

        try
        {
            if (_preparedInvoice is null || ClipboardService.Sequence != _lastClipboardSequence)
            {
                if (!await ValidateCurrentClipboardAsync(cancellation.Token, duringRun: true))
                {
                    if (!cancellation.IsCancellationRequested && !IsDisposed && !Disposing)
                    {
                        ShowError(string.IsNullOrEmpty(_lastValidationError)
                            ? "The clipboard changed during validation. Copy Excel B:N again."
                            : _lastValidationError);
                    }
                    return;
                }
            }

            if (_needsFreshCopy)
            {
                MessageBox.Show("This copy was already used. Check SAP, then copy Excel B:N again for the next run.",
                    "Check SAP first", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var calibration = _calibrationService.Load();
            if (!calibration.IsComplete)
            {
                ShowError($"Capture these SAP positions in Calibration first:\r\n\r\n{string.Join(", ", calibration.MissingFields)}");
                return;
            }

            invoice = _preparedInvoice!;
            SetWorking("Starting five-step SAP paste...");
            if (Visible)
            {
                MinimizeToTray();
                await Task.Delay(250, cancellation.Token);
            }

            runStarted = true;
            _needsFreshCopy = true;
            var result = await _automationService.RunAllAsync(
                invoice,
                calibration,
                _lastClipboardSequence,
                new Progress<string>(message =>
                {
                    if (_automationRunning && !IsDisposed && !Disposing)
                    {
                        SetWorking(message);
                    }
                }),
                cancellation.Token);

            AppLogger.Success(
                invoice.SupplierName,
                invoice.DocumentNumber,
                result.ItemRows,
                result.Duration);
            if (IsDisposed || Disposing)
            {
                return;
            }
            SetReady(invoice, $"Pasted in {result.Duration.TotalSeconds:0.00}s — check SAP; copy B:N again for the next run.");
            ShowSuccess(invoice, result);
        }
        catch (OperationCanceledException)
        {
            if (!IsDisposed && !Disposing)
            {
                SetNotReady("Paste stopped. Check SAP before copying B:N again.", keepInvoice: true);
            }
        }
        catch (Exception exception)
        {
            started.Stop();
            AppLogger.Failure(
                invoice?.SupplierName,
                invoice?.DocumentNumber,
                invoice?.Items.Count ?? 0,
                started.Elapsed,
                exception.Message);
            if (!IsDisposed && !Disposing)
            {
                SetNotReady($"Stopped: {exception.Message}", keepInvoice: true);
                ShowError(exception.Message);
            }
        }
        finally
        {
            if (runStarted)
            {
                _lastClipboardSequence = _automationService.LastClipboardSequence;
            }

            _runCancellation = null;
            _automationRunning = false;
            if (!IsDisposed)
            {
                _runButton.Enabled = true;
                _trayRunItem.Enabled = true;
                if (ClipboardService.Sequence != _lastClipboardSequence)
                {
                    _ = ValidateClipboardAfterDelayAsync();
                }
            }
        }
    }

    private async Task CheckForUpdatesAsync(bool userInitiated)
    {
        if (_updateCheckRunning)
        {
            return;
        }

        if (_automationRunning)
        {
            if (userInitiated)
            {
                _trayIcon.ShowBalloonTip(1800,
                    "Update check",
                    "Wait for the current SAP paste to finish before checking for updates.",
                    ToolTipIcon.Info);
            }

            return;
        }

        if (!userInitiated &&
            !_updateStateService.IsAutomaticCheckDue(TimeSpan.FromHours(12)))
        {
            return;
        }

        _updateCheckRunning = true;
        _updateButton.Enabled = false;
        var checkAttempted = false;
        try
        {
            checkAttempted = true;
            var update = await _updateService.CheckForUpdateAsync();
            if (_automationRunning || IsDisposed || Disposing)
            {
                checkAttempted = false;
                return;
            }
            if (update is null)
            {
                if (userInitiated)
                {
                    MessageBox.Show(
                        $"You already have the newest version ({UpdateService.CurrentVersion}).",
                        "No update available",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return;
            }

            string? installerPath;
            _updatePromptOpen = true;
            try
            {
                using var updateForm = new UpdateForm(update, _updateService);
                var result = Visible
                    ? updateForm.ShowDialog(this)
                    : updateForm.ShowDialog();
                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(updateForm.InstallerPath))
                {
                    return;
                }

                installerPath = updateForm.InstallerPath;
            }
            finally
            {
                _updatePromptOpen = false;
            }

            try
            {
                Process.Start(new ProcessStartInfo(installerPath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception exception)
            {
                AppLogger.Error("UPDATE_INSTALLER_LAUNCH_ERROR", exception.Message, exception);
                MessageBox.Show(
                    $"The update was downloaded and verified, but the setup wizard could not be opened.\r\n\r\n{exception.Message}",
                    "Unable to open update",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            _allowExit = true;
            _trayIcon.Visible = false;
            Application.Exit();
        }
        catch (Exception exception)
        {
            AppLogger.Error("UPDATE_CHECK_ERROR", exception.Message, exception);
            if (userInitiated && !_automationRunning && !IsDisposed && !Disposing)
            {
                MessageBox.Show(
                    $"Unable to check for updates.\r\n\r\n{exception.Message}",
                    "Update check failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        finally
        {
            if (checkAttempted)
            {
                _updateStateService.RecordCheck();
            }

            _updateCheckRunning = false;
            if (!IsDisposed)
            {
                _updateButton.Enabled = true;
            }
        }
    }

    private void SetReady(InvoiceClipboardData invoice, string? detail = null)
    {
        _statusLabel.Text = _needsFreshCopy ? "● Finished — copy B:N again" : "● Ready";
        _statusLabel.ForeColor = Color.FromArgb(22, 125, 72);
        if (detail is not null)
        {
            _invoiceLabel.Text = detail;
            return;
        }

        _invoiceLabel.Text =
            $"{invoice.DocumentNumber}  •  {invoice.SupplierName}  •  {invoice.Items.Count} row(s)  •  {invoice.SapDate}";
    }

    private void SetNotReady(string message, bool keepInvoice = false)
    {
        if (!keepInvoice)
        {
            _preparedInvoice = null;
        }

        _lastValidationError = message;
        _statusLabel.Text = "● Not ready";
        _statusLabel.ForeColor = Color.FromArgb(177, 63, 48);
        _invoiceLabel.Text = message;
    }

    private void SetWorking(string message)
    {
        _statusLabel.Text = "● Working";
        _statusLabel.ForeColor = Color.FromArgb(33, 99, 178);
        _invoiceLabel.Text = message;
    }

    private void ShowSuccess(InvoiceClipboardData invoice, AutomationResult result)
    {
        _trayIcon.BalloonTipTitle = $"✓ {invoice.DocumentNumber}";
        _trayIcon.BalloonTipText = $"{result.ItemRows} item row(s) pasted in {result.Duration.TotalSeconds:0.00}s. Check SAP before Add.";
        _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
        _trayIcon.ShowBalloonTip(1800);
    }

    private static void ShowError(string message) => MessageBox.Show(
        message,
        "SAP B1 Excel Helper",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);

    private void OpenCalibration()
    {
        if (_automationRunning)
        {
            return;
        }

        if (_calibrationForm is { IsDisposed: false })
        {
            _calibrationForm.Show();
            _calibrationForm.Activate();
            return;
        }

        var form = new CalibrationForm(_calibrationService);
        _calibrationForm = form;
        form.FormClosed += (_, _) =>
        {
            if (ReferenceEquals(_calibrationForm, form))
            {
                _calibrationForm = null;
            }
        };
        form.Show(this);
    }

    private void OpenHotkeySettings()
    {
        if (_automationRunning)
        {
            _trayIcon.ShowBalloonTip(1800,
                "Hotkey settings",
                "Wait for the current SAP paste to finish before changing the hotkey.",
                ToolTipIcon.Info);
            return;
        }

        var originalHotkey = _hotkey;
        UnregisterHotkey();

        using var form = new HotkeySettingsForm(originalHotkey);
        if (form.ShowDialog(this) != DialogResult.OK)
        {
            _hotkeyRegistered = RegisterHotkey(originalHotkey);
            if (!_hotkeyRegistered)
            {
                ShowHotkeyUnavailable(originalHotkey);
            }

            return;
        }

        var selectedHotkey = form.SelectedHotkey;
        if (!RegisterHotkey(selectedHotkey))
        {
            _hotkeyRegistered = RegisterHotkey(originalHotkey);
            ShowHotkeyUnavailable(selectedHotkey);
            return;
        }

        _hotkeyRegistered = true;
        try
        {
            _hotkeySettingsService.Save(selectedHotkey);
            _hotkey = selectedHotkey;
            RefreshHotkeyText();
        }
        catch (Exception exception)
        {
            AppLogger.Error("HOTKEY_SAVE_ERROR", exception.Message, exception);
            UnregisterHotkey();
            _hotkeyRegistered = RegisterHotkey(originalHotkey);
            MessageBox.Show(
                $"The hotkey could not be saved. The previous shortcut remains active.\r\n\r\n{exception.Message}",
                "Hotkey settings",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private bool RegisterHotkey(HotkeyDefinition hotkey) =>
        NativeMethods.RegisterHotKey(
            Handle,
            HotkeyId,
            (uint)hotkey.Modifiers | NativeMethods.ModNoRepeat,
            hotkey.VirtualKey);

    private void UnregisterHotkey()
    {
        if (!_hotkeyRegistered)
        {
            return;
        }

        _ = NativeMethods.UnregisterHotKey(Handle, HotkeyId);
        _hotkeyRegistered = false;
    }

    private void RefreshHotkeyText()
    {
        var hotkeyText = _hotkey.DisplayText;
        _subtitleLabel.Text = $"Copy Excel B:N, switch to SAP AP Invoice, then press {hotkeyText} once.";
        _hotkeyLabel.Text = $"Hotkey: {hotkeyText}";
        _runButton.Text = $"Run Now ({hotkeyText})";
        _trayRunItem.Text = $"Run Now ({hotkeyText})";
    }

    private static void ShowHotkeyUnavailable(HotkeyDefinition hotkey) => MessageBox.Show(
        $"{hotkey.DisplayText} is already being used by Windows or another program. Choose a different shortcut in Hotkey Settings.",
        "Hotkey unavailable",
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning);

    private static void OpenFolder(string folder)
    {
        Directory.CreateDirectory(folder);
        Process.Start(new ProcessStartInfo("explorer.exe", folder) { UseShellExecute = true });
    }

    private void MinimizeToTray()
    {
        Hide();
        ShowInTaskbar = false;
    }

    private void RestoreFromTray()
    {
        if (_automationRunning)
        {
            return;
        }

        ShowInTaskbar = true;
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void ExitApplication()
    {
        if (_automationRunning)
        {
            _runCancellation?.Cancel();
            return;
        }

        _allowExit = true;
        _trayIcon.Visible = false;
        Close();
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs eventArgs)
    {
        _clipboardValidationCancellation?.Cancel();
        if (!_allowExit && eventArgs.CloseReason == CloseReason.UserClosing)
        {
            eventArgs.Cancel = true;
            MinimizeToTray();
            return;
        }

        _runCancellation?.Cancel();

        if (_calibrationForm is { IsDisposed: false })
        {
            _calibrationForm.Close();
        }

        _trayIcon.Visible = false;
        _trayIcon.Dispose();
    }
}
