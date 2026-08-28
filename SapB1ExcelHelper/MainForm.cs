using System.Diagnostics;
using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class MainForm : Form
{
    private const int HotkeyId = 0xB101;
    private const uint VirtualKeyF8 = 0x77;

    private readonly ExcelClipboardParser _parser = new();
    private readonly SupplierMappingService _mappingService = new();
    private readonly CalibrationService _calibrationService = new();
    private readonly SapWindowService _windowService = new();
    private readonly SapAutomationService _automationService;
    private readonly NotifyIcon _trayIcon;
    private readonly System.Windows.Forms.Timer _sapStatusTimer;
    private readonly Label _statusLabel;
    private readonly Label _invoiceLabel;
    private readonly Label _sapLabel;
    private readonly Label _mappingCountLabel;
    private readonly Button _runButton;

    private InvoiceClipboardData? _preparedInvoice;
    private string _lastValidationError = "Copy Excel columns B:N first.";
    private DateTime _ignoreClipboardUntilUtc;
    private bool _automationRunning;
    private bool _allowExit;
    private bool _hotkeyRegistered;

    public MainForm()
    {
        _automationService = new SapAutomationService(_windowService);

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

        var subtitle = new Label
        {
            Text = "Copy Excel B:N. When Ready appears, switch to SAP AP Invoice and press F8.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoSize = true,
            Location = new Point(31, 68)
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
        _sapLabel = new Label
        {
            Text = "SAP: Not active",
            AutoSize = true,
            Location = new Point(22, 92)
        };
        _mappingCountLabel = new Label
        {
            Text = $"Supplier mappings: {_mappingService.Count}",
            AutoSize = true,
            Location = new Point(240, 92)
        };
        var hotkeyLabel = new Label
        {
            Text = "Hotkey: F8",
            AutoSize = true,
            Location = new Point(445, 92),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        statusPanel.Controls.AddRange(new Control[]
        {
            _statusLabel, _invoiceLabel, _sapLabel, _mappingCountLabel, hotkeyLabel
        });

        _runButton = CreateButton("Run Now (F8)", 30, 278, 170);
        _runButton.BackColor = Color.FromArgb(31, 106, 68);
        _runButton.ForeColor = Color.White;
        _runButton.FlatAppearance.BorderSize = 0;
        _runButton.Click += async (_, _) => await RunAutomationAsync();

        var mappingButton = CreateButton("Supplier Mapping", 210, 278, 180);
        mappingButton.Click += (_, _) => OpenSupplierMappings();

        var calibrationButton = CreateButton("Calibration", 400, 278, 192);
        calibrationButton.Click += (_, _) => OpenCalibration();

        var logsButton = CreateButton("Open Logs", 30, 332, 170);
        logsButton.Click += (_, _) => OpenFolder(AppPaths.LogsDirectory);

        var dataButton = CreateButton("Open Data Folder", 210, 332, 180);
        dataButton.Click += (_, _) => OpenFolder(AppPaths.DataDirectory);

        var minimizeButton = CreateButton("Minimize to Tray", 400, 332, 192);
        minimizeButton.Click += (_, _) => MinimizeToTray();

        Controls.AddRange(new Control[]
        {
            header, subtitle, statusPanel, _runButton,
            mappingButton, calibrationButton, logsButton, dataButton, minimizeButton
        });

        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Open", null, (_, _) => RestoreFromTray());
        trayMenu.Items.Add("Run Now (F8)", null, async (_, _) => await RunAutomationAsync());
        trayMenu.Items.Add("Supplier Mapping", null, (_, _) => OpenSupplierMappings());
        trayMenu.Items.Add("Calibration", null, (_, _) => OpenCalibration());
        trayMenu.Items.Add("Open Log", null, (_, _) => OpenFolder(AppPaths.LogsDirectory));
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

        _sapStatusTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _sapStatusTimer.Tick += (_, _) => UpdateSapStatus();
        _sapStatusTimer.Start();

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
            ValidateCurrentClipboard();
        };
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        _hotkeyRegistered = NativeMethods.RegisterHotKey(
            Handle,
            HotkeyId,
            NativeMethods.ModNoRepeat,
            VirtualKeyF8);
        _ = NativeMethods.AddClipboardFormatListener(Handle);

        if (!_hotkeyRegistered)
        {
            BeginInvoke(() => MessageBox.Show(
                "F8 is already being used by another program. Close that program and restart this helper.",
                "Hotkey unavailable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning));
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
        else if (message.Msg == NativeMethods.WmClipboardUpdate &&
                 DateTime.UtcNow >= _ignoreClipboardUntilUtc &&
                 !_automationRunning)
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
        await Task.Delay(45);
        ValidateCurrentClipboard();
    }

    private bool ValidateCurrentClipboard()
    {
        var text = ClipboardService.TryGetText();
        if (text is null)
        {
            SetNotReady("Clipboard does not contain Excel text.");
            return false;
        }

        try
        {
            _mappingService.Reload();
            var invoice = _parser.Parse(text);
            if (!_mappingService.TryResolve(invoice.SupplierName, out var supplierCode))
            {
                throw new ClipboardValidationException($"Supplier mapping not found: {invoice.SupplierName}");
            }

            invoice.SupplierCode = supplierCode;
            _preparedInvoice = invoice;
            _lastValidationError = string.Empty;
            SetReady(invoice);
            return true;
        }
        catch (ClipboardValidationException exception)
        {
            SetNotReady(exception.Message);
            return false;
        }
        catch (Exception exception)
        {
            SetNotReady($"Unable to read clipboard: {exception.Message}");
            return false;
        }
    }

    private async Task RunAutomationAsync()
    {
        if (_automationRunning)
        {
            return;
        }

        var currentText = ClipboardService.TryGetText();
        if (_preparedInvoice is null ||
            currentText is null ||
            !string.Equals(currentText, _preparedInvoice.OriginalClipboardText, StringComparison.Ordinal))
        {
            ValidateCurrentClipboard();
        }

        if (_preparedInvoice is null)
        {
            ShowError(_lastValidationError);
            return;
        }

        if (!_windowService.TryGetActiveApInvoice(out var sapWindow, out var windowError))
        {
            ShowError(windowError);
            return;
        }

        var invoice = _preparedInvoice;
        _automationRunning = true;
        _runButton.Enabled = false;
        _ignoreClipboardUntilUtc = DateTime.UtcNow.AddSeconds(6);
        SetWorking("Starting SAP paste...");
        var started = Stopwatch.StartNew();

        try
        {
            var calibration = _calibrationService.Load();
            var result = await _automationService.RunAsync(
                invoice,
                sapWindow!,
                calibration,
                message => SetWorking(message));

            AppLogger.Success(
                invoice.SupplierName,
                invoice.SupplierCode,
                invoice.DocumentNumber,
                result.ItemRows,
                result.Duration);
            SetReady(invoice, $"Completed in {result.Duration.TotalSeconds:0.00}s — check SAP before Add");
            ShowSuccess(invoice, result);
        }
        catch (Exception exception)
        {
            started.Stop();
            AppLogger.Failure(
                invoice.SupplierName,
                invoice.SupplierCode,
                invoice.DocumentNumber,
                invoice.Items.Count,
                started.Elapsed,
                exception.Message);
            SetNotReady($"Stopped: {exception.Message}", keepInvoice: true);
            ShowError(exception.Message);
        }
        finally
        {
            _automationRunning = false;
            _runButton.Enabled = true;
            _ignoreClipboardUntilUtc = DateTime.UtcNow.AddSeconds(2);
        }
    }

    private void SetReady(InvoiceClipboardData invoice, string? detail = null)
    {
        _statusLabel.Text = "● Ready";
        _statusLabel.ForeColor = Color.FromArgb(22, 125, 72);
        _invoiceLabel.Text = detail ??
            $"{invoice.DocumentNumber}  •  {invoice.SupplierName} → {invoice.SupplierCode}  •  {invoice.Items.Count} row(s)  •  {invoice.SapDate}";
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

    private void UpdateSapStatus()
    {
        var sapIsActive = _windowService.IsSapForeground();
        _sapLabel.Text = sapIsActive ? "SAP: Active" : "SAP: Not active";
        _sapLabel.ForeColor = sapIsActive
            ? Color.FromArgb(22, 125, 72)
            : Color.FromArgb(85, 91, 101);
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

    private void OpenSupplierMappings()
    {
        using var form = new SupplierMappingForm(_mappingService);
        if (form.ShowDialog(this) == DialogResult.OK)
        {
            _mappingService.Reload();
            _mappingCountLabel.Text = $"Supplier mappings: {_mappingService.Count}";
            ValidateCurrentClipboard();
        }
    }

    private void OpenCalibration()
    {
        using var form = new CalibrationForm(_calibrationService, _windowService);
        form.ShowDialog(this);
    }

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
        ShowInTaskbar = true;
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void ExitApplication()
    {
        _allowExit = true;
        _trayIcon.Visible = false;
        Close();
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs eventArgs)
    {
        if (!_allowExit && eventArgs.CloseReason == CloseReason.UserClosing)
        {
            eventArgs.Cancel = true;
            MinimizeToTray();
            return;
        }

        _sapStatusTimer.Stop();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
    }
}
