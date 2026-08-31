using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class CalibrationForm : Form
{
    private readonly CalibrationService _calibrationService;
    private readonly MouseCaptureService _mouseCaptureService = new();
    private readonly Dictionary<string, Label> _coordinateLabels = new(StringComparer.Ordinal);
    private readonly Label _statusLabel;
    private SapCalibration _calibration;
    private bool _capturing;

    public CalibrationForm(CalibrationService calibrationService)
    {
        _calibrationService = calibrationService;
        var savedCalibration = calibrationService.Load().Clone();
        _calibration = savedCalibration.CoordinateVersion == SapCalibration.AbsoluteDesktopCoordinateVersion
            ? savedCalibration
            : new SapCalibration();

        Text = "SAP Calibration";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(590, 530);
        MinimumSize = new Size(590, 530);
        MaximizeBox = false;
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "SAP AP Invoice Calibration",
            Font = new Font("Segoe UI Semibold", 17F),
            AutoSize = true,
            Location = new Point(22, 18)
        };
        var instructions = new Label
        {
            Text = "For each field, click Capture and then click the real input position in SAP.\nSAP must stay at the same desktop position after calibration.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoSize = true,
            Location = new Point(25, 58)
        };

        var table = new TableLayoutPanel
        {
            ColumnCount = 3,
            RowCount = 7,
            Location = new Point(25, 115),
            Size = new Size(525, 255),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        table.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
        for (var row = 1; row < 7; row++)
        {
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        }
        AddHeader(table, "SAP field", "Desktop X / Y", "Action");
        AddCalibrationRow(table, 1, "Supplier", point =>
        {
            _calibration.Supplier = point;
            _calibration.SupplierCaptured = true;
        });
        AddCalibrationRow(table, 2, "Supplier Ref.", point =>
        {
            _calibration.SupplierRef = point;
            _calibration.SupplierRefCaptured = true;
        });
        AddCalibrationRow(table, 3, "Posting Date", point =>
        {
            _calibration.PostingDate = point;
            _calibration.PostingDateCaptured = true;
        });
        AddCalibrationRow(table, 4, "Document Date", point =>
        {
            _calibration.DocumentDate = point;
            _calibration.DocumentDateCaptured = true;
        });
        AddCalibrationRow(table, 5, "Remarks", point =>
        {
            _calibration.Remarks = point;
            _calibration.RemarksCaptured = true;
        });
        AddCalibrationRow(table, 6, "First Item No.", point =>
        {
            _calibration.ItemNo = point;
            _calibration.ItemNoCaptured = true;
        });

        _statusLabel = new Label
        {
            Text = "Ready to capture.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoEllipsis = true,
            Location = new Point(27, 382),
            Size = new Size(520, 26)
        };

        var testButton = CreateButton("Test Calibration", 25, 420, 160);
        testButton.Click += async (_, _) => await TestCalibrationAsync();
        var resetButton = CreateButton("Clear All", 195, 420, 145);
        resetButton.Click += (_, _) =>
        {
            _calibration = new SapCalibration();
            RefreshCoordinateLabels();
            _statusLabel.Text = "All captured positions cleared.";
        };
        var saveButton = CreateButton("Save", 350, 420, 95);
        saveButton.BackColor = Color.FromArgb(31, 106, 68);
        saveButton.ForeColor = Color.White;
        saveButton.FlatAppearance.BorderSize = 0;
        saveButton.Click += (_, _) =>
        {
            if (!EnsureComplete("saving"))
            {
                return;
            }

            _calibrationService.Save(_calibration);
            DialogResult = DialogResult.OK;
            Close();
        };
        var cancelButton = CreateButton("Cancel", 455, 420, 95);
        cancelButton.Click += (_, _) => Close();

        Controls.AddRange(new Control[]
        {
            title, instructions, table, _statusLabel,
            testButton, resetButton, saveButton, cancelButton
        });
        RefreshCoordinateLabels();
    }

    private void AddCalibrationRow(
        TableLayoutPanel table,
        int row,
        string name,
        Action<SapPoint> setPoint)
    {
        var nameLabel = new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left };
        var coordinateLabel = new Label { AutoSize = true, Anchor = AnchorStyles.Left };
        var captureButton = new Button
        {
            Text = "Capture",
            Size = new Size(125, 30),
            Anchor = AnchorStyles.Left,
            Margin = new Padding(4, 3, 4, 3)
        };
        captureButton.Click += async (_, _) => await CaptureAsync(name, setPoint);
        table.Controls.Add(nameLabel, 0, row);
        table.Controls.Add(coordinateLabel, 1, row);
        table.Controls.Add(captureButton, 2, row);
        _coordinateLabels[name] = coordinateLabel;
    }

    private static void AddHeader(TableLayoutPanel table, string first, string second, string third)
    {
        table.Controls.Add(HeaderLabel(first), 0, 0);
        table.Controls.Add(HeaderLabel(second), 1, 0);
        table.Controls.Add(HeaderLabel(third), 2, 0);
    }

    private static Label HeaderLabel(string text) => new()
    {
        Text = text,
        Font = new Font("Segoe UI Semibold", 10F),
        AutoSize = true,
        Anchor = AnchorStyles.Left
    };

    private static Button CreateButton(string text, int x, int y, int width) => new()
    {
        Text = text,
        Location = new Point(x, y),
        Size = new Size(width, 40),
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Cursor = Cursors.Hand
    };

    private async Task CaptureAsync(string fieldName, Action<SapPoint> setPoint)
    {
        if (_capturing)
        {
            return;
        }

        _capturing = true;
        using var instruction = new CaptureInstructionForm($"Click the {fieldName} field in SAP now");
        try
        {
            _statusLabel.Text = $"Waiting for the {fieldName} click...";
            Hide();
            instruction.Show();
            await Task.Delay(150);

            var click = await _mouseCaptureService.CaptureNextLeftClickAsync(TimeSpan.FromSeconds(20));
            if (click is null)
            {
                _statusLabel.Text = "Capture timed out. Try again.";
                return;
            }

            if (!SystemInformation.VirtualScreen.Contains(click.Value))
            {
                _statusLabel.Text = "The click was outside the current Windows desktop.";
                return;
            }

            _calibration.CoordinateVersion = SapCalibration.AbsoluteDesktopCoordinateVersion;
            var absolutePoint = new SapPoint { X = click.Value.X, Y = click.Value.Y };
            setPoint(absolutePoint);
            RefreshCoordinateLabels();
            _statusLabel.Text = $"Captured {fieldName}: desktop X={absolutePoint.X}, Y={absolutePoint.Y}.";
        }
        finally
        {
            instruction.Close();
            RestoreAfterTemporaryHide();
            _capturing = false;
        }
    }

    private async Task TestCalibrationAsync()
    {
        if (!EnsureComplete("testing"))
        {
            return;
        }

        var points = new[]
        {
            ("Supplier", _calibration.Supplier),
            ("Supplier Ref.", _calibration.SupplierRef),
            ("Posting Date", _calibration.PostingDate),
            ("Document Date", _calibration.DocumentDate),
            ("Remarks", _calibration.Remarks),
            ("First Item No.", _calibration.ItemNo)
        };

        using var instruction = new CaptureInstructionForm("Testing calibration — mouse movement only, no clicks");
        Hide();
        instruction.Show();
        try
        {
            foreach (var (_, point) in points)
            {
                _ = NativeMethods.SetCursorPos(point.X, point.Y);
                await Task.Delay(650);
            }
        }
        finally
        {
            instruction.Close();
            RestoreAfterTemporaryHide();
            if (!IsDisposed && !Disposing)
            {
                _statusLabel.Text = "Test complete. No fields were clicked or changed.";
            }
        }
    }

    private void RefreshCoordinateLabels()
    {
        SetCoordinate("Supplier", _calibration.Supplier, _calibration.SupplierCaptured);
        SetCoordinate("Supplier Ref.", _calibration.SupplierRef, _calibration.SupplierRefCaptured);
        SetCoordinate("Posting Date", _calibration.PostingDate, _calibration.PostingDateCaptured);
        SetCoordinate("Document Date", _calibration.DocumentDate, _calibration.DocumentDateCaptured);
        SetCoordinate("Remarks", _calibration.Remarks, _calibration.RemarksCaptured);
        SetCoordinate("First Item No.", _calibration.ItemNo, _calibration.ItemNoCaptured);
    }

    private void SetCoordinate(string name, SapPoint point, bool captured)
    {
        var label = _coordinateLabels[name];
        label.Text = captured ? $"{point.X}, {point.Y}  ✓" : "Not captured";
        label.ForeColor = captured
            ? Color.FromArgb(22, 125, 72)
            : Color.FromArgb(177, 63, 48);
    }

    private bool EnsureComplete(string action)
    {
        if (_calibration.IsComplete)
        {
            return true;
        }

        var missing = string.Join(", ", _calibration.MissingFields);
        MessageBox.Show(
            $"Capture every SAP position before {action}.\r\n\r\nStill missing: {missing}",
            "Calibration incomplete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return false;
    }

    private void RestoreAfterTemporaryHide()
    {
        if (IsDisposed || Disposing)
        {
            return;
        }

        if (!Visible)
        {
            Show();
        }

        Activate();
    }
}

internal sealed class CaptureInstructionForm : Form
{
    public CaptureInstructionForm(string instruction)
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(26, 36, 52);
        ForeColor = Color.White;
        Size = new Size(430, 62);
        StartPosition = FormStartPosition.Manual;

        var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Location = new Point(
            workingArea.Left + (workingArea.Width - Width) / 2,
            workingArea.Top + 18);
        Controls.Add(new Label
        {
            Text = instruction,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI Semibold", 11F),
            ForeColor = Color.White
        });
    }

    protected override bool ShowWithoutActivation => true;
}
