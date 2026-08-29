using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class HotkeySettingsForm : Form
{
    private readonly TextBox _captureBox;
    private readonly Label _validationLabel;
    private readonly Button _saveButton;

    public HotkeySettingsForm(HotkeyDefinition currentHotkey)
    {
        SelectedHotkey = currentHotkey;

        Text = "Hotkey Settings";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(500, 315);
        MinimumSize = new Size(500, 315);
        MaximizeBox = false;
        MinimizeBox = false;
        Font = new Font("Segoe UI", 10F);
        BackColor = Color.FromArgb(247, 248, 250);

        var title = new Label
        {
            Text = "Choose the global paste hotkey",
            Font = new Font("Segoe UI Semibold", 17F),
            AutoSize = true,
            Location = new Point(24, 20)
        };
        var instructions = new Label
        {
            Text = "Click the box, then press your preferred shortcut.\nF1–F24 can be used alone; other keys require Ctrl or Alt.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoSize = true,
            Location = new Point(27, 61)
        };

        _captureBox = new TextBox
        {
            Text = currentHotkey.DisplayText,
            ReadOnly = true,
            ShortcutsEnabled = false,
            TextAlign = HorizontalAlignment.Center,
            Font = new Font("Segoe UI Semibold", 16F),
            Location = new Point(28, 118),
            Size = new Size(425, 36),
            TabIndex = 0
        };
        _captureBox.KeyDown += CaptureBoxOnKeyDown;
        _captureBox.MouseDown += (_, _) => _captureBox.SelectAll();

        _validationLabel = new Label
        {
            Text = "The shortcut works globally while the helper is running.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoEllipsis = true,
            Location = new Point(29, 163),
            Size = new Size(425, 25)
        };

        var resetButton = CreateButton("Reset to F8", 28, 208, 125);
        resetButton.Click += (_, _) => SetSelectedHotkey(HotkeyDefinition.Default);

        _saveButton = CreateButton("Save", 267, 208, 88);
        _saveButton.BackColor = Color.FromArgb(31, 106, 68);
        _saveButton.ForeColor = Color.White;
        _saveButton.FlatAppearance.BorderSize = 0;
        _saveButton.Click += (_, _) =>
        {
            DialogResult = DialogResult.OK;
            Close();
        };

        var cancelButton = CreateButton("Cancel", 365, 208, 88);
        cancelButton.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };

        AcceptButton = _saveButton;
        CancelButton = cancelButton;
        Controls.AddRange(new Control[]
        {
            title, instructions, _captureBox, _validationLabel,
            resetButton, _saveButton, cancelButton
        });
        Shown += (_, _) =>
        {
            _captureBox.Focus();
            _captureBox.SelectAll();
        };
    }

    public HotkeyDefinition SelectedHotkey { get; private set; }

    private static Button CreateButton(string text, int x, int y, int width) => new()
    {
        Text = text,
        Location = new Point(x, y),
        Size = new Size(width, 40),
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Cursor = Cursors.Hand
    };

    private void CaptureBoxOnKeyDown(object? sender, KeyEventArgs eventArgs)
    {
        eventArgs.Handled = true;
        eventArgs.SuppressKeyPress = true;

        if (HotkeyDefinition.TryCreate(
                eventArgs.KeyCode,
                eventArgs.Modifiers,
                out var hotkey,
                out var error))
        {
            SetSelectedHotkey(hotkey!);
            return;
        }

        _validationLabel.Text = error;
        _validationLabel.ForeColor = Color.FromArgb(177, 63, 48);
    }

    private void SetSelectedHotkey(HotkeyDefinition hotkey)
    {
        SelectedHotkey = hotkey;
        _captureBox.Text = hotkey.DisplayText;
        _captureBox.SelectAll();
        _validationLabel.Text = "Ready to save. Windows availability will be checked next.";
        _validationLabel.ForeColor = Color.FromArgb(22, 125, 72);
        _saveButton.Enabled = true;
    }
}
