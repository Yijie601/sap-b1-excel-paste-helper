using System.Diagnostics;
using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class UpdateForm : Form
{
    private readonly AppUpdateInfo _update;
    private readonly UpdateService _updateService;
    private readonly Button _installButton;
    private readonly Button _laterButton;
    private readonly Label _statusLabel;
    private readonly ProgressBar _progressBar;
    private CancellationTokenSource? _downloadCancellation;
    private bool _downloading;

    public UpdateForm(AppUpdateInfo update, UpdateService updateService)
    {
        _update = update;
        _updateService = updateService;

        Text = "SAP B1 Excel Helper Update";
        StartPosition = FormStartPosition.CenterScreen;
        Size = new Size(610, 510);
        MinimumSize = new Size(610, 510);
        MaximizeBox = false;
        MinimizeBox = false;
        Font = new Font("Segoe UI", 10F);
        BackColor = Color.FromArgb(247, 248, 250);
        TopMost = true;

        var title = new Label
        {
            Text = "A new version is available",
            Font = new Font("Segoe UI Semibold", 19F),
            AutoSize = true,
            Location = new Point(26, 22)
        };
        var versions = new Label
        {
            Text = $"Current: {UpdateService.CurrentVersion}     New: {update.Version}",
            AutoSize = true,
            ForeColor = Color.FromArgb(60, 70, 82),
            Location = new Point(29, 69)
        };
        var safety = new Label
        {
            Text = "Nothing will install until you click Download & Install and complete the visible setup wizard.",
            AutoSize = false,
            Size = new Size(545, 43),
            ForeColor = Color.FromArgb(85, 91, 101),
            Location = new Point(29, 99)
        };
        var notesLabel = new Label
        {
            Text = "Release notes",
            Font = new Font("Segoe UI Semibold", 10F),
            AutoSize = true,
            Location = new Point(29, 143)
        };
        var notes = new RichTextBox
        {
            Text = string.IsNullOrWhiteSpace(update.ReleaseNotes)
                ? update.ReleaseName
                : update.ReleaseNotes,
            ReadOnly = true,
            DetectUrls = true,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Location = new Point(30, 170),
            Size = new Size(545, 175),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        notes.LinkClicked += (_, eventArgs) =>
        {
            if (!string.IsNullOrWhiteSpace(eventArgs.LinkText))
            {
                OpenUrl(eventArgs.LinkText);
            }
        };

        var releaseLink = new LinkLabel
        {
            Text = "View this release on GitHub",
            AutoSize = true,
            Location = new Point(30, 355)
        };
        releaseLink.Click += (_, _) => OpenUrl(update.ReleasePage.ToString());

        _statusLabel = new Label
        {
            Text = $"Installer size: {FormatSize(update.InstallerSize)} · SHA-256 verification required",
            AutoEllipsis = true,
            Size = new Size(545, 24),
            ForeColor = Color.FromArgb(85, 91, 101),
            Location = new Point(30, 384)
        };
        _progressBar = new ProgressBar
        {
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Visible = false,
            Location = new Point(30, 413),
            Size = new Size(330, 24)
        };

        _installButton = CreateButton("Download & Install", 372, 403, 135);
        _installButton.BackColor = Color.FromArgb(31, 106, 68);
        _installButton.ForeColor = Color.White;
        _installButton.FlatAppearance.BorderSize = 0;
        _installButton.Click += async (_, _) => await DownloadAndInstallAsync();

        _laterButton = CreateButton("Later", 515, 403, 60);
        _laterButton.Click += (_, _) =>
        {
            if (_downloading)
            {
                _downloadCancellation?.Cancel();
            }
            else
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        };

        Controls.AddRange(new Control[]
        {
            title, versions, safety, notesLabel, notes, releaseLink,
            _statusLabel, _progressBar, _installButton, _laterButton
        });
        FormClosing += OnFormClosing;
    }

    public string? InstallerPath { get; private set; }

    private static Button CreateButton(string text, int x, int y, int width) => new()
    {
        Text = text,
        Location = new Point(x, y),
        Size = new Size(width, 38),
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Cursor = Cursors.Hand
    };

    private async Task DownloadAndInstallAsync()
    {
        if (_downloading)
        {
            return;
        }

        _downloading = true;
        _downloadCancellation = new CancellationTokenSource();
        _installButton.Enabled = false;
        _laterButton.Text = "Cancel";
        _laterButton.Width = 68;
        _laterButton.Left = 507;
        _progressBar.Visible = true;
        _statusLabel.Text = "Downloading update...";

        var progress = new Progress<int>(percentage =>
        {
            _progressBar.Value = Math.Clamp(percentage, 0, 100);
            _statusLabel.Text = percentage < 100
                ? $"Downloading update... {percentage}%"
                : "Verifying downloaded installer...";
        });

        try
        {
            InstallerPath = await _updateService.DownloadInstallerAsync(
                _update,
                progress,
                _downloadCancellation.Token);
            _statusLabel.Text = "Download verified. Opening the setup wizard...";
            _downloading = false;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (OperationCanceledException)
        {
            ResetDownloadControls("Update download cancelled.");
        }
        catch (Exception exception)
        {
            AppLogger.Error("UPDATE_DOWNLOAD_ERROR", exception.Message, exception);
            ResetDownloadControls("Download failed. You can try again or use the GitHub release page.");
            MessageBox.Show(
                exception.Message,
                "Update download failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _downloadCancellation?.Dispose();
            _downloadCancellation = null;
            _downloading = false;
        }
    }

    private void ResetDownloadControls(string status)
    {
        _statusLabel.Text = status;
        _progressBar.Value = 0;
        _progressBar.Visible = false;
        _installButton.Enabled = true;
        _laterButton.Text = "Later";
        _laterButton.Width = 60;
        _laterButton.Left = 515;
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs eventArgs)
    {
        if (!_downloading)
        {
            return;
        }

        _downloadCancellation?.Cancel();
        eventArgs.Cancel = true;
    }

    private static void OpenUrl(string url)
    {
        try
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
                uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                Process.Start(new ProcessStartInfo(uri.ToString()) { UseShellExecute = true });
            }
        }
        catch (Exception exception)
        {
            AppLogger.Error("UPDATE_LINK_ERROR", exception.Message, exception);
        }
    }

    private static string FormatSize(long bytes) => bytes switch
    {
        >= 1024L * 1024L => $"{bytes / 1024d / 1024d:0.0} MB",
        >= 1024L => $"{bytes / 1024d:0.0} KB",
        _ => $"{bytes} bytes"
    };
}
