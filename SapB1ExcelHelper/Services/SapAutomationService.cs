using System.Diagnostics;
using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

public sealed class SapAutomationException : Exception
{
    public SapAutomationException(string message) : base(message)
    {
    }
}

public sealed record AutomationResult(TimeSpan Duration, int ItemRows);

public sealed class SapAutomationService
{
    private static readonly TimeSpan FieldFocusDelay = TimeSpan.FromMilliseconds(90);
    private static readonly TimeSpan FieldPasteDelay = TimeSpan.FromMilliseconds(120);
    private static readonly TimeSpan SupplierCommitDelay = TimeSpan.FromMilliseconds(1800);

    public async Task<AutomationResult> RunAsync(
        InvoiceClipboardData invoice,
        SapCalibration calibration,
        Action<string>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var clipboard = ClipboardSnapshot.Capture();

        try
        {
            ValidateAbsoluteDesktopCoordinates(calibration);

            progress?.Invoke("Filling Supplier...");
            await PasteTextField(calibration.Supplier, invoice.SapSupplierValue);
            InputService.Tab();
            await Task.Delay(SupplierCommitDelay);

            progress?.Invoke("Filling invoice header...");
            await PasteTextField(calibration.SupplierRef, invoice.DocumentNumber);
            await PasteTextField(calibration.PostingDate, invoice.SapDate);
            await PasteTextField(calibration.DocumentDate, invoice.SapDate);
            await PasteTextField(calibration.Remarks, invoice.DocumentNumber);

            progress?.Invoke($"Pasting {invoice.Items.Count} item row(s)...");
            ClipboardService.SetText(invoice.ItemClipboardBlock);
            await ClickAbsolutePoint(calibration.ItemNo);
            InputService.Paste();
            await Task.Delay(ItemPasteDelay(invoice.Items.Count));

            stopwatch.Stop();
            return new AutomationResult(stopwatch.Elapsed, invoice.Items.Count);
        }
        finally
        {
            try
            {
                clipboard.Restore();
            }
            catch (Exception exception)
            {
                AppLogger.Error("CLIPBOARD_RESTORE_ERROR", exception.Message, exception);
            }
        }
    }

    private static async Task PasteTextField(SapPoint point, string value)
    {
        ClipboardService.SetText(value);
        await ClickAbsolutePoint(point);
        InputService.SelectAll();
        await Task.Delay(20);
        InputService.Paste();
        await Task.Delay(FieldPasteDelay);
    }

    private static async Task ClickAbsolutePoint(SapPoint point)
    {
        InputService.Click(point.X, point.Y);
        await Task.Delay(FieldFocusDelay);
    }

    private static void ValidateAbsoluteDesktopCoordinates(SapCalibration calibration)
    {
        if (!calibration.IsComplete)
        {
            throw new SapAutomationException("Absolute desktop calibration is incomplete. Capture all six SAP positions again.");
        }

        var points = new[]
        {
            calibration.Supplier,
            calibration.SupplierRef,
            calibration.PostingDate,
            calibration.DocumentDate,
            calibration.Remarks,
            calibration.ItemNo
        };
        var desktop = SystemInformation.VirtualScreen;
        if (points.Any(point => !desktop.Contains(point.X, point.Y)))
        {
            throw new SapAutomationException(
                "A calibrated point is outside the current desktop. Keep the same monitor layout or run Calibration again.");
        }
    }

    private static TimeSpan ItemPasteDelay(int rowCount) =>
        TimeSpan.FromMilliseconds(Math.Min(5000, 450 + rowCount * 35));
}
