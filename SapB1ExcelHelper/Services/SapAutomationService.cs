using System.Diagnostics;
using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

public sealed class SapAutomationException : Exception
{
    public SapAutomationException(string message) : base(message)
    {
    }
}

public sealed record AutomationStepResult(SapPasteStep Step, TimeSpan Duration, int ItemRows);

public sealed class SapAutomationService
{
    private static readonly TimeSpan FieldFocusDelay = TimeSpan.FromMilliseconds(90);
    private static readonly TimeSpan FieldPasteDelay = TimeSpan.FromMilliseconds(120);
    public async Task<AutomationStepResult> RunStepAsync(
        InvoiceClipboardData invoice,
        SapCalibration calibration,
        SapPasteStep step,
        Action<string>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var clipboard = ClipboardSnapshot.Capture();

        try
        {
            ValidateAbsoluteDesktopCoordinates(calibration);

            var itemRows = 0;
            switch (step)
            {
                case SapPasteStep.Supplier:
                    progress?.Invoke("Step 1/5: Pasting Supplier...");
                    await PasteTextField(calibration.Supplier, invoice.SapSupplierValue);
                    break;
                case SapPasteStep.PostingDate:
                    progress?.Invoke("Step 2/5: Pasting Posting Date...");
                    await PasteTextField(calibration.PostingDate, invoice.SapDate);
                    break;
                case SapPasteStep.SupplierRef:
                    progress?.Invoke("Step 3/5: Pasting Supplier Ref...");
                    await PasteTextField(calibration.SupplierRef, invoice.DocumentNumber);
                    break;
                case SapPasteStep.Remarks:
                    progress?.Invoke("Step 4/5: Pasting Remarks...");
                    await PasteTextField(calibration.Remarks, invoice.DocumentNumber);
                    break;
                case SapPasteStep.Items:
                    progress?.Invoke($"Step 5/5: Pasting the entire E:N block ({invoice.Items.Count} rows)...");
                    ClipboardService.SetText(invoice.ItemClipboardBlock);
                    await ClickAbsolutePoint(calibration.ItemNo);
                    InputService.Paste();
                    await Task.Delay(ItemPasteDelay(invoice.Items.Count));
                    itemRows = invoice.Items.Count;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(step), step, null);
            }

            stopwatch.Stop();
            return new AutomationStepResult(step, stopwatch.Elapsed, itemRows);
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
        await ClickAbsolutePoint(point);
        await PasteTextAtCurrentFocus(value);
    }

    private static async Task PasteTextAtCurrentFocus(string value)
    {
        ClipboardService.SetText(value);
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
            throw new SapAutomationException("Absolute desktop calibration is incomplete. Capture all five SAP positions again.");
        }

        var points = new[]
        {
            calibration.Supplier,
            calibration.SupplierRef,
            calibration.PostingDate,
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
