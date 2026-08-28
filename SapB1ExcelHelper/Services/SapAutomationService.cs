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
    private readonly SapWindowService _windowService;

    public SapAutomationService(SapWindowService windowService)
    {
        _windowService = windowService;
    }

    public async Task<AutomationResult> RunAsync(
        InvoiceClipboardData invoice,
        SapWindowInfo window,
        SapCalibration calibration,
        Action<string>? progress = null)
    {
        var stopwatch = Stopwatch.StartNew();
        var clipboard = ClipboardSnapshot.Capture();

        try
        {
            progress?.Invoke("Filling Supplier...");
            await FillTextField(window, calibration.Supplier, invoice.SupplierCode);
            InputService.Tab();
            await WaitUntilReady(window, TimeSpan.FromMilliseconds(220), TimeSpan.FromSeconds(2.5));

            progress?.Invoke("Filling invoice header...");
            await FillTextField(window, calibration.SupplierRef, invoice.DocumentNumber);
            await FillTextField(window, calibration.PostingDate, invoice.SapDate);
            await FillTextField(window, calibration.DocumentDate, invoice.SapDate);
            await FillTextField(window, calibration.Remarks, invoice.DocumentNumber);

            progress?.Invoke($"Pasting {invoice.Items.Count} item row(s)...");
            EnsureSapStillActive(window);
            ClipboardService.SetText(invoice.ItemClipboardBlock);
            await ClickAndVerifyEditor(window, calibration.ItemNo);
            InputService.Paste();
            await WaitUntilReady(window, TimeSpan.FromMilliseconds(180), TimeSpan.FromSeconds(3));

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

    private async Task FillTextField(SapWindowInfo window, SapPoint point, string value)
    {
        await ClickAndVerifyEditor(window, point);
        InputService.SelectAll();
        await Task.Delay(12);
        InputService.SendUnicodeText(value);
        await Task.Delay(18);
    }

    private async Task ClickAndVerifyEditor(SapWindowInfo window, SapPoint point)
    {
        EnsureSapStillActive(window);
        if (point.X < 0 || point.Y < 0 || point.X >= window.Width || point.Y >= window.Height)
        {
            throw new SapAutomationException("A calibrated field is outside the AP Invoice window. Run Calibration again.");
        }

        InputService.Click(window.Left + point.X, window.Top + point.Y);
        for (var attempt = 0; attempt < 5; attempt++)
        {
            await Task.Delay(15);
            EnsureSapStillActive(window);
            if (_windowService.GetFocusedControlClass(window)
                .Equals("TMEditTextClass", StringComparison.Ordinal))
            {
                return;
            }
        }

        throw new SapAutomationException(
            "The calibrated position did not activate an SAP input field. Run Test Calibration and recalibrate this field.");
    }

    private async Task WaitUntilReady(SapWindowInfo window, TimeSpan minimumDelay, TimeSpan timeout)
    {
        await Task.Delay(minimumDelay);
        var stopwatch = Stopwatch.StartNew();
        var stableChecks = 0;

        while (stopwatch.Elapsed < timeout)
        {
            EnsureSapStillActive(window);
            if (_windowService.IsResponsive(window) && !_windowService.IsWaitCursorVisible())
            {
                stableChecks++;
                if (stableChecks >= 3)
                {
                    return;
                }
            }
            else
            {
                stableChecks = 0;
            }

            await Task.Delay(40);
        }

        throw new SapAutomationException("SAP did not become ready before the timeout. No Add or Update action was performed.");
    }

    private void EnsureSapStillActive(SapWindowInfo window)
    {
        if (!_windowService.IsSameActiveInvoice(window))
        {
            throw new SapAutomationException("SAP AP Invoice lost focus. Automation stopped safely.");
        }
    }
}
