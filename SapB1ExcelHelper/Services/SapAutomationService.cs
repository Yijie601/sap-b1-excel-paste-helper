using System.Diagnostics;
using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

public sealed class SapAutomationException : Exception
{
    public SapAutomationException(string message) : base(message)
    {
    }

    public SapAutomationException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

public sealed record AutomationResult(TimeSpan Duration, int ItemRows);

public sealed class SapAutomationService
{
    private readonly ISapAutomationIO _io;
    private static readonly TimeSpan FieldFocusDelay = TimeSpan.FromMilliseconds(150);
    public static TimeSpan PasteInterval { get; } = TimeSpan.FromMilliseconds(800);

    internal SapAutomationService(Func<nint> clipboardOwner) : this(new WindowsAutomationIO(clipboardOwner))
    {
    }

    internal SapAutomationService(ISapAutomationIO io) => _io = io;

    public uint LastClipboardSequence { get; private set; }

    public async Task<AutomationResult> RunAllAsync(
        InvoiceClipboardData invoice,
        SapCalibration calibration,
        uint expectedClipboardSequence,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        // Trust the copy that was validated, not a later copy made while the
        // helper was hiding or waiting to acquire the clipboard.
        LastClipboardSequence = expectedClipboardSequence;
        var clipboardWritten = false;
        var itemBlock = invoice.ItemClipboardBlock;
        progress?.Report("Starting — release shortcut modifier keys...");
        await _io.WaitForModifiersReleasedAsync(cancellationToken);
        _io.CaptureTarget(calibration);
        try
        {
            for (var index = 0; index < SapPasteWorkflow.Steps.Count; index++)
            {
                var step = SapPasteWorkflow.Steps[index];
                var stepNumber = index + 1;
                var label = SapPasteWorkflow.GetLabel(step);
                var stepTimer = Stopwatch.StartNew();
                var (point, value) = step switch
                {
                    SapPasteStep.Supplier => (calibration.Supplier, invoice.SapSupplierValue),
                    SapPasteStep.PostingDate => (calibration.PostingDate, invoice.SapDate),
                    SapPasteStep.SupplierRef => (calibration.SupplierRef, invoice.DocumentNumber),
                    SapPasteStep.Remarks => (calibration.Remarks, invoice.DocumentNumber),
                    SapPasteStep.Items => (calibration.ItemNo, itemBlock),
                    _ => throw new ArgumentOutOfRangeException(nameof(step))
                };

                AppLogger.StepStarted(invoice.SupplierName, invoice.DocumentNumber, invoice.Items.Count, stepNumber, label);
                progress?.Report($"Step {stepNumber}/5: {label}...");
                try
                {
                    // Retry preparation only. Never replay keyboard input: an item
                    // paste may already have been accepted even if SAP is slow.
                    LastClipboardSequence = await _io.SetClipboardAsync(
                        value,
                        LastClipboardSequence,
                        () => progress?.Report($"Step {stepNumber}/5: Waiting for Clipboard..."),
                        cancellationToken);
                    clipboardWritten = true;
                    cancellationToken.ThrowIfCancellationRequested();
                    _io.VerifyTarget(point);
                    _io.Click(point);
                    await _io.DelayAsync(FieldFocusDelay, cancellationToken);
                    _io.VerifyTarget(point);
                    if (step != SapPasteStep.Items)
                    {
                        _io.SelectAll();
                        await _io.DelayAsync(TimeSpan.FromMilliseconds(20), cancellationToken);
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    _io.VerifyTarget(point);
                    if (_io.ClipboardSequence != LastClipboardSequence)
                    {
                        throw new ClipboardChangedException();
                    }

                    try
                    {
                        _io.Paste();
                    }
                    finally
                    {
                        // Finish the guard even after cancellation or a partial
                        // SendInput failure: SAP may already be reading the data.
                        await _io.DelayAsync(step == SapPasteStep.Items ? ItemPasteDelay(invoice.Items.Count) : PasteInterval, CancellationToken.None);
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    throw new SapAutomationException($"Step {stepNumber}/5 ({label}) stopped: {exception.Message}", exception);
                }

                AppLogger.StepSuccess(
                    invoice.SupplierName,
                    invoice.DocumentNumber,
                    invoice.Items.Count,
                    stepTimer.Elapsed,
                    stepNumber,
                    label);
            }
            stopwatch.Stop();
            return new AutomationResult(stopwatch.Elapsed, invoice.Items.Count);
        }
        finally
        {
            try
            {
                if (clipboardWritten && _io.ClipboardSequence == LastClipboardSequence)
                {
                    LastClipboardSequence = await _io.SetClipboardAsync(invoice.OriginalClipboardText, LastClipboardSequence, null, CancellationToken.None);
                }
            }
            catch (ClipboardChangedException)
            {
                // Preserve anything the user copied while the run was active.
            }
            catch (Exception exception)
            {
                AppLogger.Error("CLIPBOARD_RESTORE_ERROR", exception.Message, exception);
            }
        }
    }

    private static TimeSpan ItemPasteDelay(int rowCount) =>
        TimeSpan.FromMilliseconds(Math.Clamp(450L + rowCount * 35L, 800, 5000));
}
