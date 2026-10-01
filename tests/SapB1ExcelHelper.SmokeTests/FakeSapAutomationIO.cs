using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

internal sealed class FakeSapAutomationIO : ISapAutomationIO
{
    public uint ClipboardSequence { get; private set; } = 1;
    public string CurrentText { get; private set; } = "original";
    public List<string> Writes { get; } = new();
    public List<string> Pasted { get; } = new();
    public List<TimeSpan> Delays { get; } = new();
    public int Clicks { get; private set; }
    public int Selections { get; private set; }
    public bool TargetCovered { get; init; }
    public int FailAfterPasteNumber { get; init; }
    public Action? AfterPaste { get; set; }

    public Task<uint> SetClipboardAsync(string value, uint expectedSequence, Action? onBusy, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (expectedSequence != ClipboardSequence)
        {
            throw new ClipboardChangedException();
        }
        Writes.Add(value);
        CurrentText = value;
        return Task.FromResult(++ClipboardSequence);
    }

    public void CaptureTarget(SapCalibration calibration) => VerifyTarget(calibration.Supplier);
    public void VerifyTarget(SapPoint point)
    {
        if (TargetCovered)
        {
            throw new SapAutomationException("Target is covered.");
        }
    }
    public void Click(SapPoint point) => Clicks++;
    public void SelectAll() => Selections++;
    public void Paste()
    {
        Pasted.Add(CurrentText);
        AfterPaste?.Invoke();
        if (Pasted.Count == FailAfterPasteNumber)
        {
            throw new InvalidOperationException("Simulated partial input failure.");
        }
    }
    public Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Delays.Add(delay);
        return Task.CompletedTask;
    }
    public Task WaitForModifiersReleasedAsync(CancellationToken cancellationToken) => Task.CompletedTask;
    public void ExternalCopy(string value)
    {
        CurrentText = value;
        ClipboardSequence++;
    }
}
