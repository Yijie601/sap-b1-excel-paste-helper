using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

internal interface ISapAutomationIO
{
    uint ClipboardSequence { get; }
    Task<uint> SetClipboardAsync(string value, uint expectedSequence, Action? onBusy, CancellationToken cancellationToken);
    void CaptureTarget(SapCalibration calibration);
    void VerifyTarget(SapPoint point);
    void Click(SapPoint point);
    void SelectAll();
    void Paste();
    Task WaitForModifiersReleasedAsync(CancellationToken cancellationToken);
    Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken);
}

internal sealed class WindowsAutomationIO(Func<nint> clipboardOwner) : ISapAutomationIO
{
    private nint _target;
    private Rectangle _targetBounds;

    public uint ClipboardSequence => ClipboardService.Sequence;
    public Task<uint> SetClipboardAsync(string value, uint expectedSequence, Action? onBusy, CancellationToken cancellationToken) =>
        ClipboardService.SetTextAsync(value, clipboardOwner(), onBusy, cancellationToken, expectedSequence);
    public void Click(SapPoint point) => InputService.Click(point.X, point.Y);
    public void SelectAll() => InputService.SelectAll();
    public void Paste() => InputService.Paste();
    public Task WaitForModifiersReleasedAsync(CancellationToken cancellationToken) => InputService.WaitForModifiersReleasedAsync(cancellationToken);
    public Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken) => Task.Delay(delay, cancellationToken);

    public void CaptureTarget(SapCalibration calibration)
    {
        if (!calibration.IsComplete)
        {
            throw new SapAutomationException("Capture all five SAP positions in Calibration first.");
        }

        var points = new[] { calibration.Supplier, calibration.PostingDate, calibration.SupplierRef, calibration.Remarks, calibration.ItemNo };
        // Check actual monitor rectangles, not the virtual bounding box (which
        // can include empty gaps between monitors).
        if (points.Any(point => !Screen.AllScreens.Any(screen => screen.Bounds.Contains(point.X, point.Y))))
        {
            throw new SapAutomationException("A saved point is outside the current monitors. Run Calibration again.");
        }

        _target = RootAt(calibration.Supplier);
        if (_target == 0 || _target == Root(clipboardOwner()) || points.Any(point => RootAt(point) != _target) ||
            !NativeMethods.GetWindowRect(_target, out var rectangle))
        {
            throw new SapAutomationException("The saved fields are covered or no longer in one SAP window. Restore the calibrated SAP layout and try again.");
        }

        _targetBounds = rectangle.Rectangle;
        VerifyTarget(calibration.Supplier);
    }

    public void VerifyTarget(SapPoint point)
    {
        if (Root(NativeMethods.GetForegroundWindow()) != _target || RootAt(point) != _target ||
            !NativeMethods.GetWindowRect(_target, out var rectangle) || rectangle.Rectangle != _targetBounds)
        {
            throw new SapAutomationException("The target window moved, lost focus, or is covered by a popup. Check SAP before starting again.");
        }
    }

    private static nint Root(nint window) => window == 0 ? 0 : NativeMethods.GetAncestor(window, 2);
    private static nint RootAt(SapPoint point) => Root(NativeMethods.WindowFromPoint(new Point(point.X, point.Y)));
}
