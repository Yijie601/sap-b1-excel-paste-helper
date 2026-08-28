using System.ComponentModel;
using System.Runtime.InteropServices;

namespace SapB1ExcelHelper.Services;

public sealed class MouseCaptureService
{
    public async Task<Point?> CaptureNextLeftClickAsync(TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        var completion = new TaskCompletionSource<Point>(TaskCreationOptions.RunContinuationsAsynchronously);
        NativeMethods.LowLevelMouseProc? callback = null;
        nint hook = nint.Zero;

        callback = (code, message, data) =>
        {
            if (code >= 0 && message == NativeMethods.WmLButtonDown)
            {
                var hookData = Marshal.PtrToStructure<NativeMethods.MouseHookData>(data);
                completion.TrySetResult(hookData.Point);
            }

            return NativeMethods.CallNextHookEx(hook, code, message, data);
        };

        hook = NativeMethods.SetWindowsHookEx(NativeMethods.WhMouseLl, callback, nint.Zero, 0);
        if (hook == nint.Zero)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to listen for the calibration click.");
        }

        try
        {
            using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var timeoutTask = Task.Delay(timeout, timeoutSource.Token);
            var completed = await Task.WhenAny(completion.Task, timeoutTask);
            if (completed == completion.Task)
            {
                timeoutSource.Cancel();
                return await completion.Task;
            }

            return null;
        }
        finally
        {
            _ = NativeMethods.UnhookWindowsHookEx(hook);
            GC.KeepAlive(callback);
        }
    }
}

