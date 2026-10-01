using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace SapB1ExcelHelper.Services;

public sealed record ClipboardText(string? Text, uint Sequence);

public static class ClipboardService
{
    private const uint UnicodeText = 13;
    private const uint MoveableMemory = 0x0002;
    private const int MaximumTextBytes = 16 * 1024 * 1024;
    private static readonly object ReadSync = new();
    private static Task<ClipboardText>? _pendingRead;

    public static uint Sequence => NativeMethods.GetClipboardSequenceNumber();

    public static async Task<ClipboardText> ReadTextAsync(CancellationToken cancellationToken = default)
    {
        Task<ClipboardText> read;
        lock (ReadSync)
        {
            // Excel may delay rendering text. Reuse a pending background read so
            // the UI remains responsive without accumulating blocked workers.
            if (_pendingRead is null || _pendingRead.IsCompleted)
            {
                _pendingRead = Task.Run(() => ClipboardRetry.RunAsync(TryReadText));
            }

            read = _pendingRead;
        }

        try
        {
            return await read.WaitAsync(TimeSpan.FromSeconds(10), cancellationToken);
        }
        catch (TimeoutException exception)
        {
            throw new InvalidOperationException("Excel is still preparing clipboard text. Wait a moment and copy B:N again.", exception);
        }
    }

    public static Task<uint> SetTextAsync(
        string value,
        nint ownerWindow,
        Action? onBusy = null,
        CancellationToken cancellationToken = default,
        uint? expectedSequence = null) =>
        // EmptyClipboard can notify an external clipboard owner. Keep that
        // native call off the UI too; onBusy must be safe for a worker thread.
        Task.Run(() => ClipboardRetry.RunAsync(
            () => TryWriteText(value, ownerWindow, expectedSequence),
            cancellationToken: cancellationToken,
            onBusy: onBusy), cancellationToken);

    private static (bool Success, ClipboardText Value) TryReadText()
    {
        if (!NativeMethods.OpenClipboard(0))
        {
            return (false, new ClipboardText(null, 0));
        }

        try
        {
            if (!NativeMethods.IsClipboardFormatAvailable(UnicodeText))
            {
                return (true, new ClipboardText(null, Sequence));
            }

            var memory = NativeMethods.GetClipboardData(UnicodeText);
            if (memory == 0)
            {
                return (false, new ClipboardText(null, 0));
            }

            var size = NativeMethods.GlobalSize(memory);
            if (size == 0 || size > MaximumTextBytes)
            {
                throw new InvalidOperationException("Clipboard text is too large. Copy one invoice across B:N.");
            }

            var pointer = NativeMethods.GlobalLock(memory);
            if (pointer == 0)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to read clipboard text.");
            }

            try
            {
                var text = Marshal.PtrToStringUni(pointer, checked((int)size / 2)) ?? string.Empty;
                var terminator = text.IndexOf('\0');
                if (terminator >= 0)
                {
                    text = text[..terminator];
                }

                return (true, new ClipboardText(text, Sequence));
            }
            finally
            {
                _ = NativeMethods.GlobalUnlock(memory);
            }
        }
        finally
        {
            _ = NativeMethods.CloseClipboard();
        }
    }

    private static (bool Success, uint Value) TryWriteText(string value, nint ownerWindow, uint? expectedSequence)
    {
        if (!NativeMethods.OpenClipboard(ownerWindow))
        {
            return (false, 0);
        }

        nint memory = 0;
        try
        {
            if (expectedSequence is { } expected && Sequence != expected)
            {
                throw new ClipboardChangedException();
            }

            var bytes = Encoding.Unicode.GetBytes(value + '\0');
            memory = NativeMethods.GlobalAlloc(MoveableMemory, (nuint)bytes.Length);
            if (memory == 0)
            {
                throw new OutOfMemoryException("Unable to prepare clipboard text.");
            }

            var pointer = NativeMethods.GlobalLock(memory);
            if (pointer == 0)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to prepare clipboard text.");
            }

            try
            {
                Marshal.Copy(bytes, 0, pointer, bytes.Length);
            }
            finally
            {
                _ = NativeMethods.GlobalUnlock(memory);
            }

            if (!NativeMethods.EmptyClipboard() || NativeMethods.SetClipboardData(UnicodeText, memory) == 0)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Unable to set clipboard text.");
            }

            memory = 0; // Ownership transfers to Windows after a successful write.
            return (true, Sequence);
        }
        finally
        {
            if (memory != 0)
            {
                _ = NativeMethods.GlobalFree(memory);
            }

            _ = NativeMethods.CloseClipboard();
        }
    }
}

public sealed class ClipboardChangedException : InvalidOperationException
{
    public ClipboardChangedException() : base("The clipboard changed during the paste. Check SAP before copying the invoice again.")
    {
    }
}

internal static class ClipboardRetry
{
    internal static async Task<T> RunAsync<T>(
        Func<(bool Success, T Value)> attempt,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default,
        Action? onBusy = null)
    {
        var limit = timeout ?? TimeSpan.FromSeconds(8);
        var stopwatch = Stopwatch.StartNew();
        var notified = false;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var result = attempt();
            if (result.Success)
            {
                return result.Value;
            }

            var remaining = limit - stopwatch.Elapsed;
            if (remaining <= TimeSpan.Zero)
            {
                throw new InvalidOperationException($"The Windows clipboard stayed busy for {limit.TotalSeconds:0.#} seconds. Check SAP before trying again.");
            }

            if (!notified)
            {
                onBusy?.Invoke();
                notified = true;
            }

            await Task.Delay(remaining < TimeSpan.FromMilliseconds(100) ? remaining : TimeSpan.FromMilliseconds(100), cancellationToken);
        }
    }
}
