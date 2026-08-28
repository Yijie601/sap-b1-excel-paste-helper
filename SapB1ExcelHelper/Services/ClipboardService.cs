using System.Runtime.InteropServices;

namespace SapB1ExcelHelper.Services;

public sealed class ClipboardSnapshot
{
    private readonly IDataObject? _dataObject;
    private readonly string? _fallbackText;

    private ClipboardSnapshot(IDataObject? dataObject, string? fallbackText)
    {
        _dataObject = dataObject;
        _fallbackText = fallbackText;
    }

    public static ClipboardSnapshot Capture()
    {
        IDataObject? dataObject = null;
        string? text = null;
        ClipboardService.Retry(() =>
        {
            dataObject = Clipboard.GetDataObject();
            if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
            {
                text = Clipboard.GetText(TextDataFormat.UnicodeText);
            }
        });
        return new ClipboardSnapshot(dataObject, text);
    }

    public void Restore()
    {
        try
        {
            if (_dataObject is not null)
            {
                ClipboardService.Retry(() => Clipboard.SetDataObject(_dataObject, true));
                return;
            }
        }
        catch (Exception exception)
        {
            AppLogger.Error("CLIPBOARD_RESTORE_FALLBACK", exception.Message, exception);
        }

        if (_fallbackText is not null)
        {
            ClipboardService.SetText(_fallbackText);
        }
    }
}

public static class ClipboardService
{
    public static string? TryGetText()
    {
        string? value = null;
        try
        {
            Retry(() =>
            {
                value = Clipboard.ContainsText(TextDataFormat.UnicodeText)
                    ? Clipboard.GetText(TextDataFormat.UnicodeText)
                    : null;
            });
        }
        catch
        {
            return null;
        }

        return value;
    }

    public static void SetText(string value) => Retry(() => Clipboard.SetText(value, TextDataFormat.UnicodeText));

    internal static void Retry(Action action)
    {
        Exception? lastError = null;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                action();
                return;
            }
            catch (ExternalException exception)
            {
                lastError = exception;
                Thread.Sleep(15 + attempt * 10);
            }
        }

        throw new InvalidOperationException("The Windows clipboard is busy. Please try again.", lastError);
    }
}
