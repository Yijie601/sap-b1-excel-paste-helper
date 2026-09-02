using System.Runtime.InteropServices;

namespace SapB1ExcelHelper.Services;

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
