using System.Globalization;
using System.Text;

namespace SapB1ExcelHelper.Services;

public static class AppLogger
{
    private static readonly object Sync = new();

    public static void Success(string supplier, string documentNumber, int rows, TimeSpan duration) =>
        Write("SUCCESS", supplier, documentNumber, rows, duration, null, null);

    public static void Failure(string? supplier, string? documentNumber, int rows, TimeSpan duration, string error) =>
        Write("ERROR", supplier, documentNumber, rows, duration, error, null);

    public static void Error(string status, string message, Exception? exception = null) =>
        Write(status, null, null, 0, TimeSpan.Zero, message, exception);

    private static void Write(
        string status,
        string? supplier,
        string? documentNumber,
        int rows,
        TimeSpan duration,
        string? error,
        Exception? exception)
    {
        try
        {
            var now = DateTime.Now;
            var file = Path.Combine(AppPaths.LogsDirectory, $"sap-helper-{now:yyyy-MM}.log");
            var text = new StringBuilder()
                .AppendLine($"{now:yyyy-MM-dd HH:mm:ss.fff} {status}")
                .AppendLine($"Supplier: {supplier ?? "-"}")
                .AppendLine($"Ref: {documentNumber ?? "-"}")
                .AppendLine($"Rows: {rows.ToString(CultureInfo.InvariantCulture)}")
                .AppendLine($"Duration: {duration.TotalSeconds.ToString("0.000", CultureInfo.InvariantCulture)}s");

            if (!string.IsNullOrWhiteSpace(error))
            {
                text.AppendLine($"Error: {error}");
            }

            if (exception is not null)
            {
                text.AppendLine(exception.ToString());
            }

            text.AppendLine();
            lock (Sync)
            {
                File.AppendAllText(file, text.ToString(), Encoding.UTF8);
            }
        }
        catch
        {
            // Logging must never interrupt SAP automation.
        }
    }
}
