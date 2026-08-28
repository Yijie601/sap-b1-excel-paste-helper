using System.Globalization;
using SapB1ExcelHelper.Models;

namespace SapB1ExcelHelper.Services;

public sealed class ClipboardValidationException : Exception
{
    public ClipboardValidationException(string message) : base(message)
    {
    }
}

public sealed class ExcelClipboardParser
{
    private static readonly string[] SupportedDateFormats =
    {
        "dd-MM-yyyy",
        "dd/MM/yyyy",
        "dd.MM.yyyy",
        "yyyy-MM-dd"
    };

    public InvoiceClipboardData Parse(string clipboardText)
    {
        if (string.IsNullOrWhiteSpace(clipboardText))
        {
            throw new ClipboardValidationException("Clipboard is empty. Copy Excel columns B:N first.");
        }

        var normalized = clipboardText.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
        var lines = normalized.Split('\n').ToList();
        while (lines.Count > 0 && lines[^1].Length == 0)
        {
            lines.RemoveAt(lines.Count - 1);
        }

        if (lines.Count == 0)
        {
            throw new ClipboardValidationException("Clipboard is empty. Copy Excel columns B:N first.");
        }

        var rows = lines.Select(line => line.Split('\t')).ToList();
        if (LooksLikeHeader(rows[0]))
        {
            throw new ClipboardValidationException("Excel header detected. Please copy data rows only.");
        }

        if (rows.Any(row => row.Length != 13))
        {
            throw new ClipboardValidationException("Invalid Excel selection. Please copy exactly columns B:N (13 columns).");
        }

        var supplier = rows[0][0].Trim();
        var documentNumber = rows[0][2].Trim();
        var documentDate = ParseDate(rows[0][1]);

        if (supplier.Length == 0 || documentNumber.Length == 0)
        {
            throw new ClipboardValidationException("Supplier Name, Document Date, and Document Number are required.");
        }

        foreach (var row in rows.Skip(1))
        {
            var rowDate = ParseDate(row[1]);
            if (!string.Equals(row[0].Trim(), supplier, StringComparison.OrdinalIgnoreCase) ||
                rowDate != documentDate ||
                !string.Equals(row[2].Trim(), documentNumber, StringComparison.Ordinal))
            {
                throw new ClipboardValidationException("Multiple invoices detected.");
            }
        }

        var items = rows.Select(row => new InvoiceItem
        {
            ItemNo = row[3].Trim(),
            Outlet = row[4].Trim(),
            Qty = row[5].Trim(),
            Total = row[6].Trim(),
            UnitPrice = row[7].Trim(),
            VatCode = row[8].Trim(),
            Department = row[9].Trim(),
            Discount = row[10].Trim(),
            Uom = row[11].Trim(),
            Warehouse = row[12].Trim()
        }).ToArray();

        return new InvoiceClipboardData
        {
            SupplierName = supplier,
            DocumentDate = documentDate,
            DocumentNumber = documentNumber,
            Items = items,
            OriginalClipboardText = clipboardText
        };
    }

    private static DateTime ParseDate(string rawValue)
    {
        if (DateTime.TryParseExact(
            rawValue.Trim(),
            SupportedDateFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var result))
        {
            return result.Date;
        }

        throw new ClipboardValidationException(
            $"Unsupported date: {rawValue.Trim()}. Use dd-MM-yyyy, dd/MM/yyyy, dd.MM.yyyy, or yyyy-MM-dd.");
    }

    private static bool LooksLikeHeader(IReadOnlyList<string> row)
    {
        if (row.Count < 3)
        {
            return false;
        }

        return row[0].Contains("Supplier", StringComparison.OrdinalIgnoreCase) &&
               row[1].Contains("Date", StringComparison.OrdinalIgnoreCase) &&
               (row[2].Contains("Document", StringComparison.OrdinalIgnoreCase) ||
                row[2].Contains("Number", StringComparison.OrdinalIgnoreCase));
    }
}

