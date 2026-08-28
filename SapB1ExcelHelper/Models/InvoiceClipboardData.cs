namespace SapB1ExcelHelper.Models;

public sealed class InvoiceClipboardData
{
    public string SupplierName { get; init; } = string.Empty;
    public string SupplierCode { get; set; } = string.Empty;
    public DateTime DocumentDate { get; init; }
    public string DocumentNumber { get; init; } = string.Empty;
    public IReadOnlyList<InvoiceItem> Items { get; init; } = Array.Empty<InvoiceItem>();
    public string OriginalClipboardText { get; init; } = string.Empty;

    public string SapDate => DocumentDate.ToString("dd.MM.yy");
    public string ItemClipboardBlock => string.Join("\r\n", Items.Select(item => item.ToClipboardRow()));
}

