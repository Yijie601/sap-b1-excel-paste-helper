namespace SapB1ExcelHelper.Models;

public sealed class InvoiceItem
{
    public string ItemNo { get; init; } = string.Empty;
    public string Outlet { get; init; } = string.Empty;
    public string Qty { get; init; } = string.Empty;
    public string Total { get; init; } = string.Empty;
    public string UnitPrice { get; init; } = string.Empty;
    public string VatCode { get; init; } = string.Empty;
    public string Department { get; init; } = string.Empty;
    public string Discount { get; init; } = string.Empty;
    public string Uom { get; init; } = string.Empty;
    public string Warehouse { get; init; } = string.Empty;

    public string ToClipboardRow() => string.Join('\t',
        ItemNo, Outlet, Qty, Total, UnitPrice, VatCode,
        Department, Discount, Uom, Warehouse);
}

