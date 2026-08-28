namespace SapB1ExcelHelper.Models;

public sealed class SapPoint
{
    public int X { get; set; }
    public int Y { get; set; }

    public SapPoint Clone() => new() { X = X, Y = Y };
}

public sealed class SapCalibration
{
    public SapPoint Supplier { get; set; } = new() { X = 249, Y = 62 };
    public SapPoint SupplierRef { get; set; } = new() { X = 249, Y = 97 };
    public SapPoint PostingDate { get; set; } = new() { X = 1799, Y = 80 };
    public SapPoint DocumentDate { get; set; } = new() { X = 1799, Y = 115 };
    public SapPoint Remarks { get; set; } = new() { X = 249, Y = 817 };
    public SapPoint ItemNo { get; set; } = new() { X = 536, Y = 283 };

    public SapCalibration Clone() => new()
    {
        Supplier = Supplier.Clone(),
        SupplierRef = SupplierRef.Clone(),
        PostingDate = PostingDate.Clone(),
        DocumentDate = DocumentDate.Clone(),
        Remarks = Remarks.Clone(),
        ItemNo = ItemNo.Clone()
    };
}

