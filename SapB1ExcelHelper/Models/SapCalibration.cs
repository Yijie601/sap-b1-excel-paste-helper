using System.Text.Json.Serialization;

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

    public bool SupplierCaptured { get; set; }
    public bool SupplierRefCaptured { get; set; }
    public bool PostingDateCaptured { get; set; }
    public bool DocumentDateCaptured { get; set; }
    public bool RemarksCaptured { get; set; }
    public bool ItemNoCaptured { get; set; }

    [JsonIgnore]
    public bool IsComplete => MissingFields.Count == 0;

    [JsonIgnore]
    public IReadOnlyList<string> MissingFields
    {
        get
        {
            var fields = new List<string>(6);
            AddIfMissing(fields, SupplierCaptured, "Supplier");
            AddIfMissing(fields, SupplierRefCaptured, "Supplier Ref.");
            AddIfMissing(fields, PostingDateCaptured, "Posting Date");
            AddIfMissing(fields, DocumentDateCaptured, "Document Date");
            AddIfMissing(fields, RemarksCaptured, "Remarks");
            AddIfMissing(fields, ItemNoCaptured, "First Item No.");
            return fields;
        }
    }

    public SapCalibration Clone() => new()
    {
        Supplier = Supplier.Clone(),
        SupplierRef = SupplierRef.Clone(),
        PostingDate = PostingDate.Clone(),
        DocumentDate = DocumentDate.Clone(),
        Remarks = Remarks.Clone(),
        ItemNo = ItemNo.Clone(),
        SupplierCaptured = SupplierCaptured,
        SupplierRefCaptured = SupplierRefCaptured,
        PostingDateCaptured = PostingDateCaptured,
        DocumentDateCaptured = DocumentDateCaptured,
        RemarksCaptured = RemarksCaptured,
        ItemNoCaptured = ItemNoCaptured
    };

    private static void AddIfMissing(List<string> fields, bool captured, string name)
    {
        if (!captured)
        {
            fields.Add(name);
        }
    }
}
