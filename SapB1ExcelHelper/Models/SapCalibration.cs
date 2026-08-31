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
    public const int AbsoluteDesktopCoordinateVersion = 2;

    public int CoordinateVersion { get; set; }
    public SapPoint Supplier { get; set; } = new();
    public SapPoint PostingDate { get; set; } = new();
    public SapPoint Remarks { get; set; } = new();
    public SapPoint ItemNo { get; set; } = new();

    public bool SupplierCaptured { get; set; }
    public bool PostingDateCaptured { get; set; }
    public bool RemarksCaptured { get; set; }
    public bool ItemNoCaptured { get; set; }

    [JsonIgnore]
    public bool IsComplete => MissingFields.Count == 0;

    [JsonIgnore]
    public IReadOnlyList<string> MissingFields
    {
        get
        {
            var currentCoordinateVersion = CoordinateVersion == AbsoluteDesktopCoordinateVersion;
            var fields = new List<string>(4);
            AddIfMissing(fields, currentCoordinateVersion && SupplierCaptured, "Supplier");
            AddIfMissing(fields, currentCoordinateVersion && PostingDateCaptured, "Posting Date");
            AddIfMissing(fields, currentCoordinateVersion && RemarksCaptured, "Remarks");
            AddIfMissing(fields, currentCoordinateVersion && ItemNoCaptured, "First Item No.");
            return fields;
        }
    }

    public SapCalibration Clone() => new()
    {
        CoordinateVersion = CoordinateVersion,
        Supplier = Supplier.Clone(),
        PostingDate = PostingDate.Clone(),
        Remarks = Remarks.Clone(),
        ItemNo = ItemNo.Clone(),
        SupplierCaptured = SupplierCaptured,
        PostingDateCaptured = PostingDateCaptured,
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
