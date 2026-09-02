namespace SapB1ExcelHelper.Services;

public enum SapPasteStep
{
    Supplier,
    PostingDate,
    SupplierRef,
    Remarks,
    Items
}

public static class SapPasteWorkflow
{
    public static IReadOnlyList<SapPasteStep> Steps { get; } = Array.AsReadOnly(new[]
    {
        SapPasteStep.Supplier,
        SapPasteStep.PostingDate,
        SapPasteStep.SupplierRef,
        SapPasteStep.Remarks,
        SapPasteStep.Items
    });

    public static string GetLabel(SapPasteStep step) => step switch
    {
        SapPasteStep.Supplier => "Supplier",
        SapPasteStep.PostingDate => "Posting Date",
        SapPasteStep.SupplierRef => "Supplier Ref.",
        SapPasteStep.Remarks => "Remarks",
        SapPasteStep.Items => "Item No. (entire E:N block)",
        _ => throw new ArgumentOutOfRangeException(nameof(step), step, null)
    };
}
