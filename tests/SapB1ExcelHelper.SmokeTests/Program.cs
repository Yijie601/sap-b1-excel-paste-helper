using SapB1ExcelHelper.Services;

var tests = new (string Name, Action Run)[]
{
    ("Parses valid multi-row invoice and preserves blank columns", ParsesValidInvoice),
    ("Supports every documented date format", SupportsDateFormats),
    ("Rejects multiple invoices", RejectsMultipleInvoices),
    ("Rejects Excel headers", RejectsHeader),
    ("Rejects an incorrect column count", RejectsWrongColumnCount),
    ("Loads, saves, and resolves supplier CSV values", HandlesSupplierMappings)
};

var failures = 0;
foreach (var test in tests)
{
    try
    {
        test.Run();
        Console.WriteLine($"PASS  {test.Name}");
    }
    catch (Exception exception)
    {
        failures++;
        Console.Error.WriteLine($"FAIL  {test.Name}: {exception.Message}");
    }
}

Console.WriteLine($"{tests.Length - failures}/{tests.Length} smoke tests passed.");
return failures == 0 ? 0 : 1;

static void ParsesValidInvoice()
{
    var parser = new ExcelClipboardParser();
    var row1 = Row("LIM SOON POH TRADING", "13-08-2026", "260813/162", "ITEM-1", "OUTLET-A", "2", "20", "10", "SST", "", "0", "", "WH01");
    var row2 = Row("lim soon poh trading", "13/08/2026", "260813/162", "ITEM-2", "OUTLET-B", "1", "5", "5", "SST", "D01", "0", "EA", "WH01");
    var invoice = parser.Parse(row1 + "\r\n" + row2 + "\r\n");

    Equal("LIM SOON POH TRADING", invoice.SupplierName);
    Equal("260813/162", invoice.DocumentNumber);
    Equal("13.08.26", invoice.SapDate);
    Equal(2, invoice.Items.Count);
    Equal(string.Join('\t', "ITEM-1", "OUTLET-A", "2", "20", "10", "SST", "", "0", "", "WH01"), invoice.Items[0].ToClipboardRow());
    True(invoice.ItemClipboardBlock.Contains("\r\n", StringComparison.Ordinal), "Expected CRLF between item rows.");
}

static void SupportsDateFormats()
{
    var parser = new ExcelClipboardParser();
    foreach (var date in new[] { "13-08-2026", "13/08/2026", "13.08.2026", "2026-08-13" })
    {
        var invoice = parser.Parse(Row("Supplier", date, "REF", "I", "O", "1", "1", "1", "V", "", "0", "", "W"));
        Equal("13.08.26", invoice.SapDate);
    }
}

static void RejectsMultipleInvoices()
{
    var parser = new ExcelClipboardParser();
    var text = Row("Supplier", "13-08-2026", "REF-1", "I", "O", "1", "1", "1", "V", "", "0", "", "W") + "\r\n" +
               Row("Supplier", "13-08-2026", "REF-2", "I", "O", "1", "1", "1", "V", "", "0", "", "W");
    Throws<ClipboardValidationException>(() => parser.Parse(text), "Multiple invoices detected.");
}

static void RejectsHeader()
{
    var parser = new ExcelClipboardParser();
    var text = Row("Supplier Name", "Document Date", "Document Number", "Item", "Outlet", "QTY", "Total", "Price", "VAT", "Department", "Discount", "UoM", "Whse");
    Throws<ClipboardValidationException>(() => parser.Parse(text), "Excel header detected");
}

static void RejectsWrongColumnCount()
{
    var parser = new ExcelClipboardParser();
    Throws<ClipboardValidationException>(
        () => parser.Parse(string.Join('\t', Enumerable.Repeat("x", 12))),
        "13 columns");
}

static void HandlesSupplierMappings()
{
    var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"sap-helper-tests-{Guid.NewGuid():N}");
    Directory.CreateDirectory(temporaryDirectory);
    var file = Path.Combine(temporaryDirectory, "mapping.csv");
    try
    {
        File.WriteAllText(file, "Supplier Name,SAP Code\r\n\"Supplier, Sdn Bhd\",V100\r\n");
        var service = new SupplierMappingService(file);
        True(service.TryResolve("supplier, sdn bhd", out var code), "Case-insensitive mapping was not found.");
        Equal("V100", code);

        service.Save(new[]
        {
            new SupplierMappingEntry("  New Supplier  ", " V200 ")
        });
        True(service.TryResolve("NEW SUPPLIER", out code), "Saved mapping was not found.");
        Equal("V200", code);
    }
    finally
    {
        Directory.Delete(temporaryDirectory, true);
    }
}

static string Row(params string[] cells)
{
    Equal(13, cells.Length);
    return string.Join('\t', cells);
}

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
    {
        throw new InvalidOperationException($"Expected '{expected}', got '{actual}'.");
    }
}

static void True(bool value, string message)
{
    if (!value)
    {
        throw new InvalidOperationException(message);
    }
}

static void Throws<TException>(Action action, string expectedMessage) where TException : Exception
{
    try
    {
        action();
    }
    catch (TException exception) when (exception.Message.Contains(expectedMessage, StringComparison.Ordinal))
    {
        return;
    }

    throw new InvalidOperationException($"Expected {typeof(TException).Name} containing '{expectedMessage}'.");
}

