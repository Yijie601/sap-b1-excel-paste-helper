using System.Text.Json;
using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

var tests = new (string Name, Action Run)[]
{
    ("Parses valid multi-row invoice and preserves blank columns", ParsesValidInvoice),
    ("Starts the COL33 item paste at SAP Code instead of Supplier Name", BuildsExpectedCol33ItemBlock),
    ("Uses the first B:D header once and builds fifty E:N item rows", BuildsFiftyRowInvoice),
    ("Supports every documented date format", SupportsDateFormats),
    ("Rejects multiple invoices", RejectsMultipleInvoices),
    ("Rejects Excel headers", RejectsHeader),
    ("Rejects an incorrect column count", RejectsWrongColumnCount),
    ("Rejects a selected row without an SAP item code", RejectsMissingItemCode),
    ("Uses the Excel supplier name directly", UsesSupplierNameDirectly),
    ("Compares stable and prerelease semantic versions", ComparesSemanticVersions),
    ("Selects the newest compatible GitHub release asset", SelectsNewestUpdate),
    ("Verifies an update installer SHA-256 digest", VerifiesUpdateDigest),
    ("Validates and persists custom global hotkeys", HandlesCustomHotkeys),
    ("Requires every SAP position to be captured explicitly", RequiresCompleteCalibration)
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

static void BuildsExpectedCol33ItemBlock()
{
    var parser = new ExcelClipboardParser();
    var row1 = Row(
        "COL33 PTE.LTD",
        "03-08-2026",
        "COL26080630_F",
        "PROC-Chives&PorkDumplings",
        "O-HW",
        "15",
        "",
        "7.5",
        "TX7",
        "",
        "",
        "",
        "S-HW");
    var row2 = Row(
        "COL33 PTE.LTD",
        "03-08-2026",
        "COL26080630_F",
        "PROC-Chives&PorkDumplings",
        "O-HW",
        "3",
        "",
        "0",
        "TX7",
        "",
        "",
        "",
        "S-HW");

    var invoice = parser.Parse(row1 + "\r\n" + row2);
    Equal("COL33 PTE.LTD", invoice.SapSupplierValue);
    Equal("COL26080630_F", invoice.DocumentNumber);
    Equal("03.08.26", invoice.SapDate);
    Equal(
        "PROC-Chives&PorkDumplings\tO-HW\t15\t\t7.5\tTX7\t\t\t\tS-HW\r\n" +
        "PROC-Chives&PorkDumplings\tO-HW\t3\t\t0\tTX7\t\t\t\tS-HW",
        invoice.ItemClipboardBlock);
}

static void BuildsFiftyRowInvoice()
{
    var parser = new ExcelClipboardParser();
    var rows = Enumerable.Range(1, 50)
        .Select(index => Row(
            index == 1 ? "COL33 PTE.LTD" : "",
            index == 1 ? "03-08-2026" : "",
            index == 1 ? "COL26080630_F" : "",
            $"ITEM-{index:00}",
            "O-HW",
            index.ToString(),
            "",
            "7.5",
            "TX7",
            "",
            "",
            "",
            "S-HW"));

    var invoice = parser.Parse(string.Join("\r\n", rows));
    Equal("COL33 PTE.LTD", invoice.SupplierName);
    Equal("COL26080630_F", invoice.DocumentNumber);
    Equal("03.08.26", invoice.SapDate);
    Equal(50, invoice.Items.Count);

    var itemRows = invoice.ItemClipboardBlock.Split("\r\n", StringSplitOptions.None);
    Equal(50, itemRows.Length);
    True(itemRows[0].StartsWith("ITEM-01\tO-HW\t1\t", StringComparison.Ordinal),
        "The first item row did not start at Excel column E.");
    True(itemRows[49].StartsWith("ITEM-50\tO-HW\t50\t", StringComparison.Ordinal),
        "The fiftieth item row was not preserved.");
    True(!invoice.ItemClipboardBlock.Contains(invoice.SupplierName, StringComparison.Ordinal),
        "Supplier Name leaked into the E:N item block.");
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

static void RejectsMissingItemCode()
{
    var parser = new ExcelClipboardParser();
    Throws<ClipboardValidationException>(
        () => parser.Parse(Row("Supplier", "13-08-2026", "REF", "", "O", "1", "1", "1", "V", "", "0", "", "W")),
        "SAP Code / Item No. is required");
}

static void UsesSupplierNameDirectly()
{
    var parser = new ExcelClipboardParser();
    var invoice = parser.Parse(Row(
        "Supplier Name From Excel",
        "13-08-2026",
        "REF",
        "ITEM",
        "OUTLET",
        "1",
        "1",
        "1",
        "SST",
        "",
        "0",
        "",
        "WH01"));

    Equal("Supplier Name From Excel", invoice.SupplierName);
    Equal(invoice.SupplierName, invoice.SapSupplierValue);
}

static void ComparesSemanticVersions()
{
    var beta2 = SemanticVersion.Parse("v0.1.0-beta.2");
    var beta3 = SemanticVersion.Parse("0.1.0-beta.3+build.99");
    var stable = SemanticVersion.Parse("0.1.0");
    var nextMinorBeta = SemanticVersion.Parse("0.2.0-beta.1");

    True(beta3 > beta2, "beta.3 should be newer than beta.2.");
    True(stable > beta3, "A stable release should be newer than its prereleases.");
    True(nextMinorBeta > stable, "A newer minor prerelease should have a newer core version.");
    Equal("0.1.0-beta.3", beta3.ToString());
}

static void SelectsNewestUpdate()
{
    const string digest = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";
    var json = $$"""
        [
          {
            "tag_name": "v0.1.0-beta.3",
            "name": "Beta 3",
            "body": "Update prompt",
            "html_url": "https://github.com/Yijie601/sap-b1-excel-paste-helper/releases/tag/v0.1.0-beta.3",
            "draft": false,
            "prerelease": true,
            "assets": [
              {
                "name": "SapB1ExcelHelper-Setup-0.1.0-beta.3-win-x64.exe",
                "state": "uploaded",
                "browser_download_url": "https://github.com/Yijie601/sap-b1-excel-paste-helper/releases/download/v0.1.0-beta.3/SapB1ExcelHelper-Setup-0.1.0-beta.3-win-x64.exe",
                "size": 123456,
                "digest": "sha256:{{digest}}"
              }
            ]
          },
          {
            "tag_name": "v0.1.0-beta.1",
            "name": "Old beta",
            "body": "",
            "html_url": "https://github.com/Yijie601/sap-b1-excel-paste-helper/releases/tag/v0.1.0-beta.1",
            "draft": false,
            "prerelease": true,
            "assets": []
          }
        ]
        """;

    var update = UpdateService.SelectAvailableUpdate(
        json,
        SemanticVersion.Parse("0.1.0-beta.2"));
    True(update is not null, "Expected beta.3 update for a beta.2 installation.");
    Equal("0.1.0-beta.3", update!.Version.ToString());
    Equal(digest, update.Sha256Digest);

    var stableUserUpdate = UpdateService.SelectAvailableUpdate(
        json.Replace("v0.1.0-beta.3", "v0.2.0-beta.1", StringComparison.Ordinal)
            .Replace("0.1.0-beta.3", "0.2.0-beta.1", StringComparison.Ordinal),
        SemanticVersion.Parse("0.1.0"));
    True(stableUserUpdate is null, "Stable users must not receive prerelease updates.");
}

static void VerifiesUpdateDigest()
{
    var file = Path.Combine(Path.GetTempPath(), $"sap-helper-digest-{Guid.NewGuid():N}.tmp");
    try
    {
        File.WriteAllText(file, "abc");
        var valid = UpdateService.VerifySha256Async(
                file,
                "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD")
            .GetAwaiter()
            .GetResult();
        True(valid, "Known SHA-256 digest did not match.");
    }
    finally
    {
        File.Delete(file);
    }
}

static void HandlesCustomHotkeys()
{
    var functionKey = new HotkeyDefinition(0x78, HotkeyModifiers.None);
    True(functionKey.IsSupported(out _), "F9 should be allowed without a modifier.");
    Equal("F9", functionKey.DisplayText);

    var combination = new HotkeyDefinition(
        0x4B,
        HotkeyModifiers.Control | HotkeyModifiers.Shift);
    True(combination.IsSupported(out _), "Ctrl + Shift + K should be supported.");
    Equal("Ctrl + Shift + K", combination.DisplayText);

    var bareLetter = new HotkeyDefinition(0x4B, HotkeyModifiers.None);
    True(!bareLetter.IsSupported(out _), "Bare letter keys should not be registered globally.");

    var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"sap-helper-hotkey-{Guid.NewGuid():N}");
    var file = Path.Combine(temporaryDirectory, "hotkey.json");
    try
    {
        var service = new HotkeySettingsService(file);
        service.Save(combination);
        Equal(combination, service.Load());
    }
    finally
    {
        Directory.Delete(temporaryDirectory, true);
    }
}

static void RequiresCompleteCalibration()
{
    var calibration = new SapCalibration();
    True(!calibration.IsComplete, "Default coordinates must not count as captured positions.");
    Equal(6, calibration.MissingFields.Count);

    calibration.SupplierCaptured = true;
    calibration.SupplierRefCaptured = true;
    calibration.PostingDateCaptured = true;
    calibration.DocumentDateCaptured = true;
    calibration.RemarksCaptured = true;
    calibration.ItemNoCaptured = true;
    True(!calibration.IsComplete, "Legacy relative coordinates must not pass absolute desktop calibration.");
    calibration.CoordinateVersion = SapCalibration.AbsoluteDesktopCoordinateVersion;
    True(calibration.IsComplete, "All six captured positions should complete calibration.");
    True(calibration.Clone().IsComplete, "Cloning lost the captured-state flags.");

    const string legacyJson = """
        {
          "supplier": { "x": 249, "y": 62 },
          "itemNo": { "x": 536, "y": 283 }
        }
        """;
    var legacyCalibration = JsonSerializer.Deserialize<SapCalibration>(
        legacyJson,
        new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    True(legacyCalibration is not null && !legacyCalibration.IsComplete,
        "A legacy default-coordinate file must require fresh calibration.");
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
