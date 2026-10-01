using System.Text.Json;
using SapB1ExcelHelper.Models;
using SapB1ExcelHelper.Services;

if (args.Length == 1 && args[0] == NativeClipboardTests.ChildFlag)
{
    return NativeClipboardTests.RunIsolatedChild();
}

var tests = new List<(string Name, Action Run)>
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
    ("Requires every SAP position to be captured explicitly", RequiresCompleteCalibration),
    ("Uses five ordered paste actions with an 0.8-second guard", UsesStepByStepPasteOrder),
    ("Handles Excel quoted cells without changing item columns", HandlesQuotedCells),
    ("Rejects embedded cell separators before sending anything to SAP", RejectsEmbeddedSeparators),
    ("Formats the SAP date independently of Windows culture", FormatsInvariantDate),
    ("Recovers from clipboard contention asynchronously", RecoversFromClipboardContention),
    ("Bounds clipboard retries and supports cancellation", BoundsClipboardRetries),
    ("Waits for custom-hotkey modifiers to be released asynchronously", WaitsForModifiers),
    ("Runs five pastes with only one final clipboard restoration", RunsFivePastesWithOneRestore),
    ("Stops before input when a target is covered", StopsAtCoveredTarget),
    ("Preserves a new user copy during an active run", PreservesNewUserCopy),
    ("Rejects a new copy made between validation and run startup", RejectsChangedCopyAtStartup),
    ("Never replays the item block after a partial input failure", DoesNotReplayItems),
    ("Cancellation finishes the paste guard and stops before the next field", CancelsBeforeNextField)
};

if (args.Contains(NativeClipboardTests.IntegrationFlag, StringComparer.Ordinal))
{
    tests.Add(("Exercises Win32 clipboard in a private noninteractive window station", NativeClipboardTests.RunInIsolatedProcess));
}
else
{
    Console.WriteLine(NativeClipboardTests.NotRequestedMessage);
}

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

Console.WriteLine($"{tests.Count - failures}/{tests.Count} smoke tests passed.");
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
    Equal(5, calibration.MissingFields.Count);

    calibration.SupplierCaptured = true;
    calibration.SupplierRefCaptured = true;
    calibration.PostingDateCaptured = true;
    calibration.RemarksCaptured = true;
    calibration.ItemNoCaptured = true;
    True(!calibration.IsComplete, "Legacy relative coordinates must not pass absolute desktop calibration.");
    calibration.CoordinateVersion = SapCalibration.AbsoluteDesktopCoordinateVersion;
    True(calibration.IsComplete, "All five captured positions should complete calibration.");
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

    const string beta8Json = """
        {
          "coordinateVersion": 2,
          "supplier": { "x": 160, "y": 40 },
          "supplierRef": { "x": 160, "y": 76 },
          "postingDate": { "x": 1720, "y": 75 },
          "documentDate": { "x": 1720, "y": 110 },
          "remarks": { "x": 240, "y": 770 },
          "itemNo": { "x": 520, "y": 260 },
          "supplierCaptured": true,
          "supplierRefCaptured": true,
          "postingDateCaptured": true,
          "documentDateCaptured": true,
          "remarksCaptured": true,
          "itemNoCaptured": true
        }
        """;
    var beta8Calibration = JsonSerializer.Deserialize<SapCalibration>(
        beta8Json,
        new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    True(beta8Calibration is not null && beta8Calibration.IsComplete,
        "Existing beta 8 absolute coordinates for the five required click targets should remain valid.");

    const string beta10Json = """
        {
          "coordinateVersion": 2,
          "supplier": { "x": 160, "y": 40 },
          "postingDate": { "x": 1720, "y": 75 },
          "remarks": { "x": 240, "y": 770 },
          "itemNo": { "x": 520, "y": 260 },
          "supplierCaptured": true,
          "postingDateCaptured": true,
          "remarksCaptured": true,
          "itemNoCaptured": true
        }
        """;
    var beta10Calibration = JsonSerializer.Deserialize<SapCalibration>(
        beta10Json,
        new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    True(beta10Calibration is not null && !beta10Calibration.IsComplete,
        "A four-point beta 10 calibration must request the new Supplier Ref. click target.");
    True(beta10Calibration!.MissingFields.SequenceEqual(new[] { "Supplier Ref." }),
        "Only Supplier Ref. should be missing from a complete beta 10 calibration.");
}

static void UsesStepByStepPasteOrder()
{
    Equal(5, SapPasteWorkflow.Steps.Count);
    Equal(SapPasteStep.Supplier, SapPasteWorkflow.Steps[0]);
    Equal(SapPasteStep.PostingDate, SapPasteWorkflow.Steps[1]);
    Equal(SapPasteStep.SupplierRef, SapPasteWorkflow.Steps[2]);
    Equal(SapPasteStep.Remarks, SapPasteWorkflow.Steps[3]);
    Equal(SapPasteStep.Items, SapPasteWorkflow.Steps[4]);
    Equal("Item No. (entire E:N block)", SapPasteWorkflow.GetLabel(SapPasteWorkflow.Steps[^1]));
    Equal(TimeSpan.FromMilliseconds(800), SapAutomationService.PasteInterval);
}

static void HandlesQuotedCells()
{
    var invoice = new ExcelClipboardParser().Parse(Row("\"Supplier \"\"One\"\"\"", "03-08-2026", "\"REF-1\"", "ITEM-1", "O", "1", "", "0", "V", "", "", "", "W") + "\r\n");
    Equal("Supplier \"One\"", invoice.SupplierName);
    Equal("REF-1", invoice.DocumentNumber);
    Equal("ITEM-1\tO\t1\t\t0\tV\t\t\t\tW", invoice.ItemClipboardBlock);
}

static void RejectsEmbeddedSeparators()
{
    foreach (var value in new[] { "\"ITEM\nSECOND\"", "\"ITEM\tSECOND\"", "ITEM\0SECOND" })
    {
        Throws<ClipboardValidationException>(() => new ExcelClipboardParser().Parse(
            Row("Supplier", "03-08-2026", "REF", value, "O", "1", "", "0", "V", "", "", "", "W")), "inside a cell");
    }
    Throws<ClipboardValidationException>(() => new ExcelClipboardParser().Parse("\"incomplete"), "incomplete");
}

static void FormatsInvariantDate()
{
    var previous = System.Globalization.CultureInfo.CurrentCulture;
    try
    {
        System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("th-TH");
        Equal("03.08.26", ExampleInvoice().SapDate);
    }
    finally
    {
        System.Globalization.CultureInfo.CurrentCulture = previous;
    }
}

static void RecoversFromClipboardContention()
{
    var attempts = 0;
    var notifications = 0;
    var task = ClipboardRetry.RunAsync(() => (++attempts >= 4, 42), onBusy: () => notifications++);
    True(!task.IsCompleted, "A busy clipboard retry must yield rather than block the caller.");
    Equal(42, task.GetAwaiter().GetResult());
    Equal(4, attempts);
    Equal(1, notifications);
}

static void BoundsClipboardRetries()
{
    Throws<InvalidOperationException>(() => ClipboardRetry.RunAsync(() => (false, 0), TimeSpan.FromMilliseconds(30)).GetAwaiter().GetResult(), "stayed busy");
    using var cancellation = new CancellationTokenSource();
    var task = ClipboardRetry.RunAsync(() => (false, 0), cancellationToken: cancellation.Token);
    cancellation.Cancel();
    try
    {
        task.GetAwaiter().GetResult();
        throw new InvalidOperationException("Expected clipboard retry cancellation.");
    }
    catch (OperationCanceledException)
    {
    }
}

static void WaitsForModifiers()
{
    var checks = 0;
    var task = InputService.WaitForModifiersReleasedAsync(() => ++checks < 3, CancellationToken.None);
    True(!task.IsCompleted, "Held shortcut modifiers must yield rather than block the UI.");
    task.GetAwaiter().GetResult();
    Equal(3, checks);
    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    try
    {
        InputService.WaitForModifiersReleasedAsync(() => true, cancellation.Token).GetAwaiter().GetResult();
        throw new InvalidOperationException("Expected modifier-wait cancellation.");
    }
    catch (OperationCanceledException)
    {
    }
}

static void RunsFivePastesWithOneRestore()
{
    var invoice = ExampleInvoice();
    var io = new FakeSapAutomationIO();
    var service = new SapAutomationService(io);
    var result = service.RunAllAsync(invoice, new SapCalibration(), io.ClipboardSequence).GetAwaiter().GetResult();
    Equal(invoice.Items.Count, result.ItemRows);
    Equal(6, io.Writes.Count);
    True(io.Pasted.SequenceEqual(new[] { invoice.SupplierName, invoice.SapDate, invoice.DocumentNumber, invoice.DocumentNumber, invoice.ItemClipboardBlock }),
        "Supplier or header values entered the wrong action or item block.");
    Equal(invoice.OriginalClipboardText, io.Writes[^1]);
    Equal(5, io.Clicks);
    Equal(4, io.Selections);
    Equal(5, io.Delays.Count(delay => delay >= TimeSpan.FromMilliseconds(800)));
}

static void StopsAtCoveredTarget()
{
    var io = new FakeSapAutomationIO { TargetCovered = true };
    Throws<SapAutomationException>(() => new SapAutomationService(io).RunAllAsync(ExampleInvoice(), new SapCalibration(), io.ClipboardSequence).GetAwaiter().GetResult(), "covered");
    Equal(0, io.Clicks);
    Equal(0, io.Pasted.Count);
    Equal(0, io.Writes.Count);
}

static void PreservesNewUserCopy()
{
    var io = new FakeSapAutomationIO();
    io.AfterPaste = () =>
    {
        if (io.Pasted.Count == 2)
        {
            io.ExternalCopy("USER NEW COPY");
        }
    };
    Throws<SapAutomationException>(() => new SapAutomationService(io).RunAllAsync(ExampleInvoice(), new SapCalibration(), io.ClipboardSequence).GetAwaiter().GetResult(), "clipboard changed");
    Equal("USER NEW COPY", io.CurrentText);
    Equal(2, io.Pasted.Count);
    Equal(2, io.Writes.Count);
}

static void RejectsChangedCopyAtStartup()
{
    var io = new FakeSapAutomationIO();
    var validatedSequence = io.ClipboardSequence;
    io.ExternalCopy("NEW COPY AFTER VALIDATION");
    Throws<SapAutomationException>(() => new SapAutomationService(io).RunAllAsync(
        ExampleInvoice(), new SapCalibration(), validatedSequence).GetAwaiter().GetResult(), "clipboard changed");
    Equal("NEW COPY AFTER VALIDATION", io.CurrentText);
    Equal(0, io.Writes.Count);
    Equal(0, io.Pasted.Count);
}

static void DoesNotReplayItems()
{
    var io = new FakeSapAutomationIO { FailAfterPasteNumber = 5 };
    Throws<SapAutomationException>(() => new SapAutomationService(io).RunAllAsync(ExampleInvoice(), new SapCalibration(), io.ClipboardSequence).GetAwaiter().GetResult(), "partial input");
    Equal(5, io.Pasted.Count);
    Equal(1, io.Pasted.Count(value => value == ExampleInvoice().ItemClipboardBlock));
    True(io.Delays[^1] >= TimeSpan.FromMilliseconds(800), "An uncertain item paste must finish its guard before restoration.");
}

static void CancelsBeforeNextField()
{
    using var cancellation = new CancellationTokenSource();
    var io = new FakeSapAutomationIO { AfterPaste = cancellation.Cancel };
    try
    {
        new SapAutomationService(io).RunAllAsync(ExampleInvoice(), new SapCalibration(), io.ClipboardSequence, cancellationToken: cancellation.Token).GetAwaiter().GetResult();
        throw new InvalidOperationException("Expected paste cancellation.");
    }
    catch (OperationCanceledException)
    {
    }
    Equal(1, io.Pasted.Count);
    Equal(2, io.Writes.Count);
    Equal(ExampleInvoice().OriginalClipboardText, io.CurrentText);
    Equal(TimeSpan.FromMilliseconds(800), io.Delays[^1]);
}

static InvoiceClipboardData ExampleInvoice() => new ExcelClipboardParser().Parse(
    Row("COL33 PTE.LTD", "03-08-2026", "COL26080630_F", "ITEM-1", "O-HW", "15", "", "7.5", "TX7", "", "", "", "S-HW"));

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
