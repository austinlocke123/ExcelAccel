using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Names;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Names;
using ExcelAccel.Persistence.Names;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class NameInventoryCoordinatorTests
{
    private static NameRecord Workbook(string name, string refersTo) =>
        new NameRecord(name, NameScopeKind.Workbook, null, refersTo, true, false);

    [Fact]
    public void OpeningCapturesOnceAndProjectsThePresentation()
    {
        var port = new FakePort("Book.xlsx", Workbook("Rate", "=Sheet1!$A$1"), Workbook("Cap", "=42"));

        var result = new NameInventoryCoordinator().Open(port, new NameInventoryRequest(true, true));

        Assert.Equal(2, result.Inventory.Entries.Count);
        Assert.Equal("Book.xlsx", result.WorkbookId);
        Assert.Equal(2, result.Presentation.Rows.Count);
        Assert.Equal(1, port.CaptureCount);
    }

    /// <summary>
    /// Searching must filter the captured snapshot. Re-reading would let results
    /// change under a filter the user did not touch, and would put a workbook
    /// scan behind every keystroke.
    /// </summary>
    [Fact]
    public void SearchingFiltersTheSnapshotWithoutTouchingTheWorkbookAgain()
    {
        var port = new FakePort("Book.xlsx", Workbook("GrowthRate", "=Sheet1!$A$1"), Workbook("Discount", "=42"));
        var coordinator = new NameInventoryCoordinator();
        var opened = coordinator.Open(port, new NameInventoryRequest(true, true));

        var filtered = coordinator.Search(opened, "growth");

        Assert.Single(filtered.Presentation.Rows);
        Assert.Equal(2, filtered.Inventory.Entries.Count);
        Assert.Equal(1, port.CaptureCount);
    }

    [Fact]
    public void AWorkbookThatCannotBeIdentifiedIsRefusedRatherThanGuessed()
    {
        var port = new FakePort(string.Empty, Workbook("Rate", "=Sheet1!$A$1"));

        var refusal = Assert.Throws<CommandRefusedException>(
            () => new NameInventoryCoordinator().Open(port, new NameInventoryRequest(true, true)));

        Assert.Equal(RefusalCodes.SelectionUnsupported, refusal.RefusalCode);
    }

    [Fact]
    public void AnEmptyWorkbookOpensCleanlyRatherThanRefusing()
    {
        var port = new FakePort("Book.xlsx");

        var result = new NameInventoryCoordinator().Open(port, new NameInventoryRequest(true, true));

        Assert.Empty(result.Inventory.Entries);
        Assert.True(result.Inventory.IsComplete);
    }

    [Fact]
    public void TheInventoryCommandIsRegisteredAsReadOnlyWithNoUndo()
    {
        var descriptor = BuiltInCommandRegistry.GetRequired(NamesCommandCatalog.InventoryId);

        Assert.Equal(CommandImpact.ReadOnly, descriptor.Impact);
        Assert.Empty(descriptor.ChangedProperties);
        Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
        Assert.Equal("CAP-NAME-001", descriptor.CapabilityId);
    }

    /// <summary>
    /// The spec forbids these ids existing at all until new requirements and ADRs
    /// are accepted, so their absence is asserted rather than assumed.
    /// </summary>
    [Theory]
    [InlineData("names.rename")]
    [InlineData("names.delete")]
    [InlineData("links.repoint")]
    [InlineData("links.break")]
    public void MutationCommandsDoNotExist(string commandId) =>
        Assert.DoesNotContain(BuiltInCommandRegistry.All, command => command.Id == commandId);

    private sealed class FakePort : INameInventoryPort
    {
        private readonly string _workbookId;
        private readonly NameRecord[] _records;

        public FakePort(string workbookId, params NameRecord[] records)
        {
            _workbookId = workbookId;
            _records = records;
        }

        public int CaptureCount { get; private set; }

        public string CaptureWorkbookId() => _workbookId;

        public IReadOnlyList<NameRecord> CaptureNames()
        {
            CaptureCount++;
            return _records;
        }
    }
}

public sealed class NameInventoryExporterTests
{
    private static NameInventory Inventory(params NameRecord[] records) =>
        NameInventoryBuilder.Build(records, true, true);

    private static NameRecord Workbook(string name, string refersTo) =>
        new NameRecord(name, NameScopeKind.Workbook, null, refersTo, true, false);

    [Fact]
    public void ExpressionsAreExcludedByDefaultAndIncludedOnlyOnRequest()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = Inventory(Workbook("Secret", "='[Budget.xlsx]Q3'!$A$1"));
        var exporter = new NameInventoryExporter();
        var redactedPath = Path.Combine(sandbox.Path, "redacted.csv");
        var fullPath = Path.Combine(sandbox.Path, "full.csv");

        exporter.Export(exporter.Plan(redactedPath, inventory, false), inventory);
        exporter.Export(exporter.Plan(fullPath, inventory, true), inventory);

        Assert.DoesNotContain("Budget.xlsx", File.ReadAllText(redactedPath), StringComparison.Ordinal);
        Assert.Contains("Budget.xlsx", File.ReadAllText(fullPath), StringComparison.Ordinal);
    }

    [Fact]
    public void TheManifestStatesDestinationCountFieldsAndCoverage()
    {
        var inventory = Inventory(Workbook("Rate", "=Sheet1!$A$1"));

        var manifest = new NameInventoryExporter().Plan(Path.Combine(Path.GetTempPath(), "names.csv"), inventory, false);

        Assert.Equal(1, manifest.NameCount);
        Assert.False(manifest.IncludeExpressions);
        Assert.DoesNotContain("refers_to", manifest.IncludedFields);
        Assert.Contains("Complete", manifest.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheExportCarriesItsCoverageStatementAsAHeader()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = Inventory(Workbook("Odd", "=!!!"));
        var exporter = new NameInventoryExporter();
        var path = Path.Combine(sandbox.Path, "names.csv");

        exporter.Export(exporter.Plan(path, inventory, false), inventory);

        var text = File.ReadAllText(path);
        Assert.StartsWith("# ", text, StringComparison.Ordinal);
        Assert.Contains("Partial", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RewritingAnExistingExportReplacesItWhole()
    {
        using var sandbox = new TemporaryDirectory();
        var exporter = new NameInventoryExporter();
        var path = Path.Combine(sandbox.Path, "names.csv");
        var first = Inventory(Workbook("First", "=Sheet1!$A$1"));
        var second = Inventory(Workbook("Second", "=Sheet1!$A$2"));

        exporter.Export(exporter.Plan(path, first, false), first);
        exporter.Export(exporter.Plan(path, second, false), second);

        var text = File.ReadAllText(path);
        Assert.Contains("Second", text, StringComparison.Ordinal);
        Assert.DoesNotContain("First", text, StringComparison.Ordinal);
        Assert.Empty(Directory.GetFiles(sandbox.Path, "*.tmp"));
    }

    [Fact]
    public void AnEmptyDestinationIsRefused() =>
        Assert.Throws<ArgumentException>(() => new NameInventoryExporter().Plan(" ", Inventory(), false));

    /// <summary>
    /// The existing helpers of this name are private to their own test classes,
    /// so this one is local rather than reaching across.
    /// </summary>
    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excelaccel-names-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
        }
    }

    [Fact]
    public void FieldsAreQuotedSoACommaCannotShiftAColumn()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = Inventory(new NameRecord("Comma,Name", NameScopeKind.Workbook, null, "=Sheet1!$A$1", true, false));
        var exporter = new NameInventoryExporter();
        var path = Path.Combine(sandbox.Path, "names.csv");

        exporter.Export(exporter.Plan(path, inventory, false), inventory);

        Assert.Contains("\"Comma,Name\"", File.ReadAllText(path), StringComparison.Ordinal);
    }
}
