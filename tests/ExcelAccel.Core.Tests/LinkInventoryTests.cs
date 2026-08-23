using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Links;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Links;
using ExcelAccel.Persistence.Links;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class LinkInventoryTests
{
    private static LinkUsageRecord Formula(string sourceToken, string worksheet, string address, bool broken = false) =>
        new LinkUsageRecord(LinkCategory.FormulaReference, sourceToken, worksheet, address, sourceToken, broken);

    private static LinkUsageRecord Name(string sourceToken, string detail, bool broken = false) =>
        new LinkUsageRecord(LinkCategory.DefinedName, sourceToken, null, null, detail, broken);

    private static LinkInventory Build(
        IEnumerable<LinkSourceRecord>? sources = null,
        IEnumerable<LinkUsageRecord>? usages = null) =>
        LinkInventoryBuilder.Build(
            sources ?? Array.Empty<LinkSourceRecord>(),
            usages ?? Array.Empty<LinkUsageRecord>());

    /// <summary>
    /// A formula token and a full path naming the same workbook must land in one
    /// group, or the same link appears twice under different spellings.
    /// </summary>
    [Theory]
    [InlineData("='[Budget.xlsx]Sheet1'!$A$1", "budget.xlsx")]
    [InlineData("C:\\Models\\Budget.xlsx", "budget.xlsx")]
    [InlineData("\\\\server\\share\\Budget.xlsx", "budget.xlsx")]
    [InlineData("'C:\\Models\\[Budget.xlsx]Sheet1'", "budget.xlsx")]
    [InlineData("BUDGET.XLSX", "budget.xlsx")]
    [InlineData("", "")]
    public void SourcesNormalizeToTheirFileName(string raw, string expected) =>
        Assert.Equal(expected, LinkInventoryBuilder.NormalizeToken(raw));

    [Fact]
    public void UsagesGroupUnderOneSourceRegardlessOfHowTheyWereSpelled()
    {
        var inventory = Build(
            sources: new[] { new LinkSourceRecord("C:\\Models\\Budget.xlsx", false) },
            usages: new[]
            {
                Formula("='[Budget.xlsx]Sheet1'!$A$1", "Sheet1", "A1"),
                Name("='[BUDGET.XLSX]Sheet2'!$B$2", "Name: Rate"),
            });

        var source = Assert.Single(inventory.Sources);
        Assert.Equal("budget.xlsx", source.Token);
        Assert.Equal(2, source.UsageCount);
        Assert.Equal("C:\\Models\\Budget.xlsx", source.DisplaySource);
    }

    /// <summary>
    /// Nothing checks whether the file exists or is reachable, so the label must
    /// not imply that it does (AC-LINK-003).
    /// </summary>
    [Fact]
    public void StatusDescribesThisSessionAndClaimsNoNetworkOrDiskCheck()
    {
        var closed = Build(sources: new[] { new LinkSourceRecord("Budget.xlsx", false) });
        var open = Build(sources: new[] { new LinkSourceRecord("Budget.xlsx", true) });

        Assert.Equal(LinkStatus.NotOpenInSession, closed.Sources.Single().Status);
        Assert.Equal(LinkStatus.OpenInSession, open.Sources.Single().Status);

        var label = LinkInventoryPresentation.Label(LinkStatus.NotOpenInSession);
        Assert.Contains("not checked", label, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("missing", label, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("inaccessible", label, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("unavailable", label, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ABrokenUsageMakesItsWholeSourceBroken()
    {
        var inventory = Build(
            sources: new[] { new LinkSourceRecord("Budget.xlsx", true) },
            usages: new[]
            {
                Formula("='[Budget.xlsx]Sheet1'!$A$1", "Sheet1", "A1"),
                Formula("='[Budget.xlsx]Sheet1'!#REF!", "Sheet1", "A2", broken: true),
            });

        Assert.Equal(LinkStatus.Broken, inventory.Sources.Single().Status);
    }

    /// <summary>
    /// Excel reports the source, so hiding it because nothing scanned explains it
    /// would contradict what the user sees in Excel's own dialog.
    /// </summary>
    [Fact]
    public void ASourceWithNoScannedUsageStaysVisible()
    {
        var inventory = Build(sources: new[] { new LinkSourceRecord("C:\\Models\\Orphan.xlsx", false) });

        var source = Assert.Single(inventory.Sources);
        Assert.Empty(source.Usages);

        var presentation = LinkInventoryPresentation.Create(inventory, "Book.xlsx");
        Assert.Contains("No usage found", presentation.Rows.Single().Values.Last(), StringComparison.Ordinal);
    }

    [Fact]
    public void AUsageWhoseSourceExcelDidNotReportStillAppears()
    {
        var inventory = Build(usages: new[] { Formula("='[Unreported.xlsx]Sheet1'!$A$1", "Sheet1", "A1") });

        Assert.Equal("unreported.xlsx", inventory.Sources.Single().Token);
    }

    [Fact]
    public void UnscannedCategoriesAreReportedAsGapsRatherThanCountedAsZero()
    {
        var inventory = Build(usages: new[] { Formula("='[Budget.xlsx]S'!$A$1", "Sheet1", "A1") });

        Assert.False(inventory.IsComplete);
        Assert.Contains("category_not_scanned:chartseries", inventory.CoverageGaps);
        Assert.Contains("category_not_scanned:queryconnection", inventory.CoverageGaps);
        Assert.Contains("category_not_scanned:datavalidation", inventory.CoverageGaps);
        Assert.Contains("Partial", LinkInventorySearch.CoverageStatement(inventory), StringComparison.Ordinal);
    }

    [Fact]
    public void ScanningEveryCategoryReportsComplete()
    {
        var inventory = LinkInventoryBuilder.Build(
            Array.Empty<LinkSourceRecord>(),
            new[] { Formula("='[Budget.xlsx]S'!$A$1", "Sheet1", "A1") },
            Enum.GetValues(typeof(LinkCategory)).Cast<LinkCategory>());

        Assert.True(inventory.IsComplete);
        Assert.Contains("Complete", LinkInventorySearch.CoverageStatement(inventory), StringComparison.Ordinal);
    }

    [Fact]
    public void OrderingIsDeterministicRegardlessOfCaptureOrder()
    {
        var usages = new[]
        {
            Formula("='[Zeta.xlsx]S'!$A$1", "Sheet2", "B2"),
            Formula("='[Alpha.xlsx]S'!$A$1", "Sheet1", "A1"),
            Name("='[Alpha.xlsx]S'!$A$2", "Name: Rate"),
        };

        var forward = Build(usages: usages).Sources.Select(source => source.Token);
        var reversed = Build(usages: usages.Reverse().ToArray()).Sources.Select(source => source.Token);

        Assert.Equal(forward, reversed);
        Assert.Equal(new[] { "alpha.xlsx", "zeta.xlsx" }, forward);
    }

    [Fact]
    public void OnlyACellAnchoredUsageIsNavigable()
    {
        var inventory = Build(usages: new[]
        {
            Formula("='[Budget.xlsx]S'!$A$1", "Sheet1", "A1"),
            Name("='[Budget.xlsx]S'!$A$2", "Name: Rate"),
        });

        var presentation = LinkInventoryPresentation.Create(inventory, "Book.xlsx");

        Assert.Equal(1, presentation.Rows.Count(row => row.IsNavigable));
        var target = presentation.Rows.Single(row => row.IsNavigable).NavigationTarget!;
        Assert.Equal("Sheet1", target.WorksheetName);
        Assert.Equal("A1", target.Address);
        Assert.Equal("Book.xlsx", target.WorkbookId);
    }

    [Fact]
    public void ANonNavigableUsageStatesWhy()
    {
        var inventory = Build(usages: new[] { Name("='[Budget.xlsx]S'!$A$1", "Name: Rate") });

        var usage = inventory.Sources.Single().Usages.Single();

        Assert.False(usage.IsNavigable);
        Assert.False(string.IsNullOrWhiteSpace(usage.NonNavigableReason));
    }

    [Fact]
    public void ABrokenUsageRefusesNavigation()
    {
        var inventory = Build(usages: new[] { Formula("='[Budget.xlsx]S'!#REF!", "Sheet1", "A1", broken: true) });

        var usage = inventory.Sources.Single().Usages.Single();

        Assert.False(usage.IsNavigable);
        Assert.Contains("no longer resolves", usage.NonNavigableReason!, StringComparison.Ordinal);
    }

    [Fact]
    public void ExceedingTheUsageBoundIsReportedRatherThanSilentlyTruncated()
    {
        var usages = Enumerable.Range(0, LinkInventoryBuilder.MaximumUsages + 3)
            .Select(index => Formula("='[Budget.xlsx]S'!$A$1", "Sheet1", "A" + index))
            .ToArray();

        var inventory = Build(usages: usages);

        Assert.Equal(LinkInventoryBuilder.MaximumUsages, inventory.UsageCount);
        Assert.Equal(3, inventory.ExcludedByBound);
        Assert.Contains("usage_count_exceeded_qualified_bound", inventory.CoverageGaps);
    }

    [Fact]
    public void SearchFiltersDeterministicallyOverTheCapturedResult()
    {
        var inventory = Build(usages: new[]
        {
            Formula("='[Budget.xlsx]S'!$A$1", "Model", "A1"),
            Formula("='[Forecast.xlsx]S'!$A$1", "Other", "B2"),
        });

        Assert.Single(LinkInventorySearch.Filter(inventory, "budget"));
        Assert.Single(LinkInventorySearch.Filter(inventory, "model"));
        Assert.Equal(2, LinkInventorySearch.Filter(inventory, "  ").Count);
        Assert.Equal(2, LinkInventorySearch.Filter(inventory, null, status: LinkStatus.NotOpenInSession).Count);
        Assert.Empty(LinkInventorySearch.Filter(inventory, null, category: LinkCategory.ChartSeries));
    }

    [Fact]
    public void EveryPresentationRowSuppliesOneValuePerColumn()
    {
        var inventory = Build(
            sources: new[] { new LinkSourceRecord("Orphan.xlsx", false) },
            usages: new[] { Formula("='[Budget.xlsx]S'!$A$1", "Sheet1", "A1") });

        var presentation = LinkInventoryPresentation.Create(inventory, "Book.xlsx");

        Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Values.Count));
        Assert.Equal(AuditTraceStatus.Partial, presentation.Status);
    }
}

public sealed class LinkInventoryCoordinatorTests
{
    [Fact]
    public void OpeningCapturesOnceAndSearchingDoesNotTouchExcelAgain()
    {
        var port = new FakePort();
        var coordinator = new LinkInventoryCoordinator();

        var opened = coordinator.Open(port);
        var filtered = coordinator.Search(opened, "budget");

        Assert.Equal(1, port.UsageCaptureCount);
        Assert.Equal(2, opened.Inventory.Sources.Count);
        Assert.Single(filtered.Presentation.Rows);
    }

    [Fact]
    public void AWorkbookThatCannotBeIdentifiedIsRefused()
    {
        var refusal = Assert.Throws<CommandRefusedException>(
            () => new LinkInventoryCoordinator().Open(new FakePort { WorkbookId = string.Empty }));

        Assert.Equal(RefusalCodes.SelectionUnsupported, refusal.RefusalCode);
    }

    [Fact]
    public void AWorkbookWithNoLinksOpensCleanly()
    {
        var result = new LinkInventoryCoordinator().Open(new FakePort { Empty = true });

        Assert.Empty(result.Inventory.Sources);
    }

    [Fact]
    public void TheInventoryCommandIsRegisteredAsReadOnlyWithNoUndo()
    {
        var descriptor = BuiltInCommandRegistry.GetRequired(LinksCommandCatalog.InventoryId);

        Assert.Equal(CommandImpact.ReadOnly, descriptor.Impact);
        Assert.Empty(descriptor.ChangedProperties);
        Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
        Assert.Equal("CAP-LINK-001", descriptor.CapabilityId);
    }

    private sealed class FakePort : ILinkInventoryPort
    {
        public string WorkbookId { get; set; } = "Book.xlsx";
        public bool Empty { get; set; }
        public int UsageCaptureCount { get; private set; }

        public IReadOnlyList<LinkCategory> ScannedCategories { get; } =
            new[] { LinkCategory.FormulaReference, LinkCategory.DefinedName };

        public string CaptureWorkbookId() => WorkbookId;

        public IReadOnlyList<LinkSourceRecord> CaptureSources() => Empty
            ? Array.Empty<LinkSourceRecord>()
            : new[] { new LinkSourceRecord("C:\\Models\\Budget.xlsx", false), new LinkSourceRecord("Forecast.xlsx", true) };

        public IReadOnlyList<LinkUsageRecord> CaptureUsages()
        {
            UsageCaptureCount++;
            return Empty
                ? Array.Empty<LinkUsageRecord>()
                : new[]
                {
                    new LinkUsageRecord(LinkCategory.FormulaReference, "='[Budget.xlsx]S'!$A$1", "Sheet1", "A1", "budget", false),
                };
        }
    }
}

public sealed class LinkInventoryExporterTests
{
    private static LinkInventory Sample() => LinkInventoryBuilder.Build(
        new[] { new LinkSourceRecord("C:\\Deals\\ProjectFalcon.xlsx", false) },
        new[]
        {
            new LinkUsageRecord(
                LinkCategory.FormulaReference,
                "='[ProjectFalcon.xlsx]Sheet1'!$A$1",
                "Sheet1",
                "A1",
                "='[ProjectFalcon.xlsx]Sheet1'!$A$1",
                false),
        });

    /// <summary>
    /// A link path routinely names a drive, a client, or a deal, so it is
    /// redacted to the file name unless the user opts in (AC-LINK-011).
    /// </summary>
    [Fact]
    public void SourcePathsAreRedactedByDefault()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = Sample();
        var exporter = new LinkInventoryExporter();
        var redacted = Path.Combine(sandbox.Path, "redacted.csv");
        var full = Path.Combine(sandbox.Path, "full.csv");

        exporter.Export(exporter.Plan(redacted, inventory, false), inventory);
        exporter.Export(exporter.Plan(full, inventory, true), inventory);

        var redactedText = File.ReadAllText(redacted);
        Assert.DoesNotContain("C:\\Deals", redactedText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("projectfalcon.xlsx", redactedText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("C:\\Deals", File.ReadAllText(full), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheManifestStatesCountsFieldsAndCoverage()
    {
        var manifest = new LinkInventoryExporter()
            .Plan(Path.Combine(Path.GetTempPath(), "links.csv"), Sample(), false);

        Assert.Equal(1, manifest.SourceCount);
        Assert.Equal(1, manifest.UsageCount);
        Assert.False(manifest.IncludePaths);
        Assert.Contains("redacted to file name", manifest.ToString(), StringComparison.Ordinal);
        Assert.Contains("Partial", manifest.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void ASourceWithNoUsageStillGetsARow()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = LinkInventoryBuilder.Build(
            new[] { new LinkSourceRecord("Orphan.xlsx", false) }, Array.Empty<LinkUsageRecord>());
        var exporter = new LinkInventoryExporter();
        var path = Path.Combine(sandbox.Path, "links.csv");

        exporter.Export(exporter.Plan(path, inventory, false), inventory);

        Assert.Contains("orphan.xlsx", File.ReadAllText(path), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RewritingAnExistingExportLeavesNoTemporaryFile()
    {
        using var sandbox = new TemporaryDirectory();
        var inventory = Sample();
        var exporter = new LinkInventoryExporter();
        var path = Path.Combine(sandbox.Path, "links.csv");

        exporter.Export(exporter.Plan(path, inventory, false), inventory);
        exporter.Export(exporter.Plan(path, inventory, false), inventory);

        Assert.Empty(Directory.GetFiles(sandbox.Path, "*.tmp"));
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excelaccel-links-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
        }
    }
}
