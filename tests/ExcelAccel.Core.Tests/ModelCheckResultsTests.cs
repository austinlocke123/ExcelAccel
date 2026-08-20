using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using ExcelAccel.Persistence.ModelCheck;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ModelCheckCoordinatorTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void ASelectionScanReadsOnlyTheSelectedRectangle()
    {
        var port = new FakePort(Bounds(1, 1, 100, 4));
        var request = Request(ModelCheckScopeKind.Selection, "A1:B3");

        new ModelCheckCoordinator().Execute(port, request);

        var band = Assert.Single(port.RequestedBands);
        Assert.Equal("A1:B3", band.ToString());
    }

    [Fact]
    public void AWorksheetScanBandsTheUsedRegionAndReportsProgress()
    {
        var tracker = new OperationProgressTracker();
        var observed = new List<OperationProgress>();
        tracker.Changed += observed.Add;
        // 2,500 rows x 8 columns is 20,000 cells, which exceeds the 10,000-cell
        // block ceiling and so must be read as more than one band.
        var port = new FakePort(Bounds(1, 1, 2500, 8));

        new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"), tracker);

        Assert.True(port.RequestedBands.Count > 1, $"expected more than one band, read {port.RequestedBands.Count}");
        Assert.Contains(observed, progress => progress.Phase == OperationPhase.Snapshot);
        Assert.Contains(observed, progress => progress.Phase == OperationPhase.Analyze);
        Assert.Equal(OperationPhase.Completed, observed[observed.Count - 1].Phase);
    }

    [Fact]
    public void WorkbookScopeIsRefusedWithoutReadingAnything()
    {
        var port = new FakePort(Bounds(1, 1, 10, 2));

        var result = new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Workbook, "A1"));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(ModelCheckRefusalCodes.ScopeTooLarge, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void AnInflatedUsedRegionIsRefusedWithoutReadingAnything()
    {
        var port = new FakePort(Bounds(1, 1, AuditAddress.MaximumRow, AuditAddress.MaximumColumn));

        var result = new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void ALargeWorksheetScanIsRefusedWhenNotConfirmed()
    {
        var port = new FakePort(Bounds(1, 1, 30_000, 1));

        var result = new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"), null, _ => false);

        Assert.Equal(ModelCheckRefusalCodes.PreviewRequired, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void AConfirmedLargeWorksheetScanDescribesThePlannedRead()
    {
        var port = new FakePort(Bounds(1, 1, 30_000, 1));
        ModelCheckScanPreview? observed = null;

        new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"), null, preview =>
        {
            observed = preview;
            return true;
        });

        Assert.NotNull(observed);
        Assert.Equal(Sheet, observed!.WorksheetName);
        Assert.Equal(30_000, observed.CellCount);
        Assert.Equal(ModelCheckRuleCatalog.All.Count, observed.RuleCount);
    }

    [Fact]
    public void CancellationRefusesAndReadsNoFurtherBlocks()
    {
        var tracker = new OperationProgressTracker();
        var port = new FakePort(Bounds(1, 1, 2500, 4)) { OnBlock = index => { if (index == 0) tracker.RequestCancellation(); } };

        var result = new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"), tracker);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(ModelCheckRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Empty(result.Findings);
    }

    [Fact]
    public void ARescanRepeatsTheExactScopeAndRuleSet()
    {
        var request = Request(ModelCheckScopeKind.Worksheet, "A1");
        var ignores = ModelCheckConfiguration.Default.WithIgnoredFingerprints(new[] { new string('a', 64) });

        var repeated = request.WithConfiguration(ignores);

        Assert.Equal(request.Scope, repeated.Scope);
        Assert.Equal(request.Target.ToString(), repeated.Target.ToString());
        Assert.Equal(request.RuleIds, repeated.RuleIds);
        Assert.Same(ignores, repeated.Configuration);
    }

    [Fact]
    public void TheReportProjectsIntoTheSharedViewWithNavigableFindings()
    {
        var port = new FakePort(Bounds(1, 1, 10, 2), Cell("B1", "=A1*7"));

        var result = new ModelCheckCoordinator().Execute(port, Request(ModelCheckScopeKind.Worksheet, "A1"));
        var presentation = ModelCheckReport.Create(result).ToPresentation();

        Assert.Equal("ExcelAccel Model Check", presentation.Title);
        Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Values.Count));
        Assert.All(presentation.Rows, row => Assert.True(row.IsNavigable));
        Assert.Contains(presentation.SummaryLines, line => line.StartsWith("Cells scanned:", StringComparison.Ordinal));
    }

    [Fact]
    public void EveryModelCheckCommandIsRegisteredWithItsAcceptanceIds()
    {
        foreach (var id in new[]
        {
            ModelCheckCommandCatalog.RunSelectionId,
            ModelCheckCommandCatalog.RunWorksheetId,
            ModelCheckCommandCatalog.RescanId,
            ModelCheckCommandCatalog.IgnoreLocalId,
            ModelCheckCommandCatalog.UnignoreLocalId,
            ModelCheckCommandCatalog.ExportId,
        })
        {
            var descriptor = BuiltInCommandRegistry.GetRequired(id);
            Assert.NotEmpty(descriptor.AcceptanceIds);
            Assert.Equal("CAP-CHECK-001", descriptor.CapabilityId);
        }

        Assert.Equal(
            Core.Commands.CommandImpact.ReadOnly,
            BuiltInCommandRegistry.GetRequired(ModelCheckCommandCatalog.RunSelectionId).Impact);
        Assert.Equal(
            PreviewPolicy.Mandatory,
            BuiltInCommandRegistry.GetRequired(ModelCheckCommandCatalog.ExportId).PreviewPolicy);
    }

    [Fact]
    public void TheCoordinatorRejectsNullArguments()
    {
        var coordinator = new ModelCheckCoordinator();
        Assert.Throws<ArgumentNullException>(() => coordinator.Execute(null!, Request(ModelCheckScopeKind.Selection, "A1")));
        Assert.Throws<ArgumentNullException>(() => coordinator.Execute(new FakePort(Bounds(1, 1, 1, 1)), null!));
    }

    private static AuditCellIdentity Identity(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static ModelCheckCell Cell(string address, string formula) => new ModelCheckCell(
        Identity(address), formula, AuditCellClassification.Formula, "General");

    private static UsedRegionBounds Bounds(int firstRow, int firstColumn, int rowCount, int columnCount) =>
        new UsedRegionBounds(Sheet, firstRow, firstColumn, rowCount, columnCount);

    private static ModelCheckScanRequest Request(ModelCheckScopeKind scope, string address) =>
        new ModelCheckScanRequest(scope, Identity(address), ModelCheckRuleCatalog.AllRuleIds, ModelCheckConfiguration.Default);

    private sealed class FakePort : IModelCheckSnapshotPort
    {
        private readonly UsedRegionBounds _bounds;
        private readonly ModelCheckCell[] _cells;

        public FakePort(UsedRegionBounds bounds, params ModelCheckCell[] cells)
        {
            _bounds = bounds;
            _cells = cells;
        }

        public List<AuditRectangle> RequestedBands { get; } = new List<AuditRectangle>();

        public Action<int>? OnBlock { get; set; }

        public AuditCellIdentity CaptureTarget() => Identity("A1");

        public UsedRegionBounds CaptureUsedRegion(string worksheetName) => _bounds;

        public IReadOnlyList<ModelCheckCell> CaptureBlock(string worksheetName, AuditRectangle band)
        {
            OnBlock?.Invoke(RequestedBands.Count);
            RequestedBands.Add(band);
            return _cells
                .Where(cell => AuditAddress.TryParse(cell.Identity.Address, out var rectangle) && band.Intersects(rectangle))
                .ToArray();
        }
    }
}

public sealed class ModelCheckIgnoreStoreTests
{
    [Fact]
    public void IgnoresRoundTripAndStoreNoWorkbookContent()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "ignores.tsv");
        var store = new ModelCheckIgnoreStore();
        var entry = new ModelCheckIgnoreEntry("check.formula.error", 1, new string('a', 64), "Model!B2");

        store.SaveAtomic(path, new[] { entry });
        var loaded = store.Load(path);

        var reloaded = Assert.Single(loaded);
        Assert.Equal(entry.RuleId, reloaded.RuleId);
        Assert.Equal(entry.Fingerprint, reloaded.Fingerprint);
        var text = File.ReadAllText(path);
        Assert.DoesNotContain("=", text);
    }

    [Fact]
    public void ARejectedFingerprintIsNeverStored() =>
        Assert.Throws<ArgumentException>(() => new ModelCheckIgnoreEntry("check.formula.error", 1, "not-a-digest"));

    [Fact]
    public void DuplicateFingerprintsCollapseAndOrderIsCanonical()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "ignores.tsv");
        var store = new ModelCheckIgnoreStore();
        var first = new ModelCheckIgnoreEntry("check.b", 1, new string('b', 64));
        var second = new ModelCheckIgnoreEntry("check.a", 1, new string('a', 64));

        store.SaveAtomic(path, new[] { first, second, first });
        var loaded = store.Load(path);

        Assert.Equal(2, loaded.Count);
        Assert.Equal("check.a", loaded[0].RuleId);
    }

    [Fact]
    public void AMissingFileLoadsAsAnEmptySet()
    {
        using var sandbox = new TempDirectory();

        Assert.Empty(new ModelCheckIgnoreStore().Load(Path.Combine(sandbox.Path, "absent.tsv")));
    }
}

public sealed class ModelCheckResultExporterTests
{
    [Fact]
    public void TheDefaultManifestExcludesFormulasValuesAndEvidence()
    {
        var manifest = new ModelCheckResultExporter().Plan("C:\\out\\findings.csv", new[] { Finding() }, includeEvidence: false);

        Assert.Contains("cell formulas", manifest.ExcludedFields);
        Assert.Contains("cell values", manifest.ExcludedFields);
        Assert.Contains("rule evidence", manifest.ExcludedFields);
        Assert.DoesNotContain("evidence", manifest.IncludedFields);
        Assert.Contains(manifest.Lines, line => line.Contains("Nothing is transmitted"));
    }

    [Fact]
    public void ADefaultExportWritesNoEvidenceColumn()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "findings.csv");
        var exporter = new ModelCheckResultExporter();
        var findings = new[] { Finding() };

        exporter.Export(exporter.Plan(path, findings, false), findings);

        var lines = File.ReadAllLines(path);
        Assert.Equal(2, lines.Length);
        Assert.DoesNotContain("evidence", lines[0]);
        Assert.DoesNotContain("Peer region", lines[1]);
    }

    [Fact]
    public void OptingInAddsTheEvidenceColumn()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "findings.csv");
        var exporter = new ModelCheckResultExporter();
        var findings = new[] { Finding() };

        exporter.Export(exporter.Plan(path, findings, true), findings);

        var lines = File.ReadAllLines(path);
        Assert.Contains("evidence", lines[0]);
        Assert.Contains("Peer region", lines[1]);
    }

    [Fact]
    public void AnExportReplacesTheDestinationDeterministically()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "findings.csv");
        File.WriteAllText(path, "stale");
        var exporter = new ModelCheckResultExporter();
        var findings = new[] { Finding() };

        exporter.Export(exporter.Plan(path, findings, false), findings);
        var first = File.ReadAllText(path);
        exporter.Export(exporter.Plan(path, findings, false), findings);

        Assert.Equal(first, File.ReadAllText(path));
        Assert.DoesNotContain("stale", first);
    }

    [Fact]
    public void AFailedExportLeavesExistingDestinationDataIntact()
    {
        using var sandbox = new TempDirectory();
        var path = Path.Combine(sandbox.Path, "findings.csv");
        File.WriteAllText(path, "original");
        var exporter = new ModelCheckResultExporter();
        var manifest = exporter.Plan(path, Array.Empty<ExportableFinding>(), false);

        Assert.Throws<InvalidDataException>(() => exporter.Export(
            manifest,
            Enumerable.Range(0, ModelCheckResultExporter.MaximumFindings + 1).Select(_ => Finding()).ToArray()));
        Assert.Equal("original", File.ReadAllText(path));
    }

    [Fact]
    public void AnEmptyDestinationIsRejected() =>
        Assert.Throws<ArgumentException>(() => new ModelCheckResultExporter().Plan(" ", Array.Empty<ExportableFinding>(), false));

    private static ExportableFinding Finding() => new ExportableFinding(
        "check.formula.pattern_inconsistency",
        1,
        "Attention",
        "Model",
        "B3",
        "A formula differs from the shape its peer region shares.",
        "Exact",
        new string('a', 64),
        new[] { "Peer region: Model!B1:B4" });
}

internal sealed class TempDirectory : IDisposable
{
    public TempDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excelaccel-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Path)) Directory.Delete(Path, true);
        }
        catch (IOException)
        {
        }
    }
}
