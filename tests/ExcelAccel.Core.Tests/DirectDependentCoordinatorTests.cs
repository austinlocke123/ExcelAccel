using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DependentScanRegionTests
{
    [Fact]
    public void AnEmptyUsedRegionPlansNoBlocksInsteadOfRefusing()
    {
        Assert.True(DependentScanRegion.TryCreate(Bounds(1, 1, 0, 0), out var region, out var code, out _));

        Assert.Null(code);
        Assert.Equal(0, region!.BlockCount);
        Assert.Equal(0, region.CellCount);
    }

    [Fact]
    public void BandsCoverTheRegionExactlyWithoutOverlapOrGap()
    {
        Assert.True(DependentScanRegion.TryCreate(Bounds(5, 2, 2500, 8), out var region, out _, out _));

        var covered = new HashSet<(int Row, int Column)>();
        for (var index = 0; index < region!.BlockCount; index++)
        {
            var band = region.Block(index);
            Assert.Equal(region.FirstColumn, band.FirstColumn);
            Assert.Equal(region.LastColumn, band.LastColumn);
            for (var row = band.FirstRow; row <= band.LastRow; row++)
            {
                for (var column = band.FirstColumn; column <= band.LastColumn; column++)
                {
                    Assert.True(covered.Add((row, column)), $"Cell {column}/{row} was read twice.");
                }
            }
        }

        Assert.Equal(region.CellCount, covered.Count);
        Assert.Equal(region.FirstRow, covered.Min(cell => cell.Row));
        Assert.Equal(region.LastRow, covered.Max(cell => cell.Row));
    }

    [Fact]
    public void NoSingleBandExceedsTheBoundedBlockCeiling()
    {
        Assert.True(DependentScanRegion.TryCreate(Bounds(1, 1, 20_000, 7), out var region, out _, out _));

        for (var index = 0; index < region!.BlockCount; index++)
        {
            var band = region.Block(index);
            var cells = (long)(band.LastRow - band.FirstRow + 1) * (band.LastColumn - band.FirstColumn + 1);
            Assert.InRange(cells, 1, DependentScanRegion.MaximumBlockCells);
        }
    }

    /// <summary>
    /// Excel routinely reports a used range far larger than the real content
    /// because of stray formatting. The reported range is never a resource bound.
    /// </summary>
    [Fact]
    public void AnInflatedUsedRangeIsRefusedRatherThanRead()
    {
        Assert.False(DependentScanRegion.TryCreate(
            Bounds(1, 1, AuditAddress.MaximumRow, AuditAddress.MaximumColumn), out var region, out var code, out var message));

        Assert.Null(region);
        Assert.Equal(AuditRefusalCodes.ScanRegionTooLarge, code);
        Assert.Contains("used range", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ARegionWiderThanOneBandIsRefusedRatherThanSplitUnsafely()
    {
        Assert.False(DependentScanRegion.TryCreate(
            Bounds(1, 1, 2, DependentScanRegion.MaximumBlockCells + 1), out _, out var code, out _));

        Assert.Equal(AuditRefusalCodes.ScanRegionTooLarge, code);
    }

    [Fact]
    public void ARegionOutsideTheAddressableGridIsRefused()
    {
        Assert.False(DependentScanRegion.TryCreate(Bounds(AuditAddress.MaximumRow, 1, 5, 1), out _, out var code, out _));

        Assert.Equal(AuditRefusalCodes.ScanRegionUnsupported, code);
    }

    [Fact]
    public void TheRegionCeilingIsEnforcedExactlyAtTheBoundary()
    {
        Assert.True(DependentScanRegion.TryCreate(Bounds(1, 1, 25_000, 10), out _, out _, out _));
        Assert.False(DependentScanRegion.TryCreate(Bounds(1, 1, 25_001, 10), out _, out var code, out _));

        Assert.Equal(AuditRefusalCodes.ScanRegionTooLarge, code);
    }

    private static UsedRegionBounds Bounds(int firstRow, int firstColumn, int rowCount, int columnCount) =>
        new UsedRegionBounds("Model", firstRow, firstColumn, rowCount, columnCount);
}

public sealed class DirectDependentCoordinatorTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void TheScanReadsOnlyTheBandedRegionAndReportsDependents()
    {
        var port = new FakePort(Target("A1"), Bounds(1, 1, 30, 4), ("B1", "=A1"), ("C5", "=SUM(A1:A10)"));

        var result = new DirectDependentCoordinator().Execute(port);

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.Equal(new[] { "Book.xlsx|Model|B1", "Book.xlsx|Model|C5" },
            result.Dependents.Select(value => value.Dependent.ToString()));
        Assert.Equal("worksheet", result.ScanScope);
        Assert.All(port.RequestedBands, band =>
        {
            Assert.InRange(band.FirstRow, 1, 30);
            Assert.InRange(band.LastRow, 1, 30);
            Assert.Equal(1, band.FirstColumn);
            Assert.Equal(4, band.LastColumn);
        });
    }

    [Fact]
    public void ProgressIsReportedMonotonicallyThroughSnapshotAnalyzeAndCompletion()
    {
        var tracker = new OperationProgressTracker();
        var observed = new List<OperationProgress>();
        tracker.Changed += observed.Add;
        var port = new FakePort(Target("A1"), Bounds(1, 1, 2500, 8), ("B1", "=A1"));

        new DirectDependentCoordinator().Execute(port, tracker);

        Assert.Contains(observed, progress => progress.Phase == OperationPhase.Snapshot);
        Assert.Contains(observed, progress => progress.Phase == OperationPhase.Analyze);
        Assert.Equal(OperationPhase.Completed, observed[observed.Count - 1].Phase);
        var snapshots = observed.Where(progress => progress.Phase == OperationPhase.Snapshot).ToArray();
        Assert.Equal(port.RequestedBands.Count, snapshots[snapshots.Length - 1].Completed);
        Assert.Equal(port.RequestedBands.Count, snapshots[snapshots.Length - 1].Total);
    }

    [Fact]
    public void CancellationStopsTheScanAndNeverReportsAPartialScanAsAResult()
    {
        var tracker = new OperationProgressTracker();
        var port = new FakePort(Target("A1"), Bounds(1, 1, 2500, 8), ("B1", "=A1"));
        port.OnBlockRead = index =>
        {
            if (index == 0) tracker.RequestCancellation();
        };

        var result = new DirectDependentCoordinator().Execute(port, tracker);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Empty(result.Dependents);
        Assert.True(port.RequestedBands.Count < port.PlannedBandCount);
        Assert.Equal(OperationPhase.Cancelled, tracker.Current.Phase);
    }

    [Fact]
    public void CancellationRequestedBeforeAnyReadStopsImmediately()
    {
        var tracker = new OperationProgressTracker();
        tracker.RequestCancellation();
        var port = new FakePort(Target("A1"), Bounds(1, 1, 2500, 8), ("B1", "=A1"));

        var result = new DirectDependentCoordinator().Execute(port, tracker);

        Assert.Equal(AuditRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void AnInflatedUsedRangeIsRefusedWithoutReadingAnyBlock()
    {
        var port = new FakePort(Target("A1"), Bounds(1, 1, AuditAddress.MaximumRow, AuditAddress.MaximumColumn));

        var result = new DirectDependentCoordinator().Execute(port, null, _ => true);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.ScanRegionTooLarge, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void AnEmptyWorksheetCompletesWithNoDependentsAndNoReads()
    {
        var port = new FakePort(Target("A1"), Bounds(1, 1, 0, 0));

        var result = new DirectDependentCoordinator().Execute(port);

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.Empty(result.Dependents);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void TheScanScopeFollowsTheTargetWorksheetAndIsNeverWidened()
    {
        var port = new FakePort(Target("A1"), Bounds(1, 1, 10, 2), ("B1", "=A1"));

        new DirectDependentCoordinator().Execute(port);

        Assert.Equal(DependentScanScopeKind.Worksheet, port.ObservedScope!.Kind);
        Assert.Equal(Sheet, port.ObservedScope.WorksheetName);
        Assert.Equal(Workbook, port.ObservedScope.WorkbookId);
    }

    [Fact]
    public void ExecuteRejectsANullPort() =>
        Assert.Throws<ArgumentNullException>(() => new DirectDependentCoordinator().Execute(null!));

    private static AuditCellIdentity Target(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static UsedRegionBounds Bounds(int firstRow, int firstColumn, int rowCount, int columnCount) =>
        new UsedRegionBounds(Sheet, firstRow, firstColumn, rowCount, columnCount);

    private sealed class FakePort : IDependentScanPort
    {
        private readonly AuditCellIdentity _target;
        private readonly UsedRegionBounds _bounds;
        private readonly IReadOnlyList<(string Address, string Formula)> _formulas;

        public FakePort(
            AuditCellIdentity target,
            UsedRegionBounds bounds,
            params (string Address, string Formula)[] formulas)
        {
            _target = target;
            _bounds = bounds;
            _formulas = formulas;
            DependentScanRegion.TryCreate(bounds, out var region, out _, out _);
            PlannedBandCount = region?.BlockCount ?? 0;
        }

        public List<AuditRectangle> RequestedBands { get; } = new List<AuditRectangle>();

        public DependentScanScope? ObservedScope { get; private set; }

        public int PlannedBandCount { get; }

        public Action<int>? OnBlockRead { get; set; }

        public AuditCellIdentity CaptureTarget() => _target;

        public IReadOnlyList<string> CaptureWorksheetNames() => new[] { Sheet };

        public UsedRegionBounds CaptureUsedRegion(string worksheetName) => _bounds;

        public IReadOnlyList<AuditFormulaCell> CaptureBlock(string worksheetName, AuditRectangle band)
        {
            OnBlockRead?.Invoke(RequestedBands.Count);
            RequestedBands.Add(band);
            return _formulas
                .Where(cell => Within(band, cell.Address))
                .Select(cell => new AuditFormulaCell(new AuditCellIdentity(Workbook, Sheet, cell.Address), cell.Formula))
                .ToArray();
        }

        public IReadOnlyList<AuditNameBinding> CaptureNames(DependentScanScope scope)
        {
            ObservedScope = scope;
            return Array.Empty<AuditNameBinding>();
        }

        private static bool Within(AuditRectangle band, string address)
        {
            Assert.True(AuditAddress.TryParse(address, out var cell));
            return band.Intersects(cell);
        }
    }
}
