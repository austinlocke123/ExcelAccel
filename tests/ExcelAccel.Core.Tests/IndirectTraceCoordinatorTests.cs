using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class IndirectTraceCoordinatorTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void APrecedentChainIsFollowedThroughFormulaLookups()
    {
        var port = new FakeFormulaLookup(("D1", "=C1"), ("C1", "=B1"), ("B1", "=A1"));

        var result = Traverse(new PrecedentTraceExpansion(port), "D1", TraceDirection.Precedents);

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.Equal(new[] { "Model!C1", "Model!B1", "Model!A1" }, result.Nodes.Select(node => node.DisplayTarget));
        Assert.Equal(new[] { 1, 2, 3 }, result.Nodes.Select(node => node.Depth));
    }

    [Fact]
    public void ACellWithoutAFormulaEndsTheChainWithoutBeingACoverageGap()
    {
        var port = new FakeFormulaLookup(("B1", "=A1"));

        var result = Traverse(new PrecedentTraceExpansion(port), "B1", TraceDirection.Precedents);

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.Equal(0, result.CoverageGapCount);
        Assert.Single(result.Nodes);
    }

    [Fact]
    public void ACircularPrecedentChainTerminates()
    {
        var port = new FakeFormulaLookup(("A1", "=B1"), ("B1", "=A1"));

        var result = Traverse(new PrecedentTraceExpansion(port), "A1", TraceDirection.Precedents);

        Assert.True(result.ContainsCycle);
        Assert.Contains(result.Nodes, node => node.IsCycle);
    }

    [Fact]
    public void ADependentChainIsFollowedThroughOneWorksheetIndex()
    {
        var scope = DependentScanScope.Worksheet(Workbook, Sheet);
        var index = ReverseReferenceIndex.Build(scope, new[]
        {
            Formula("B1", "=A1"),
            Formula("C1", "=B1"),
            Formula("D1", "=C1"),
        });

        var result = Traverse(new DependentTraceExpansion(index), "A1", TraceDirection.Dependents);

        Assert.Equal(new[] { "Model!B1", "Model!C1", "Model!D1" }, result.Nodes.Select(node => node.DisplayTarget));
        Assert.Equal("indirect_dependents", result.ScanScope);
    }

    [Fact]
    public void ProgressAdvancesAndCompletes()
    {
        var observed = new List<OperationProgress>();
        var tracker = new OperationProgressTracker();
        tracker.Changed += observed.Add;
        var port = new FakeFormulaLookup(("C1", "=B1"), ("B1", "=A1"));

        new IndirectTraceCoordinator().Execute(
            Cell("C1"), new PrecedentTraceExpansion(port), TraceDirection.Precedents, IndirectTraceOptions.Default, tracker);

        Assert.Contains(observed, progress => progress.Phase == OperationPhase.Analyze);
        Assert.Equal(OperationPhase.Completed, observed[observed.Count - 1].Phase);
    }

    [Fact]
    public void CancellationRefusesAndReportsTheCancelledPhase()
    {
        var tracker = new OperationProgressTracker();
        tracker.RequestCancellation();
        var port = new FakeFormulaLookup(("C1", "=B1"), ("B1", "=A1"));

        var result = new IndirectTraceCoordinator().Execute(
            Cell("C1"), new PrecedentTraceExpansion(port), TraceDirection.Precedents, IndirectTraceOptions.Default, tracker);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Equal(OperationPhase.Cancelled, tracker.Current.Phase);
    }

    [Fact]
    public void TheReportProjectsIntoTheSharedViewShapeAndMarksNavigableRows()
    {
        var port = new FakeFormulaLookup(("B1", "=A1+'[Other.xlsx]Data'!C3"));

        var presentation = IndirectTraceReport.Create(
            Traverse(new PrecedentTraceExpansion(port), "B1", TraceDirection.Precedents)).ToPresentation();

        Assert.Equal("ExcelAccel Indirect Precedents", presentation.Title);
        Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Values.Count));
        Assert.Contains(presentation.Rows, row => row.IsNavigable);
        Assert.Contains(presentation.Rows, row => !row.IsNavigable);
    }

    [Fact]
    public void ACycleEdgeIsNeverNavigable()
    {
        var port = new FakeFormulaLookup(("A1", "=B1"), ("B1", "=A1"));

        var presentation = IndirectTraceReport.Create(
            Traverse(new PrecedentTraceExpansion(port), "A1", TraceDirection.Precedents)).ToPresentation();

        var cycleRow = Assert.Single(presentation.Rows, row => row.Values.Contains("Cycle (already visited)"));
        Assert.False(cycleRow.IsNavigable);
    }

    [Fact]
    public void TheReportNamesEveryCapAndCycleInItsSummary()
    {
        var port = new FakeFormulaLookup(("D1", "=C1"), ("C1", "=B1"), ("B1", "=A1"));

        var report = IndirectTraceReport.Create(
            Traverse(new PrecedentTraceExpansion(port), "D1", TraceDirection.Precedents, new IndirectTraceOptions(2, 50)));

        Assert.Contains("Truncated by depth: yes", report.SummaryLines);
        Assert.Contains("Truncated by node cap: no", report.SummaryLines);
        Assert.Contains("Contains cycle: no", report.SummaryLines);
        Assert.Contains("Completeness: not claimed", report.SummaryLines);
        Assert.Contains("depth cap was reached", report.Headline);
    }

    [Fact]
    public void BothIndirectCommandsAreRegisteredAsReadOnlyWithTraversalAcceptance()
    {
        foreach (var id in new[] { AuditingCommandCatalog.IndirectPrecedentsId, AuditingCommandCatalog.IndirectDependentsId })
        {
            var descriptor = BuiltInCommandRegistry.GetRequired(id);
            Assert.Equal(Core.Commands.CommandImpact.ReadOnly, descriptor.Impact);
            Assert.Empty(descriptor.ChangedProperties);
            Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
            Assert.Contains("AC-AUD-010", descriptor.AcceptanceIds);
            Assert.Contains("AC-AUD-014", descriptor.AcceptanceIds);
        }
    }

    [Fact]
    public void TheCoordinatorRejectsNullArguments()
    {
        var coordinator = new IndirectTraceCoordinator();
        Assert.Throws<ArgumentNullException>(() => coordinator.Execute(null!, new PrecedentTraceExpansion(new FakeFormulaLookup()), TraceDirection.Precedents));
        Assert.Throws<ArgumentNullException>(() => coordinator.Execute(Cell("A1"), null!, TraceDirection.Precedents));
        Assert.Throws<ArgumentNullException>(() => new PrecedentTraceExpansion(null!));
        Assert.Throws<ArgumentNullException>(() => new DependentTraceExpansion(null!));
    }

    private static AuditCellIdentity Cell(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static AuditFormulaCell Formula(string address, string formula) =>
        new AuditFormulaCell(Cell(address), formula);

    private static IndirectTraceResult Traverse(
        ITraceExpansionPort expansion,
        string root,
        TraceDirection direction,
        IndirectTraceOptions? options = null) =>
        new IndirectTraceCoordinator().Execute(Cell(root), expansion, direction, options ?? IndirectTraceOptions.Default);

    private sealed class FakeFormulaLookup : IFormulaLookupPort
    {
        private readonly Dictionary<string, string> _formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public FakeFormulaLookup(params (string Address, string Formula)[] formulas)
        {
            foreach (var (address, formula) in formulas) _formulas[address] = formula;
        }

        public string? TryReadFormula(AuditCellIdentity cell) =>
            _formulas.TryGetValue(cell.Address, out var formula) ? formula : null;

        public ReferenceSnapshotIndex CaptureIndex(DirectPrecedentCapturePlan plan) =>
            new ReferenceSnapshotIndex(plan.LocalTargets.Select(target =>
                new KeyValuePair<AuditCellIdentity, AuditCellClassification>(
                    target,
                    _formulas.ContainsKey(target.Address) ? AuditCellClassification.Formula : AuditCellClassification.Value)));
    }
}
