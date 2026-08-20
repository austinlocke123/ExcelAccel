using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class IndirectTraceEngineTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void TraversalIsBreadthFirstAndRetainsDepthAndOrigin()
    {
        var port = Graph(("A1", new[] { "B1", "C1" }), ("B1", new[] { "D1" }), ("C1", new[] { "E1" }));

        var result = Traverse(port, "A1");

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.Equal(
            new[] { "Model!B1", "Model!C1", "Model!D1", "Model!E1" },
            result.Nodes.Select(node => node.DisplayTarget));
        Assert.Equal(new[] { 1, 1, 2, 2 }, result.Nodes.Select(node => node.Depth));
        Assert.Equal("Model!A1", result.Nodes[0].ViaDisplay);
        Assert.Equal("Model!B1", result.Nodes[2].ViaDisplay);
        Assert.Equal(2, result.DeepestDepthReached);
        Assert.True(result.CanClaimCompleteness);
    }

    [Fact]
    public void EveryIndirectEdgeRetainsTheDirectEvidenceItCameFrom()
    {
        var port = Graph(("A1", new[] { "B1" }), ("B1", new[] { "C1" }));

        var result = Traverse(port, "A1");

        Assert.All(result.Nodes, node => Assert.NotEmpty(node.Evidence));
        Assert.Equal("ref:B1", result.Nodes[0].Evidence[0].SourceText);
        Assert.Equal("ref:C1", result.Nodes[1].Evidence[0].SourceText);
    }

    [Fact]
    public void ACycleTerminatesAndIsRepresentedWithoutBeingExpandedAgain()
    {
        var port = Graph(("A1", new[] { "B1" }), ("B1", new[] { "C1" }), ("C1", new[] { "A1" }));

        var result = Traverse(port, "A1");

        Assert.True(result.ContainsCycle);
        var cycle = Assert.Single(result.Nodes, node => node.IsCycle);
        Assert.Equal("Model!A1", cycle.DisplayTarget);
        Assert.Equal("Model!C1", cycle.ViaDisplay);
        Assert.Equal(3, result.ExpandedNodeCount);
        Assert.Equal(1, port.ExpansionCount("A1"));
    }

    [Fact]
    public void ASelfReferenceIsASingleCycleEdgeRatherThanAnInfiniteLoop()
    {
        var port = Graph(("A1", new[] { "A1" }));

        var result = Traverse(port, "A1");

        var node = Assert.Single(result.Nodes);
        Assert.True(node.IsCycle);
        Assert.Equal(1, port.ExpansionCount("A1"));
    }

    [Fact]
    public void ADiamondVisitsTheSharedNodeOnceAndMarksTheSecondPathAsACycle()
    {
        var port = Graph(("A1", new[] { "B1", "C1" }), ("B1", new[] { "D1" }), ("C1", new[] { "D1" }));

        var result = Traverse(port, "A1");

        Assert.Equal(1, port.ExpansionCount("D1"));
        Assert.Equal(2, result.Nodes.Count(node => node.DisplayTarget == "Model!D1"));
        Assert.Single(result.Nodes, node => node.DisplayTarget == "Model!D1" && node.IsCycle);
    }

    [Fact]
    public void TheDepthCapProducesAnExplicitTruncatedResultWithAFrontier()
    {
        var port = Graph(("A1", new[] { "B1" }), ("B1", new[] { "C1" }), ("C1", new[] { "D1" }));

        var result = Traverse(port, "A1", new IndirectTraceOptions(2, 100));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.True(result.TruncatedByDepth);
        Assert.False(result.CanClaimCompleteness);
        Assert.Equal(AuditRefusalCodes.DepthLimitReached, result.LimitationCode);
        Assert.Equal(new[] { "Model!B1", "Model!C1" }, result.Nodes.Select(node => node.DisplayTarget));
        Assert.True(result.Nodes.Single(node => node.DisplayTarget == "Model!C1").IsUnexpandedFrontier);
        Assert.Equal(0, port.ExpansionCount("C1"));
    }

    [Fact]
    public void TheNodeCapProducesAnExplicitTruncatedResultRatherThanASilentOmission()
    {
        var port = Graph(("A1", new[] { "B1", "C1", "D1", "E1" }));

        var result = Traverse(port, "A1", new IndirectTraceOptions(4, 2));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.True(result.TruncatedByNodeCap);
        Assert.Equal(AuditRefusalCodes.ScanTruncated, result.LimitationCode);
        Assert.Equal(2, result.Nodes.Count);
    }

    [Fact]
    public void AnIncompleteExpansionIsCountedAsACoverageGapAndBlocksCompleteness()
    {
        var port = new FakeExpansion();
        port.Add("A1", new TraceExpansion(new[] { Edge("B1") }, isComplete: false));

        var result = Traverse(port, "A1");

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.Equal(1, result.CoverageGapCount);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void ExternalAndUnresolvedEdgesTerminateWithoutBeingExpanded()
    {
        var port = new FakeExpansion();
        port.Add("A1", new TraceExpansion(
            new[]
            {
                new TraceEdge(null, "[Other.xlsx]Model!A1", AuditReferenceKind.External, Evidence("ext"), isExternal: true),
                new TraceEdge(null, "Rate", AuditReferenceKind.Unresolved, Evidence("Rate"), isUnresolved: true),
            },
            isComplete: true));

        var result = Traverse(port, "A1");

        Assert.Equal(2, result.Nodes.Count);
        Assert.True(result.Nodes[0].IsExternal);
        Assert.True(result.Nodes[1].IsUnresolved);
        Assert.Equal(1, result.ExpandedNodeCount);
    }

    [Fact]
    public void CancellationRefusesRatherThanReportingAPartialTraversalAsATrace()
    {
        var port = Graph(("A1", new[] { "B1" }), ("B1", new[] { "C1" }));
        var calls = 0;

        var result = new IndirectTraceEngine().Traverse(
            Cell("A1"), port, TraceDirection.Precedents, IndirectTraceOptions.Default, () => calls++ > 0);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Empty(result.Nodes);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void RepeatTraversalsOfAnIdenticalGraphProduceIdenticalResults()
    {
        var first = Traverse(Graph(("A1", new[] { "C1", "B1" }), ("B1", new[] { "D1" }), ("C1", new[] { "D1" })), "A1");
        var second = Traverse(Graph(("A1", new[] { "C1", "B1" }), ("B1", new[] { "D1" }), ("C1", new[] { "D1" })), "A1");

        Assert.Equal(
            first.Nodes.Select(node => node.DisplayTarget + "@" + node.Depth + "/" + node.ViaDisplay + "/" + node.IsCycle),
            second.Nodes.Select(node => node.DisplayTarget + "@" + node.Depth + "/" + node.ViaDisplay + "/" + node.IsCycle));
    }

    [Fact]
    public void ADeepChainTerminatesAtTheHardCeilingWithoutRunningAway()
    {
        var port = new FakeExpansion();
        for (var index = 1; index <= 200; index++)
        {
            port.Add("A" + index, new TraceExpansion(new[] { Edge("A" + (index + 1)) }, true));
        }

        var result = Traverse(port, "A1", new IndirectTraceOptions(IndirectTraceOptions.DepthCeiling, 100));

        Assert.True(result.TruncatedByDepth);
        Assert.Equal(IndirectTraceOptions.DepthCeiling, result.DeepestDepthReached);
        Assert.Equal(IndirectTraceOptions.DepthCeiling, result.Nodes.Count);
    }

    [Fact]
    public void OptionsRejectLimitsOutsideTheHardCeilings()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new IndirectTraceOptions(0, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IndirectTraceOptions(IndirectTraceOptions.DepthCeiling + 1, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IndirectTraceOptions(4, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new IndirectTraceOptions(4, IndirectTraceOptions.NodeCeiling + 1));
    }

    [Fact]
    public void TheEngineRejectsNullArguments()
    {
        var engine = new IndirectTraceEngine();
        Assert.Throws<ArgumentNullException>(() => engine.Traverse(null!, new FakeExpansion(), TraceDirection.Precedents));
        Assert.Throws<ArgumentNullException>(() => engine.Traverse(Cell("A1"), null!, TraceDirection.Precedents));
    }

    [Fact]
    public void TheDirectionIsReportedAsTheScanScope()
    {
        Assert.Equal("indirect_precedents", Traverse(Graph(("A1", new string[0])), "A1").ScanScope);
        Assert.Equal(
            "indirect_dependents",
            new IndirectTraceEngine().Traverse(Cell("A1"), Graph(("A1", new string[0])), TraceDirection.Dependents).ScanScope);
    }

    private static AuditCellIdentity Cell(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static IndirectTraceResult Traverse(FakeExpansion port, string root, IndirectTraceOptions? options = null) =>
        new IndirectTraceEngine().Traverse(Cell(root), port, TraceDirection.Precedents, options ?? IndirectTraceOptions.Default);

    private static IEnumerable<AuditReferenceEvidence> Evidence(string text) =>
        new[] { new AuditReferenceEvidence(text, new FormulaSourceSpan(0, text.Length), AuditReferenceKind.Cell) };

    private static TraceEdge Edge(string address) =>
        new TraceEdge(Cell(address), address, AuditReferenceKind.Cell, Evidence("ref:" + address));

    private static FakeExpansion Graph(params (string Node, string[] Targets)[] edges)
    {
        var port = new FakeExpansion();
        foreach (var (node, targets) in edges)
        {
            port.Add(node, new TraceExpansion(targets.Select(Edge), true));
        }

        return port;
    }

    private sealed class FakeExpansion : ITraceExpansionPort
    {
        private readonly Dictionary<string, TraceExpansion> _edges = new Dictionary<string, TraceExpansion>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _expansions = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public void Add(string address, TraceExpansion expansion) => _edges[address] = expansion;

        public int ExpansionCount(string address) => _expansions.TryGetValue(address, out var count) ? count : 0;

        public TraceExpansion Expand(AuditCellIdentity node)
        {
            _expansions[node.Address] = ExpansionCount(node.Address) + 1;
            return _edges.TryGetValue(node.Address, out var expansion) ? expansion : TraceExpansion.Empty;
        }
    }
}
