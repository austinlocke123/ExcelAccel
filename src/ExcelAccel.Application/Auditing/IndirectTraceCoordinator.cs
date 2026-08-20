using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Auditing;

/// <summary>
/// Reads the formula of one arbitrary cell. Used to walk a precedent chain,
/// where each step needs the formula of a cell the user did not select.
/// </summary>
public interface IFormulaLookupPort
{
    /// <summary>Returns null when the cell holds no formula, which ends the chain.</summary>
    string? TryReadFormula(AuditCellIdentity cell);

    ReferenceSnapshotIndex CaptureIndex(DirectPrecedentCapturePlan plan);
}

/// <summary>
/// Expands a node into its direct precedents. A cell with no formula is a leaf,
/// not a gap: it genuinely has no precedents.
/// </summary>
public sealed class PrecedentTraceExpansion : ITraceExpansionPort
{
    private readonly IFormulaLookupPort _port;
    private readonly DirectPrecedentAnalyzer _analyzer = new DirectPrecedentAnalyzer();

    public PrecedentTraceExpansion(IFormulaLookupPort port) =>
        _port = port ?? throw new ArgumentNullException(nameof(port));

    public TraceExpansion Expand(AuditCellIdentity node)
    {
        if (node is null) throw new ArgumentNullException(nameof(node));
        var formula = _port.TryReadFormula(node);
        if (string.IsNullOrEmpty(formula)) return TraceExpansion.Empty;

        var plan = _analyzer.CreateCapturePlan(node, formula!);
        var index = _port.CaptureIndex(plan);
        var result = _analyzer.Analyze(new FormulaReferenceSnapshot(node, formula!, index, FormulaDialect.InvariantA1));
        if (result.Status == AuditTraceStatus.Refused)
        {
            return new TraceExpansion(Array.Empty<TraceEdge>(), isComplete: false);
        }

        return new TraceExpansion(
            result.Precedents.Select(precedent => new TraceEdge(
                precedent.Target,
                precedent.DisplayTarget,
                precedent.Kind,
                precedent.Evidence,
                precedent.IsExternal,
                precedent.IsUnresolved)),
            result.CanClaimCompleteness);
    }
}

/// <summary>
/// Expands a node into its direct dependents using one prebuilt worksheet index,
/// so a traversal never rescans the worksheet per step.
/// </summary>
public sealed class DependentTraceExpansion : ITraceExpansionPort
{
    private readonly ReverseReferenceIndex _index;

    public DependentTraceExpansion(ReverseReferenceIndex index) =>
        _index = index ?? throw new ArgumentNullException(nameof(index));

    public TraceExpansion Expand(AuditCellIdentity node)
    {
        if (node is null) throw new ArgumentNullException(nameof(node));
        var result = _index.FindDirectDependents(node);
        if (result.Status == AuditTraceStatus.Refused)
        {
            return new TraceExpansion(Array.Empty<TraceEdge>(), isComplete: false);
        }

        return new TraceExpansion(
            result.Dependents.Select(dependent => new TraceEdge(
                dependent.Dependent,
                AuditPresentationLabels.Location(dependent.Dependent),
                dependent.Evidence[0].Kind,
                dependent.Evidence)),
            result.CanClaimCompleteness);
    }
}

/// <summary>
/// Drives a bounded indirect traversal, owning the progress and cancellation
/// policy so it stays testable without Excel.
/// </summary>
public sealed class IndirectTraceCoordinator
{
    public IndirectTraceResult Execute(
        AuditCellIdentity root,
        ITraceExpansionPort expansion,
        TraceDirection direction,
        IndirectTraceOptions? options = null,
        OperationProgressTracker? tracker = null)
    {
        if (root is null) throw new ArgumentNullException(nameof(root));
        if (expansion is null) throw new ArgumentNullException(nameof(expansion));
        var limits = options ?? IndirectTraceOptions.Default;
        var progress = tracker ?? new OperationProgressTracker();
        var expanded = 0;

        progress.Report(new OperationProgress(OperationPhase.Analyze, 0, limits.MaximumNodes, "Traversing references."));
        var counted = new CountingExpansion(expansion, () =>
        {
            expanded = Math.Min(expanded + 1, limits.MaximumNodes);
            progress.Report(new OperationProgress(OperationPhase.Analyze, expanded, limits.MaximumNodes, "Traversing references."));
        });

        var result = new IndirectTraceEngine().Traverse(root, counted, direction, limits, () => progress.CancellationRequested);
        progress.Report(result.Status == AuditTraceStatus.Refused
            ? new OperationProgress(OperationPhase.Cancelled, expanded, limits.MaximumNodes, "Traversal cancelled.")
            : new OperationProgress(OperationPhase.Completed, 1, 1, "Traversal complete."));
        return result;
    }

    private sealed class CountingExpansion : ITraceExpansionPort
    {
        private readonly ITraceExpansionPort _inner;
        private readonly Action _onExpanded;

        public CountingExpansion(ITraceExpansionPort inner, Action onExpanded)
        {
            _inner = inner;
            _onExpanded = onExpanded;
        }

        public TraceExpansion Expand(AuditCellIdentity node)
        {
            var expansion = _inner.Expand(node);
            _onExpanded();
            return expansion;
        }
    }
}
