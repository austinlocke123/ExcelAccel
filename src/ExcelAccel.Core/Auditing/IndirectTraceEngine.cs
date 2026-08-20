using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Auditing;

public enum TraceDirection
{
    Precedents,
    Dependents,
}

/// <summary>One resolved or terminal edge out of a node.</summary>
public sealed class TraceEdge
{
    public TraceEdge(
        AuditCellIdentity? target,
        string displayTarget,
        AuditReferenceKind kind,
        IEnumerable<AuditReferenceEvidence> evidence,
        bool isExternal = false,
        bool isUnresolved = false)
    {
        Target = target;
        DisplayTarget = !string.IsNullOrWhiteSpace(displayTarget)
            ? displayTarget
            : throw new ArgumentException("A display target is required.", nameof(displayTarget));
        Kind = kind;
        Evidence = Array.AsReadOnly((evidence ?? throw new ArgumentNullException(nameof(evidence))).ToArray());
        IsExternal = isExternal;
        IsUnresolved = isUnresolved;
    }

    public AuditCellIdentity? Target { get; }

    public string DisplayTarget { get; }

    public AuditReferenceKind Kind { get; }

    public IReadOnlyList<AuditReferenceEvidence> Evidence { get; }

    public bool IsExternal { get; }

    public bool IsUnresolved { get; }
}

/// <summary>
/// The edges out of one node, plus whether that expansion was fully covered.
/// An incomplete expansion is a coverage gap and blocks a completeness claim.
/// </summary>
public sealed class TraceExpansion
{
    public TraceExpansion(IEnumerable<TraceEdge> edges, bool isComplete)
    {
        Edges = Array.AsReadOnly((edges ?? throw new ArgumentNullException(nameof(edges))).ToArray());
        IsComplete = isComplete;
    }

    public IReadOnlyList<TraceEdge> Edges { get; }

    public bool IsComplete { get; }

    public static TraceExpansion Empty { get; } = new TraceExpansion(Array.Empty<TraceEdge>(), true);
}

/// <summary>Expands one node in the traversal direction. Read-only.</summary>
public interface ITraceExpansionPort
{
    TraceExpansion Expand(AuditCellIdentity node);
}

/// <summary>Explicit, bounded traversal limits. Both are hard-capped.</summary>
public sealed class IndirectTraceOptions
{
    /// <summary>Hard ceiling on requested depth. No caller may exceed it.</summary>
    public const int DepthCeiling = 16;

    /// <summary>Hard ceiling on reported nodes. No caller may exceed it.</summary>
    public const int NodeCeiling = 5_000;

    public IndirectTraceOptions(int maximumDepth, int maximumNodes)
    {
        MaximumDepth = maximumDepth >= 1 && maximumDepth <= DepthCeiling
            ? maximumDepth
            : throw new ArgumentOutOfRangeException(nameof(maximumDepth), $"Depth must be between 1 and {DepthCeiling}.");
        MaximumNodes = maximumNodes >= 1 && maximumNodes <= NodeCeiling
            ? maximumNodes
            : throw new ArgumentOutOfRangeException(nameof(maximumNodes), $"The node cap must be between 1 and {NodeCeiling}.");
    }

    public int MaximumDepth { get; }

    public int MaximumNodes { get; }

    public static IndirectTraceOptions Default { get; } = new IndirectTraceOptions(8, 1_000);
}

public sealed class IndirectTraceNode
{
    internal IndirectTraceNode(
        AuditCellIdentity? target,
        string displayTarget,
        int depth,
        string viaDisplay,
        AuditReferenceKind kind,
        IReadOnlyList<AuditReferenceEvidence> evidence,
        bool isCycle,
        bool isExternal,
        bool isUnresolved,
        bool isUnexpandedFrontier)
    {
        Target = target;
        DisplayTarget = displayTarget;
        Depth = depth;
        ViaDisplay = viaDisplay;
        Kind = kind;
        Evidence = evidence;
        IsCycle = isCycle;
        IsExternal = isExternal;
        IsUnresolved = isUnresolved;
        IsUnexpandedFrontier = isUnexpandedFrontier;
    }

    public AuditCellIdentity? Target { get; }

    public string DisplayTarget { get; }

    /// <summary>Edges out of the root are depth 1.</summary>
    public int Depth { get; }

    /// <summary>The node this one was reached from.</summary>
    public string ViaDisplay { get; }

    public AuditReferenceKind Kind { get; }

    /// <summary>
    /// The direct evidence of the edge that reached this node, retained at every
    /// depth so an indirect result never loses the reference it came from.
    /// </summary>
    public IReadOnlyList<AuditReferenceEvidence> Evidence { get; }

    /// <summary>This edge closes back onto an already-visited node.</summary>
    public bool IsCycle { get; }

    public bool IsExternal { get; }

    public bool IsUnresolved { get; }

    /// <summary>Reached, but not expanded because a cap stopped the traversal.</summary>
    public bool IsUnexpandedFrontier { get; }
}

public sealed class IndirectTraceResult
{
    internal IndirectTraceResult(
        AuditTraceStatus status,
        AuditCellIdentity root,
        TraceDirection direction,
        IndirectTraceOptions options,
        IEnumerable<IndirectTraceNode> nodes,
        int expandedNodeCount,
        int deepestDepthReached,
        int coverageGapCount,
        bool truncatedByDepth,
        bool truncatedByNodeCap,
        bool containsCycle,
        string? limitationCode,
        string? refusalCode,
        string? message)
    {
        Status = status;
        Root = root ?? throw new ArgumentNullException(nameof(root));
        Direction = direction;
        Options = options;
        Nodes = Array.AsReadOnly((nodes ?? throw new ArgumentNullException(nameof(nodes))).ToArray());
        ExpandedNodeCount = expandedNodeCount;
        DeepestDepthReached = deepestDepthReached;
        CoverageGapCount = coverageGapCount;
        TruncatedByDepth = truncatedByDepth;
        TruncatedByNodeCap = truncatedByNodeCap;
        ContainsCycle = containsCycle;
        LimitationCode = limitationCode;
        RefusalCode = refusalCode;
        Message = message;
    }

    public AuditTraceStatus Status { get; }

    public AuditCellIdentity Root { get; }

    public TraceDirection Direction { get; }

    public IndirectTraceOptions Options { get; }

    public IReadOnlyList<IndirectTraceNode> Nodes { get; }

    public int ExpandedNodeCount { get; }

    public int DeepestDepthReached { get; }

    public int CoverageGapCount { get; }

    public bool TruncatedByDepth { get; }

    public bool TruncatedByNodeCap { get; }

    public bool ContainsCycle { get; }

    public string ScanScope => Direction == TraceDirection.Precedents ? "indirect_precedents" : "indirect_dependents";

    public string? LimitationCode { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public bool CanClaimCompleteness => Status == AuditTraceStatus.Complete;

    public static IndirectTraceResult Refused(
        AuditCellIdentity root,
        TraceDirection direction,
        IndirectTraceOptions options,
        string code,
        string message) =>
        new IndirectTraceResult(
            AuditTraceStatus.Refused,
            root,
            direction,
            options ?? throw new ArgumentNullException(nameof(options)),
            Array.Empty<IndirectTraceNode>(),
            0,
            0,
            0,
            false,
            false,
            false,
            null,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A refusal code is required.", nameof(code)),
            message ?? string.Empty);
}

/// <summary>
/// Deterministic breadth-first traversal over an expansion port, bounded by an
/// explicit depth and node cap.
///
/// Traversal is breadth-first and preserves the expansion's own ordering, so an
/// identical snapshot always produces an identical result. A node is expanded at
/// most once: revisiting one is represented as a cycle edge rather than expanded
/// again, so a circular model terminates instead of looping. Hitting either cap
/// produces an explicit truncated result, never a silent omission.
/// </summary>
public sealed class IndirectTraceEngine
{
    public IndirectTraceResult Traverse(
        AuditCellIdentity root,
        ITraceExpansionPort port,
        TraceDirection direction,
        IndirectTraceOptions? options = null,
        Func<bool>? cancellationRequested = null)
    {
        if (root is null) throw new ArgumentNullException(nameof(root));
        if (port is null) throw new ArgumentNullException(nameof(port));
        var limits = options ?? IndirectTraceOptions.Default;

        var nodes = new List<IndirectTraceNode>();
        var visited = new HashSet<AuditCellIdentity> { root };
        var queue = new Queue<QueueEntry>();
        queue.Enqueue(new QueueEntry(root, AuditPresentationLabels.Location(root), 0));

        var expanded = 0;
        var gaps = 0;
        var deepest = 0;
        var truncatedByDepth = false;
        var truncatedByNodeCap = false;
        var containsCycle = false;

        while (queue.Count != 0)
        {
            if (cancellationRequested?.Invoke() == true)
            {
                return IndirectTraceResult.Refused(
                    root,
                    direction,
                    limits,
                    AuditRefusalCodes.ScanCancelled,
                    "The traversal was cancelled; no partial traversal is reported as a trace.");
            }

            var current = queue.Dequeue();
            if (current.Depth >= limits.MaximumDepth)
            {
                truncatedByDepth = true;
                continue;
            }

            var expansion = port.Expand(current.Node);
            expanded++;
            if (!expansion.IsComplete) gaps++;

            foreach (var edge in expansion.Edges)
            {
                if (nodes.Count >= limits.MaximumNodes)
                {
                    truncatedByNodeCap = true;
                    break;
                }

                var depth = current.Depth + 1;
                deepest = Math.Max(deepest, depth);
                var isCycle = edge.Target is not null && !visited.Add(edge.Target);
                if (isCycle) containsCycle = true;
                var expandable = edge.Target is not null && !isCycle && !edge.IsExternal && !edge.IsUnresolved;
                var frontier = expandable && depth >= limits.MaximumDepth;
                if (frontier) truncatedByDepth = true;

                nodes.Add(new IndirectTraceNode(
                    edge.Target,
                    edge.Target is null ? edge.DisplayTarget : AuditPresentationLabels.Location(edge.Target),
                    depth,
                    current.Display,
                    edge.Kind,
                    edge.Evidence,
                    isCycle,
                    edge.IsExternal,
                    edge.IsUnresolved,
                    frontier));

                if (expandable && !frontier)
                {
                    queue.Enqueue(new QueueEntry(edge.Target!, AuditPresentationLabels.Location(edge.Target!), depth));
                }
            }

            if (truncatedByNodeCap) break;
        }

        var truncated = truncatedByDepth || truncatedByNodeCap;
        var incomplete = truncated || gaps > 0;
        var limitationCode = truncatedByNodeCap
            ? AuditRefusalCodes.ScanTruncated
            : truncatedByDepth
                ? AuditRefusalCodes.DepthLimitReached
                : null;

        return new IndirectTraceResult(
            incomplete ? AuditTraceStatus.Partial : AuditTraceStatus.Complete,
            root,
            direction,
            limits,
            nodes,
            expanded,
            deepest,
            gaps,
            truncatedByDepth,
            truncatedByNodeCap,
            containsCycle,
            limitationCode,
            null,
            incomplete
                ? "The traversal is partial; inspect the reported caps and coverage gaps."
                : null);
    }

    private readonly struct QueueEntry
    {
        public QueueEntry(AuditCellIdentity node, string display, int depth)
        {
            Node = node;
            Display = display;
            Depth = depth;
        }

        public AuditCellIdentity Node { get; }

        public string Display { get; }

        public int Depth { get; }
    }
}
