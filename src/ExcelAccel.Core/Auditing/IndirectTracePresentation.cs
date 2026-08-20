using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// Deterministic read-only projection of an <see cref="IndirectTraceResult"/>.
/// It formats an existing traversal and never expands, resolves, or reorders
/// anything the engine did not already establish.
/// </summary>
public sealed class IndirectTraceReport
{
    private IndirectTraceReport(
        IndirectTraceResult result,
        string rootDisplay,
        string headline,
        string completenessStatement,
        IReadOnlyList<string> summaryLines)
    {
        Result = result;
        RootDisplay = rootDisplay;
        Headline = headline;
        CompletenessStatement = completenessStatement;
        SummaryLines = summaryLines;
    }

    public IndirectTraceResult Result { get; }

    public string RootDisplay { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public static IndirectTraceReport Create(IndirectTraceResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        var rootDisplay = AuditPresentationLabels.Location(result.Root);
        var directionLabel = result.Direction == TraceDirection.Precedents ? "precedents" : "dependents";
        var completeness = result.CanClaimCompleteness
            ? $"Completeness is claimed for indirect {directionLabel} within the requested depth and node caps only."
            : "Completeness cannot be claimed for this traversal.";

        var summary = new List<string>
        {
            "Root: " + rootDisplay,
            "Workbook: " + result.Root.WorkbookId,
            "Direction: indirect " + directionLabel,
            "Status: " + AuditPresentationLabels.Status(result.Status),
            "Scan scope: " + result.ScanScope,
            "Requested depth cap: " + AuditPresentationLabels.Count(result.Options.MaximumDepth),
            "Requested node cap: " + AuditPresentationLabels.Count(result.Options.MaximumNodes),
        };
        if (!string.IsNullOrWhiteSpace(result.LimitationCode)) summary.Add("Limitation code: " + result.LimitationCode);
        if (!string.IsNullOrWhiteSpace(result.RefusalCode)) summary.Add("Refusal code: " + result.RefusalCode);
        summary.Add("Nodes reported: " + AuditPresentationLabels.Count(result.Nodes.Count));
        summary.Add("Nodes expanded: " + AuditPresentationLabels.Count(result.ExpandedNodeCount));
        summary.Add("Deepest depth reached: " + AuditPresentationLabels.Count(result.DeepestDepthReached));
        summary.Add("Coverage gaps: " + AuditPresentationLabels.Count(result.CoverageGapCount));
        summary.Add("Truncated by depth: " + (result.TruncatedByDepth ? "yes" : "no"));
        summary.Add("Truncated by node cap: " + (result.TruncatedByNodeCap ? "yes" : "no"));
        summary.Add("Contains cycle: " + (result.ContainsCycle ? "yes" : "no"));
        summary.Add(AuditPresentationLabels.CompletenessLine(result.CanClaimCompleteness));
        if (!string.IsNullOrWhiteSpace(result.Message)) summary.Add("Message: " + result.Message);

        return new IndirectTraceReport(result, rootDisplay, Summarize(result, rootDisplay, directionLabel), completeness, summary.AsReadOnly());
    }

    public TraceResultPresentation ToPresentation() => new TraceResultPresentation(
        Result.Direction == TraceDirection.Precedents
            ? "ExcelAccel Indirect Precedents"
            : "ExcelAccel Indirect Dependents",
        Result.Status,
        Headline,
        CompletenessStatement,
        new[]
        {
            new TraceColumn("Depth", 60),
            new TraceColumn("Node", 180),
            new TraceColumn("Reached from", 160),
            new TraceColumn("Kind", 90),
            new TraceColumn("State", 150),
            new TraceColumn("Source reference", 240),
        },
        Result.Nodes.Select(node => new TraceRow(
            new[]
            {
                node.Depth.ToString(CultureInfo.InvariantCulture),
                node.DisplayTarget,
                node.ViaDisplay,
                AuditPresentationLabels.Kind(node.Kind),
                State(node),
                AuditPresentationLabels.EvidenceList(node.Evidence),
            },
            node.IsCycle || node.IsExternal || node.IsUnresolved ? null : node.Target)),
        SummaryLines,
        Result.RefusalCode);

    private static string Summarize(IndirectTraceResult result, string rootDisplay, string directionLabel)
    {
        if (result.Status == AuditTraceStatus.Refused)
        {
            return $"Indirect {directionLabel} refused for {rootDisplay}: {result.Message} (code {result.RefusalCode})";
        }

        var prefix = $"{AuditPresentationLabels.Count(result.Nodes.Count)} indirect {directionLabel} node" +
            (result.Nodes.Count == 1 ? string.Empty : "s") +
            $" from {rootDisplay} to depth {AuditPresentationLabels.Count(result.DeepestDepthReached)}.";
        if (result.Status == AuditTraceStatus.Complete) return prefix + " Complete within the requested caps.";
        var reasons = new List<string>();
        if (result.TruncatedByDepth) reasons.Add("the depth cap was reached");
        if (result.TruncatedByNodeCap) reasons.Add("the node cap was reached");
        if (result.CoverageGapCount > 0) reasons.Add($"{AuditPresentationLabels.Count(result.CoverageGapCount)} coverage gaps");
        return prefix + " Partial: " + string.Join(", ", reasons) + ".";
    }

    private static string State(IndirectTraceNode node)
    {
        if (node.IsCycle) return "Cycle (already visited)";
        if (node.IsExternal) return "External (closed; never opened)";
        if (node.IsUnresolved) return "Unresolved";
        return node.IsUnexpandedFrontier ? "Not expanded (cap reached)" : "Expanded";
    }
}
