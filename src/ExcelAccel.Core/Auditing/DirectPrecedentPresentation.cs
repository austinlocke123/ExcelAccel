using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

public sealed class DirectPrecedentRow
{
    internal DirectPrecedentRow(
        string nodeId,
        string displayTarget,
        string workbookDisplay,
        string kind,
        string classification,
        string state,
        int edgeCount,
        string sourceEvidence)
    {
        NodeId = nodeId;
        DisplayTarget = displayTarget;
        WorkbookDisplay = workbookDisplay;
        Kind = kind;
        Classification = classification;
        State = state;
        EdgeCount = edgeCount;
        SourceEvidence = sourceEvidence;
    }

    public string NodeId { get; }
    public string DisplayTarget { get; }
    public string WorkbookDisplay { get; }
    public string Kind { get; }
    public string Classification { get; }
    public string State { get; }
    public int EdgeCount { get; }
    public string SourceEvidence { get; }
}

/// <summary>
/// Deterministic read-only projection of a <see cref="DirectPrecedentResult"/>.
/// It formats an existing result and never resolves, evaluates, reorders, or
/// reclassifies anything the analyzer did not already establish.
/// </summary>
public sealed class DirectPrecedentReport
{
    private DirectPrecedentReport(
        AuditTraceStatus status,
        string sourceDisplay,
        string workbookDisplay,
        string statusLabel,
        string coverageLabel,
        string scanScope,
        string headline,
        string completenessStatement,
        bool canClaimCompleteness,
        int precedentCount,
        int unresolvedEdgeCount,
        int externalEdgeCount,
        string? limitationCode,
        string? refusalCode,
        string? message,
        IReadOnlyList<DirectPrecedentRow> rows,
        IReadOnlyList<string> summaryLines)
    {
        Status = status;
        SourceDisplay = sourceDisplay;
        WorkbookDisplay = workbookDisplay;
        StatusLabel = statusLabel;
        CoverageLabel = coverageLabel;
        ScanScope = scanScope;
        Headline = headline;
        CompletenessStatement = completenessStatement;
        CanClaimCompleteness = canClaimCompleteness;
        PrecedentCount = precedentCount;
        UnresolvedEdgeCount = unresolvedEdgeCount;
        ExternalEdgeCount = externalEdgeCount;
        LimitationCode = limitationCode;
        RefusalCode = refusalCode;
        Message = message;
        Rows = rows;
        SummaryLines = summaryLines;
    }

    public AuditTraceStatus Status { get; }

    public string SourceDisplay { get; }

    public string WorkbookDisplay { get; }

    public string StatusLabel { get; }

    public string CoverageLabel { get; }

    public string ScanScope { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public bool CanClaimCompleteness { get; }

    public int PrecedentCount { get; }

    public int UnresolvedEdgeCount { get; }

    public int ExternalEdgeCount { get; }

    public string? LimitationCode { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public IReadOnlyList<DirectPrecedentRow> Rows { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public static DirectPrecedentReport Create(DirectPrecedentResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        var sourceDisplay = AuditPresentationLabels.Location(result.Source);
        var rows = result.Precedents.Select(ToRow).ToArray();
        var statusLabel = AuditPresentationLabels.Status(result.Status);
        var coverageLabel = AuditPresentationLabels.Coverage(result.Coverage);
        var completeness = result.CanClaimCompleteness
            ? "Completeness is claimed for the target formula only; it says nothing about indirect precedents."
            : "Completeness cannot be claimed for this result.";
        var summary = new List<string>
        {
            "Source: " + sourceDisplay,
            "Workbook: " + result.Source.WorkbookId,
            "Status: " + statusLabel,
            "Scan scope: " + result.ScanScope,
            "Parser coverage: " + coverageLabel,
        };
        if (!string.IsNullOrWhiteSpace(result.LimitationCode)) summary.Add("Limitation code: " + result.LimitationCode);
        if (!string.IsNullOrWhiteSpace(result.RefusalCode)) summary.Add("Refusal code: " + result.RefusalCode);
        summary.Add("Direct precedents: " + AuditPresentationLabels.Count(rows.Length));
        summary.Add("Unresolved edges: " + AuditPresentationLabels.Count(result.UnresolvedEdgeCount));
        summary.Add("External edges: " + AuditPresentationLabels.Count(result.ExternalEdgeCount));
        summary.Add(AuditPresentationLabels.CompletenessLine(result.CanClaimCompleteness));
        if (!string.IsNullOrWhiteSpace(result.Message)) summary.Add("Message: " + result.Message);

        return new DirectPrecedentReport(
            result.Status,
            sourceDisplay,
            result.Source.WorkbookId,
            statusLabel,
            coverageLabel,
            result.ScanScope,
            Summarize(result, sourceDisplay, rows.Length),
            completeness,
            result.CanClaimCompleteness,
            rows.Length,
            result.UnresolvedEdgeCount,
            result.ExternalEdgeCount,
            result.LimitationCode,
            result.RefusalCode,
            result.Message,
            Array.AsReadOnly(rows),
            summary.AsReadOnly());
    }

    private static string Summarize(DirectPrecedentResult result, string sourceDisplay, int count)
    {
        if (result.Status == AuditTraceStatus.Refused)
        {
            return $"Direct precedents refused for {sourceDisplay}: {result.Message} (code {result.RefusalCode})";
        }

        var prefix = $"{AuditPresentationLabels.Count(count)} direct precedent{(count == 1 ? string.Empty : "s")} for {sourceDisplay}.";
        return result.Status == AuditTraceStatus.Complete
            ? prefix + " Complete for the target formula."
            : prefix + $" Partial: {AuditPresentationLabels.Count(result.UnresolvedEdgeCount)} unresolved and {AuditPresentationLabels.Count(result.ExternalEdgeCount)} external edges.";
    }

    private static DirectPrecedentRow ToRow(DirectPrecedent precedent) => new DirectPrecedentRow(
        precedent.NodeId,
        precedent.Target is null ? precedent.DisplayTarget : AuditPresentationLabels.Location(precedent.Target),
        precedent.Target?.WorkbookId ?? string.Empty,
        AuditPresentationLabels.Kind(precedent.Kind),
        AuditPresentationLabels.Classification(precedent.Classification),
        StateLabelOf(precedent),
        precedent.Evidence.Count,
        AuditPresentationLabels.EvidenceList(precedent.Evidence));

    private static string StateLabelOf(DirectPrecedent precedent)
    {
        if (precedent.IsClosedExternal) return "External (closed; never opened)";
        if (precedent.IsExternal) return "External";
        return precedent.IsUnresolved ? "Unresolved" : "Resolved";
    }
}
