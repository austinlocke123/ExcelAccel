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
        var sourceDisplay = Location(result.Source);
        var rows = result.Precedents.Select(ToRow).ToArray();
        var statusLabel = StatusLabelOf(result.Status);
        var coverageLabel = CoverageLabelOf(result.Coverage);
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
        summary.Add("Direct precedents: " + Count(rows.Length));
        summary.Add("Unresolved edges: " + Count(result.UnresolvedEdgeCount));
        summary.Add("External edges: " + Count(result.ExternalEdgeCount));
        summary.Add("Completeness: " + (result.CanClaimCompleteness ? "claimed" : "not claimed"));
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

        var prefix = $"{Count(count)} direct precedent{(count == 1 ? string.Empty : "s")} for {sourceDisplay}.";
        return result.Status == AuditTraceStatus.Complete
            ? prefix + " Complete for the target formula."
            : prefix + $" Partial: {Count(result.UnresolvedEdgeCount)} unresolved and {Count(result.ExternalEdgeCount)} external edges.";
    }

    private static DirectPrecedentRow ToRow(DirectPrecedent precedent) => new DirectPrecedentRow(
        precedent.NodeId,
        precedent.Target is null ? precedent.DisplayTarget : Location(precedent.Target),
        precedent.Target?.WorkbookId ?? string.Empty,
        KindLabelOf(precedent.Kind),
        ClassificationLabelOf(precedent.Classification),
        StateLabelOf(precedent),
        precedent.Evidence.Count,
        string.Join("; ", precedent.Evidence.Select(Evidence)));

    private static string Evidence(AuditReferenceEvidence evidence) =>
        evidence.SourceText + " [" +
        evidence.SourceSpan.Start.ToString(CultureInfo.InvariantCulture) + "+" +
        evidence.SourceSpan.Length.ToString(CultureInfo.InvariantCulture) + "]";

    private static string Location(AuditCellIdentity identity) => identity.WorksheetName + "!" + identity.Address;

    private static string Count(int value) => value.ToString("N0", CultureInfo.InvariantCulture);

    private static string StatusLabelOf(AuditTraceStatus status) => status switch
    {
        AuditTraceStatus.Complete => "Complete",
        AuditTraceStatus.Partial => "Partial",
        AuditTraceStatus.Refused => "Refused",
        _ => "Unknown",
    };

    private static string CoverageLabelOf(FormulaCoverageDisposition coverage) => coverage switch
    {
        FormulaCoverageDisposition.Transform => "Fully parsed",
        FormulaCoverageDisposition.RoundTrip => "Fully parsed (round-trip only)",
        FormulaCoverageDisposition.InspectOnly => "Parser coverage gap (inspect only)",
        FormulaCoverageDisposition.Refuse => "Refused",
        _ => "Unknown",
    };

    private static string KindLabelOf(AuditReferenceKind kind) => kind switch
    {
        AuditReferenceKind.Cell => "Cell",
        AuditReferenceKind.Range => "Range",
        AuditReferenceKind.Name => "Name",
        AuditReferenceKind.External => "External",
        AuditReferenceKind.Unresolved => "Unresolved",
        _ => "Unknown",
    };

    private static string ClassificationLabelOf(AuditCellClassification classification) => classification switch
    {
        AuditCellClassification.Formula => "Formula",
        AuditCellClassification.Value => "Value",
        AuditCellClassification.Error => "Error",
        AuditCellClassification.Blank => "Blank",
        AuditCellClassification.Mixed => "Mixed",
        _ => "Not captured",
    };

    private static string StateLabelOf(DirectPrecedent precedent)
    {
        if (precedent.IsClosedExternal) return "External (closed; never opened)";
        if (precedent.IsExternal) return "External";
        return precedent.IsUnresolved ? "Unresolved" : "Resolved";
    }
}
