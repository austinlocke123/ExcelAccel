using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Auditing;

public sealed class DirectDependentRow
{
    internal DirectDependentRow(string displayTarget, string workbookDisplay, string kinds, int edgeCount, string sourceEvidence)
    {
        DisplayTarget = displayTarget;
        WorkbookDisplay = workbookDisplay;
        Kinds = kinds;
        EdgeCount = edgeCount;
        SourceEvidence = sourceEvidence;
    }

    public string DisplayTarget { get; }

    public string WorkbookDisplay { get; }

    /// <summary>How the dependent reaches the target, for example "Cell" or "Cell, Name".</summary>
    public string Kinds { get; }

    public int EdgeCount { get; }

    public string SourceEvidence { get; }
}

/// <summary>
/// Deterministic read-only projection of a <see cref="DirectDependentResult"/>.
/// It formats an existing result and never scans, resolves, evaluates, reorders,
/// or reclassifies anything the index did not already establish. Wording comes
/// from <see cref="AuditPresentationLabels"/> so it cannot drift from the
/// precedent presentation.
/// </summary>
public sealed class DirectDependentReport
{
    private DirectDependentReport(
        AuditTraceStatus status,
        string targetDisplay,
        string workbookDisplay,
        string statusLabel,
        string scanScope,
        string headline,
        string completenessStatement,
        bool canClaimCompleteness,
        int dependentCount,
        int scannedFormulaCount,
        int coverageGapCount,
        bool truncated,
        string? limitationCode,
        string? refusalCode,
        string? message,
        IReadOnlyList<DirectDependentRow> rows,
        IReadOnlyList<string> summaryLines)
    {
        Status = status;
        TargetDisplay = targetDisplay;
        WorkbookDisplay = workbookDisplay;
        StatusLabel = statusLabel;
        ScanScope = scanScope;
        Headline = headline;
        CompletenessStatement = completenessStatement;
        CanClaimCompleteness = canClaimCompleteness;
        DependentCount = dependentCount;
        ScannedFormulaCount = scannedFormulaCount;
        CoverageGapCount = coverageGapCount;
        Truncated = truncated;
        LimitationCode = limitationCode;
        RefusalCode = refusalCode;
        Message = message;
        Rows = rows;
        SummaryLines = summaryLines;
    }

    public AuditTraceStatus Status { get; }

    public string TargetDisplay { get; }

    public string WorkbookDisplay { get; }

    public string StatusLabel { get; }

    public string ScanScope { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public bool CanClaimCompleteness { get; }

    public int DependentCount { get; }

    public int ScannedFormulaCount { get; }

    public int CoverageGapCount { get; }

    public bool Truncated { get; }

    public string? LimitationCode { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public IReadOnlyList<DirectDependentRow> Rows { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public static DirectDependentReport Create(DirectDependentResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        var targetDisplay = AuditPresentationLabels.Location(result.Target);
        var rows = result.Dependents.Select(ToRow).ToArray();
        var completeness = result.CanClaimCompleteness
            ? "Completeness is claimed for the declared " + result.ScanScope + " scope only; it says nothing about other worksheets or indirect dependents."
            : "Completeness cannot be claimed for this result.";
        var summary = new List<string>
        {
            "Target: " + targetDisplay,
            "Workbook: " + result.Target.WorkbookId,
            "Status: " + AuditPresentationLabels.Status(result.Status),
            "Scan scope: " + result.ScanScope,
        };
        if (!string.IsNullOrWhiteSpace(result.Scope.WorksheetName)) summary.Add("Worksheet: " + result.Scope.WorksheetName);
        if (!string.IsNullOrWhiteSpace(result.LimitationCode)) summary.Add("Limitation code: " + result.LimitationCode);
        if (!string.IsNullOrWhiteSpace(result.RefusalCode)) summary.Add("Refusal code: " + result.RefusalCode);
        summary.Add("Direct dependents: " + AuditPresentationLabels.Count(rows.Length));
        summary.Add("Formulas scanned: " + AuditPresentationLabels.Count(result.ScannedFormulaCount));
        summary.Add("Coverage gaps: " + AuditPresentationLabels.Count(result.CoverageGapCount));
        summary.Add("Truncated: " + (result.Truncated ? "yes" : "no"));
        summary.Add(AuditPresentationLabels.CompletenessLine(result.CanClaimCompleteness));
        if (!string.IsNullOrWhiteSpace(result.Message)) summary.Add("Message: " + result.Message);

        return new DirectDependentReport(
            result.Status,
            targetDisplay,
            result.Target.WorkbookId,
            AuditPresentationLabels.Status(result.Status),
            result.ScanScope,
            Summarize(result, targetDisplay, rows.Length),
            completeness,
            result.CanClaimCompleteness,
            rows.Length,
            result.ScannedFormulaCount,
            result.CoverageGapCount,
            result.Truncated,
            result.LimitationCode,
            result.RefusalCode,
            result.Message,
            Array.AsReadOnly(rows),
            summary.AsReadOnly());
    }

    /// <summary>Projects this report into the shared trace view shape.</summary>
    public TraceResultPresentation ToPresentation() => new TraceResultPresentation(
        "ExcelAccel Direct Dependents",
        Status,
        Headline,
        CompletenessStatement,
        new[]
        {
            new TraceColumn("Dependent", 200),
            new TraceColumn("Reached by", 120),
            new TraceColumn("Edges", 60),
            new TraceColumn("Source reference", 300),
        },
        Rows.Select(row => (IReadOnlyList<string>)new[]
        {
            row.DisplayTarget,
            row.Kinds,
            AuditPresentationLabels.Count(row.EdgeCount),
            row.SourceEvidence,
        }),
        SummaryLines,
        RefusalCode);

    private static string Summarize(DirectDependentResult result, string targetDisplay, int count)
    {
        if (result.Status == AuditTraceStatus.Refused)
        {
            return $"Direct dependents refused for {targetDisplay}: {result.Message} (code {result.RefusalCode})";
        }

        var prefix = $"{AuditPresentationLabels.Count(count)} direct dependent{(count == 1 ? string.Empty : "s")} of {targetDisplay} in the {result.ScanScope} scope.";
        return result.Status == AuditTraceStatus.Complete
            ? prefix + $" Complete across {AuditPresentationLabels.Count(result.ScannedFormulaCount)} scanned formulas."
            : prefix + $" Partial: {AuditPresentationLabels.Count(result.CoverageGapCount)} coverage gaps across {AuditPresentationLabels.Count(result.ScannedFormulaCount)} scanned formulas" +
                (result.Truncated ? ", and the scan was truncated." : ".");
    }

    private static DirectDependentRow ToRow(DirectDependent dependent) => new DirectDependentRow(
        AuditPresentationLabels.Location(dependent.Dependent),
        dependent.Dependent.WorkbookId,
        string.Join(", ", dependent.Evidence
            .Select(evidence => AuditPresentationLabels.Kind(evidence.Kind))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)),
        dependent.Evidence.Count,
        AuditPresentationLabels.EvidenceList(dependent.Evidence));
}
