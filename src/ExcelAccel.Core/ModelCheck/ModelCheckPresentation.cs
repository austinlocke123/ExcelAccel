using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// Deterministic read-only projection of a scan result. It formats findings the
/// engine already produced and never re-evaluates a rule.
/// </summary>
public sealed class ModelCheckReport
{
    private ModelCheckReport(
        ModelCheckScanResult result,
        string headline,
        string completenessStatement,
        IReadOnlyList<string> summaryLines)
    {
        Result = result;
        Headline = headline;
        CompletenessStatement = completenessStatement;
        SummaryLines = summaryLines;
    }

    public ModelCheckScanResult Result { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public static ModelCheckReport Create(ModelCheckScanResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        var completeness = result.CanClaimCompleteness
            ? "Every enabled rule ran to exact coverage over the scanned scope."
            : "Completeness cannot be claimed for this scan.";

        var summary = new List<string>
        {
            "Workbook: " + result.Snapshot.WorkbookId,
            "Scan scope: " + result.ScanScope,
            "Status: " + AuditPresentationLabels.Status(result.Status),
            "Cells scanned: " + AuditPresentationLabels.Count(result.Snapshot.Cells.Count),
            "Rules enabled: " + AuditPresentationLabels.Count(result.EnabledRuleIds.Count),
            "Findings: " + AuditPresentationLabels.Count(result.Findings.Count),
            "Suppressed by local ignores: " + AuditPresentationLabels.Count(result.SuppressedFindingCount),
            "Truncated: " + (result.Truncated ? "yes" : "no"),
        };
        if (!string.IsNullOrWhiteSpace(result.RefusalCode)) summary.Add("Refusal code: " + result.RefusalCode);
        foreach (var coverage in result.RuleCoverage.OrderBy(entry => entry.RuleId, StringComparer.Ordinal))
        {
            summary.Add("  " + coverage.RuleId + ": " + coverage.Coverage + ", " +
                AuditPresentationLabels.Count(coverage.FindingCount) + " findings over " +
                AuditPresentationLabels.Count(coverage.EvaluatedCellCount) + " cells");
        }

        foreach (var failure in result.RuleFailures)
        {
            summary.Add("  FAILED " + failure.RuleId + ": " + failure.Diagnostic);
        }

        summary.Add(AuditPresentationLabels.CompletenessLine(result.CanClaimCompleteness));
        if (!string.IsNullOrWhiteSpace(result.Message)) summary.Add("Message: " + result.Message);

        return new ModelCheckReport(result, Summarize(result), completeness, summary.AsReadOnly());
    }

    public TraceResultPresentation ToPresentation() => new TraceResultPresentation(
        "ExcelAccel Model Check",
        Result.Status,
        Headline,
        CompletenessStatement,
        new[]
        {
            new TraceColumn("Severity", 90),
            new TraceColumn("Rule", 250),
            new TraceColumn("Location", 150),
            new TraceColumn("Coverage", 90),
            new TraceColumn("Evidence", 320),
        },
        Result.Findings.Select(finding => new TraceRow(
            new[]
            {
                finding.Severity.ToString(),
                finding.RuleId + " v" + finding.RuleVersion.ToString(CultureInfo.InvariantCulture),
                AuditPresentationLabels.Location(finding.Target),
                finding.Coverage.ToString(),
                string.Join(" | ", finding.Evidence),
            },
            finding.Target)),
        SummaryLines,
        Result.RefusalCode);

    private static string Summarize(ModelCheckScanResult result)
    {
        if (result.Status == AuditTraceStatus.Refused)
        {
            return $"Model Check refused: {result.Message} (code {result.RefusalCode})";
        }

        var prefix = $"{AuditPresentationLabels.Count(result.Findings.Count)} finding" +
            (result.Findings.Count == 1 ? string.Empty : "s") +
            $" from {AuditPresentationLabels.Count(result.EnabledRuleIds.Count)} rules over " +
            $"{AuditPresentationLabels.Count(result.Snapshot.Cells.Count)} cells.";
        if (result.Status == AuditTraceStatus.Complete) return prefix + " Complete for the scanned scope.";
        var reasons = new List<string>();
        if (result.RuleFailures.Count != 0) reasons.Add($"{AuditPresentationLabels.Count(result.RuleFailures.Count)} rule failures");
        if (result.Truncated) reasons.Add("the finding cap was reached");
        var partialRules = result.RuleCoverage.Count(entry => entry.Coverage != RuleCoverage.Exact);
        if (partialRules != 0) reasons.Add($"{AuditPresentationLabels.Count(partialRules)} rules with incomplete coverage");
        return prefix + " Partial: " + string.Join(", ", reasons) + ".";
    }
}
