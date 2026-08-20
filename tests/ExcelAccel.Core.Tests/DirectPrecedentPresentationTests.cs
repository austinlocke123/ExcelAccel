using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DirectPrecedentPresentationTests
{
    private static readonly AuditCellIdentity Source = new AuditCellIdentity("Book.xlsx", "Model", "D10");

    [Fact]
    public void CompleteResultReportsClaimedCompletenessAndExactRows()
    {
        var report = DirectPrecedentReport.Create(Analyze("=A1+B2", "A1", "B2"));

        Assert.Equal(AuditTraceStatus.Complete, report.Status);
        Assert.Equal("Complete", report.StatusLabel);
        Assert.True(report.CanClaimCompleteness);
        Assert.Equal("Model!D10", report.SourceDisplay);
        Assert.Equal("Book.xlsx", report.WorkbookDisplay);
        Assert.Equal("target_formula", report.ScanScope);
        Assert.Equal(new[] { "Model!A1", "Model!B2" }, report.Rows.Select(row => row.DisplayTarget));
        Assert.All(report.Rows, row =>
        {
            Assert.Equal("Cell", row.Kind);
            Assert.Equal("Value", row.Classification);
            Assert.Equal("Resolved", row.State);
            Assert.Equal(1, row.EdgeCount);
        });
        Assert.Contains("Completeness: claimed", report.SummaryLines);
        Assert.Contains("Complete for the target formula.", report.Headline);
    }

    [Fact]
    public void SourceEvidenceRetainsEveryDeduplicatedEdgeWithItsSpan()
    {
        var report = DirectPrecedentReport.Create(Analyze("=A1+$A$1", "A1"));

        var row = Assert.Single(report.Rows);
        Assert.Equal("Model!A1", row.DisplayTarget);
        Assert.Equal(2, row.EdgeCount);
        Assert.Equal("A1 [1+2]; $A$1 [4+4]", row.SourceEvidence);
    }

    [Fact]
    public void ClosedExternalEdgeIsLabelledAsNeverOpenedAndBlocksCompleteness()
    {
        var report = DirectPrecedentReport.Create(Analyze("='[Closed.xlsx]Data'!C3"));

        Assert.Equal(AuditTraceStatus.Partial, report.Status);
        Assert.False(report.CanClaimCompleteness);
        var row = Assert.Single(report.Rows);
        Assert.Equal("External", row.Kind);
        Assert.Equal("External (closed; never opened)", row.State);
        Assert.Equal("Not captured", row.Classification);
        Assert.Equal("[Closed.xlsx]Data!C3", row.DisplayTarget);
        Assert.Equal(string.Empty, row.WorkbookDisplay);
        Assert.Contains("Completeness: not claimed", report.SummaryLines);
        Assert.Contains("External edges: 1", report.SummaryLines);
    }

    [Fact]
    public void UnresolvedNameIsReportedWithoutAWorkbookTarget()
    {
        var report = DirectPrecedentReport.Create(Analyze("=Rate"));

        Assert.Equal(AuditTraceStatus.Partial, report.Status);
        var row = Assert.Single(report.Rows);
        Assert.Equal("Unresolved", row.Kind);
        Assert.Equal("Unresolved", row.State);
        Assert.Equal("Rate", row.DisplayTarget);
        Assert.Contains("Unresolved edges: 1", report.SummaryLines);
        Assert.DoesNotContain("Complete", report.Headline);
    }

    [Fact]
    public void MissingCaptureClassificationIsNeverPresentedAsResolved()
    {
        var report = DirectPrecedentReport.Create(Analyze("=A1"));

        var row = Assert.Single(report.Rows);
        Assert.Equal("Not captured", row.Classification);
        Assert.Equal("Unresolved", row.State);
        Assert.False(report.CanClaimCompleteness);
    }

    [Fact]
    public void RefusedResultCarriesItsCodeAndPresentsNoPrecedentRow()
    {
        var refusal = DirectPrecedentResult.Refused(Source, AuditRefusalCodes.StaleTarget, "The source formula changed.");

        var report = DirectPrecedentReport.Create(refusal);

        Assert.Equal(AuditTraceStatus.Refused, report.Status);
        Assert.Empty(report.Rows);
        Assert.False(report.CanClaimCompleteness);
        Assert.Equal(AuditRefusalCodes.StaleTarget, report.RefusalCode);
        Assert.Contains("Refusal code: " + AuditRefusalCodes.StaleTarget, report.SummaryLines);
        Assert.Contains("refused for Model!D10", report.Headline);
        Assert.Equal("Refused", report.CoverageLabel);
    }

    [Fact]
    public void ParserCoverageGapIsNamedInTheSummary()
    {
        var report = DirectPrecedentReport.Create(Analyze("=SUM(Table1[Amount])"));

        Assert.Equal("Parser coverage gap (inspect only)", report.CoverageLabel);
        Assert.False(report.CanClaimCompleteness);
        Assert.Contains(report.SummaryLines, line => line.StartsWith("Limitation code: ", StringComparison.Ordinal));
    }

    [Fact]
    public void ProjectionIsDeterministicForAnIdenticalResult()
    {
        var first = DirectPrecedentReport.Create(Analyze("=A1+B2+'Inputs'!C3", "A1", "B2"));
        var second = DirectPrecedentReport.Create(Analyze("=A1+B2+'Inputs'!C3", "A1", "B2"));

        Assert.Equal(first.Headline, second.Headline);
        Assert.Equal(first.SummaryLines, second.SummaryLines);
        Assert.Equal(
            first.Rows.Select(row => row.DisplayTarget + "|" + row.State + "|" + row.SourceEvidence),
            second.Rows.Select(row => row.DisplayTarget + "|" + row.State + "|" + row.SourceEvidence));
    }

    [Fact]
    public void PresentationRejectsANullResult() =>
        Assert.Throws<ArgumentNullException>(() => DirectPrecedentReport.Create(null!));

    [Fact]
    public void TheRegisteredCommandIsReadOnlyAndCarriesItsAuditAcceptanceIds()
    {
        var descriptor = BuiltInCommandRegistry.GetRequired(AuditingCommandCatalog.DirectPrecedentsId);

        Assert.Equal(Core.Commands.CommandImpact.ReadOnly, descriptor.Impact);
        Assert.Empty(descriptor.ChangedProperties);
        Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
        Assert.Equal("CAP-AUD-001", descriptor.CapabilityId);
        Assert.Equal(
            new[] { "AC-AUD-001", "AC-AUD-002", "AC-AUD-003", "AC-AUD-004", "AC-AUD-005" },
            descriptor.AcceptanceIds);
        Assert.Equal("Alt, X, A, AP, D", descriptor.ShortcutLabel);
    }

    private static DirectPrecedentResult Analyze(string formula, params string[] capturedValueAddresses)
    {
        var index = new ReferenceSnapshotIndex(capturedValueAddresses.Select(address =>
            new KeyValuePair<AuditCellIdentity, AuditCellClassification>(
                new AuditCellIdentity(Source.WorkbookId, Source.WorksheetName, address),
                AuditCellClassification.Value)));
        return new DirectPrecedentAnalyzer().Analyze(new FormulaReferenceSnapshot(Source, formula, index));
    }
}
