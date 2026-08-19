using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DirectPrecedentAnalyzerTests
{
    private readonly DirectPrecedentAnalyzer _analyzer = new DirectPrecedentAnalyzer();
    private readonly AuditCellIdentity _source = new AuditCellIdentity("Book.xlsx", "Model", "D10");

    [Fact]
    public void DirectReferencesRetainIdentityKindClassificationAndEvidence()
    {
        var index = Index(
            Cell("Model", "A1", AuditCellClassification.Value),
            Cell("Inputs", "B2:C3", AuditCellClassification.Mixed),
            Cell("Model", "C5", AuditCellClassification.Error));
        var result = _analyzer.Analyze(Snapshot("=A1+'Inputs'!$B$2:$C$3+C5", index));

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.True(result.CanClaimCompleteness);
        Assert.Equal("target_formula", result.ScanScope);
        Assert.Equal(3, result.Precedents.Count);
        Assert.Contains(result.Precedents, item =>
            item.Target!.WorksheetName == "Model" && item.Target.Address == "A1" &&
            item.Kind == AuditReferenceKind.Cell && item.Classification == AuditCellClassification.Value);
        Assert.Contains(result.Precedents, item =>
            item.Target!.WorksheetName == "Inputs" && item.Target.Address == "B2:C3" &&
            item.Kind == AuditReferenceKind.Range && item.Classification == AuditCellClassification.Mixed);
        Assert.Contains(result.Precedents, item =>
            item.Target!.Address == "C5" && item.Classification == AuditCellClassification.Error);
        Assert.All(result.Precedents, item => Assert.Equal(1, item.Depth));
    }

    [Fact]
    public void EquivalentReferencesDeduplicateWithoutLosingSourceEdges()
    {
        var result = _analyzer.Analyze(Snapshot("=A1+$A$1+Model!A1", Index(Cell("Model", "A1", AuditCellClassification.Formula))));

        var precedent = Assert.Single(result.Precedents);
        Assert.Equal("A1", precedent.Target!.Address);
        Assert.Equal(3, precedent.Evidence.Count);
        Assert.Equal(new[] { "A1", "$A$1", "Model!A1" }, precedent.Evidence.Select(item => item.SourceText));
        Assert.Equal(new[] { 1, 4, 9 }, precedent.Evidence.Select(item => item.SourceSpan.Start));
        Assert.All(precedent.Evidence, item => Assert.Equal(AuditReferenceKind.Cell, item.Kind));
    }

    [Fact]
    public void NameAndCellResolvingToSameTargetShareNodeAndRetainEdgeKinds()
    {
        var target = new AuditCellIdentity("Book.xlsx", "Model", "A1");
        var index = new ReferenceSnapshotIndex(
            new[] { Pair(target, AuditCellClassification.Value) },
            new[] { new AuditNameBinding("Rate", AuditNameScope.Workbook, target) });

        var result = _analyzer.Analyze(Snapshot("=A1+Rate", index));

        var precedent = Assert.Single(result.Precedents);
        Assert.Equal(2, precedent.Evidence.Count);
        Assert.Equal(
            new[] { AuditReferenceKind.Cell, AuditReferenceKind.Name },
            precedent.Evidence.Select(item => item.Kind));
    }

    [Fact]
    public void ReversedRangeIsCanonicalizedForEquivalentDeduplication()
    {
        var result = _analyzer.Analyze(Snapshot("=B2:A1+$A$1:$B$2", Index(Cell("Model", "A1:B2", AuditCellClassification.Mixed))));

        var precedent = Assert.Single(result.Precedents);
        Assert.Equal(AuditReferenceKind.Range, precedent.Kind);
        Assert.Equal("A1:B2", precedent.Target!.Address);
        Assert.Equal(2, precedent.Evidence.Count);
    }

    [Fact]
    public void ClosedExternalReferenceIsVisibleAndNeverComplete()
    {
        var result = _analyzer.Analyze(Snapshot("='[Closed.xlsx]Data'!A1+B2", Index(Cell("Model", "B2", AuditCellClassification.Value))));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        Assert.Equal(1, result.ExternalEdgeCount);
        var external = Assert.Single(result.Precedents, item => item.IsExternal);
        Assert.True(external.IsClosedExternal);
        Assert.False(external.IsUnresolved);
        Assert.Null(external.Target);
        Assert.Contains("[Closed.xlsx]Data!A1", external.DisplayTarget, StringComparison.Ordinal);
        Assert.Equal(FormulaRefusalCodes.ExternalReferenceInspectOnly, result.LimitationCode);
    }

    [Fact]
    public void WorksheetNameBindingOverridesWorkbookBinding()
    {
        var localTarget = new AuditCellIdentity("Book.xlsx", "Inputs", "B2");
        var workbookTarget = new AuditCellIdentity("Book.xlsx", "Inputs", "B3");
        var index = new ReferenceSnapshotIndex(
            new[]
            {
                Pair(localTarget, AuditCellClassification.Value),
                Pair(workbookTarget, AuditCellClassification.Formula),
            },
            new[]
            {
                new AuditNameBinding("Rate", AuditNameScope.Workbook, workbookTarget),
                new AuditNameBinding("Rate", AuditNameScope.Worksheet, localTarget, "Model"),
            });

        var result = _analyzer.Analyze(Snapshot("=Rate+A1", index));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        var name = Assert.Single(result.Precedents, item => item.Kind == AuditReferenceKind.Name);
        Assert.Equal(localTarget, name.Target);
        Assert.Equal(AuditCellClassification.Value, name.Classification);
        Assert.Equal(FormulaRefusalCodes.NameInspectOnly, result.LimitationCode);
    }

    [Fact]
    public void UnresolvedNameIsProminentAndPreventsCompletenessClaim()
    {
        var result = _analyzer.Analyze(Snapshot("=MissingRate+A1", Index(Cell("Model", "A1", AuditCellClassification.Value))));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.Equal(1, result.UnresolvedEdgeCount);
        var unresolved = Assert.Single(result.Precedents, item => item.IsUnresolved);
        Assert.Equal("MissingRate", unresolved.DisplayTarget);
        Assert.Equal(AuditReferenceKind.Unresolved, unresolved.Kind);
    }

    [Fact]
    public void StructuredReferenceReturnsExplicitPartialResult()
    {
        var result = _analyzer.Analyze(Snapshot("=SUM(Table1[Amount])", Index()));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        Assert.Equal(FormulaRefusalCodes.StructuredReferenceInspectOnly, result.LimitationCode);
        Assert.Single(result.Precedents);
        Assert.True(result.Precedents[0].IsUnresolved);
    }

    [Fact]
    public void MissingCapturedClassificationPreventsFalseCompletenessClaim()
    {
        var result = _analyzer.Analyze(Snapshot("=A1", Index()));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        var precedent = Assert.Single(result.Precedents);
        Assert.Equal(AuditCellClassification.Unknown, precedent.Classification);
        Assert.True(precedent.IsUnresolved);
    }

    [Theory]
    [InlineData("hello", AuditRefusalCodes.TargetNotFormula)]
    [InlineData("=XFE1", FormulaRefusalCodes.InvalidReference)]
    public void NonFormulaOrInvalidFormulaReturnsCategorizedRefusal(string formula, string expectedCode)
    {
        var result = _analyzer.Analyze(Snapshot(formula, Index()));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.False(result.CanClaimCompleteness);
        Assert.Empty(result.Precedents);
        Assert.Equal(expectedCode, result.RefusalCode);
    }

    [Fact]
    public void R1C1SnapshotRefusesWithoutPartialTrace()
    {
        var snapshot = new FormulaReferenceSnapshot(_source, "=R1C1", Index(), FormulaDialect.InvariantR1C1);

        var result = _analyzer.Analyze(snapshot);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.NotationUnsupported, result.RefusalCode);
        Assert.Empty(result.Precedents);
    }

    [Fact]
    public void SnapshotDefensivelyCopiesCellsAndNames()
    {
        var cells = new List<KeyValuePair<AuditCellIdentity, AuditCellClassification>>
        {
            Cell("Model", "A1", AuditCellClassification.Value),
        };
        var names = new List<AuditNameBinding>
        {
            new AuditNameBinding("Rate", AuditNameScope.Workbook, cells[0].Key),
        };
        var index = new ReferenceSnapshotIndex(cells, names);
        cells.Clear();
        names.Clear();

        Assert.Single(index.Cells);
        Assert.Single(index.Names);
    }

    [Fact]
    public void AnalysisIsDeterministicForIdenticalSnapshot()
    {
        var snapshot = Snapshot("=C3+A1+C3", Index(
            Cell("Model", "A1", AuditCellClassification.Value),
            Cell("Model", "C3", AuditCellClassification.Formula)));

        var first = _analyzer.Analyze(snapshot);
        var second = _analyzer.Analyze(snapshot);

        Assert.Equal(first.Status, second.Status);
        Assert.Equal(first.Precedents.Select(item => item.NodeId), second.Precedents.Select(item => item.NodeId));
        Assert.Equal(first.Precedents.SelectMany(item => item.Evidence).Select(item => item.SourceSpan.Start),
            second.Precedents.SelectMany(item => item.Evidence).Select(item => item.SourceSpan.Start));
    }

    private FormulaReferenceSnapshot Snapshot(string formula, ReferenceSnapshotIndex index) =>
        new FormulaReferenceSnapshot(_source, formula, index);

    private static ReferenceSnapshotIndex Index(params KeyValuePair<AuditCellIdentity, AuditCellClassification>[] cells) =>
        new ReferenceSnapshotIndex(cells);

    private static KeyValuePair<AuditCellIdentity, AuditCellClassification> Cell(
        string sheet,
        string address,
        AuditCellClassification classification) =>
        Pair(new AuditCellIdentity("Book.xlsx", sheet, address), classification);

    private static KeyValuePair<AuditCellIdentity, AuditCellClassification> Pair(
        AuditCellIdentity identity,
        AuditCellClassification classification) =>
        new KeyValuePair<AuditCellIdentity, AuditCellClassification>(identity, classification);
}
