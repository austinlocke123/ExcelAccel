using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ModelCheckRuleTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    // --- pattern inconsistency -------------------------------------------------

    [Fact]
    public void CopiedFormulasShareAShapeAndTheOddOneOutIsReported()
    {
        var findings = Run<PatternInconsistencyRule>(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Formula("B3", "=A3*3"), Formula("B4", "=A4*2"));

        var finding = Assert.Single(findings);
        Assert.Equal("B3", finding.Target.Address);
        Assert.Contains(finding.Evidence, line => line.StartsWith("Baseline shape:", StringComparison.Ordinal));
        Assert.Contains(finding.Evidence, line => line.StartsWith("Peer region:", StringComparison.Ordinal));
    }

    [Fact]
    public void AbsoluteAndRelativeReferencesNormalizeDifferently()
    {
        var findings = Run<PatternInconsistencyRule>(
            Formula("B1", "=A1*$C$1"), Formula("B2", "=A2*$C$1"), Formula("B3", "=A3*C3"), Formula("B4", "=A4*$C$1"));

        Assert.Equal("B3", Assert.Single(findings).Target.Address);
    }

    [Fact]
    public void AConsistentRegionProducesNoPatternFinding() =>
        Assert.Empty(Run<PatternInconsistencyRule>(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Formula("B3", "=A3*2")));

    [Fact]
    public void ARegionSmallerThanThePeerMinimumIsNotJudged() =>
        Assert.Empty(Run<PatternInconsistencyRule>(Formula("B1", "=A1*2"), Formula("B2", "=A2*9")));

    [Fact]
    public void AParserGapPreventsAFalseConsistencyResultForTheRegion()
    {
        var snapshot = Snapshot(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Formula("B3", "=SUM(Table1[Amount])"), Formula("B4", "=A4*2"));

        var result = new ModelCheckEngine().Run(snapshot, new[] { new PatternInconsistencyRule() });

        Assert.Equal(RuleCoverage.Partial, result.RuleCoverage[0].Coverage);
        Assert.Empty(result.Findings);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void ABlankBreaksAPeerRegionUnderTheApprovedDefault()
    {
        var findings = Run<PatternInconsistencyRule>(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Formula("B3", "=A3*2"),
            Blank("B4"),
            Formula("B5", "=A5*9"), Formula("B6", "=A6*9"), Formula("B7", "=A7*9"));

        Assert.Empty(findings);
    }

    // --- constant interrupting a formula region --------------------------------

    [Fact]
    public void AConstantInsideAFormulaRegionIsReportedSeparatelyFromAnEmbeddedLiteral()
    {
        var snapshot = Snapshot(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Value("B3"), Formula("B4", "=A4*2"));

        var result = new ModelCheckEngine().Run(snapshot, ModelCheckRuleCatalog.All);

        var constant = Assert.Single(result.Findings, finding => finding.RuleId == ConstantInterruptsRegionRule.Id);
        Assert.Equal("B3", constant.Target.Address);
        Assert.DoesNotContain(result.Findings, finding =>
            finding.RuleId == EmbeddedNumericConstantRule.Id && finding.Target.Address == "B3");
    }

    [Fact]
    public void ABlankIsNotReportedAsAnInterruptingConstant() =>
        Assert.Empty(Run<ConstantInterruptsRegionRule>(
            Formula("B1", "=A1*2"), Formula("B2", "=A2*2"), Blank("B3"), Formula("B4", "=A4*2")));

    // --- embedded numeric constants --------------------------------------------

    [Fact]
    public void AnEmbeddedLiteralIsIdentifiedWithItsExactSourceSpan()
    {
        var finding = Assert.Single(Run<EmbeddedNumericConstantRule>(Formula("B1", "=A1*1.075")));

        Assert.Contains("Literal: 1.075", finding.Evidence);
        Assert.Contains(finding.Evidence, line => line.StartsWith("Source span: [", StringComparison.Ordinal));
    }

    [Fact]
    public void AllowlistedLiteralsAreExcludedDeterministically() =>
        Assert.Empty(Run<EmbeddedNumericConstantRule>(Formula("B1", "=A1*2"), Formula("B2", "=A2/100")));

    [Fact]
    public void AStructurallyLiteralTakingFunctionArgumentIsExcluded() =>
        Assert.Empty(Run<EmbeddedNumericConstantRule>(Formula("B1", "=ROUND(A1*B1,3)")));

    [Fact]
    public void ChangingTheAllowlistProducesTheExpectedRescanDelta()
    {
        var cells = new[] { Formula("B1", "=A1*7") };
        var before = new ModelCheckEngine().Run(Snapshot(cells), new[] { new EmbeddedNumericConstantRule() });
        var after = new ModelCheckEngine().Run(
            Snapshot(cells),
            new[] { new EmbeddedNumericConstantRule() },
            ModelCheckConfiguration.Default.WithAllowedEmbeddedLiterals(new[] { 0d, 1d, 7d }, 2));

        Assert.Single(before.Findings);
        Assert.Empty(after.Findings);
    }

    [Fact]
    public void NumericTextAndCellConstantsAreNotTreatedAsEmbeddedLiterals()
    {
        var snapshot = Snapshot(
            new ModelCheckCell(Cell("B1"), null, AuditCellClassification.Value, "General"),
            Formula("B2", "=\"12345\"&A2"));

        var result = new ModelCheckEngine().Run(snapshot, new[] { new EmbeddedNumericConstantRule() });

        Assert.Empty(result.Findings);
    }

    // --- errors ----------------------------------------------------------------

    [Fact]
    public void ABrokenReferenceAndACapturedErrorAreBothClassified()
    {
        var snapshot = Snapshot(
            Formula("B1", "=#REF!*2"),
            new ModelCheckCell(Cell("B2"), "=A2/0", AuditCellClassification.Error, "General", "#DIV/0!"));

        var findings = new ModelCheckEngine().Run(snapshot, new[] { new FormulaErrorRule() }).Findings;

        Assert.Equal(2, findings.Count);
        Assert.Contains(findings, finding => finding.Evidence.Contains("Error kind: #REF!"));
        Assert.Contains(findings, finding => finding.Evidence.Contains("Error kind: #DIV/0!"));
        Assert.All(findings, finding =>
            Assert.Contains("The workbook was not recalculated for this rule.", finding.Evidence));
    }

    [Fact]
    public void AHealthyFormulaProducesNoErrorFinding() =>
        Assert.Empty(Run<FormulaErrorRule>(Formula("B1", "=A1")));

    // --- external references ---------------------------------------------------

    [Fact]
    public void AnExternalReferenceIsReportedWithoutOpeningTheSource()
    {
        var finding = Assert.Single(Run<ExternalReferenceRule>(Formula("B1", "='[Other.xlsx]Data'!A1")));

        Assert.Contains(finding.Evidence, line => line.StartsWith("External sources:", StringComparison.Ordinal));
        Assert.Contains("No external workbook was opened, contacted, or refreshed.", finding.Evidence);
    }

    [Fact]
    public void ALocalReferenceIsNotReportedAsExternal() =>
        Assert.Empty(Run<ExternalReferenceRule>(Formula("B1", "='Inputs'!A1")));

    // --- circular references ---------------------------------------------------

    [Fact]
    public void ADirectCycleIsIdentifiedAndTerminates()
    {
        var findings = Run<CircularReferenceRule>(Formula("A1", "=B1"), Formula("B1", "=A1"));

        Assert.Equal(2, findings.Count);
        Assert.All(findings, finding => Assert.Contains(finding.Evidence, line => line.StartsWith("Cycle:", StringComparison.Ordinal)));
        Assert.All(findings, finding =>
            Assert.Contains("Iterative calculation settings were not read or changed.", finding.Evidence));
    }

    [Fact]
    public void ASelfReferenceIsIdentifiedAsACycle()
    {
        var findings = Run<CircularReferenceRule>(Formula("A1", "=A1+1"));

        Assert.Single(findings);
    }

    [Fact]
    public void ALongerCycleTerminatesAndIsReportedOnce()
    {
        var findings = Run<CircularReferenceRule>(
            Formula("A1", "=B1"), Formula("B1", "=C1"), Formula("C1", "=A1"));

        Assert.Equal(3, findings.Count);
        Assert.Single(findings.Select(finding => finding.PeerContext).Distinct());
    }

    [Fact]
    public void AnAcyclicChainProducesNoCircularFinding() =>
        Assert.Empty(Run<CircularReferenceRule>(Formula("B1", "=A1"), Formula("C1", "=B1")));

    // --- number formats --------------------------------------------------------

    [Fact]
    public void ANumberFormatExceptionIsReportedAgainstItsBaseline()
    {
        var findings = Run<NumberFormatInconsistencyRule>(
            Value("B1", "0.00"), Value("B2", "0.00"), Value("B3", "0%"), Value("B4", "0.00"));

        var finding = Assert.Single(findings);
        Assert.Equal("B3", finding.Target.Address);
        Assert.Contains("Baseline format: 0.00", finding.Evidence);
        Assert.Contains("This cell's format: 0%", finding.Evidence);
    }

    [Fact]
    public void FormatComparisonIgnoresCaseAndSurroundingWhitespace() =>
        Assert.Empty(Run<NumberFormatInconsistencyRule>(
            Value("B1", "0.00"), Value("B2", " 0.00 "), Value("B3", "0.00")));

    [Fact]
    public void FormulaAndValueCellsShareAFormatPeerRegion()
    {
        var findings = Run<NumberFormatInconsistencyRule>(
            new ModelCheckCell(Cell("B1"), "=A1", AuditCellClassification.Formula, "0.00"),
            Value("B2", "0.00"),
            Value("B3", "0.00"),
            new ModelCheckCell(Cell("B4"), "=A4", AuditCellClassification.Formula, "0%"));

        Assert.Equal("B4", Assert.Single(findings).Target.Address);
    }

    [Fact]
    public void EveryCatalogRuleHasAStableIdVersionAndStatement()
    {
        Assert.Equal(ModelCheckRuleCatalog.All.Count, ModelCheckRuleCatalog.AllRuleIds.Count);
        Assert.All(ModelCheckRuleCatalog.All, rule =>
        {
            Assert.StartsWith("check.", rule.RuleId, StringComparison.Ordinal);
            Assert.True(rule.Version >= 1);
            Assert.False(string.IsNullOrWhiteSpace(rule.Statement));
        });
        Assert.Equal(
            ModelCheckRuleCatalog.AllRuleIds,
            ModelCheckRuleCatalog.AllRuleIds.Distinct(StringComparer.Ordinal));
    }

    /// <summary>
    /// Findings state a rule, a location, and evidence. They never declare a
    /// verdict on the model or carry a score.
    /// </summary>
    [Fact]
    public void NoRuleStatementDeclaresCorrectnessOrCarriesAScore()
    {
        var banned = new[] { "wrong", "incorrect", "error-free", "score", "quality", "grade", "healthy", "bad" };
        Assert.All(ModelCheckRuleCatalog.All, rule =>
            Assert.DoesNotContain(banned, word => rule.Statement.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0));
    }

    private static AuditCellIdentity Cell(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static ModelCheckCell Formula(string address, string formula) =>
        new ModelCheckCell(Cell(address), formula, AuditCellClassification.Formula, "General");

    private static ModelCheckCell Value(string address, string numberFormat = "General") =>
        new ModelCheckCell(Cell(address), null, AuditCellClassification.Value, numberFormat);

    private static ModelCheckCell Blank(string address) =>
        new ModelCheckCell(Cell(address), null, AuditCellClassification.Blank, "General");

    private static ModelCheckSnapshot Snapshot(params ModelCheckCell[] cells) =>
        new ModelCheckSnapshot(ModelCheckScopeKind.Worksheet, Workbook, cells);

    private static IReadOnlyList<ModelCheckFinding> Run<TRule>(params ModelCheckCell[] cells)
        where TRule : IModelCheckRule, new() =>
        new ModelCheckEngine().Run(Snapshot(cells), new IModelCheckRule[] { new TRule() }).Findings;
}
