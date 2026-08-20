using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ModelCheckEngineTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void AScanEvaluatesExactlyTheEnabledRulesAgainstOneSnapshot()
    {
        var snapshot = Snapshot(Formula("B1", "=A1"), Formula("B2", "=A2"), Formula("B3", "=A3"));

        var result = new ModelCheckEngine().Run(snapshot, ModelCheckRuleCatalog.Select(new[] { PatternInconsistencyRule.Id }));

        Assert.Equal(new[] { PatternInconsistencyRule.Id }, result.EnabledRuleIds);
        Assert.Single(result.RuleCoverage);
        Assert.Equal(PatternInconsistencyRule.Id, result.RuleCoverage[0].RuleId);
    }

    [Fact]
    public void EveryFindingCarriesRuleVersionSeverityTargetEvidenceCoverageAndFingerprint()
    {
        var result = Run(Formula("B1", "=A1*7"), Formula("B2", "=A2"), Formula("B3", "=A3"), Formula("B4", "=A4"));

        var finding = Assert.Single(result.Findings, value => value.RuleId == PatternInconsistencyRule.Id);
        Assert.Equal(1, finding.RuleVersion);
        Assert.Equal(ModelCheckSeverity.Attention, finding.Severity);
        Assert.Equal("Model!B1", AuditPresentationLabels.Location(finding.Target));
        Assert.NotEmpty(finding.Evidence);
        Assert.Equal(RuleCoverage.Exact, finding.Coverage);
        Assert.False(string.IsNullOrWhiteSpace(finding.Fingerprint));
        Assert.True(finding.IsNavigable);
    }

    [Fact]
    public void AFingerprintContainsNoRawFormulaOrValueContent()
    {
        var result = Run(Formula("B1", "=A1*987654"), Formula("B2", "=A2"), Formula("B3", "=A3"), Formula("B4", "=A4"));

        Assert.All(result.Findings, finding =>
        {
            Assert.Matches("^[0-9a-f]{64}$", finding.Fingerprint);
            Assert.DoesNotContain("987654", finding.Fingerprint);
            Assert.DoesNotContain("=", finding.Fingerprint);
        });
    }

    [Fact]
    public void ReorderingRuleExecutionDoesNotChangeCanonicalFindings()
    {
        var cells = new[] { Formula("B1", "=A1*7"), Formula("B2", "=A2"), Formula("B3", "=A3"), Formula("B4", "=A4") };
        var rules = ModelCheckRuleCatalog.All;

        var forward = new ModelCheckEngine().Run(Snapshot(cells), rules);
        var reversed = new ModelCheckEngine().Run(Snapshot(cells.Reverse().ToArray()), rules.Reverse().ToArray());

        Assert.Equal(
            forward.Findings.Select(finding => finding.CanonicalKey),
            reversed.Findings.Select(finding => finding.CanonicalKey));
    }

    [Fact]
    public void IdenticalInputsProduceIdenticalFingerprints()
    {
        var cells = new[] { Formula("B1", "=A1*7"), Formula("B2", "=A2"), Formula("B3", "=A3"), Formula("B4", "=A4") };

        var first = new ModelCheckEngine().Run(Snapshot(cells), ModelCheckRuleCatalog.All);
        var second = new ModelCheckEngine().Run(Snapshot(cells), ModelCheckRuleCatalog.All);

        Assert.Equal(
            first.Findings.Select(finding => finding.Fingerprint),
            second.Findings.Select(finding => finding.Fingerprint));
    }

    [Fact]
    public void ARuleFailureMakesTheScanIncompleteAndNamesTheRule()
    {
        var result = new ModelCheckEngine().Run(Snapshot(Formula("B1", "=A1")), new IModelCheckRule[] { new ThrowingRule() });

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        var failure = Assert.Single(result.RuleFailures);
        Assert.Equal("check.test.throwing", failure.RuleId);
        Assert.Equal("InvalidOperationException", failure.Diagnostic);
        Assert.Contains(result.RuleCoverage, entry => entry.RuleId == "check.test.throwing");
    }

    [Fact]
    public void CancellationRefusesAndPresentsNoPartialResults()
    {
        var result = new ModelCheckEngine().Run(
            Snapshot(Formula("B1", "=A1")), ModelCheckRuleCatalog.All, null, () => true);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(ModelCheckRefusalCodes.ScanCancelled, result.RefusalCode);
        Assert.Empty(result.Findings);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void AnActiveIgnoreSuppressesOnlyTheMatchingFingerprint()
    {
        var cells = new[] { Formula("B1", "=A1*7"), Formula("B2", "=A2"), Formula("B3", "=A3"), Formula("B4", "=A4") };
        var baseline = new ModelCheckEngine().Run(Snapshot(cells), ModelCheckRuleCatalog.All);
        var target = baseline.Findings.First();

        var suppressed = new ModelCheckEngine().Run(
            Snapshot(cells),
            ModelCheckRuleCatalog.All,
            ModelCheckConfiguration.Default.WithIgnoredFingerprints(new[] { target.Fingerprint }));

        Assert.Equal(1, suppressed.SuppressedFindingCount);
        Assert.DoesNotContain(suppressed.Findings, finding => finding.Fingerprint == target.Fingerprint);
        Assert.Equal(baseline.Findings.Count - 1, suppressed.Findings.Count);
    }

    [Fact]
    public void PartialRuleCoverageBlocksTheCompletenessClaim()
    {
        var snapshot = Snapshot(
            Formula("B1", "=SUM(Table1[Amount])"),
            Formula("B2", "=A2"),
            Formula("B3", "=A3"),
            Formula("B4", "=A4"));

        var result = new ModelCheckEngine().Run(snapshot, ModelCheckRuleCatalog.Select(new[] { PatternInconsistencyRule.Id }));

        Assert.Equal(RuleCoverage.Partial, result.RuleCoverage[0].Coverage);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void ACleanScanClaimsCompleteness()
    {
        var result = new ModelCheckEngine().Run(
            Snapshot(Formula("B1", "=A1"), Formula("B2", "=A2"), Formula("B3", "=A3")),
            ModelCheckRuleCatalog.Select(new[] { PatternInconsistencyRule.Id }));

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.True(result.CanClaimCompleteness);
        Assert.Empty(result.Findings);
    }

    [Fact]
    public void TheEngineRejectsNullArguments()
    {
        var engine = new ModelCheckEngine();
        Assert.Throws<ArgumentNullException>(() => engine.Run(null!, ModelCheckRuleCatalog.All));
        Assert.Throws<ArgumentNullException>(() => engine.Run(Snapshot(), null!));
    }

    [Fact]
    public void ASnapshotRejectsDuplicateCellsAndOversizeInput()
    {
        var duplicate = new[] { Formula("A1", "=B1"), Formula("A1", "=C1") };
        Assert.Throws<ArgumentException>(() => Snapshot(duplicate));
    }

    internal static ModelCheckCell Formula(string address, string formula) => new ModelCheckCell(
        new AuditCellIdentity(Workbook, Sheet, address), formula, AuditCellClassification.Formula, "General");

    internal static ModelCheckCell Value(string address, string numberFormat = "General") => new ModelCheckCell(
        new AuditCellIdentity(Workbook, Sheet, address), null, AuditCellClassification.Value, numberFormat);

    internal static ModelCheckSnapshot Snapshot(params ModelCheckCell[] cells) =>
        new ModelCheckSnapshot(ModelCheckScopeKind.Worksheet, Workbook, cells);

    private static ModelCheckScanResult Run(params ModelCheckCell[] cells) =>
        new ModelCheckEngine().Run(Snapshot(cells), ModelCheckRuleCatalog.All);

    private sealed class ThrowingRule : IModelCheckRule
    {
        public string RuleId => "check.test.throwing";

        public int Version => 1;

        public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Information;

        public string Statement => "Always fails.";

        public ModelCheckRuleResult Evaluate(ModelCheckContext context) =>
            throw new InvalidOperationException("seeded failure");
    }
}
