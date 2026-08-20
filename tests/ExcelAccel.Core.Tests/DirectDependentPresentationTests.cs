using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DirectDependentPresentationTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void ACompleteScanReportsClaimedCompletenessScopeAndExactRows()
    {
        var report = DirectDependentReport.Create(Scan("A1", ("B1", "=A1"), ("C1", "=SUM(A1:A9)")));

        Assert.Equal("Complete", report.StatusLabel);
        Assert.True(report.CanClaimCompleteness);
        Assert.Equal("Model!A1", report.TargetDisplay);
        Assert.Equal("worksheet", report.ScanScope);
        Assert.Equal(new[] { "Model!B1", "Model!C1" }, report.Rows.Select(row => row.DisplayTarget));
        Assert.Equal("Cell", report.Rows[0].Kinds);
        Assert.Equal("Range", report.Rows[1].Kinds);
        Assert.Equal("A1 [1+2]", report.Rows[0].SourceEvidence);
        Assert.Contains("Completeness: claimed", report.SummaryLines);
        Assert.Contains("Worksheet: Model", report.SummaryLines);
        Assert.Contains("Formulas scanned: 2", report.SummaryLines);
        Assert.Contains("worksheet scope", report.CompletenessStatement);
    }

    [Fact]
    public void CoverageGapsAreNamedAndBlockTheCompletenessClaim()
    {
        var report = DirectDependentReport.Create(Scan("A1", ("B1", "=A1"), ("C1", "=Rate")));

        Assert.Equal("Partial", report.StatusLabel);
        Assert.False(report.CanClaimCompleteness);
        Assert.Contains("Coverage gaps: 1", report.SummaryLines);
        Assert.Contains("Completeness: not claimed", report.SummaryLines);
        Assert.Contains("coverage gaps", report.Headline);
    }

    [Fact]
    public void ADependentReachedByTwoKindsListsBothWithEveryEdge()
    {
        var names = new[] { new AuditNameBinding("Rate", AuditNameScope.Workbook, Cell("A1")) };
        var index = ReverseReferenceIndex.Build(
            DependentScanScope.Worksheet(Workbook, Sheet),
            new[] { new AuditFormulaCell(Cell("B1"), "=A1+Rate") },
            names);

        var report = DirectDependentReport.Create(index.FindDirectDependents(Cell("A1")));

        var row = Assert.Single(report.Rows);
        Assert.Equal("Cell, Name", row.Kinds);
        Assert.Equal(2, row.EdgeCount);
        Assert.Equal("A1 [1+2]; Rate [4+4]", row.SourceEvidence);
    }

    [Fact]
    public void ARefusedScanCarriesItsCodeAndPresentsNoDependentRow()
    {
        var scope = DependentScanScope.Worksheet(Workbook, Sheet);
        var refusal = DirectDependentResult.Refused(Cell("A1"), scope, AuditRefusalCodes.ScanCancelled, "The scan was cancelled.");

        var report = DirectDependentReport.Create(refusal);

        Assert.Equal("Refused", report.StatusLabel);
        Assert.Empty(report.Rows);
        Assert.False(report.CanClaimCompleteness);
        Assert.Contains("Refusal code: " + AuditRefusalCodes.ScanCancelled, report.SummaryLines);
        Assert.Contains("refused for Model!A1", report.Headline);
    }

    [Fact]
    public void ProjectionIsDeterministicForAnIdenticalResult()
    {
        var first = DirectDependentReport.Create(Scan("A1", ("B1", "=A1"), ("C1", "=A1")));
        var second = DirectDependentReport.Create(Scan("A1", ("B1", "=A1"), ("C1", "=A1")));

        Assert.Equal(first.Headline, second.Headline);
        Assert.Equal(first.SummaryLines, second.SummaryLines);
        Assert.Equal(first.Rows.Select(row => row.DisplayTarget), second.Rows.Select(row => row.DisplayTarget));
    }

    [Fact]
    public void PresentationRejectsANullResult() =>
        Assert.Throws<ArgumentNullException>(() => DirectDependentReport.Create(null!));

    /// <summary>
    /// Both auditing presentations must describe the same state in the same
    /// words, which is why the label maps have a single definition.
    /// </summary>
    [Fact]
    public void PrecedentAndDependentPresentationsShareTheirWording()
    {
        var dependent = DirectDependentReport.Create(Scan("A1", ("B1", "=A1")));
        var precedentIndex = new ReferenceSnapshotIndex(new[]
        {
            new KeyValuePair<AuditCellIdentity, AuditCellClassification>(Cell("A1"), AuditCellClassification.Value),
        });
        var precedent = DirectPrecedentReport.Create(new DirectPrecedentAnalyzer().Analyze(
            new FormulaReferenceSnapshot(Cell("B1"), "=A1", precedentIndex)));

        Assert.Equal(precedent.StatusLabel, dependent.StatusLabel);
        Assert.Equal("Model!A1", precedent.Rows[0].DisplayTarget);
        Assert.Equal("Model!B1", dependent.Rows[0].DisplayTarget);
        Assert.Equal(precedent.Rows[0].SourceEvidence, dependent.Rows[0].SourceEvidence);
        Assert.Contains("Completeness: claimed", precedent.SummaryLines);
        Assert.Contains("Completeness: claimed", dependent.SummaryLines);
    }

    [Fact]
    public void TheRegisteredCommandIsReadOnlyAndDeclaresItsThresholdPreview()
    {
        var descriptor = BuiltInCommandRegistry.GetRequired(AuditingCommandCatalog.DirectDependentsId);

        Assert.Equal(Core.Commands.CommandImpact.ReadOnly, descriptor.Impact);
        Assert.Empty(descriptor.ChangedProperties);
        Assert.Equal(PreviewPolicy.Threshold, descriptor.PreviewPolicy);
        Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
        Assert.Equal("CAP-AUD-001", descriptor.CapabilityId);
        Assert.Equal(new[] { "AC-AUD-006", "AC-AUD-007", "AC-AUD-008", "AC-AUD-009" }, descriptor.AcceptanceIds);
        Assert.Equal("Alt, X, A, A, DD", descriptor.ShortcutLabel);
    }

    private static AuditCellIdentity Cell(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static DirectDependentResult Scan(string target, params (string Address, string Formula)[] formulas)
    {
        var scope = DependentScanScope.Worksheet(Workbook, Sheet);
        var index = ReverseReferenceIndex.Build(
            scope,
            formulas.Select(cell => new AuditFormulaCell(Cell(cell.Address), cell.Formula)));
        return index.FindDirectDependents(Cell(target));
    }
}

public sealed class DependentScanPreviewTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";

    [Fact]
    public void ASmallWorksheetScansWithoutAskingForConfirmation()
    {
        var port = new PreviewPort(Rows(100));
        var confirmed = false;

        var result = new DirectDependentCoordinator().Execute(port, null, _ => { confirmed = true; return true; });

        Assert.False(confirmed);
        Assert.NotEqual(AuditTraceStatus.Refused, result.Status);
        Assert.NotEmpty(port.RequestedBands);
    }

    [Fact]
    public void ALargeWorksheetIsRefusedWhenTheScanIsNotConfirmed()
    {
        var port = new PreviewPort(Rows(30_000));

        var result = new DirectDependentCoordinator().Execute(port, null, _ => false);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.PreviewRequired, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void ALargeWorksheetIsRefusedWhenNoConfirmationIsAvailableAtAll()
    {
        var port = new PreviewPort(Rows(30_000));

        var result = new DirectDependentCoordinator().Execute(port);

        Assert.Equal(AuditRefusalCodes.PreviewRequired, result.RefusalCode);
        Assert.Empty(port.RequestedBands);
    }

    [Fact]
    public void AConfirmedLargeWorksheetScansAndDescribesThePlannedRead()
    {
        var port = new PreviewPort(Rows(30_000));
        DependentScanPreview? observed = null;

        var result = new DirectDependentCoordinator().Execute(port, null, preview => { observed = preview; return true; });

        Assert.NotNull(observed);
        Assert.Equal(Sheet, observed!.WorksheetName);
        Assert.Equal("Model!A1", observed.TargetDisplay);
        Assert.Equal(30_000, observed.CellCount);
        Assert.Equal(port.RequestedBands.Count, observed.BlockCount);
        Assert.NotEqual(AuditTraceStatus.Refused, result.Status);
    }

    [Fact]
    public void TheThresholdIsAppliedAtItsExactBoundary()
    {
        var below = new PreviewPort(Rows((int)DirectDependentCoordinator.PreviewThresholdCells));
        var above = new PreviewPort(Rows((int)DirectDependentCoordinator.PreviewThresholdCells + 1));
        var asked = 0;

        new DirectDependentCoordinator().Execute(below, null, _ => { asked++; return true; });
        Assert.Equal(0, asked);

        new DirectDependentCoordinator().Execute(above, null, _ => { asked++; return true; });
        Assert.Equal(1, asked);
    }

    private static UsedRegionBounds Rows(int rowCount) => new UsedRegionBounds(Sheet, 1, 1, rowCount, 1);

    private sealed class PreviewPort : IDependentScanPort
    {
        private readonly UsedRegionBounds _bounds;

        public PreviewPort(UsedRegionBounds bounds) => _bounds = bounds;

        public List<AuditRectangle> RequestedBands { get; } = new List<AuditRectangle>();

        public AuditCellIdentity CaptureTarget() => new AuditCellIdentity(Workbook, Sheet, "A1");

        public UsedRegionBounds CaptureUsedRegion(DependentScanScope scope) => _bounds;

        public IReadOnlyList<AuditFormulaCell> CaptureBlock(DependentScanScope scope, AuditRectangle band)
        {
            RequestedBands.Add(band);
            return Array.Empty<AuditFormulaCell>();
        }

        public IReadOnlyList<AuditNameBinding> CaptureNames(DependentScanScope scope) => Array.Empty<AuditNameBinding>();
    }
}
