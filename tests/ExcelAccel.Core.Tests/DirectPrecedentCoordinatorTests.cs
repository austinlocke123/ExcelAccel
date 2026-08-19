using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DirectPrecedentCoordinatorTests
{
    [Fact]
    public void CoordinatorCapturesOnlyPlannedLocalTargetsAndRevalidatesSource()
    {
        var target = new AuditCellIdentity("Book.xlsx", "Model", "D10");
        var port = new FakePort(new FormulaTargetCapture(target, "=A1+'Inputs'!B2+'[Closed.xlsx]Data'!C3"));

        var result = new DirectPrecedentCoordinator().Execute(port);

        Assert.True(port.SourceWasRevalidated);
        Assert.Equal(
            new[] { "Book.xlsx|Inputs|B2", "Book.xlsx|Model|A1" },
            port.CapturedTargets.Select(value => value.ToString()));
        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.Equal(3, result.Precedents.Count);
        Assert.Equal(1, result.ExternalEdgeCount);
    }

    [Fact]
    public void CoordinatorReturnsCategorizedStaleRefusalWithoutPublishingAnalysis()
    {
        var target = new AuditCellIdentity("Book.xlsx", "Model", "D10");
        var port = new FakePort(new FormulaTargetCapture(target, "=A1")) { SourceIsCurrent = false };

        var result = new DirectPrecedentCoordinator().Execute(port);

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.StaleTarget, result.RefusalCode);
        Assert.Empty(result.Precedents);
    }

    [Fact]
    public void CapturePlanIncludesNamesEvenWhenAnotherInspectOnlyLimitationWins()
    {
        var target = new AuditCellIdentity("Book.xlsx", "Model", "D10");
        var analyzer = new DirectPrecedentAnalyzer();

        var plan = analyzer.CreateCapturePlan(target, "='[Closed.xlsx]Data'!C3+Rate+A1");

        Assert.Equal(new[] { "Book.xlsx|Model|A1" }, plan.LocalTargets.Select(value => value.ToString()));
        Assert.Equal(new[] { "Rate" }, plan.NameCandidates);
    }

    private sealed class FakePort : IDirectPrecedentSnapshotPort
    {
        private readonly FormulaTargetCapture _capture;

        public FakePort(FormulaTargetCapture capture) => _capture = capture;

        public bool SourceIsCurrent { get; set; } = true;
        public bool SourceWasRevalidated { get; private set; }
        public IReadOnlyList<AuditCellIdentity> CapturedTargets { get; private set; } = Array.Empty<AuditCellIdentity>();

        public FormulaTargetCapture CaptureTarget() => _capture;

        public ReferenceSnapshotIndex CaptureIndex(DirectPrecedentCapturePlan plan)
        {
            CapturedTargets = plan.LocalTargets;
            return new ReferenceSnapshotIndex(plan.LocalTargets.Select(target =>
                new KeyValuePair<AuditCellIdentity, AuditCellClassification>(target, AuditCellClassification.Value)));
        }

        public bool SourceMatches(FormulaTargetCapture capture)
        {
            SourceWasRevalidated = true;
            return SourceIsCurrent;
        }
    }
}
