using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

/// <summary>
/// The trace view lifecycle. This state machine previously lived inside the
/// WinForms view, which the test project cannot reference, so it had no unit
/// coverage at all and was exercised only by the hidden-Excel smoke.
/// </summary>
public sealed class TraceViewSessionTests
{
    private const string CommandId = "audit.precedents.direct";
    private const string Workbook = "Book.xlsx";

    [Fact]
    public void ANewSessionHoldsNoResultAndRevalidatesToKeep()
    {
        var session = new TraceViewSession(CommandId);

        Assert.False(session.HasResult);
        Assert.Null(session.Presentation);
        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
    }

    [Fact]
    public void AnOpenWorkbookKeepsTheResultWithNoNotice()
    {
        var session = Presented(WorkbookPresence.Open);

        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
        Assert.True(session.HasResult);
        Assert.Null(session.Notice);
    }

    [Fact]
    public void AClosedWorkbookDiscardsTheResultSoAStaleTraceCannotStayOnScreen()
    {
        var session = Presented(WorkbookPresence.Closed);

        Assert.Equal(TraceViewAction.Discard, session.Revalidate());
        Assert.False(session.HasResult);
        Assert.Null(session.Presentation);
        Assert.Equal(string.Empty, session.WorkbookId);
    }

    [Fact]
    public void AnUnverifiableWorkbookKeepsTheResultUnderAnExplicitNotice()
    {
        var session = Presented(WorkbookPresence.Unknown);

        Assert.Equal(TraceViewAction.Warn, session.Revalidate());
        Assert.True(session.HasResult);
        Assert.Equal(TraceViewSession.UnverifiedNotice, session.Notice);
    }

    [Fact]
    public void AProbeFailureIsTreatedAsUnverifiedAndSurfacedForLogging()
    {
        var port = new FakePresence(WorkbookPresence.Open) { Failure = new InvalidOperationException("COM boom") };
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), Workbook, port);

        var action = session.Revalidate();

        Assert.Equal(TraceViewAction.Warn, action);
        Assert.True(session.HasResult);
        Assert.IsType<InvalidOperationException>(session.LastProbeError);
    }

    [Fact]
    public void ARecoveredWorkbookClearsAPreviousNotice()
    {
        var port = new FakePresence(WorkbookPresence.Unknown);
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), Workbook, port);
        Assert.Equal(TraceViewAction.Warn, session.Revalidate());

        port.Presence = WorkbookPresence.Open;

        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
        Assert.Null(session.Notice);
        Assert.Null(session.LastProbeError);
    }

    [Fact]
    public void PresentingAgainResetsTheNoticeAndRetargetsTheWorkbook()
    {
        var port = new FakePresence(WorkbookPresence.Unknown);
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), Workbook, port);
        session.Revalidate();

        session.Present(Presentation(), "Other.xlsx", new FakePresence(WorkbookPresence.Open));

        Assert.Null(session.Notice);
        Assert.Equal("Other.xlsx", session.WorkbookId);
        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
    }

    [Fact]
    public void ADiscardedSessionStopsProbingUntilSomethingIsPresentedAgain()
    {
        var port = new FakePresence(WorkbookPresence.Closed);
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), Workbook, port);

        Assert.Equal(TraceViewAction.Discard, session.Revalidate());
        var probesAfterDiscard = port.ProbeCount;

        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
        Assert.Equal(probesAfterDiscard, port.ProbeCount);
    }

    [Fact]
    public void ReentrantRevalidationDoesNotProbeTwice()
    {
        var session = new TraceViewSession(CommandId);
        var port = new FakePresence(WorkbookPresence.Open);
        port.OnProbe = () => session.Revalidate();
        session.Present(Presentation(), Workbook, port);

        session.Revalidate();

        Assert.Equal(1, port.ProbeCount);
    }

    [Fact]
    public void AnEmptyWorkbookIdIsNeverProbed()
    {
        var port = new FakePresence(WorkbookPresence.Closed);
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), string.Empty, port);

        Assert.Equal(TraceViewAction.Keep, session.Revalidate());
        Assert.Equal(0, port.ProbeCount);
    }

    [Fact]
    public void TheSessionRejectsInvalidConstructionAndPresentation()
    {
        Assert.Throws<ArgumentException>(() => new TraceViewSession(" "));
        var session = new TraceViewSession(CommandId);
        Assert.Throws<ArgumentNullException>(() => session.Present(null!, Workbook, new FakePresence(WorkbookPresence.Open)));
        Assert.Throws<ArgumentNullException>(() => session.Present(Presentation(), Workbook, null!));
    }

    private static TraceViewSession Presented(WorkbookPresence presence)
    {
        var session = new TraceViewSession(CommandId);
        session.Present(Presentation(), Workbook, new FakePresence(presence));
        return session;
    }

    private static TraceResultPresentation Presentation() => new TraceResultPresentation(
        "ExcelAccel Direct Precedents",
        AuditTraceStatus.Complete,
        "1 direct precedent for Model!D10.",
        "Completeness is claimed.",
        new[] { new TraceColumn("Target", 100) },
        new[] { (IReadOnlyList<string>)new[] { "Model!A1" } },
        new[] { "Status: Complete" },
        null);

    private sealed class FakePresence : IWorkbookPresencePort
    {
        public FakePresence(WorkbookPresence presence) => Presence = presence;

        public WorkbookPresence Presence { get; set; }

        public Exception? Failure { get; set; }

        public int ProbeCount { get; private set; }

        public Action? OnProbe { get; set; }

        public WorkbookPresence Probe(string workbookId)
        {
            ProbeCount++;
            OnProbe?.Invoke();
            if (Failure is not null) throw Failure;
            return Presence;
        }
    }
}

public sealed class TraceResultPresentationTests
{
    [Fact]
    public void BothAuditingReportsProjectIntoTheSharedShape()
    {
        var precedent = PrecedentReport().ToPresentation();
        var dependent = DependentReport().ToPresentation();

        Assert.Equal("ExcelAccel Direct Precedents", precedent.Title);
        Assert.Equal("ExcelAccel Direct Dependents", dependent.Title);
        foreach (var presentation in new[] { precedent, dependent })
        {
            Assert.NotEmpty(presentation.Columns);
            Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Count));
            Assert.NotEmpty(presentation.SummaryLines);
        }
    }

    [Fact]
    public void TheProjectionCarriesTheHeadlineCompletenessAndRefusalCode()
    {
        var refusal = DirectPrecedentResult.Refused(
            new AuditCellIdentity("Book.xlsx", "Model", "D10"), AuditRefusalCodes.StaleTarget, "changed");

        var presentation = DirectPrecedentReport.Create(refusal).ToPresentation();

        Assert.Equal(AuditTraceStatus.Refused, presentation.Status);
        Assert.Equal(AuditRefusalCodes.StaleTarget, presentation.RefusalCode);
        Assert.Empty(presentation.Rows);
        Assert.Contains("refused", presentation.Headline, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ARowThatDoesNotMatchTheColumnsIsRejected() =>
        Assert.Throws<ArgumentException>(() => new TraceResultPresentation(
            "Title",
            AuditTraceStatus.Complete,
            "headline",
            "completeness",
            new[] { new TraceColumn("A", 10), new TraceColumn("B", 10) },
            new[] { (IReadOnlyList<string>)new[] { "only one" } },
            new[] { "summary" },
            null));

    private static DirectPrecedentReport PrecedentReport()
    {
        var source = new AuditCellIdentity("Book.xlsx", "Model", "D10");
        var index = new ReferenceSnapshotIndex(new[]
        {
            new KeyValuePair<AuditCellIdentity, AuditCellClassification>(
                new AuditCellIdentity("Book.xlsx", "Model", "A1"), AuditCellClassification.Value),
        });
        return DirectPrecedentReport.Create(new DirectPrecedentAnalyzer().Analyze(
            new FormulaReferenceSnapshot(source, "=A1", index)));
    }

    private static DirectDependentReport DependentReport()
    {
        var scope = DependentScanScope.Worksheet("Book.xlsx", "Model");
        var index = ReverseReferenceIndex.Build(scope, new[]
        {
            new AuditFormulaCell(new AuditCellIdentity("Book.xlsx", "Model", "B1"), "=A1"),
        });
        return DirectDependentReport.Create(index.FindDirectDependents(new AuditCellIdentity("Book.xlsx", "Model", "A1")));
    }
}
