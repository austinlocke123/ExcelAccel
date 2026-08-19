using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Collaboration;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CollaborationPolicyTests
{
    private static readonly DateTimeOffset PlannedAt = new DateTimeOffset(2026, 8, 18, 20, 0, 0, TimeSpan.Zero);
    private static readonly CollaborationPolicy Policy = new CollaborationPolicy(
        new CollaborationPolicyLimits(
            TimeSpan.FromSeconds(2),
            TimeSpan.FromSeconds(10),
            TimeSpan.FromSeconds(30)));

    public static TheoryData<CommandImpact, WorkbookCollaborationState, bool, bool, string?> PolicyMatrix =>
        new TheoryData<CommandImpact, WorkbookCollaborationState, bool, bool, string?>
        {
            { CommandImpact.ReadOnly, WorkbookCollaborationState.Unknown, false, true, null },
            { CommandImpact.Low, State(AutoSaveState.Unknown, CoauthoringState.NotDetected), false, false, CollaborationRefusalCodes.AutoSaveUnknown },
            { CommandImpact.Low, State(AutoSaveState.OffOrDisabled, CoauthoringState.Unknown), false, false, CollaborationRefusalCodes.CoauthoringUnknown },
            { CommandImpact.Low, State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected), false, true, null },
            { CommandImpact.Low, State(AutoSaveState.On, CoauthoringState.PotentialModern), false, true, null },
            { CommandImpact.Low, State(AutoSaveState.OffOrDisabled, CoauthoringState.LegacyShared), false, false, CollaborationRefusalCodes.LegacySharedUnqualified },
            { CommandImpact.Low, State(AutoSaveState.On, CoauthoringState.RemoteChangeInProgress, true), false, false, CollaborationRefusalCodes.RemoteChangeInProgress },
            { CommandImpact.Medium, State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected), false, false, CollaborationRefusalCodes.ReceiptRequired },
            { CommandImpact.Medium, State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected), true, true, null },
            { CommandImpact.Medium, State(AutoSaveState.On, CoauthoringState.PotentialModern), true, false, CollaborationRefusalCodes.RemoteEventsRequired },
            { CommandImpact.Medium, State(AutoSaveState.On, CoauthoringState.PotentialModern, true), true, true, null },
            { CommandImpact.High, State(AutoSaveState.On, CoauthoringState.PotentialModern, true), true, false, CollaborationRefusalCodes.CoauthoringUnqualified },
            { CommandImpact.High, State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected), false, true, null }
        };

    [Theory]
    [MemberData(nameof(PolicyMatrix))]
    public void ImpactPolicyIsExplicitForEveryCollaborationState(
        CommandImpact impact,
        WorkbookCollaborationState state,
        bool receiptEligible,
        bool expectedAllowed,
        string? expectedRefusalCode)
    {
        var decision = Policy.CanPlan(impact, state, receiptEligible);

        Assert.Equal(expectedAllowed, decision.Allowed);
        Assert.Equal(expectedRefusalCode, decision.RefusalCode);
    }

    [Fact]
    public void UnchangedFreshLowImpactLeaseIsExecutable()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.On, CoauthoringState.PotentialModern);
        var plannedStamp = tracker.Capture(Fingerprint("General"), state);
        var lease = new CollaborationPlanLease("format.number.currency", CommandImpact.Low, plannedStamp, PlannedAt);
        var currentStamp = tracker.Capture(Fingerprint("General"), state);

        var decision = Policy.ValidateForExecution(lease, currentStamp, PlannedAt.AddSeconds(1), false);

        Assert.True(decision.Allowed);
    }

    [Fact]
    public void LocalInterveningEventInvalidatesPlanBeforeMutation()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected);
        var lease = Lease(tracker.Capture(Fingerprint("General"), state), CommandImpact.Low);
        tracker.Record(WorkbookChangeKind.LocalEdit);

        var decision = Policy.ValidateForExecution(
            lease,
            tracker.Capture(Fingerprint("General"), state),
            PlannedAt.AddMilliseconds(10),
            false);

        Assert.False(decision.Allowed);
        Assert.Equal(CollaborationRefusalCodes.StaleWorkbookVersion, decision.RefusalCode);
    }

    [Fact]
    public void RemoteChangeBeginningRefusesAndCompletionKeepsOldPlanStale()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.On, CoauthoringState.PotentialModern, true);
        var lease = Lease(tracker.Capture(Fingerprint("General"), state), CommandImpact.Low);

        tracker.Record(WorkbookChangeKind.RemoteChangeBeginning);
        var duringRemoteChange = tracker.Capture(Fingerprint("General"), state);
        Assert.Equal(CoauthoringState.RemoteChangeInProgress, duringRemoteChange.Collaboration.Coauthoring);
        Assert.Equal(
            CollaborationRefusalCodes.StaleWorkbookVersion,
            Policy.ValidateForExecution(lease, duringRemoteChange, PlannedAt.AddMilliseconds(10), false).RefusalCode);

        tracker.Record(WorkbookChangeKind.RemoteChangeCompleted);
        var afterRemoteChange = tracker.Capture(Fingerprint("General"), state);
        Assert.Equal(CoauthoringState.RemoteChangeObserved, afterRemoteChange.Collaboration.Coauthoring);
        Assert.Equal(
            CollaborationRefusalCodes.StaleWorkbookVersion,
            Policy.ValidateForExecution(lease, afterRemoteChange, PlannedAt.AddMilliseconds(20), false).RefusalCode);
    }

    [Fact]
    public void FingerprintRevalidationCatchesChangesMissingFromEventCoverage()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.On, CoauthoringState.PotentialModern);
        var lease = Lease(tracker.Capture(Fingerprint("General"), state), CommandImpact.Low);
        var current = tracker.Capture(Fingerprint("0.00"), state);

        var decision = Policy.ValidateForExecution(lease, current, PlannedAt.AddMilliseconds(10), false);

        Assert.False(decision.Allowed);
        Assert.Equal(CollaborationRefusalCodes.StalePrecondition, decision.RefusalCode);
    }

    [Fact]
    public void LeaseExpiresAtImpactSpecificBound()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected);
        var stamp = tracker.Capture(Fingerprint("General"), state);
        var lowLease = Lease(stamp, CommandImpact.Low);
        var mediumLease = Lease(stamp, CommandImpact.Medium);

        Assert.Equal(
            CollaborationRefusalCodes.PlanExpired,
            Policy.ValidateForExecution(lowLease, stamp, PlannedAt.AddSeconds(2).AddTicks(1), false).RefusalCode);
        Assert.True(Policy.ValidateForExecution(mediumLease, stamp, PlannedAt.AddSeconds(2).AddTicks(1), true).Allowed);
    }

    [Fact]
    public void ClockRegressionRefusesPlan()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected);
        var stamp = tracker.Capture(Fingerprint("General"), state);

        var decision = Policy.ValidateForExecution(
            Lease(stamp, CommandImpact.Low),
            stamp,
            PlannedAt.AddTicks(-1),
            false);

        Assert.Equal(CollaborationRefusalCodes.ClockInvalid, decision.RefusalCode);
    }

    [Fact]
    public void WorkbookIdentityAndClosureInvalidateLease()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        var state = State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected);
        var stamp = tracker.Capture(Fingerprint("General"), state);
        var lease = Lease(stamp, CommandImpact.Low);
        var otherWorkbook = new WorkbookChangeTracker("Other.xlsx").Capture(Fingerprint("General"), state);

        Assert.Equal(
            CollaborationRefusalCodes.StaleWorkbookVersion,
            Policy.ValidateForExecution(lease, otherWorkbook, PlannedAt, false).RefusalCode);

        tracker.Close();
        var closed = tracker.Capture(Fingerprint("General"), state);
        Assert.Equal(
            CollaborationRefusalCodes.WorkbookClosed,
            Policy.ValidateForExecution(lease, closed, PlannedAt, false).RefusalCode);
        Assert.Throws<InvalidOperationException>(() => tracker.Record(WorkbookChangeKind.LocalEdit));
        tracker.Close();
    }

    [Fact]
    public void ConcurrentInvalidationsDoNotLoseRevisionIncrements()
    {
        var tracker = new WorkbookChangeTracker("Book.xlsx");
        const int eventCount = 1000;

        Parallel.For(0, eventCount, _ => tracker.Record(WorkbookChangeKind.LocalEdit));

        var stamp = tracker.Capture(string.Empty, State(AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected));
        Assert.Equal(eventCount, stamp.Revision);
    }

    [Fact]
    public void FingerprintsAreDeterministicAndComponentBounded()
    {
        Assert.Equal(Fingerprint("A", "BC"), Fingerprint("A", "BC"));
        Assert.NotEqual(Fingerprint("A", "BC"), Fingerprint("AB", "C"));
        Assert.NotEqual(Fingerprint(null, string.Empty), Fingerprint(string.Empty, null));
        Assert.Equal(64, Fingerprint("General").Length);
    }

    [Fact]
    public void FingerprintsAreLocaleInvariantAndHardLimited()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("fr-FR");
            var french = Fingerprint("1234", "0.00");
            CultureInfo.CurrentCulture = new CultureInfo("ar-SA");
            var arabic = Fingerprint("1234", "0.00");
            Assert.Equal(french, arabic);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PreconditionFingerprint.Create(new string[PreconditionFingerprint.MaximumComponentCount + 1]));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Fingerprint(new string('x', PreconditionFingerprint.MaximumTotalCharacters + 1)));
    }

    [Fact]
    public void PlanLeaseTimestampIsSeparateFromCanonicalCommandPlan()
    {
        var commandPlanProperties = typeof(CommandPlan).GetProperties().Select(property => property.Name).ToArray();
        Assert.DoesNotContain("CreatedAtUtc", commandPlanProperties);
        Assert.DoesNotContain("ExpiresAtUtc", commandPlanProperties);
    }

    [Fact]
    public void ReadOnlyLeaseStillRefusesClosedOrDifferentWorkbook()
    {
        var state = State(AutoSaveState.Unknown, CoauthoringState.Unknown);
        var plannedTracker = new WorkbookChangeTracker("Book.xlsx");
        var plannedStamp = plannedTracker.Capture(string.Empty, state);
        var lease = Lease(plannedStamp, CommandImpact.ReadOnly);
        var other = new WorkbookChangeTracker("Other.xlsx").Capture(string.Empty, state);

        Assert.Equal(
            CollaborationRefusalCodes.StaleWorkbookVersion,
            Policy.ValidateForExecution(lease, other, PlannedAt, false).RefusalCode);

        plannedTracker.Close();
        Assert.Equal(
            CollaborationRefusalCodes.WorkbookClosed,
            Policy.ValidateForExecution(
                lease,
                plannedTracker.Capture(string.Empty, state),
                PlannedAt,
                false).RefusalCode);
    }

    [Fact]
    public void InvalidEnumValuesAreRejectedAtPublicBoundaries()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new WorkbookCollaborationState((AutoSaveState)99, CoauthoringState.NotDetected, false));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new WorkbookCollaborationState(AutoSaveState.On, (CoauthoringState)99, false));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            Policy.CanPlan((CommandImpact)99, WorkbookCollaborationState.Unknown, false));

        var tracker = new WorkbookChangeTracker("Book.xlsx");
        Assert.Throws<ArgumentOutOfRangeException>(() => tracker.Record((WorkbookChangeKind)99));
    }

    [Theory]
    [InlineData(null, false, WorkbookLocationKind.Unsaved, AutoSaveState.Unknown, CoauthoringState.NotDetected)]
    [InlineData(false, false, WorkbookLocationKind.Unsaved, AutoSaveState.OffOrDisabled, CoauthoringState.NotDetected)]
    [InlineData(false, false, WorkbookLocationKind.LocalOrSyncedPath, AutoSaveState.OffOrDisabled, CoauthoringState.Unknown)]
    [InlineData(false, false, WorkbookLocationKind.CloudUrl, AutoSaveState.OffOrDisabled, CoauthoringState.PotentialModern)]
    [InlineData(true, false, WorkbookLocationKind.LocalOrSyncedPath, AutoSaveState.On, CoauthoringState.PotentialModern)]
    [InlineData(false, true, WorkbookLocationKind.LocalOrSyncedPath, AutoSaveState.OffOrDisabled, CoauthoringState.LegacyShared)]
    [InlineData(false, null, WorkbookLocationKind.Unsaved, AutoSaveState.OffOrDisabled, CoauthoringState.Unknown)]
    public void ReadOnlyWorkbookSignalsClassifyWithoutInventingExclusivity(
        bool? autoSaveOn,
        bool? legacyShared,
        WorkbookLocationKind location,
        AutoSaveState expectedAutoSave,
        CoauthoringState expectedCoauthoring)
    {
        var state = WorkbookCollaborationClassifier.Classify(
            new WorkbookCollaborationProbe(autoSaveOn, legacyShared, location, false));

        Assert.Equal(expectedAutoSave, state.AutoSave);
        Assert.Equal(expectedCoauthoring, state.Coauthoring);
        Assert.False(state.RemoteChangeEventsSupported);
    }

    [Fact]
    public void SelectionSnapshotDefaultsToUnknownCollaborationState()
    {
        var snapshot = new SelectionSnapshot(
            new SelectionContext("Book.xlsx", "Sheet1", "A1"),
            1,
            false,
            "General");

        Assert.Equal(AutoSaveState.Unknown, snapshot.Collaboration.AutoSave);
        Assert.Equal(CoauthoringState.Unknown, snapshot.Collaboration.Coauthoring);
    }

    private static WorkbookCollaborationState State(
        AutoSaveState autoSave,
        CoauthoringState coauthoring,
        bool remoteEvents = false) =>
        new WorkbookCollaborationState(autoSave, coauthoring, remoteEvents);

    private static CollaborationPlanLease Lease(WorkbookConcurrencyStamp stamp, CommandImpact impact) =>
        new CollaborationPlanLease("test.command", impact, stamp, PlannedAt);

    private static string Fingerprint(params string?[] values) => PreconditionFingerprint.Create(values);
}
