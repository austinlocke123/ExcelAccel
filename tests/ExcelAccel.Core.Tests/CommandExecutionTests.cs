using System;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

/// <summary>
/// Pipeline-level coverage that outlived <c>ApplyCurrencyFormatCommand</c>. The
/// bespoke one-shot currency command was removed when currency became a profile
/// cycle, so the safety, ceiling, and revalidation contracts it used to prove are
/// asserted here against <see cref="ProfileFormattingCommand"/>, which is now the
/// only implementation of that pipeline.
/// </summary>
public sealed class CommandExecutionTests
{
    [Fact]
    public void InspectSelectionIsReadOnly()
    {
        var port = new FakeSelectionPort(Snapshot("A1:B2", 4));
        var command = new InspectSelectionCommand();
        var plan = command.Plan(port.CaptureSelection());

        var result = command.Execute(plan, port);

        Assert.True(result.Succeeded);
        Assert.Equal(CommandImpact.ReadOnly, plan.Impact);
        Assert.Empty(plan.ChangedProperties);
        Assert.Equal(0, port.NumberFormatWriteCount);
    }

    [Fact]
    public void FormattingRefusesOversizedSelectionDuringPlanningWithoutWriting()
    {
        var port = new FakeFormattingPort("General")
        {
            Selection = Snapshot("A1:XFD1048576", ProfileFormattingCommand.MaximumCellCount + 1),
        };
        var command = Phase1AFormattingCatalog.Create("format.number.currency");

        var refusal = Assert.Throws<CommandRefusedException>(
            () => command.Plan(new ProfileStore().LoadDefault(), port));

        Assert.Equal(RefusalCodes.ResourceLimit, refusal.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Theory]
    [InlineData(2, false, false, false, false, false, RefusalCodes.SelectionUnsupported)]
    [InlineData(1, true, false, false, false, false, RefusalCodes.SelectionUnsupported)]
    [InlineData(1, false, true, false, false, false, RefusalCodes.ProtectedTarget)]
    [InlineData(1, false, false, true, false, false, RefusalCodes.ProtectedTarget)]
    [InlineData(1, false, false, false, true, false, RefusalCodes.ArrayOrSpillUnsafe)]
    [InlineData(1, false, false, false, false, true, RefusalCodes.ArrayOrSpillUnsafe)]
    public void FormattingReturnsStableSafetyRefusalCodes(
        int areaCount,
        bool merged,
        bool worksheetProtected,
        bool workbookReadOnly,
        bool legacyArray,
        bool dynamicSpill,
        string expected)
    {
        var port = new FakeFormattingPort("General")
        {
            Selection = SnapshotWithSafety(new SelectionSafetyState(
                areaCount, merged, worksheetProtected, workbookReadOnly, legacyArray, dynamicSpill)),
        };
        var command = Phase1AFormattingCatalog.Create("format.number.currency");

        var refusal = Assert.Throws<CommandRefusedException>(
            () => command.Plan(new ProfileStore().LoadDefault(), port));

        Assert.Equal(expected, refusal.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void FormattingRefusesWhenTheSpillCheckIsUnavailable()
    {
        var port = new FakeFormattingPort("General")
        {
            Selection = SnapshotWithSafety(
                new SelectionSafetyState(1, false, false, false, false, false, false)),
        };
        var command = Phase1AFormattingCatalog.Create("format.number.currency");

        var refusal = Assert.Throws<CommandRefusedException>(
            () => command.Plan(new ProfileStore().LoadDefault(), port));

        Assert.Equal(RefusalCodes.ArrayOrSpillUnsafe, refusal.RefusalCode);
    }

    [Fact]
    public void FormattingRevalidatesSafetyImmediatelyBeforeMutation()
    {
        var port = new FakeFormattingPort("General");
        var command = Phase1AFormattingCatalog.Create("format.number.currency");
        var plan = command.Plan(new ProfileStore().LoadDefault(), port);
        port.Selection = SnapshotWithSafety(new SelectionSafetyState(1, false, true, false, false, false));

        var result = command.Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(RefusalCodes.ProtectedTarget, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void FormattingRefusesAMovedSelectionWithoutWriting()
    {
        var port = new FakeFormattingPort("General");
        var command = Phase1AFormattingCatalog.Create("format.number.currency");
        var plan = command.Plan(new ProfileStore().LoadDefault(), port);
        port.Selection = Snapshot("B2", 1);

        var result = command.Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    private static SelectionSnapshot Snapshot(string address, long count) =>
        new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", address), count, false, "General");

    private static SelectionSnapshot SnapshotWithSafety(SelectionSafetyState safety) =>
        new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General", safety);

    private sealed class FakeFormattingPort : IFormattingPort
    {
        public FakeFormattingPort(string value)
        {
            Value = value;
            Selection = new SelectionSnapshot(
                new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General");
        }

        public string Value { get; set; }

        public SelectionSnapshot Selection { get; set; }

        public int WriteCount { get; private set; }

        public SelectionSnapshot CaptureSelection() => Selection;

        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();

        public string ReadFormattingProperty(string propertyId) => Value;

        public void WriteFormattingProperty(string propertyId, string invariantValue)
        {
            WriteCount++;
            Value = invariantValue;
        }
    }

    private sealed class FakeSelectionPort : ISelectionPort
    {
        public FakeSelectionPort(SelectionSnapshot current)
        {
            Current = current;
        }

        public SelectionSnapshot Current { get; set; }

        public int NumberFormatWriteCount { get; private set; }

        public SelectionSnapshot CaptureSelection() => Current;

        public void SetNumberFormat(string formatCode)
        {
            NumberFormatWriteCount++;
            Current = new SelectionSnapshot(
                Current.Context, Current.CellCount, Current.HasFormula, formatCode, Current.Safety, Current.Collaboration);
        }
    }
}
