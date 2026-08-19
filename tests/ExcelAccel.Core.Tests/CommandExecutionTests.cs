using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

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
    public void CurrencyFormatWritesOnlyDeclaredNumberFormat()
    {
        var port = new FakeSelectionPort(Snapshot("A1:B2", 4));
        var command = new ApplyCurrencyFormatCommand();
        var plan = command.Plan(port.CaptureSelection());

        var result = command.Execute(plan, port);

        Assert.True(result.Succeeded);
        Assert.Equal(new[] { ApplyCurrencyFormatCommand.ChangedProperty }, plan.ChangedProperties);
        Assert.Equal(1, port.NumberFormatWriteCount);
        Assert.Equal(ApplyCurrencyFormatCommand.DefaultFormatCode, port.LastNumberFormat);
    }

    [Fact]
    public void CurrencyFormatRefusesStaleSelectionWithoutWriting()
    {
        var port = new FakeSelectionPort(Snapshot("A1", 1));
        var command = new ApplyCurrencyFormatCommand();
        var plan = command.Plan(port.CaptureSelection());
        port.Current = Snapshot("B2", 1);

        var result = command.Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(0, port.NumberFormatWriteCount);
    }

    [Fact]
    public void CurrencyFormatRefusesOversizedSelectionDuringPlanning()
    {
        var command = new ApplyCurrencyFormatCommand();
        var snapshot = Snapshot("A1:XFD1048576", ApplyCurrencyFormatCommand.MaximumCellCount + 1);

        Assert.Throws<CommandRefusedException>(() => command.Plan(snapshot));
    }

    [Fact]
    public void CurrencyFormatReturnsStableSafetyRefusalCodes()
    {
        var command = new ApplyCurrencyFormatCommand();

        Assert.Equal(RefusalCodes.MultiAreaUnsupported, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(2, false, false, false, false, false))).RefusalCode);
        Assert.Equal(RefusalCodes.SelectionUnsupported, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, true, false, false, false, false))).RefusalCode);
        Assert.Equal(RefusalCodes.ProtectedTarget, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, false, true, false, false, false))).RefusalCode);
        Assert.Equal(RefusalCodes.ReadOnlyWorkbook, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, false, false, true, false, false))).RefusalCode);
        Assert.Equal(RefusalCodes.ArrayOrSpillUnsafe, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, false, false, false, true, false))).RefusalCode);
        Assert.Equal(RefusalCodes.ArrayOrSpillUnsafe, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, false, false, false, false, true))).RefusalCode);
        Assert.Equal(RefusalCodes.ExcelCapabilityMissing, command.CanExecute(SnapshotWithSafety(new SelectionSafetyState(1, false, false, false, false, false, false))).RefusalCode);
    }

    [Fact]
    public void CurrencyFormatRevalidatesSafetyImmediatelyBeforeMutation()
    {
        var port = new FakeSelectionPort(Snapshot("A1", 1));
        var command = new ApplyCurrencyFormatCommand();
        var plan = command.Plan(port.CaptureSelection());
        port.Current = SnapshotWithSafety(new SelectionSafetyState(1, false, true, false, false, false));

        var result = command.Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(RefusalCodes.ProtectedTarget, result.RefusalCode);
        Assert.Equal(0, port.NumberFormatWriteCount);
    }

    [Fact]
    public void CurrencyFormatRefusesInterveningPlannedPropertyChange()
    {
        var port = new FakeSelectionPort(Snapshot("A1", 1));
        var command = new ApplyCurrencyFormatCommand();
        var plan = command.Plan(port.CaptureSelection());
        port.Current = new SelectionSnapshot(
            new SelectionContext("Book.xlsx", "Sheet1", "A1"),
            1,
            false,
            "0.00");

        var result = command.Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.NumberFormatWriteCount);
    }

    private static SelectionSnapshot Snapshot(string address, long count) =>
        new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", address), count, false, "General");

    private static SelectionSnapshot SnapshotWithSafety(SelectionSafetyState safety) =>
        new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General", safety);

    private sealed class FakeSelectionPort : ISelectionPort
    {
        public FakeSelectionPort(SelectionSnapshot current)
        {
            Current = current;
        }

        public SelectionSnapshot Current { get; set; }

        public int NumberFormatWriteCount { get; private set; }

        public string? LastNumberFormat { get; private set; }

        public SelectionSnapshot CaptureSelection() => Current;

        public void SetNumberFormat(string formatCode)
        {
            NumberFormatWriteCount++;
            LastNumberFormat = formatCode;
        }
    }
}
