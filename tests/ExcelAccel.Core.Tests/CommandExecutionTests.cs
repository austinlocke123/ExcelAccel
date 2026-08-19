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

    private static SelectionSnapshot Snapshot(string address, long count) =>
        new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", address), count, false, "General");

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
