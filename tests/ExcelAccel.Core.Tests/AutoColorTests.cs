using System.Linq;
using ExcelAccel.Application.AutoColor;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class AutoColorTests
{
    [Theory]
    [InlineData(CellScalarKind.Text, "", AutoColorCategory.Text)]
    [InlineData(CellScalarKind.Number, "", AutoColorCategory.NumericHardcode)]
    [InlineData(CellScalarKind.Number, "=A1+1", AutoColorCategory.SameSheetFormula)]
    [InlineData(CellScalarKind.Number, "='Other Sheet'!A1", AutoColorCategory.CrossSheetFormula)]
    [InlineData(CellScalarKind.Number, "='[Other.xlsx]Sheet1'!A1", AutoColorCategory.ExternalFormula)]
    [InlineData(CellScalarKind.Error, "=1/0", AutoColorCategory.Error)]
    [InlineData(CellScalarKind.Boolean, "", AutoColorCategory.Unsupported)]
    public void ClassifierDistinguishesGoldenCategories(CellScalarKind kind, string formula, AutoColorCategory expected) =>
        Assert.Equal(expected, AutoColorPlanner.Classify(new AutoColorCellSnapshot("A1", kind, formula, "#FFFFFF")));

    [Fact]
    public void PlanIsDeterministicChangesOnlyFontColorAndCountsUnsupported()
    {
        var cells = new[]
        {
            new AutoColorCellSnapshot("B1", CellScalarKind.Number, "=Sheet2!A1", "#000000"),
            new AutoColorCellSnapshot("A1", CellScalarKind.Number, "", "#000000"),
            new AutoColorCellSnapshot("C1", CellScalarKind.Boolean, "", "#000000"),
        };
        var snapshot = new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1:C1"), 3, null, "General");
        var profile = new ProfileStore().LoadDefault();

        var first = AutoColorPlanner.Plan(profile, snapshot, cells, AutoColorScope.Selection);
        var second = AutoColorPlanner.Plan(profile, snapshot, cells.Reverse(), AutoColorScope.Selection);

        Assert.Equal(first.Fingerprint, second.Fingerprint);
        Assert.Equal(new[] { "A1", "B1" }, first.Changes.Select(value => value.Address));
        Assert.All(first.Changes, value => Assert.False(string.IsNullOrWhiteSpace(value.AfterColor)));
        Assert.Equal(1, first.UnsupportedCount);
        Assert.False(first.RequiresPreview);
    }

    [Fact]
    public void WorksheetAlwaysRequiresPreviewAndExecutionRemainsGateClosed()
    {
        var snapshot = new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General");
        var plan = AutoColorPlanner.Plan(new ProfileStore().LoadDefault(), snapshot,
            new[] { new AutoColorCellSnapshot("A1", CellScalarKind.Text, "", "#000000") }, AutoColorScope.Worksheet);

        Assert.True(plan.RequiresPreview);
        var gate = AutoColorPlanner.ExecutionGate();
        Assert.False(gate.Allowed);
        Assert.Equal("PERFORMANCE_QUALIFICATION_REQUIRED", gate.RefusalCode);
    }

    [Fact]
    public void UnsafeScopeRefusesBeforeAnyExecutionCanExist()
    {
        var snapshot = new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General",
            new SelectionSafetyState(1, false, true, false, false, false));
        Assert.Throws<CommandRefusedException>(() => AutoColorPlanner.Plan(new ProfileStore().LoadDefault(), snapshot,
            new[] { new AutoColorCellSnapshot("A1", CellScalarKind.Text, "", "#000000") }, AutoColorScope.Selection));
    }
}
