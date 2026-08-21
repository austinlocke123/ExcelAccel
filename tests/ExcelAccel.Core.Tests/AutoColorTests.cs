using System.Linq;
using ExcelAccel.Application.AutoColor;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.ModelCheck;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class AutoColorTests
{
    [Theory]
    [InlineData(CellScalarKind.Text, "", AutoColorCategory.Text)]
    [InlineData(CellScalarKind.Number, "", AutoColorCategory.NumericHardcode)]
    [InlineData(CellScalarKind.Number, "=A1+B1", AutoColorCategory.SameSheetFormula)]
    [InlineData(CellScalarKind.Number, "=Sheet1!A1", AutoColorCategory.SameSheetFormula)]
    [InlineData(CellScalarKind.Number, "='Other Sheet'!A1", AutoColorCategory.CrossSheetFormula)]
    [InlineData(CellScalarKind.Number, "='[Other.xlsx]Sheet1'!A1", AutoColorCategory.ExternalFormula)]
    [InlineData(CellScalarKind.Error, "=1/0", AutoColorCategory.Error)]
    [InlineData(CellScalarKind.Boolean, "", AutoColorCategory.Unsupported)]
    public void ClassifierDistinguishesGoldenCategories(CellScalarKind kind, string formula, AutoColorCategory expected) =>
        Assert.Equal(expected, AutoColorPlanner.Classify(
            new AutoColorCellSnapshot("A1", kind, formula, "#FFFFFF"), "Sheet1"));

    /// <summary>
    /// The rule this feature exists for. A number typed into a formula is a
    /// hardcode however it is dressed up, and no allowlist softens it.
    /// </summary>
    [Theory]
    [InlineData("=A1*2")]
    [InlineData("=A1+1")]
    [InlineData("=SUM(A1:A9)/12")]
    [InlineData("=DATE(2026,1,1)")]
    public void AnyEmbeddedNumericLiteralMakesACellAHardcode(string formula) =>
        Assert.Equal(
            AutoColorCategory.NumericHardcode,
            AutoColorPlanner.Classify(new AutoColorCellSnapshot("A1", CellScalarKind.Number, formula, "#FFFFFF"), "Sheet1"));

    /// <summary>
    /// A hardcode outranks cross-sheet and external. The red external signal is
    /// given up on that one cell so the typed number stays findable; the external
    /// reference is still reported independently by Model Check.
    /// </summary>
    [Theory]
    [InlineData("='[Other.xlsx]Sheet1'!A1+5")]
    [InlineData("='Other Sheet'!A1*3")]
    public void AHardcodeOutranksCrossSheetAndExternal(string formula) =>
        Assert.Equal(
            AutoColorCategory.NumericHardcode,
            AutoColorPlanner.Classify(new AutoColorCellSnapshot("A1", CellScalarKind.Number, formula, "#FFFFFF"), "Sheet1"));

    /// <summary>
    /// AutoColor and Model Check answer different questions and are correct to
    /// disagree. Unifying them would either flood the findings list or blind the
    /// colour map, so the divergence is asserted rather than left to drift.
    /// </summary>
    [Fact]
    public void AutoColorIsStricterThanModelCheckOnTheSameLiteral()
    {
        const string formula = "=A1*2";

        var category = AutoColorPlanner.Classify(
            new AutoColorCellSnapshot("A1", CellScalarKind.Number, formula, "#FFFFFF"), "Sheet1");
        var literals = FormulaShape.ReadEmbeddedLiterals(formula);

        Assert.Equal(AutoColorCategory.NumericHardcode, category);
        Assert.Contains(literals, literal => literal.Value == 2);
        Assert.Contains(2d, new ModelCheckConfiguration().AllowedEmbeddedLiterals);
    }

    /// <summary>
    /// A formula outside qualified parser coverage cannot be classified honestly.
    /// Guessing would either hide a hardcode or invent one, so it is counted as
    /// unsupported and left exactly as it is.
    /// </summary>
    [Fact]
    public void AnUnparseableFormulaIsUnsupportedRatherThanGuessed() =>
        Assert.Equal(
            AutoColorCategory.Unsupported,
            AutoColorPlanner.Classify(
                new AutoColorCellSnapshot("A1", CellScalarKind.Number, "=((((", "#FFFFFF"), "Sheet1"));

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

    /// <summary>
    /// Selection scope is bounded by the selection the user made, so it is not
    /// blocked by the gate that worksheet scope waits behind.
    /// </summary>
    [Fact]
    public void SelectionScopeIsPermittedWhileWorksheetScopeStaysGated()
    {
        Assert.True(AutoColorPlanner.ExecutionGate(AutoColorScope.Selection).Allowed);

        var worksheet = AutoColorPlanner.ExecutionGate(AutoColorScope.Worksheet);
        Assert.False(worksheet.Allowed);
        Assert.Equal("PERFORMANCE_QUALIFICATION_REQUIRED", worksheet.RefusalCode);
    }

    [Fact]
    public void WorksheetAlwaysRequiresPreview()
    {
        var snapshot = new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General");

        var plan = AutoColorPlanner.Plan(new ProfileStore().LoadDefault(), snapshot,
            new[] { new AutoColorCellSnapshot("A1", CellScalarKind.Text, "", "#000000") }, AutoColorScope.Worksheet);

        Assert.True(plan.RequiresPreview);
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
