using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.DataCleaning;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DataCleaningCommandTests
{
    [Theory]
    [InlineData(" \t\u00a0hello\u3000\r\n", "hello")]
    [InlineData("\u2003世界\u202f", "世界")]
    [InlineData("internal  space", "internal  space")]
    public void TrimUsesOnlyTheVersionedUnicodeWhitespaceTable(string input, string expected)
    {
        Assert.Equal(expected, DataCleaningCommand.TrimOuter(input));
        Assert.Equal(expected, DataCleaningCommand.TrimOuter(DataCleaningCommand.TrimOuter(input)));
    }

    [Theory]
    [InlineData("  alpha\t\u00a0beta\u2003 gamma  ", "alpha beta gamma")]
    [InlineData("日本語\u3000テキスト", "日本語 テキスト")]
    public void CollapseWhitespaceIsInternationalAndIdempotent(string input, string expected)
    {
        var output = DataCleaningCommand.CollapseWhitespace(input);
        Assert.Equal(expected, output);
        Assert.Equal(output, DataCleaningCommand.CollapseWhitespace(output));
    }

    [Fact]
    public void RemoveNonprintingPreservesInternationalTextAndConfiguredLines()
    {
        const string input = "A\u0000\t世界\nB\u007f🙂";
        Assert.Equal("A\t世界\nB🙂", DataCleaningCommand.RemoveNonprinting(input, preserveTabsAndNewlines: true));
        Assert.Equal("A世界B🙂", DataCleaningCommand.RemoveNonprinting(input, preserveTabsAndNewlines: false));
    }

    [Fact]
    public void TextPlanChangesOnlyTextConstantsAndCategorizesMixedSelection()
    {
        var snapshot = Snapshot(Block(1, 5,
            FormulaCellValue.Text("  x  "), FormulaCellValue.Formula("=\" y \""), FormulaCellValue.Number(2),
            FormulaCellValue.Blank(), FormulaCellValue.Text("clean")));
        var command = Command("clean.text.trim_outer");

        var plan = command.PlanTrimOuter(snapshot);

        Assert.Equal("x", plan.After[0, 0].InvariantValue);
        Assert.Equal("=\" y \"", plan.After[0, 1].InvariantValue);
        Assert.Equal(2, plan.After[0, 2].AsNumber());
        Assert.True(plan.After[0, 3].IsBlank);
        Assert.Equal("clean", plan.After[0, 4].InvariantValue);
        Assert.Equal(1, plan.ChangedCount);
        Assert.Equal(4, plan.SkippedCount);
        Assert.True(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void ZeroToBlankMatchesNumericConstantsOnly()
    {
        var snapshot = Snapshot(Block(1, 6,
            FormulaCellValue.Number(0), FormulaCellValue.Number(-0.0), FormulaCellValue.Text("0"),
            FormulaCellValue.Formula("=0"), FormulaCellValue.Blank(), FormulaCellValue.Number(1)));

        var plan = Command("clean.display.zero_to_blank").PlanDisplayConversion(snapshot, DisplayValueConversion.ZeroToBlank);

        Assert.True(plan.After[0, 0].IsBlank);
        Assert.True(plan.After[0, 1].IsBlank);
        Assert.Equal("0", plan.After[0, 2].InvariantValue);
        Assert.Equal("=0", plan.After[0, 3].InvariantValue);
        Assert.True(plan.After[0, 4].IsBlank);
        Assert.Equal(1, plan.After[0, 5].AsNumber());
        Assert.Equal(2, plan.ChangedCount);
        Assert.Equal(4, plan.SkippedCount);
        Assert.True(plan.CommandPlan.RequiresPreview);
    }

    [Theory]
    [InlineData(DisplayValueConversion.NaTextToBlank, "N/A", true)]
    [InlineData(DisplayValueConversion.NaTextToBlank, "n/a", false)]
    [InlineData(DisplayValueConversion.NmTextToBlank, "NM", true)]
    [InlineData(DisplayValueConversion.NmTextToBlank, "N.M.", false)]
    [InlineData(DisplayValueConversion.DashTextToBlank, "-", true)]
    [InlineData(DisplayValueConversion.DashTextToBlank, "–", false)]
    public void TextToBlankConversionsUseExactCaseSensitiveSpellings(DisplayValueConversion conversion, string input, bool changes)
    {
        var plan = Command("clean.display.test").PlanDisplayConversion(
            Snapshot(Block(1, 1, FormulaCellValue.Text(input))), conversion);
        Assert.Equal(changes, plan.After[0, 0].IsBlank);
        Assert.Equal(changes ? 1 : 0, plan.ChangedCount);
    }

    [Fact]
    public void DataMutationUsesTheQualifiedTransactionalExecutorAndUndo()
    {
        var snapshot = Snapshot(Block(1, 2, FormulaCellValue.Text(" x "), FormulaCellValue.Formula("=A1")));
        var descriptor = DataCleaningCommandCatalog.GetRequired("clean.text.trim_outer");
        var plan = new DataCleaningCommand(descriptor).PlanTrimOuter(snapshot);
        var port = new FakePort(snapshot);
        var store = new SessionUndoStore();

        var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, store);

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal("x", port.Current[0, 0].InvariantValue);
        Assert.Equal("=A1", port.Current[0, 1].InvariantValue);
        Assert.True(store.TryUndo("book", port, DateTimeOffset.UtcNow).Succeeded);
        Assert.Equal(" x ", port.Current[0, 0].InvariantValue);
    }

    private static DataCleaningCommand Command(string id) =>
        new DataCleaningCommand(DataCleaningCommandCatalog.All.FindDescriptor(id));
    private static FormulaCellBlock Block(int rows, int columns, params FormulaCellValue[] values) => new FormulaCellBlock(rows, columns, values);
    private static FormulaBlockSnapshot Snapshot(FormulaCellBlock block) => new FormulaBlockSnapshot(
        new SelectionSnapshot(new SelectionContext("book", "Sheet1", "A1:Z99"), block.CellCount, null, "General", SelectionSafetyState.Safe()), 1, 1, block);

    private sealed class FakePort : IFormulaBlockPort
    {
        private readonly FormulaBlockSnapshot _snapshot;
        public FakePort(FormulaBlockSnapshot snapshot) { _snapshot = snapshot; Current = snapshot.Contents; }
        public FormulaCellBlock Current { get; private set; }
        public SelectionSnapshot CaptureSelection() => CaptureFormulaBlock().Selection;
        public FormulaBlockSnapshot CaptureFormulaBlock() => new FormulaBlockSnapshot(_snapshot.Selection, _snapshot.FirstRow, _snapshot.FirstColumn, Current);
        public FormulaBlockSnapshot CaptureFormulaBlock(SelectionContext target) { if (!target.Equals(_snapshot.Selection.Context)) throw new InvalidOperationException(); return CaptureFormulaBlock(); }
        public void WriteFormulaBlock(FormulaCellBlock contents) => Current = contents;
        public void WriteFormulaBlock(SelectionContext target, FormulaCellBlock contents) { if (!target.Equals(_snapshot.Selection.Context)) throw new InvalidOperationException(); Current = contents; }
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public bool TryRead(SelectionContext target, string propertyId, out string value) { value = Current.Serialize(); return target.Equals(_snapshot.Selection.Context); }
        public bool TryWrite(SelectionContext target, string propertyId, string value) { if (!target.Equals(_snapshot.Selection.Context)) return false; Current = FormulaCellBlock.Deserialize(value); return true; }
    }
}

internal static class DataCleaningTestExtensions
{
    public static CommandDescriptor FindDescriptor(this IReadOnlyList<CommandDescriptor> descriptors, string id)
    {
        foreach (var descriptor in descriptors) if (descriptor.Id == id) return descriptor;
        return new CommandDescriptor(id, 1, id, CommandImpact.Medium, new[] { "value" }, false, "test");
    }
}
