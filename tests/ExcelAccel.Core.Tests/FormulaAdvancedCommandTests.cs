using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormulaAdvancedCommandTests
{
    [Fact]
    public void RowSpacingChangesOnlyExactIntervalDestinations()
    {
        var snapshot = Snapshot(Block(5, 1,
            FormulaCellValue.Formula("=B10"), FormulaCellValue.Blank(), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Blank()), 10, 1);
        var command = Command("formula.spacing.rows", new[] { "formula" });

        var plan = command.PlanSpacing(snapshot, FormulaSpacingDirection.Rows, 2);

        Assert.Equal("=B10", plan.After[0, 0].InvariantValue);
        Assert.True(plan.After[1, 0].IsBlank);
        Assert.Equal("=B12", plan.After[2, 0].InvariantValue);
        Assert.True(plan.After[3, 0].IsBlank);
        Assert.Equal("=B14", plan.After[4, 0].InvariantValue);
        Assert.Equal(2, plan.ChangedCount);
    }

    [Fact]
    public void TransposeMapsPositionsConstantsAndReferenceAxes()
    {
        var source = Snapshot(Block(2, 3,
            FormulaCellValue.Formula("=C10"), FormulaCellValue.Number(2), FormulaCellValue.Text("x"),
            FormulaCellValue.Boolean(true), FormulaCellValue.Formula("=$D11"), FormulaCellValue.Blank()), 10, 2, "$B$10:$D$11");
        var destination = Snapshot(Block(3, 2,
            FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank()), 20, 5, "$E$20:$F$22");
        var command = Command("formula.transpose", new[] { "formula", "value" }, ChangedPropertyPolicy.DeclaredSubset);

        var plan = command.PlanTranspose(source, destination);

        Assert.Equal("=E21", plan.After[0, 0].InvariantValue);
        Assert.Equal(FormulaCellKind.Boolean, plan.After[0, 1].Kind);
        Assert.Equal(2, plan.After[1, 0].AsNumber());
        Assert.Equal("=F$4", plan.After[1, 1].InvariantValue);
        Assert.Equal("x", plan.After[2, 0].InvariantValue);
        Assert.True(plan.After[2, 1].IsBlank);
        Assert.True(plan.CommandPlan.RequiresPreview);
        Assert.True(plan.RequiresExternalSourceRevalidation);
    }

    [Fact]
    public void ExternalSourcePlanRevalidatesThenExecutesWithoutChangingSelection()
    {
        var source = Snapshot(Block(1, 1, FormulaCellValue.Formula("=A1")), 1, 1, "$A$1");
        var destination = Snapshot(Block(1, 1, FormulaCellValue.Blank()), 3, 3, "$C$3");
        var descriptor = Descriptor("formula.transpose", new[] { "formula", "value" }, ChangedPropertyPolicy.DeclaredSubset);
        var advanced = new FormulaAdvancedCommand(descriptor);
        var plan = advanced.PlanTranspose(source, destination);
        var port = new FakePort(destination, source);

        var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal("=C3", port.Current[0, 0].InvariantValue);
        Assert.Equal(1, port.WriteCount);
        Assert.Equal(destination.Selection.Context, port.CaptureSelection().Context);
    }

    [Fact]
    public void ChangedExternalSourceRefusesBeforeDestinationWrite()
    {
        var source = Snapshot(Block(1, 1, FormulaCellValue.Formula("=A1")), 1, 1, "$A$1");
        var destination = Snapshot(Block(1, 1, FormulaCellValue.Blank()), 3, 3, "$C$3");
        var descriptor = Descriptor("formula.transpose", new[] { "formula", "value" }, ChangedPropertyPolicy.DeclaredSubset);
        var plan = new FormulaAdvancedCommand(descriptor).PlanTranspose(source, destination);
        var changedSource = Snapshot(Block(1, 1, FormulaCellValue.Formula("=B1")), 1, 1, "$A$1");
        var port = new FakePort(destination, changedSource);

        var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.Equal(CommandResultStatus.Refused, result.Status);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void FormulaFromAboveUsesEveryColumnSourceAndExactRowDisplacement()
    {
        var source = Snapshot(Block(1, 2, FormulaCellValue.Formula("=A4+$B$1"), FormulaCellValue.Formula("=C4")), 4, 1, "$A$4:$B$4");
        var destination = Snapshot(Block(2, 2, FormulaCellValue.Blank(), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Blank()), 5, 1, "$A$5:$B$6");

        var plan = Command("fill.formula_from_above", new[] { "formula" }).PlanFormulaFromAbove(source, destination);

        Assert.Equal("=A5+$B$1", plan.After[0, 0].InvariantValue);
        Assert.Equal("=C5", plan.After[0, 1].InvariantValue);
        Assert.Equal("=A6+$B$1", plan.After[1, 0].InvariantValue);
        Assert.Equal("=C6", plan.After[1, 1].InvariantValue);
        Assert.True(plan.RequiresExternalSourceRevalidation);
    }

    [Theory]
    [InlineData(SequenceFillDirection.Right, "1,3,5,7,9,11")]
    [InlineData(SequenceFillDirection.Down, "1,5,9,3,7,11")]
    public void NumericFillUsesExplicitDirectionAndNoNeighborInference(
        SequenceFillDirection direction, string expectedText)
    {
        var destination = Snapshot(Block(2, 3,
            FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank()), 1, 1);

        var plan = Command("fill.numeric_sequence", new[] { "value" }).PlanNumericSequence(destination, 1, 2, direction);

        var expected = Array.ConvertAll(expectedText.Split(','), double.Parse);
        for (var index = 0; index < expected.Length; index++)
            Assert.Equal(expected[index], plan.After.Cells[index].AsNumber());
        Assert.False(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void DateSerialPolicyHandlesExcelLeapBugAnd1904EpochExplicitly()
    {
        Assert.Equal(59, FormulaAdvancedCommand.ToExcelSerial(new DateTime(1900, 2, 28), ExcelDateSystem.Excel1900));
        Assert.Equal(61, FormulaAdvancedCommand.ToExcelSerial(new DateTime(1900, 3, 1), ExcelDateSystem.Excel1900));
        Assert.Equal(0, FormulaAdvancedCommand.ToExcelSerial(new DateTime(1904, 1, 1), ExcelDateSystem.Excel1904));

        var destination = Snapshot(Block(2, 1, FormulaCellValue.Blank(), FormulaCellValue.Blank()), 1, 1);
        var command = Command("fill.date_sequence", new[] { "value" });
        var plan = command.PlanDateSequence(destination, new DateTime(2026, 8, 19), 7, SequenceFillDirection.Down, ExcelDateSystem.Excel1900);
        Assert.Equal(7, plan.After[1, 0].AsNumber() - plan.After[0, 0].AsNumber());
        Assert.Equal("2026-08-19", plan.CommandPlan.Arguments["start_date"]);
        Assert.Throws<CommandRefusedException>(() => command.PlanDateSequence(destination,
            new DateTime(2026, 8, 19, 12, 0, 0), 1, SequenceFillDirection.Down, ExcelDateSystem.Excel1900));
    }

    [Fact]
    public void FormulaSourceStoreExpiresAndClearsWithoutWorkbookAccess()
    {
        var store = new FormulaSourceStore();
        var snapshot = Snapshot(Block(1, 1, FormulaCellValue.Formula("=A1")), 1, 1, "$A$1");
        var now = new DateTimeOffset(2026, 8, 19, 12, 0, 0, TimeSpan.Zero);
        store.Capture(snapshot, now);

        Assert.True(store.TryGet(now.AddMinutes(29), TimeSpan.FromMinutes(30), out var captured, out _));
        Assert.Same(snapshot, captured);
        Assert.False(store.TryGet(now.AddMinutes(31), TimeSpan.FromMinutes(30), out _, out var expired));
        Assert.Contains("expired", expired, StringComparison.OrdinalIgnoreCase);
        Assert.False(store.TryGet(now.AddMinutes(31), TimeSpan.FromMinutes(30), out _, out var missing));
        Assert.Contains("No formula source", missing, StringComparison.Ordinal);
    }

    [Fact]
    public void FormulasOnlyPasteUsesWholeMultipleRepetitionAndSkipsSourceConstants()
    {
        var source = Snapshot(Block(1, 2, FormulaCellValue.Formula("=A1+$B$1"), FormulaCellValue.Number(9)), 1, 1, "$A$1:$B$1");
        var destination = Snapshot(Block(2, 4,
            FormulaCellValue.Blank(), FormulaCellValue.Text("keep"), FormulaCellValue.Blank(), FormulaCellValue.Text("keep"),
            FormulaCellValue.Blank(), FormulaCellValue.Text("keep"), FormulaCellValue.Blank(), FormulaCellValue.Text("keep")),
            5, 4, "$D$5:$G$6");
        var descriptor = Descriptor("paste.formulas_only", new[] { "formula" }, ChangedPropertyPolicy.Exact);
        var plan = new FormulaAdvancedCommand(descriptor).PlanPasteFormulas(source, destination);

        Assert.Equal("=D5+$B$1", plan.After[0, 0].InvariantValue);
        Assert.Equal("keep", plan.After[0, 1].InvariantValue);
        Assert.Equal("=F5+$B$1", plan.After[0, 2].InvariantValue);
        Assert.Equal("=D6+$B$1", plan.After[1, 0].InvariantValue);
        Assert.Equal("=F6+$B$1", plan.After[1, 2].InvariantValue);
        Assert.Equal(4, plan.ChangedCount);
        Assert.Equal(4, plan.SkippedCount);
        Assert.True(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void FormulasOnlyPasteRefusesShapeMismatchAndOverlapBeforeMutation()
    {
        var source = Snapshot(Block(1, 2, FormulaCellValue.Formula("=A1"), FormulaCellValue.Formula("=B1")), 1, 1, "$A$1:$B$1");
        var mismatch = Snapshot(Block(1, 3, FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank()), 5, 1, "$A$5:$C$5");
        var overlap = Snapshot(Block(2, 2, FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank()), 1, 2, "$B$1:$C$2");
        var command = Command("paste.formulas_only", new[] { "formula" });

        Assert.Throws<CommandRefusedException>(() => command.PlanPasteFormulas(source, mismatch));
        Assert.Throws<CommandRefusedException>(() => command.PlanPasteFormulas(source, overlap));
    }

    [Fact]
    public void ValuesOnlyPasteUsesUnderlyingFormulaResultsAndReportsFormulaReplacement()
    {
        var source = Snapshot(Block(1, 2, FormulaCellValue.Formula("=1+1"), FormulaCellValue.Text("x")), 1, 1, "$A$1:$B$1",
            Values(1, 2, UnderlyingCellValue.Number(2), UnderlyingCellValue.Text("x")));
        var destination = Snapshot(Block(2, 2,
            FormulaCellValue.Formula("=99"), FormulaCellValue.Blank(),
            FormulaCellValue.Number(8), FormulaCellValue.Text("old")), 5, 4, "$D$5:$E$6");

        var plan = Command("paste.values_only", new[] { "value" }).PlanPasteValues(source, destination);

        Assert.Equal(2, plan.After[0, 0].AsNumber());
        Assert.Equal("x", plan.After[0, 1].InvariantValue);
        Assert.Equal(2, plan.After[1, 0].AsNumber());
        Assert.Equal("x", plan.After[1, 1].InvariantValue);
        Assert.DoesNotContain(plan.After.Cells, value => value.IsFormula);
        Assert.Equal("2", plan.CommandPlan.Arguments["source_formula_mappings"]);
        Assert.Equal("1", plan.CommandPlan.Arguments["destination_formula_replacements"]);
        Assert.True(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void ValuesOnlyPasteRevalidatesCalculatedSourceValuesBeforeWriting()
    {
        var source = Snapshot(Block(1, 1, FormulaCellValue.Formula("=A1")), 1, 1, "$A$1",
            Values(1, 1, UnderlyingCellValue.Number(2)));
        var changedSource = Snapshot(source.Contents, 1, 1, "$A$1", Values(1, 1, UnderlyingCellValue.Number(3)));
        var destination = Snapshot(Block(1, 1, FormulaCellValue.Blank()), 5, 4, "$D$5");
        var descriptor = Descriptor("paste.values_only", new[] { "value" }, ChangedPropertyPolicy.Exact);
        var plan = new FormulaAdvancedCommand(descriptor).PlanPasteValues(source, destination);
        var port = new FakePort(destination, changedSource);

        var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.Equal(CommandResultStatus.Refused, result.Status);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void ValueFromAboveUsesEachUnderlyingValueAndRefusesErrors()
    {
        var source = Snapshot(Block(1, 2, FormulaCellValue.Formula("=1+1"), FormulaCellValue.Text("x")), 4, 1, "$A$4:$B$4",
            Values(1, 2, UnderlyingCellValue.Number(2), UnderlyingCellValue.Text("x")));
        var destination = Snapshot(Block(2, 2,
            FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Blank()), 5, 1, "$A$5:$B$6");

        var plan = Command("fill.value_from_above", new[] { "value" }).PlanValueFromAbove(source, destination);

        Assert.Equal(2, plan.After[0, 0].AsNumber());
        Assert.Equal("x", plan.After[0, 1].InvariantValue);
        Assert.Equal(2, plan.After[1, 0].AsNumber());
        Assert.Equal("x", plan.After[1, 1].InvariantValue);
        Assert.True(plan.CommandPlan.RequiresPreview);

        var errorSource = Snapshot(Block(1, 1, FormulaCellValue.Formula("=1/0")), 4, 1, "$A$4",
            Values(1, 1, UnderlyingCellValue.Error(-2146826281)));
        var oneCell = Snapshot(Block(1, 1, FormulaCellValue.Blank()), 5, 1, "$A$5");
        Assert.Throws<CommandRefusedException>(() =>
            Command("fill.value_from_above", new[] { "value" }).PlanValueFromAbove(errorSource, oneCell));
    }

    private static FormulaAdvancedCommand Command(string id, IEnumerable<string> properties,
        ChangedPropertyPolicy policy = ChangedPropertyPolicy.Exact) => new FormulaAdvancedCommand(Descriptor(id, properties, policy));

    private static CommandDescriptor Descriptor(string id, IEnumerable<string> properties, ChangedPropertyPolicy policy) =>
        new CommandDescriptor(id, 1, id, CommandImpact.Medium, properties, true, "test", "CAP-FORM-001",
            CommandContextRequirement.Selection, PreviewPolicy.Threshold, UndoPolicy.SessionPropertyReceipt,
            changedPropertyPolicy: policy);

    private static FormulaCellBlock Block(int rows, int columns, params FormulaCellValue[] values) => new FormulaCellBlock(rows, columns, values);
    private static UnderlyingValueBlock Values(int rows, int columns, params UnderlyingCellValue[] values) => new UnderlyingValueBlock(rows, columns, values);
    private static FormulaBlockSnapshot Snapshot(FormulaCellBlock block, int firstRow, int firstColumn, string address = "$A$1:$Z$99",
        UnderlyingValueBlock? underlyingValues = null) =>
        new FormulaBlockSnapshot(new SelectionSnapshot(new SelectionContext("book", "Sheet1", address), block.CellCount, null, "General", SelectionSafetyState.Safe()), firstRow, firstColumn, block, underlyingValues);

    private sealed class FakePort : IFormulaBlockPort
    {
        private readonly FormulaBlockSnapshot _snapshot;
        private readonly FormulaBlockSnapshot? _source;
        public FakePort(FormulaBlockSnapshot snapshot, FormulaBlockSnapshot? source = null) { _snapshot = snapshot; _source = source; Current = snapshot.Contents; }
        public FormulaCellBlock Current { get; private set; }
        public int WriteCount { get; private set; }
        public SelectionSnapshot CaptureSelection() => CaptureFormulaBlock().Selection;
        public FormulaBlockSnapshot CaptureFormulaBlock() => new FormulaBlockSnapshot(_snapshot.Selection, _snapshot.FirstRow, _snapshot.FirstColumn, Current, _snapshot.UnderlyingValues);
        public FormulaBlockSnapshot CaptureFormulaBlock(SelectionContext target)
        {
            if (target.Equals(_snapshot.Selection.Context)) return CaptureFormulaBlock();
            if (_source is not null && target.Equals(_source.Selection.Context)) return _source;
            throw new InvalidOperationException("Unexpected range target.");
        }
        public void WriteFormulaBlock(FormulaCellBlock contents) { WriteCount++; Current = contents; }
        public void WriteFormulaBlock(SelectionContext target, FormulaCellBlock contents)
        {
            if (!target.Equals(_snapshot.Selection.Context)) throw new InvalidOperationException("Unexpected write target.");
            WriteFormulaBlock(contents);
        }
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public bool TryRead(SelectionContext target, string propertyId, out string value) { value = Current.Serialize(); return target.Equals(_snapshot.Selection.Context); }
        public bool TryWrite(SelectionContext target, string propertyId, string value) { if (!target.Equals(_snapshot.Selection.Context)) return false; Current = FormulaCellBlock.Deserialize(value); return true; }
    }
}
