using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.SelectionTools;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class SelectionMatchCommandTests
{
    [Fact]
    public void PredicatesDistinguishFormulasConstantsTrueBlanksAndNumericHardcodes()
    {
        var source = Snapshot(20, 3, Block(2, 3,
            FormulaCellValue.Formula("=1"), FormulaCellValue.Number(7), FormulaCellValue.Text("7"),
            FormulaCellValue.Blank(), FormulaCellValue.Boolean(true), FormulaCellValue.Formula("=\"\"")));

        Assert.Equal(new[] { "C20", "E21" }, Plan(source, SelectionPredicate.Formulas).Areas.Select(value => value.Address));
        Assert.Equal(new[] { "D20:E20", "D21" }, Plan(source, SelectionPredicate.Constants).Areas.Select(value => value.Address));
        Assert.Equal(new[] { "C21" }, Plan(source, SelectionPredicate.Blanks).Areas.Select(value => value.Address));
        Assert.Equal(new[] { "D20" }, Plan(source, SelectionPredicate.NumericHardcodes).Areas.Select(value => value.Address));
    }

    [Fact]
    public void ExternalPredicateUsesParsedReferencesAndIgnoresBracketText()
    {
        var source = Snapshot(1, 1, Block(1, 3,
            FormulaCellValue.Formula("='[Book One.xlsx]Model'!A1"),
            FormulaCellValue.Formula("=\"[Book.xlsx]Model!A1\""),
            FormulaCellValue.Formula("=Sheet2!A1")));

        Assert.Equal(new[] { "A1" }, Plan(source, SelectionPredicate.ExternalFormulas).Areas.Select(value => value.Address));
    }

    [Fact]
    public void CompressionMergesIdenticalHorizontalRunsVertically()
    {
        var source = Snapshot(5, 2, Block(3, 4,
            FormulaCellValue.Blank(), FormulaCellValue.Number(1), FormulaCellValue.Number(2), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Number(3), FormulaCellValue.Number(4), FormulaCellValue.Blank(),
            FormulaCellValue.Number(5), FormulaCellValue.Blank(), FormulaCellValue.Blank(), FormulaCellValue.Number(6)));

        Assert.Equal(new[] { "C5:D6", "B7", "E7" }, Plan(source, SelectionPredicate.NumericHardcodes).Areas.Select(value => value.Address));
    }

    [Fact]
    public void FragmentedMatchSetFailsClosedAboveAreaLimit()
    {
        var cells = new List<FormulaCellValue>();
        for (var row = 0; row < 12; row++)
            for (var column = 0; column < 12; column++)
                cells.Add((row + column) % 2 == 0 ? FormulaCellValue.Number(1) : FormulaCellValue.Blank());

        var exception = Assert.Throws<CommandRefusedException>(() =>
            Plan(Snapshot(1, 1, new FormulaCellBlock(12, 12, cells)), SelectionPredicate.NumericHardcodes));

        Assert.Equal(RefusalCodes.ResourceLimit, exception.RefusalCode);
    }

    [Fact]
    public void ExecuteRevalidatesSourceThenVerifiesExactSelectionWithoutWriting()
    {
        var source = Snapshot(2, 2, Block(2, 2,
            FormulaCellValue.Number(1), FormulaCellValue.Blank(),
            FormulaCellValue.Number(2), FormulaCellValue.Text("x")));
        var plan = Plan(source, SelectionPredicate.NumericHardcodes);
        var port = new FakePort(source);

        var result = Command().Execute(plan, port);

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(new[] { "B2:B3" }, port.Selected);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void ExecuteRefusesStaleSourceBeforeChangingSelection()
    {
        var source = Snapshot(2, 2, Block(1, 2, FormulaCellValue.Number(1), FormulaCellValue.Blank()));
        var plan = Plan(source, SelectionPredicate.NumericHardcodes);
        var port = new FakePort(source) { Current = Block(1, 2, FormulaCellValue.Number(2), FormulaCellValue.Blank()) };

        var result = Command().Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Empty(port.Selected);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void ExecuteFailsWhenExcelReportsASelectionDifferentFromThePlan()
    {
        var source = Snapshot(2, 2, Block(1, 2, FormulaCellValue.Number(1), FormulaCellValue.Blank()));
        var plan = Plan(source, SelectionPredicate.NumericHardcodes);
        var port = new FakePort(source) { ObservedOverride = new[] { "Z99" } };

        var result = Command().Execute(plan, port);

        Assert.False(result.Succeeded);
        Assert.Equal("SELECTION_POSTCONDITION_MISMATCH", result.DiagnosticId);
        Assert.Equal(0, port.WriteCount);
    }

    private static SelectionMatchCommand Command() =>
        new SelectionMatchCommand(SelectionCommandCatalog.GetRequired("selection.select.numeric_hardcodes"));

    private static SelectionMatchPlan Plan(FormulaBlockSnapshot source, SelectionPredicate predicate) => Command().Plan(source, predicate);

    private static FormulaCellBlock Block(int rows, int columns, params FormulaCellValue[] values) => new FormulaCellBlock(rows, columns, values);

    private static FormulaBlockSnapshot Snapshot(int firstRow, int firstColumn, FormulaCellBlock block) => new FormulaBlockSnapshot(
        new SelectionSnapshot(new SelectionContext("book", "Sheet1", "A1:Z99"), block.CellCount, null, "General", SelectionSafetyState.Safe()),
        firstRow, firstColumn, block);

    private sealed class FakePort : ISelectionMatchPort
    {
        private readonly FormulaBlockSnapshot _source;
        public FakePort(FormulaBlockSnapshot source) { _source = source; Current = source.Contents; }
        public FormulaCellBlock Current { get; set; }
        public List<string> Selected { get; } = new List<string>();
        public IReadOnlyList<string>? ObservedOverride { get; set; }
        public int WriteCount { get; private set; }
        public FormulaBlockSnapshot CaptureFormulaBlock() => CaptureFormulaBlock(_source.Selection.Context);
        public FormulaBlockSnapshot CaptureFormulaBlock(SelectionContext target)
        {
            if (!target.Equals(_source.Selection.Context)) throw new InvalidOperationException();
            return new FormulaBlockSnapshot(_source.Selection, _source.FirstRow, _source.FirstColumn, Current);
        }
        public SelectionSnapshot CaptureSelection() => _source.Selection;
        public void SelectAreas(SelectionContext sourceContext, IReadOnlyList<SelectionArea> areas)
        {
            if (!sourceContext.Equals(_source.Selection.Context)) throw new InvalidOperationException();
            Selected.Clear();
            Selected.AddRange(areas.Select(value => value.Address));
        }
        public IReadOnlyList<string> CaptureSelectedAreaAddresses() => ObservedOverride ?? Selected.AsReadOnly();
        public void WriteFormulaBlock(FormulaCellBlock contents) { WriteCount++; Current = contents; }
        public void WriteFormulaBlock(SelectionContext target, FormulaCellBlock contents) { WriteCount++; Current = contents; }
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public bool TryRead(SelectionContext target, string propertyId, out string value) { value = string.Empty; return false; }
        public bool TryWrite(SelectionContext target, string propertyId, string value) => false;
    }
}
