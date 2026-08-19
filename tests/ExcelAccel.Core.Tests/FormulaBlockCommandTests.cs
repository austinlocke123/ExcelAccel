using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormulaBlockCommandTests
{
    [Fact]
    public void CellBlockSerializationIsBoundedExactAndRoundTrips()
    {
        var block = Block(2, 3,
            FormulaCellValue.Formula("=A1+1"), FormulaCellValue.Text("hello|:世界"), FormulaCellValue.Number(12.25),
            FormulaCellValue.Boolean(true), FormulaCellValue.Blank(), FormulaCellValue.Number(-0.0));

        var restored = FormulaCellBlock.Deserialize(block.Serialize());

        Assert.True(block.ContentEquals(restored));
        Assert.Equal(block.Fingerprint, restored.Fingerprint);
        Assert.Throws<FormatException>(() => FormulaCellBlock.Deserialize("FCB1|1|1|1:not-base64!|"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaCellBlock(101, 100, Enumerable.Repeat(FormulaCellValue.Blank(), 10100)));
    }

    [Fact]
    public void SmartCopyDownTranslatesFromTopEdgeWithoutChangingTheSource()
    {
        var snapshot = Snapshot(Block(3, 2,
            FormulaCellValue.Formula("=A1+$B1"), FormulaCellValue.Formula("=C$1+$D$1"),
            FormulaCellValue.Blank(), FormulaCellValue.Blank(),
            FormulaCellValue.Blank(), FormulaCellValue.Blank()));
        var command = Command("formula.copy.down");

        var plan = command.PlanCopy(snapshot, FormulaCopyDirection.Down);

        Assert.Equal("=A1+$B1", plan.After[0, 0].InvariantValue);
        Assert.Equal("=A2+$B2", plan.After[1, 0].InvariantValue);
        Assert.Equal("=A3+$B3", plan.After[2, 0].InvariantValue);
        Assert.Equal("=C$1+$D$1", plan.After[2, 1].InvariantValue);
        Assert.Equal(4, plan.ChangedCount);
        Assert.False(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void SmartCopyRequiresPreviewForAnyNonblankOverwrite()
    {
        var snapshot = Snapshot(Block(2, 1,
            FormulaCellValue.Formula("=A1"), FormulaCellValue.Text("existing")));

        var plan = Command("formula.copy.down").PlanCopy(snapshot, FormulaCopyDirection.Down);

        Assert.True(plan.CommandPlan.RequiresPreview);
        Assert.Single(plan.Samples);
        Assert.Contains("existing", plan.Samples[0]);
    }

    [Theory]
    [InlineData("=Table1[Amount]", FormulaRefusalCodes.StructuredReferenceInspectOnly)]
    [InlineData("=NamedRate*A1", FormulaRefusalCodes.NameInspectOnly)]
    [InlineData("=A1#", FormulaRefusalCodes.DynamicArrayInspectOnly)]
    public void OneUnsupportedCopySourceRefusesTheWholePlan(string formula, string expectedCode)
    {
        var snapshot = Snapshot(Block(2, 1, FormulaCellValue.Formula(formula), FormulaCellValue.Blank()));

        var exception = Assert.Throws<CommandRefusedException>(() =>
            Command("formula.copy.down").PlanCopy(snapshot, FormulaCopyDirection.Down));

        Assert.Equal(expectedCode, exception.RefusalCode);
        Assert.Equal(formula, snapshot.Contents[0, 0].InvariantValue);
        Assert.True(snapshot.Contents[1, 0].IsBlank);
    }

    [Fact]
    public void IfErrorPlanSkipsConstantsAndNeverPartiallyPlansUnsupportedFormula()
    {
        var valid = Snapshot(Block(1, 3,
            FormulaCellValue.Formula("=A1+B1"), FormulaCellValue.Number(3), FormulaCellValue.Formula("=IFERROR(C1,0)")));
        var command = Command("formula.iferror.toggle");

        var plan = command.PlanIfError(valid, "0");

        Assert.Equal("=IFERROR(A1+B1,0)", plan.After[0, 0].InvariantValue);
        Assert.Equal(FormulaCellKind.Number, plan.After[0, 1].Kind);
        Assert.Equal("=C1", plan.After[0, 2].InvariantValue);
        Assert.Equal(2, plan.ChangedCount);
        Assert.Equal(1, plan.SkippedCount);

        var unsupported = Snapshot(Block(1, 2,
            FormulaCellValue.Formula("=A1"), FormulaCellValue.Formula("=Table1[Amount]")));
        Assert.Throws<CommandRefusedException>(() => command.PlanIfError(unsupported, "0"));
        Assert.Equal("=A1", unsupported.Contents[0, 0].InvariantValue);
    }

    [Fact]
    public void SignAndUnitsRequireExplicitConstantInclusion()
    {
        var snapshot = Snapshot(Block(1, 3,
            FormulaCellValue.Formula("=A1+B1"), FormulaCellValue.Number(2500), FormulaCellValue.Text("2500")));

        var sign = Command("formula.sign.reverse").PlanReverseSign(snapshot, includeNumericConstants: true);
        var units = Command("formula.units.to_thousands").PlanScale(snapshot, 1000, divide: true, includeNumericConstants: true);
        var formulasOnly = Command("formula.units.to_thousands").PlanScale(snapshot, 1000, divide: true, includeNumericConstants: false);

        Assert.Equal("=-(A1+B1)", sign.After[0, 0].InvariantValue);
        Assert.Equal(-2500, sign.After[0, 1].AsNumber());
        Assert.Equal("=(A1+B1)/1000", units.After[0, 0].InvariantValue);
        Assert.Equal(2.5, units.After[0, 1].AsNumber());
        Assert.Equal(2500, formulasOnly.After[0, 1].AsNumber());
        Assert.Equal(2, formulasOnly.SkippedCount);
        Assert.Equal(new[] { "formula", "value" }, units.CommandPlan.ChangedProperties);
    }

    [Fact]
    public void ExecuteRevalidatesWritesVerifiesAndCreatesOptimisticUndo()
    {
        var before = Block(2, 1, FormulaCellValue.Formula("=A1"), FormulaCellValue.Blank());
        var port = new FakeFormulaPort(Snapshot(before));
        var command = Command("formula.copy.down");
        var plan = command.PlanCopy(port.CaptureFormulaBlock(), FormulaCopyDirection.Down);
        var store = new SessionUndoStore();

        var result = command.Execute(plan, port, null, store);

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal("=A2", port.Current[1, 0].InvariantValue);
        Assert.Equal(1, store.Count("book"));
        var undo = store.TryUndo("book", port, DateTimeOffset.UtcNow);
        Assert.True(undo.Succeeded, undo.Message);
        Assert.True(port.Current[1, 0].IsBlank);
    }

    [Fact]
    public void ExecuteRefusesStaleOrUnconfirmedPlanBeforeAnyWrite()
    {
        var before = Block(2, 1, FormulaCellValue.Formula("=A1"), FormulaCellValue.Text("overwrite"));
        var port = new FakeFormulaPort(Snapshot(before));
        var command = Command("formula.copy.down");
        var plan = command.PlanCopy(port.CaptureFormulaBlock(), FormulaCopyDirection.Down);
        var store = new SessionUndoStore();

        var unconfirmed = command.Execute(plan, port, null, store);
        Assert.Equal(CommandResultStatus.Refused, unconfirmed.Status);
        Assert.Equal(RefusalCodes.PreviewRequired, unconfirmed.RefusalCode);
        Assert.Equal(0, port.WriteCount);

        port.Current = Block(2, 1, FormulaCellValue.Formula("=B1"), FormulaCellValue.Text("overwrite"));
        var stale = command.Execute(plan, port, plan.CommandPlan.PlanHash, store);
        Assert.Equal(CommandResultStatus.Refused, stale.Status);
        Assert.Equal(RefusalCodes.StaleContext, stale.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void FailedWriteRestoresAndVerifiesTheCompleteBeforeState()
    {
        var before = Block(2, 1, FormulaCellValue.Formula("=A1"), FormulaCellValue.Blank());
        var port = new FakeFormulaPort(Snapshot(before)) { MutateThenThrowOnce = true };
        var command = Command("formula.copy.down");
        var plan = command.PlanCopy(port.CaptureFormulaBlock(), FormulaCopyDirection.Down);

        var result = command.Execute(plan, port, null, new SessionUndoStore());

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.True(before.ContentEquals(port.Current));
        Assert.Equal(2, port.WriteCount);
    }

    [Fact]
    public void MissingReceiptStoreRefusesBeforeMutation()
    {
        var port = new FakeFormulaPort(Snapshot(Block(2, 1, FormulaCellValue.Formula("=A1"), FormulaCellValue.Blank())));
        var command = Command("formula.copy.down");
        var plan = command.PlanCopy(port.CaptureFormulaBlock(), FormulaCopyDirection.Down);

        var result = command.Execute(plan, port, null, null);

        Assert.Equal(CommandResultStatus.Refused, result.Status);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void FormulaReceiptPostStateComparisonIsCaseSensitive()
    {
        var before = Block(1, 1, FormulaCellValue.Formula("=A1"));
        var after = Block(1, 1, FormulaCellValue.Formula("=a1"));
        var port = new FakeFormulaPort(Snapshot(after));
        var store = new SessionUndoStore();
        var now = DateTimeOffset.UtcNow;
        store.Add(new PropertyReceipt("receipt", "formula.test", 1, port.CaptureSelection().Context,
            FormulaBlockCommand.ReceiptPropertyId, before.Serialize(), after.Serialize().ToUpperInvariant(),
            "plan", now, now.AddHours(1)));

        var result = store.TryUndo("book", port, now);

        Assert.Equal(UndoOutcome.Stale, result.Outcome);
        Assert.True(after.ContentEquals(port.Current));
    }

    private static FormulaBlockCommand Command(string id) => new FormulaBlockCommand(FormulaCommandCatalog.GetRequired(id));

    private static FormulaCellBlock Block(int rows, int columns, params FormulaCellValue[] cells) =>
        new FormulaCellBlock(rows, columns, cells);

    private static FormulaBlockSnapshot Snapshot(FormulaCellBlock block) =>
        new FormulaBlockSnapshot(
            new SelectionSnapshot(new SelectionContext("book", "Sheet1", "$A$1:$Z$99"), block.CellCount, null, "General", SelectionSafetyState.Safe()),
            1, 1, block);

    private sealed class FakeFormulaPort : IFormulaBlockPort
    {
        private readonly FormulaBlockSnapshot _template;
        public FakeFormulaPort(FormulaBlockSnapshot snapshot) { _template = snapshot; Current = snapshot.Contents; }
        public FormulaCellBlock Current { get; set; }
        public int WriteCount { get; private set; }
        public bool MutateThenThrowOnce { get; set; }
        public SelectionSnapshot CaptureSelection() => CaptureFormulaBlock().Selection;
        public FormulaBlockSnapshot CaptureFormulaBlock() => new FormulaBlockSnapshot(_template.Selection, _template.FirstRow, _template.FirstColumn, Current);
        public void WriteFormulaBlock(FormulaCellBlock contents)
        {
            WriteCount++;
            Current = contents;
            if (MutateThenThrowOnce) { MutateThenThrowOnce = false; throw new InvalidOperationException("Injected partial write failure."); }
        }
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public bool TryRead(SelectionContext target, string propertyId, out string value)
        {
            value = string.Empty;
            if (!target.Equals(_template.Selection.Context) || propertyId != FormulaBlockCommand.ReceiptPropertyId) return false;
            value = Current.Serialize();
            return true;
        }
        public bool TryWrite(SelectionContext target, string propertyId, string value)
        {
            if (!target.Equals(_template.Selection.Context) || propertyId != FormulaBlockCommand.ReceiptPropertyId) return false;
            try { WriteFormulaBlock(FormulaCellBlock.Deserialize(value)); return true; }
            catch { return false; }
        }
    }
}
