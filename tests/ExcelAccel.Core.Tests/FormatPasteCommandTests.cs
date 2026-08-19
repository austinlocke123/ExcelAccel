using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormatPasteCommandTests
{
    [Fact]
    public void FormatBlockRoundTripsEveryApprovedPropertyExactly()
    {
        var block = new FormatBlock(1, 2, new[]
        {
            Format("#,##0.00", "Aptos", 11, true, false, "single", "right", "bottom", 1),
            Format("yyyy-mm-dd", "Arial", 12.5, false, true, "double", "center", "top", 0),
        });

        var restored = FormatBlock.Deserialize(block.Serialize());

        Assert.True(block.ContentEquals(restored));
        Assert.Equal(block.Fingerprint, restored.Fingerprint);
    }

    [Fact]
    public void PlanUsesExactWholeMultipleRepetitionAndApprovedPropertiesOnly()
    {
        var left = Format("0.0", "Aptos", 11, true, false, "none", "right", "bottom", 0);
        var right = Format("0%", "Arial", 10, false, true, "single", "center", "top", 2);
        var source = Snapshot("A1:B1", 1, 1, new FormatBlock(1, 2, new[] { left, right }));
        var destination = Snapshot("D5:G6", 5, 4, Uniform(2, 4, Format("General", "Calibri", 11, false, false, "none", "general", "bottom", 0)));

        var plan = Command().Plan(source, destination);

        Assert.Same(left, plan.After[0, 0]);
        Assert.Same(right, plan.After[0, 1]);
        Assert.Same(left, plan.After[1, 2]);
        Assert.Same(right, plan.After[1, 3]);
        Assert.Equal(8, plan.ChangedCells);
        Assert.Equal(new[] { FormatPasteCommand.ReceiptPropertyId }, plan.CommandPlan.ChangedProperties);
        Assert.True(plan.CommandPlan.RequiresPreview);
    }

    [Fact]
    public void PlanRefusesOverlapAndNonMultipleShape()
    {
        var value = Format("General", "Aptos", 11, false, false, "none", "general", "bottom", 0);
        var source = Snapshot("A1:B1", 1, 1, Uniform(1, 2, value));
        var overlap = Snapshot("B1:C1", 1, 2, Uniform(1, 2, value));
        var mismatch = Snapshot("D1:F1", 1, 4, Uniform(1, 3, value));

        Assert.Throws<CommandRefusedException>(() => Command().Plan(source, overlap));
        Assert.Throws<CommandRefusedException>(() => Command().Plan(source, mismatch));
    }

    [Fact]
    public void ExecuteRevalidatesWritesVerifiesAndSupportsExactUndo()
    {
        var sourceValue = Format("0.00", "Aptos", 12, true, true, "single", "right", "center", 1);
        var beforeValue = Format("General", "Calibri", 11, false, false, "none", "general", "bottom", 0);
        var source = Snapshot("A1", 1, 1, Uniform(1, 1, sourceValue));
        var destination = Snapshot("D5:E5", 5, 4, Uniform(1, 2, beforeValue));
        var port = new FakePort(source, destination);
        var store = new SessionUndoStore();
        var command = Command();
        var plan = command.Plan(source, destination);

        var result = command.Execute(plan, port, plan.CommandPlan.PlanHash, store);

        Assert.True(result.Succeeded, result.Message);
        Assert.True(plan.After.ContentEquals(port.Current));
        Assert.Equal(1, port.WriteCount);
        Assert.True(store.TryUndo("book", port, DateTimeOffset.UtcNow).Succeeded);
        Assert.True(destination.Contents.ContentEquals(port.Current));
    }

    [Fact]
    public void ChangedSourceRefusesBeforeDestinationWrite()
    {
        var original = Format("0.00", "Aptos", 11, false, false, "none", "right", "bottom", 0);
        var changed = Format("0%", "Aptos", 11, false, false, "none", "right", "bottom", 0);
        var destinationValue = Format("General", "Calibri", 11, false, false, "none", "general", "bottom", 0);
        var source = Snapshot("A1", 1, 1, Uniform(1, 1, original));
        var destination = Snapshot("D5", 5, 4, Uniform(1, 1, destinationValue));
        var plan = Command().Plan(source, destination);
        var port = new FakePort(Snapshot("A1", 1, 1, Uniform(1, 1, changed)), destination);

        var result = Command().Execute(plan, port, plan.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.Equal(CommandResultStatus.Refused, result.Status);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void ReceiptStoreFailureRollsBackVerifiedFormatting()
    {
        var sourceValue = Format("0.00", "Arial", 12, true, false, "single", "right", "center", 1);
        var beforeValue = Format("General", "Aptos", 11, false, false, "none", "left", "bottom", 0);
        var source = Snapshot("A1", 1, 1, Uniform(1, 1, sourceValue));
        var destination = Snapshot("D5", 5, 4, Uniform(1, 1, beforeValue));
        var port = new FakePort(source, destination);
        var command = Command();
        var plan = command.Plan(source, destination);

        var result = command.Execute(plan, port, plan.CommandPlan.PlanHash, new ThrowingSink());

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.Equal("RECEIPT_STORE_ROLLED_BACK", result.DiagnosticId);
        Assert.True(destination.Contents.ContentEquals(port.Current));
    }

    private static FormatPasteCommand Command() => new FormatPasteCommand(new CommandDescriptor(
        "paste.formats_only", 1, "Paste Formats Only", CommandImpact.Medium,
        new[] { FormatPasteCommand.ReceiptPropertyId }, true, "test", "CAP-FORM-001",
        CommandContextRequirement.Selection, PreviewPolicy.Mandatory, UndoPolicy.SessionPropertyReceipt));
    private static CellFormatValue Format(string numberFormat, string fontName, double size, bool bold, bool italic,
        string underline, string horizontal, string vertical, int indent) =>
        new CellFormatValue(numberFormat, fontName, size, bold, italic, underline, horizontal, vertical, indent);
    private static FormatBlock Uniform(int rows, int columns, CellFormatValue value)
    {
        var values = new CellFormatValue[rows * columns];
        for (var index = 0; index < values.Length; index++) values[index] = value;
        return new FormatBlock(rows, columns, values);
    }
    private static FormatBlockSnapshot Snapshot(string address, int row, int column, FormatBlock block) => new FormatBlockSnapshot(
        new SelectionSnapshot(new SelectionContext("book", "Sheet1", address), block.CellCount, null, "General", SelectionSafetyState.Safe()),
        row, column, block);

    private sealed class FakePort : IFormatBlockPort
    {
        private readonly FormatBlockSnapshot _source;
        private readonly FormatBlockSnapshot _destination;
        public FakePort(FormatBlockSnapshot source, FormatBlockSnapshot destination)
        { _source = source; _destination = destination; Current = destination.Contents; }
        public FormatBlock Current { get; private set; }
        public int WriteCount { get; private set; }
        public SelectionSnapshot CaptureSelection() => _destination.Selection;
        public FormatBlockSnapshot CaptureFormatBlock() => CaptureFormatBlock(_destination.Selection.Context);
        public FormatBlockSnapshot CaptureFormatBlock(SelectionContext target)
        {
            if (target.Equals(_source.Selection.Context)) return _source;
            if (target.Equals(_destination.Selection.Context)) return new FormatBlockSnapshot(_destination.Selection,
                _destination.FirstRow, _destination.FirstColumn, Current);
            throw new InvalidOperationException();
        }
        public void WriteFormatBlock(SelectionContext target, FormatBlock contents)
        {
            if (!target.Equals(_destination.Selection.Context)) throw new InvalidOperationException();
            WriteCount++;
            Current = contents;
        }
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public bool TryRead(SelectionContext target, string propertyId, out string value)
        {
            value = string.Empty;
            if (!target.Equals(_destination.Selection.Context) || propertyId != FormatPasteCommand.ReceiptPropertyId) return false;
            value = Current.Serialize(); return true;
        }
        public bool TryWrite(SelectionContext target, string propertyId, string value)
        {
            if (!target.Equals(_destination.Selection.Context) || propertyId != FormatPasteCommand.ReceiptPropertyId) return false;
            WriteFormatBlock(target, FormatBlock.Deserialize(value)); return true;
        }
    }

    private sealed class ThrowingSink : IPropertyReceiptSink
    {
        public void Add(PropertyReceipt receipt) => throw new InvalidOperationException("Injected receipt failure.");
    }
}
