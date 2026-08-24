using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.AutoColor;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class AutoColorCommandTests
{
    private static AutoColorCellSnapshot Cell(string address, CellScalarKind kind, string formula, string color) =>
        new AutoColorCellSnapshot(address, kind, formula, color);

    private static ExcelAccel.Application.Profiles.ProfileDefinition Profile() => new ProfileStore().LoadDefault();

    private static AutoColorCommand Selection() =>
        AutoColorCommandCatalog.Create(AutoColorCommandCatalog.SelectionId);

    [Fact]
    public void ARecolourWritesOnlyFontColourAndRecordsOneReceipt()
    {
        var port = new FakePort(
            Cell("A1", CellScalarKind.Number, string.Empty, "#000000"),
            Cell("A2", CellScalarKind.Number, "=A1+1", "#000000"));
        var store = new SessionUndoStore();
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, store);

        Assert.True(result.Succeeded);
        Assert.Equal(new[] { FontColorBlock.ReceiptPropertyId }, execution.CommandPlan.ChangedProperties);
        Assert.Equal(1, store.Count("Book.xlsx"));
        // Both are hardcodes: A1 is a typed number, A2 embeds the literal 1.
        Assert.Equal("#0000FF", port.ColorOf("A1"));
        Assert.Equal("#0000FF", port.ColorOf("A2"));
    }

    /// <summary>
    /// The whole recolour is one undo value, so one undo puts every cell back.
    /// </summary>
    [Fact]
    public void OneUndoRestoresEveryCell()
    {
        var port = new FakePort(
            Cell("A1", CellScalarKind.Number, string.Empty, "#111111"),
            Cell("A2", CellScalarKind.Text, string.Empty, "#222222"));
        var store = new SessionUndoStore();
        var command = Selection();
        var execution = command.Plan(Profile(), port);
        command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, store);

        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);

        Assert.Equal(UndoOutcome.Success, undo.Outcome);
        Assert.Equal("#111111", port.ColorOf("A1"));
        Assert.Equal("#222222", port.ColorOf("A2"));
    }

    /// <summary>
    /// The receipt carries only cells that changed. Already-correct and
    /// unsupported cells in the original selection must not make that sparse
    /// receipt look stale when undo validates it.
    /// </summary>
    [Fact]
    public void OneUndoRestoresASparseChangeInsideAMixedSelection()
    {
        var port = new FakePort(
            Cell("A1", CellScalarKind.Number, string.Empty, "#111111"),
            Cell("A2", CellScalarKind.Number, string.Empty, "#0000FF"),
            Cell("A3", CellScalarKind.Empty, string.Empty, "#333333"));
        var store = new SessionUndoStore();
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, store);
        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);

        Assert.True(result.Succeeded);
        Assert.Single(execution.Plan.Changes);
        Assert.Equal(UndoOutcome.Success, undo.Outcome);
        Assert.Equal("#111111", port.ColorOf("A1"));
        Assert.Equal("#0000FF", port.ColorOf("A2"));
        Assert.Equal("#333333", port.ColorOf("A3"));
    }

    [Fact]
    public void AnAlreadyCorrectSelectionWritesNothing()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#0000FF"));
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.True(result.Succeeded);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void AChangeAfterPlanningIsRefusedWithoutWriting()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#000000"));
        var command = Selection();
        var execution = command.Plan(Profile(), port);
        port.Replace(Cell("A1", CellScalarKind.Text, string.Empty, "#000000"));

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    /// <summary>
    /// A failed write must leave every cell exactly as it was, not partly
    /// recoloured.
    /// </summary>
    [Fact]
    public void AFailedWriteRestoresTheExactPriorColours()
    {
        var port = new FakePort(
            Cell("A1", CellScalarKind.Number, string.Empty, "#111111"),
            Cell("A2", CellScalarKind.Number, string.Empty, "#222222"))
        {
            FailWriteNumber = 1,
        };
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.Equal("AUTO_COLOR_ROLLED_BACK", result.DiagnosticId);
        Assert.Equal("#111111", port.ColorOf("A1"));
        Assert.Equal("#222222", port.ColorOf("A2"));
    }

    [Fact]
    public void APostconditionMismatchCannotReportSuccess()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#111111")) { IgnoreWrites = true };
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, new SessionUndoStore());

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void AReceiptStoreFailureRollsTheWriteBack()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#111111"));
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, new ThrowingSink());

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.Equal("RECEIPT_STORE_ROLLED_BACK", result.DiagnosticId);
        Assert.Equal("#111111", port.ColorOf("A1"));
    }

    [Fact]
    public void WithoutAnUndoStoreNothingIsWritten()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#111111"));
        var command = Selection();
        var execution = command.Plan(Profile(), port);

        var result = command.Execute(execution, Profile(), port, execution.CommandPlan.PlanHash, null);

        Assert.Equal(RefusalCodes.CommandUnavailable, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    /// <summary>
    /// A recolour that cannot fit in one undo value is refused rather than
    /// written, because it could not be reversed.
    /// </summary>
    [Fact]
    public void ARecolourTooLargeToUndoIsRefusedBeforeWriting()
    {
        var cells = Enumerable.Range(1, FontColorBlock.MaximumCells + 1)
            .Select(row => Cell("A" + row, CellScalarKind.Number, string.Empty, "#000000"))
            .ToArray();
        var port = new FakePort(cells);

        var refusal = Assert.Throws<CommandRefusedException>(() => Selection().Plan(Profile(), port));

        Assert.Equal(RefusalCodes.ResourceLimit, refusal.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void WorksheetScopeStaysGatedAtPlanningTime()
    {
        var port = new FakePort(Cell("A1", CellScalarKind.Number, string.Empty, "#000000"));
        var command = AutoColorCommandCatalog.Create(AutoColorCommandCatalog.WorksheetId);

        var refusal = Assert.Throws<CommandRefusedException>(() => command.Plan(Profile(), port));

        Assert.Equal("PERFORMANCE_QUALIFICATION_REQUIRED", refusal.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void BothCommandsAreRegisteredWithUndoAndOneChangedProperty()
    {
        foreach (var id in new[] { AutoColorCommandCatalog.SelectionId, AutoColorCommandCatalog.WorksheetId })
        {
            var descriptor = BuiltInCommandRegistry.GetRequired(id);
            Assert.Equal(CommandImpact.Medium, descriptor.Impact);
            Assert.Equal(new[] { FontColorBlock.ReceiptPropertyId }, descriptor.ChangedProperties);
            Assert.Equal(UndoPolicy.SessionPropertyReceipt, descriptor.UndoPolicy);
            Assert.Equal("CAP-FMT-002", descriptor.CapabilityId);
        }

        Assert.Equal(
            PreviewPolicy.Mandatory,
            BuiltInCommandRegistry.GetRequired(AutoColorCommandCatalog.WorksheetId).PreviewPolicy);
    }

    private sealed class ThrowingSink : IPropertyBatchReceiptSink
    {
        public void Add(PropertyBatchReceipt receipt) => throw new InvalidOperationException("Injected receipt failure.");
    }

    private sealed class FakePort : IAutoColorPort
    {
        private readonly Dictionary<string, AutoColorCellSnapshot> _cells;

        public FakePort(params AutoColorCellSnapshot[] cells) =>
            _cells = cells.ToDictionary(cell => cell.Address, StringComparer.OrdinalIgnoreCase);

        public int WriteCount { get; private set; }
        public bool IgnoreWrites { get; set; }

        /// <summary>
        /// The 1-based write that throws. The command writes once for the whole
        /// recolour and once more to roll back, so failing write 1 and allowing
        /// write 2 is what exercises a successful rollback.
        /// </summary>
        public int FailWriteNumber { get; set; } = int.MaxValue;

        public string ColorOf(string address) => _cells[address].FontColor;

        public void Replace(AutoColorCellSnapshot cell) => _cells[cell.Address] = cell;

        public SelectionSnapshot CaptureSelection() => new SelectionSnapshot(
            new SelectionContext("Book.xlsx", "Sheet1", "A1:A" + _cells.Count), _cells.Count, false, "General");

        public IReadOnlyList<AutoColorCellSnapshot> CaptureCells(AutoColorScope scope) =>
            _cells.Values.OrderBy(cell => cell.Address, StringComparer.Ordinal).ToArray();

        public void WriteFontColors(IEnumerable<AutoColorChange> changes)
        {
            WriteCount++;
            if (WriteCount == FailWriteNumber)
            {
                throw new InvalidOperationException("Injected write failure.");
            }

            if (IgnoreWrites) return;
            foreach (var change in changes)
            {
                var existing = _cells[change.Address];
                _cells[change.Address] = new AutoColorCellSnapshot(
                    existing.Address, existing.ScalarKind, existing.Formula, change.AfterColor);
            }
        }

        public bool TryRead(SelectionContext target, string propertyId, out string value)
        {
            value = FontColorBlock.Serialize(
                _cells.Values.Select(cell => new KeyValuePair<string, string>(cell.Address, cell.FontColor)));
            return propertyId == FontColorBlock.ReceiptPropertyId;
        }

        public bool TryRead(
            SelectionContext target,
            string propertyId,
            string referenceValue,
            out string value)
        {
            var addresses = FontColorBlock.Deserialize(referenceValue)
                .Select(cell => cell.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            value = FontColorBlock.Serialize(_cells.Values
                .Where(cell => addresses.Contains(cell.Address))
                .Select(cell => new KeyValuePair<string, string>(cell.Address, cell.FontColor)));
            return propertyId == FontColorBlock.ReceiptPropertyId;
        }

        public bool TryWrite(SelectionContext target, string propertyId, string value)
        {
            if (propertyId != FontColorBlock.ReceiptPropertyId) return false;
            foreach (var cell in FontColorBlock.Deserialize(value))
            {
                var existing = _cells[cell.Key];
                _cells[cell.Key] = new AutoColorCellSnapshot(
                    existing.Address, existing.ScalarKind, existing.Formula, cell.Value);
            }

            return true;
        }
    }
}
