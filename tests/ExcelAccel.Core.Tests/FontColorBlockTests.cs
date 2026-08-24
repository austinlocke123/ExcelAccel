using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.AutoColor;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FontColorBlockTests
{
    private static KeyValuePair<string, string> Cell(string address, string color) =>
        new KeyValuePair<string, string>(address, color);

    [Fact]
    public void ABlockRoundTripsExactly()
    {
        var cells = new[]
        {
            Cell("A1", "#FF0000"),
            Cell("B2", "#0000FF"),
            Cell("C3", "#FF0000"),
        };

        var restored = FontColorBlock.Deserialize(FontColorBlock.Serialize(cells));

        Assert.Equal(
            cells.OrderBy(cell => cell.Key, StringComparer.Ordinal).ToArray(),
            restored.OrderBy(cell => cell.Key, StringComparer.Ordinal).ToArray());
    }

    /// <summary>
    /// Undo compares the stored value against what it reads back, so two captures
    /// of one state must produce identical bytes regardless of capture order.
    /// </summary>
    [Fact]
    public void SerializationIsDeterministicRegardlessOfInputOrder()
    {
        var cells = new[] { Cell("C3", "#FF0000"), Cell("A1", "#0000FF"), Cell("B2", "#FF0000") };

        var forward = FontColorBlock.Serialize(cells);
        var reversed = FontColorBlock.Serialize(cells.Reverse().ToArray());

        Assert.Equal(forward, reversed);
    }

    /// <summary>
    /// The point of grouping: a real model holds a few font colours across many
    /// cells, so the undo value stays small enough to store.
    /// </summary>
    [Fact]
    public void ManyCellsSharingAColourStayCompact()
    {
        var cells = Enumerable.Range(1, 5_000)
            .Select(row => Cell("A" + row, row % 2 == 0 ? "#000000" : "#0000FF"))
            .ToArray();

        var serialized = FontColorBlock.Serialize(cells);

        Assert.Equal(2, serialized.Split(';').Length);
        Assert.True(
            serialized.Length < SessionUndoStore.MaximumValueCharacters,
            $"5,000 cells serialized to {serialized.Length:N0} characters.");
        Assert.Equal(5_000, FontColorBlock.Deserialize(serialized).Count);
    }

    /// <summary>
    /// The ceiling exists so the undo value fits. If either constant moves without
    /// the other being reconsidered, this fails.
    /// </summary>
    [Fact]
    public void TheCellCeilingKeepsTheWorstCaseInsideTheStoreLimit() =>
        Assert.True(
            FontColorBlock.WorstCaseCharacters(FontColorBlock.MaximumCells) < SessionUndoStore.MaximumValueCharacters,
            $"Worst case {FontColorBlock.WorstCaseCharacters(FontColorBlock.MaximumCells):N0} characters "
            + $"exceeds the {SessionUndoStore.MaximumValueCharacters:N0} limit.");

    [Fact]
    public void EveryCellCarryingADistinctColourStillFits()
    {
        var cells = Enumerable.Range(1, 20_000)
            .Select(row => Cell("A" + row, "#" + (row % 0xFFFFFF).ToString("X6")))
            .ToArray();

        var serialized = FontColorBlock.Serialize(cells);

        Assert.True(
            serialized.Length < SessionUndoStore.MaximumValueCharacters,
            $"20,000 distinct colours serialized to {serialized.Length:N0} characters.");
    }

    [Fact]
    public void ExceedingTheCellCeilingIsRefused()
    {
        var cells = Enumerable.Range(1, FontColorBlock.MaximumCells + 1)
            .Select(row => Cell("A" + row, "#000000"));

        var error = Assert.Throws<ArgumentException>(() => FontColorBlock.Serialize(cells));

        Assert.Contains(FontColorBlock.MaximumCells.ToString("N0"), error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptyBlockRoundTripsAsEmpty()
    {
        Assert.Equal(string.Empty, FontColorBlock.Serialize(Array.Empty<KeyValuePair<string, string>>()));
        Assert.Empty(FontColorBlock.Deserialize(string.Empty));
        Assert.Empty(FontColorBlock.Deserialize(null));
    }

    [Theory]
    [InlineData("not-a-colour")]
    [InlineData("#FFF")]
    [InlineData("#GGGGGG")]
    [InlineData("")]
    public void AnInvalidColourIsRefused(string color) =>
        Assert.Throws<ArgumentException>(() => FontColorBlock.Serialize(new[] { Cell("A1", color) }));

    /// <summary>
    /// A separator inside an address would silently split one cell into two on
    /// the way back, so it is refused on the way in.
    /// </summary>
    [Theory]
    [InlineData("A1;B2")]
    [InlineData("A1,B2")]
    [InlineData("A1=B2")]
    [InlineData("  ")]
    public void AnAddressCarryingASeparatorIsRefused(string address) =>
        Assert.Throws<ArgumentException>(() => FontColorBlock.Serialize(new[] { Cell(address, "#000000") }));

    [Fact]
    public void ARepeatedCellIsRefusedRatherThanSilentlyCollapsed() =>
        Assert.Throws<ArgumentException>(() => FontColorBlock.Serialize(new[]
        {
            Cell("A1", "#000000"),
            Cell("A1", "#FF0000"),
        }));

    [Fact]
    public void AMalformedGroupIsRefusedRatherThanPartiallyRead()
    {
        Assert.Throws<FormatException>(() => FontColorBlock.Deserialize("#FF0000"));
        Assert.Throws<FormatException>(() => FontColorBlock.Deserialize("=A1"));
    }

    /// <summary>
    /// A recolour of any size is one change on one receipt, which is what makes
    /// the 32-change batch limit irrelevant to AutoColor.
    /// </summary>
    [Fact]
    public void AWholeRecolourIsOneChangeOnOneReceipt()
    {
        var before = FontColorBlock.Serialize(
            Enumerable.Range(1, 1_000).Select(row => Cell("A" + row, "#000000")).ToArray());
        var after = FontColorBlock.Serialize(
            Enumerable.Range(1, 1_000).Select(row => Cell("A" + row, "#0000FF")).ToArray());
        var store = new SessionUndoStore();
        var context = new SelectionContext("Book.xlsx", "Sheet1", "A1:A1000");

        store.Add(new PropertyBatchReceipt(
            "receipt", "format.auto_color.selection", 1, context,
            new[] { new PropertyChange(FontColorBlock.ReceiptPropertyId, before, after) },
            "hash", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(8)));

        Assert.Equal(1, store.Count("Book.xlsx"));
    }

    /// <summary>
    /// The block is compared ordinally, because its exact bytes are the value.
    /// Case-insensitive comparison would let a differently-cased capture look
    /// unchanged and undo the wrong state.
    /// </summary>
    [Fact]
    public void UndoRestoresTheWholeBlockAndDetectsAnInterveningChange()
    {
        var before = FontColorBlock.Serialize(new[] { Cell("A1", "#000000"), Cell("A2", "#000000") });
        var after = FontColorBlock.Serialize(new[] { Cell("A1", "#0000FF"), Cell("A2", "#FF0000") });
        var context = new SelectionContext("Book.xlsx", "Sheet1", "A1:A2");
        var port = new FakePort { Value = after };
        var store = new SessionUndoStore();
        store.Add(new PropertyBatchReceipt(
            "receipt", "format.auto_color.selection", 1, context,
            new[] { new PropertyChange(FontColorBlock.ReceiptPropertyId, before, after) },
            "hash", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(8)));

        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);

        Assert.Equal(UndoOutcome.Success, undo.Outcome);
        Assert.Equal(before, port.Value);
    }

    [Fact]
    public void UndoRefusesWhenTheBlockChangedSincePlanning()
    {
        var before = FontColorBlock.Serialize(new[] { Cell("A1", "#000000") });
        var after = FontColorBlock.Serialize(new[] { Cell("A1", "#0000FF") });
        var context = new SelectionContext("Book.xlsx", "Sheet1", "A1");
        var port = new FakePort { Value = FontColorBlock.Serialize(new[] { Cell("A1", "#008000") }) };
        var store = new SessionUndoStore();
        store.Add(new PropertyBatchReceipt(
            "receipt", "format.auto_color.selection", 1, context,
            new[] { new PropertyChange(FontColorBlock.ReceiptPropertyId, before, after) },
            "hash", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(8)));

        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);

        Assert.Equal(UndoOutcome.Stale, undo.Outcome);
        Assert.Equal(FontColorBlock.Serialize(new[] { Cell("A1", "#008000") }), port.Value);
    }

    private sealed class FakePort : IPropertyReceiptPort
    {
        public string Value { get; set; } = string.Empty;

        public bool TryRead(SelectionContext target, string propertyId, out string value)
        {
            value = Value;
            return propertyId == FontColorBlock.ReceiptPropertyId;
        }

        public bool TryWrite(SelectionContext target, string propertyId, string value)
        {
            if (propertyId != FontColorBlock.ReceiptPropertyId) return false;
            Value = value;
            return true;
        }
    }
}
