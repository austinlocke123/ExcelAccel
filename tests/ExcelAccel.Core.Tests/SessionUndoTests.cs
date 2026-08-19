using System;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class SessionUndoTests
{
    private static readonly SelectionContext Target = new SelectionContext("Book.xlsx", "Sheet1", "A1:B2");
    private static readonly DateTimeOffset Now = DateTimeOffset.Parse("2026-08-19T12:00:00Z");

    [Fact]
    public void UndoRestoresOnlyWhenCurrentValueMatchesPostState()
    {
        var store = new SessionUndoStore();
        store.Add(Receipt("before", "after"));
        var port = new FakeReceiptPort("after");

        var result = store.TryUndo(Target.WorkbookId, port, Now);

        Assert.True(result.Succeeded);
        Assert.Equal("before", port.Value);
        Assert.Equal(1, port.WriteCount);
    }

    [Fact]
    public void InterveningChangeRefusesWithoutOverwritingAndInvalidatesReceipt()
    {
        var store = new SessionUndoStore();
        store.Add(Receipt("before", "after"));
        var port = new FakeReceiptPort("later edit");

        var result = store.TryUndo(Target.WorkbookId, port, Now);

        Assert.Equal(UndoOutcome.Stale, result.Outcome);
        Assert.Equal("later edit", port.Value);
        Assert.Equal(0, port.WriteCount);
        Assert.Equal(0, store.Count(Target.WorkbookId));
    }

    [Fact]
    public void StoreIsBoundedPerWorkbookAndClearIsContentFree()
    {
        var store = new SessionUndoStore();
        for (var index = 0; index < 30; index++) store.Add(Receipt(index.ToString(), "after", index.ToString()));
        Assert.Equal(SessionUndoStore.MaximumReceiptsPerWorkbook, store.Count(Target.WorkbookId));
        store.ClearWorkbook(Target.WorkbookId);
        Assert.Equal(0, store.Count(Target.WorkbookId));
    }

    [Fact]
    public void ExpiredAndFailedWritesCannotReportSuccess()
    {
        var expired = new SessionUndoStore();
        expired.Add(Receipt("before", "after"));
        Assert.Equal(UndoOutcome.Expired, expired.TryUndo(Target.WorkbookId, new FakeReceiptPort("after"), Now.AddHours(9)).Outcome);

        var failed = new SessionUndoStore();
        failed.Add(Receipt("before", "after"));
        Assert.Equal(UndoOutcome.WriteFailed, failed.TryUndo(Target.WorkbookId, new FakeReceiptPort("after") { FailWrite = true }, Now).Outcome);
    }

    private static PropertyReceipt Receipt(string before, string after, string id = "receipt") =>
        new PropertyReceipt(id, "test.command", 1, Target, "font_color", before, after, "ABC", Now, Now.AddHours(8));

    private sealed class FakeReceiptPort : IPropertyReceiptPort
    {
        public FakeReceiptPort(string value) => Value = value;
        public string Value { get; private set; }
        public int WriteCount { get; private set; }
        public bool FailWrite { get; set; }
        public bool TryRead(SelectionContext target, string propertyId, out string value) { value = Value; return target.Equals(Target); }
        public bool TryWrite(SelectionContext target, string propertyId, string value)
        {
            if (FailWrite || !target.Equals(Target)) return false;
            WriteCount++;
            Value = value;
            return true;
        }
    }
}
