using System;
using ExcelAccel.Application.Commands;

namespace ExcelAccel.Application.Undo;

public static class UndoLastCommand
{
    public const string Id = "undo.last_excelaccel_property";

    public static CommandResult Execute(string workbookId, SessionUndoStore store, IPropertyReceiptPort port, DateTimeOffset now)
    {
        if (string.IsNullOrWhiteSpace(workbookId)) return CommandResult.Refused(Id, "An active workbook is required.", RefusalCodes.SelectionUnsupported);
        var result = store.TryUndo(workbookId, port, now);
        if (result.Succeeded) return CommandResult.Success(Id, result.Message, 1, result.ReceiptId);
        var code = result.Outcome == UndoOutcome.Empty ? "UNDO_EMPTY" : result.Outcome == UndoOutcome.Stale ? RefusalCodes.StaleContext : "UNDO_" + result.Outcome.ToString().ToUpperInvariant();
        return CommandResult.Refused(Id, result.Message, code);
    }
}
