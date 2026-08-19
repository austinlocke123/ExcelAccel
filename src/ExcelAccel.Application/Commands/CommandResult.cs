using System;

namespace ExcelAccel.Application.Commands;

public sealed class CommandResult
{
    private CommandResult(
        CommandResultStatus status,
        string commandId,
        string message,
        long changedCount,
        long skippedCount,
        string refusalCode,
        string diagnosticId,
        string receiptId)
    {
        if (changedCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(changedCount));
        }

        if (skippedCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(skippedCount));
        }

        Status = status;
        CommandId = commandId;
        Message = message;
        ChangedCount = changedCount;
        SkippedCount = skippedCount;
        RefusalCode = refusalCode;
        DiagnosticId = diagnosticId;
        ReceiptId = receiptId;
    }

    public bool Succeeded => Status == CommandResultStatus.Success;

    public CommandResultStatus Status { get; }

    public string CommandId { get; }

    public string Message { get; }

    public long AffectedCellCount => ChangedCount;

    public long ChangedCount { get; }

    public long SkippedCount { get; }

    public string RefusalCode { get; }

    public string DiagnosticId { get; }

    public string ReceiptId { get; }

    public static CommandResult Success(CommandPlan plan, string message, string receiptId = "") =>
        new CommandResult(CommandResultStatus.Success, plan.CommandId, message, plan.AffectedCellCount, 0, string.Empty, string.Empty, receiptId ?? string.Empty);

    public static CommandResult Success(string commandId, string message, long changedCount = 0, string receiptId = "") =>
        new CommandResult(CommandResultStatus.Success, commandId, message, changedCount, 0, string.Empty, string.Empty, receiptId ?? string.Empty);

    public static CommandResult Refused(CommandPlan plan, string reason, string refusalCode = RefusalCodes.SelectionUnsupported) =>
        new CommandResult(CommandResultStatus.Refused, plan.CommandId, reason, 0, 0, refusalCode, string.Empty, string.Empty);

    public static CommandResult Refused(string commandId, string reason, string refusalCode = RefusalCodes.SelectionUnsupported) =>
        new CommandResult(CommandResultStatus.Refused, commandId ?? throw new ArgumentNullException(nameof(commandId)), reason, 0, 0, refusalCode, string.Empty, string.Empty);

    public static CommandResult Failed(string commandId, string reason, string diagnosticId) =>
        new CommandResult(
            CommandResultStatus.Failed,
            commandId ?? throw new ArgumentNullException(nameof(commandId)),
            reason ?? string.Empty,
            0,
            0,
            string.Empty,
            diagnosticId ?? string.Empty,
            string.Empty);

    public static CommandResult Partial(string commandId, string reason, long changedCount, long skippedCount, string diagnosticId) =>
        new CommandResult(
            CommandResultStatus.Partial,
            commandId ?? throw new ArgumentNullException(nameof(commandId)),
            reason ?? string.Empty,
            changedCount,
            skippedCount,
            string.Empty,
            diagnosticId ?? string.Empty,
            string.Empty);
}
