using System;

namespace ExcelAccel.Application.Commands;

public sealed class CommandResult
{
    private CommandResult(bool succeeded, string commandId, string message, long affectedCellCount, string refusalCode)
    {
        Succeeded = succeeded;
        CommandId = commandId;
        Message = message;
        AffectedCellCount = affectedCellCount;
        RefusalCode = refusalCode;
    }

    public bool Succeeded { get; }

    public string CommandId { get; }

    public string Message { get; }

    public long AffectedCellCount { get; }

    public string RefusalCode { get; }

    public static CommandResult Success(CommandPlan plan, string message) =>
        new CommandResult(true, plan.CommandId, message, plan.AffectedCellCount, string.Empty);

    public static CommandResult Refused(CommandPlan plan, string reason, string refusalCode = RefusalCodes.SelectionUnsupported) =>
        new CommandResult(false, plan.CommandId, reason, 0, refusalCode);

    public static CommandResult Refused(string commandId, string reason, string refusalCode = RefusalCodes.SelectionUnsupported) =>
        new CommandResult(false, commandId ?? throw new ArgumentNullException(nameof(commandId)), reason, 0, refusalCode);
}
