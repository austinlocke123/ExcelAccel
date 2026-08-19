using System;

namespace ExcelAccel.Core.Commands;

public sealed class CommandResult
{
    private CommandResult(bool succeeded, string commandId, string message, long affectedCellCount)
    {
        Succeeded = succeeded;
        CommandId = commandId;
        Message = message;
        AffectedCellCount = affectedCellCount;
    }

    public bool Succeeded { get; }

    public string CommandId { get; }

    public string Message { get; }

    public long AffectedCellCount { get; }

    public static CommandResult Success(CommandPlan plan, string message) =>
        new CommandResult(true, plan.CommandId, message, plan.AffectedCellCount);

    public static CommandResult Refused(CommandPlan plan, string reason) =>
        new CommandResult(false, plan.CommandId, reason, 0);

    public static CommandResult Refused(string commandId, string reason) =>
        new CommandResult(false, commandId ?? throw new ArgumentNullException(nameof(commandId)), reason, 0);
}
