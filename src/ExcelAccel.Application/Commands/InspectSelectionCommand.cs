using System;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public sealed class InspectSelectionCommand
{
    public const string Id = "inspect.selection.summary";

    public CommandPlan Plan(SelectionSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        return new CommandPlan(
            Id,
            CommandImpact.ReadOnly,
            snapshot.Context,
            Array.Empty<string>(),
            snapshot.CellCount,
            "Read selection metadata without changing the workbook.");
    }

    public CommandResult Execute(CommandPlan plan, ISelectionPort port)
    {
        if (plan is null)
        {
            throw new ArgumentNullException(nameof(plan));
        }

        if (port is null)
        {
            throw new ArgumentNullException(nameof(port));
        }

        var authorization = CommandExecutionGate.Authorize(BuiltInCommandRegistry.GetRequired(Id), plan);
        if (!authorization.Allowed)
        {
            return CommandResult.Refused(plan, authorization.Message, authorization.RefusalCode);
        }

        var current = port.CaptureSelection();
        if (!plan.Context.Equals(current.Context))
        {
            return CommandResult.Refused(plan, "The Excel selection changed after planning; run the command again.", RefusalCodes.StaleContext);
        }

        var formulaState = current.HasFormula.HasValue
            ? (current.HasFormula.Value ? "contains formulas" : "contains no formulas")
            : "contains a mix of formula and non-formula cells";

        return CommandResult.Success(
            plan,
            $"{current.Context.WorksheetName}!{current.Context.Address}: {current.CellCount:N0} cell(s), {formulaState}, number format '{current.NumberFormat}'.");
    }
}
