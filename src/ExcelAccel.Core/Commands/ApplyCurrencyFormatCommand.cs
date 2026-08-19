using System;

namespace ExcelAccel.Core.Commands;

public sealed class ApplyCurrencyFormatCommand
{
    public const string Id = "format.number.currency";
    public const string ChangedProperty = "number_format";
    public const string DefaultFormatCode = "$#,##0.00;($#,##0.00);-";
    public const long MaximumCellCount = 50_000;

    public CommandPlan Plan(SelectionSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        if (snapshot.CellCount > MaximumCellCount)
        {
            throw new CommandRefusedException($"The selection contains {snapshot.CellCount:N0} cells; the Phase 0 safety limit is {MaximumCellCount:N0}.");
        }

        return new CommandPlan(
            Id,
            CommandImpact.Low,
            snapshot.Context,
            new[] { ChangedProperty },
            snapshot.CellCount,
            $"Set only NumberFormat on {snapshot.CellCount:N0} selected cell(s).");
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

        var current = port.CaptureSelection();
        if (!plan.Context.Equals(current.Context))
        {
            return CommandResult.Refused(plan, "The Excel selection changed after planning; no formatting was applied.");
        }

        if (current.CellCount != plan.AffectedCellCount)
        {
            return CommandResult.Refused(plan, "The selected range changed size after planning; no formatting was applied.");
        }

        port.SetNumberFormat(DefaultFormatCode);
        return CommandResult.Success(plan, $"Applied the currency number format to {plan.AffectedCellCount:N0} cell(s).");
    }
}
