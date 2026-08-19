using System;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public sealed class ApplyCurrencyFormatCommand
{
    public const string Id = "format.number.currency";
    public const string ChangedProperty = "number_format";
    public const string DefaultFormatCode = "$#,##0.00;($#,##0.00);-";
    public const long MaximumCellCount = 50_000;

    public CanExecuteResult CanExecute(SelectionSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        if (snapshot.Safety.AreaCount != 1)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.MultiAreaUnsupported,
                "Currency formatting currently requires one contiguous selection.",
                "Select one rectangular cell range and try again.");
        }

        if (snapshot.Safety.HasMergedCells)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.SelectionUnsupported,
                "Currency formatting is not qualified for merged cells.",
                "Select an unmerged range and try again.");
        }

        if (!snapshot.Safety.DynamicArraySpillCheckSupported)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ExcelCapabilityMissing,
                "This Excel build does not expose the required spilled-array safety check.",
                "Use a qualified Microsoft 365 Excel build.");
        }

        if (snapshot.Safety.HasLegacyArray || snapshot.Safety.HasDynamicArraySpill)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ArrayOrSpillUnsafe,
                "Currency formatting is not qualified for array or spilled-array selections.",
                "Select cells outside the array or spill range and try again.");
        }

        if (snapshot.Safety.WorksheetProtected)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ProtectedTarget,
                "The active worksheet is protected.",
                "Unprotect the worksheet or select an unprotected target.");
        }

        if (snapshot.Safety.WorkbookReadOnly)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ReadOnlyWorkbook,
                "The active workbook is read-only.",
                "Use an editable copy of the workbook.");
        }

        if (snapshot.CellCount > MaximumCellCount)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ResourceLimit,
                $"The selection contains {snapshot.CellCount:N0} cells; the Phase 0 safety limit is {MaximumCellCount:N0}.",
                "Select a smaller range and try again.");
        }

        return CanExecuteResult.Permit();
    }

    public CommandPlan Plan(SelectionSnapshot snapshot)
    {
        if (snapshot is null)
        {
            throw new ArgumentNullException(nameof(snapshot));
        }

        var canExecute = CanExecute(snapshot);
        if (!canExecute.Allowed)
        {
            throw new CommandRefusedException(canExecute.RefusalCode, canExecute.Message, canExecute.Remediation);
        }

        return new CommandPlan(
            Id,
            CommandImpact.Low,
            snapshot.Context,
            new[] { ChangedProperty },
            snapshot.CellCount,
            $"Set only NumberFormat on {snapshot.CellCount:N0} selected cell(s).",
            PreconditionFingerprint.Create(snapshot.NumberFormat));
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
            return CommandResult.Refused(plan, "The Excel selection changed after planning; no formatting was applied.", RefusalCodes.StaleContext);
        }

        if (current.CellCount != plan.AffectedCellCount)
        {
            return CommandResult.Refused(plan, "The selected range changed size after planning; no formatting was applied.", RefusalCodes.StaleContext);
        }

        if (!string.Equals(
                plan.PreconditionFingerprint,
                PreconditionFingerprint.Create(current.NumberFormat),
                StringComparison.Ordinal))
        {
            return CommandResult.Refused(
                plan,
                "The planned number-format property changed after planning; no formatting was applied.",
                RefusalCodes.StaleContext);
        }

        var canExecute = CanExecute(current);
        if (!canExecute.Allowed)
        {
            return CommandResult.Refused(plan, canExecute.Message, canExecute.RefusalCode);
        }

        port.SetNumberFormat(DefaultFormatCode);
        return CommandResult.Success(plan, $"Applied the currency number format to {plan.AffectedCellCount:N0} cell(s).");
    }
}
