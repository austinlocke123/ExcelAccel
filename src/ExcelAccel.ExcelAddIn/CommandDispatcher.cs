using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;

namespace ExcelAccel.ExcelAddIn;

internal static class CommandDispatcher
{
    public static CommandResult InspectSelection()
    {
        var port = CreateSelectionAdapter();
        var command = new InspectSelectionCommand();
        var plan = command.Plan(port.CaptureSelection());
        return command.Execute(plan, port);
    }

    public static CommandResult ApplyCurrencyFormat()
    {
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(ApplyCurrencyFormatCommand.Id))
        {
            return CommandResult.Refused(
                ApplyCurrencyFormatCommand.Id,
                "ExcelAccel is in safe mode after an unclean prior session. Restart Excel cleanly before using mutation commands.",
                RefusalCodes.CommandQuarantined);
        }

        var port = CreateSelectionAdapter();
        var command = new ApplyCurrencyFormatCommand();
        var snapshot = port.CaptureSelection();
        var canExecute = command.CanExecute(snapshot);
        if (!canExecute.Allowed)
        {
            return CommandResult.Refused(
                ApplyCurrencyFormatCommand.Id,
                $"{canExecute.Message} {canExecute.Remediation}".Trim(),
                canExecute.RefusalCode);
        }

        var plan = command.Plan(snapshot);
        return command.Execute(plan, port);
    }

    private static ExcelSelectionAdapter CreateSelectionAdapter() =>
        new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
}
