using ExcelAccel.Core.Commands;
using ExcelAccel.ExcelAddIn.Interop;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

internal static class CommandDispatcher
{
    public static CommandResult InspectSelection()
    {
        var port = new ExcelSelectionAdapter();
        var command = new InspectSelectionCommand();
        var plan = command.Plan(port.CaptureSelection());
        return command.Execute(plan, port);
    }

    public static CommandResult ApplyCurrencyFormat()
    {
        if (RuntimeState.IsSafeMode)
        {
            return CommandResult.Refused(
                ApplyCurrencyFormatCommand.Id,
                "ExcelAccel is in safe mode after an unclean prior session. Restart Excel cleanly before using mutation commands.");
        }

        var port = new ExcelSelectionAdapter();
        var command = new ApplyCurrencyFormatCommand();
        var plan = command.Plan(port.CaptureSelection());
        return command.Execute(plan, port);
    }
}
