using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Navigation;
using ExcelAccel.Core.Commands;
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

    public static CommandResult ApplyProfileFormatting(string commandId)
    {
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(commandId))
        {
            return CommandResult.Refused(commandId, "Mutation commands are disabled in safe mode or quarantine.", RefusalCodes.CommandQuarantined);
        }

        var port = CreateSelectionAdapter();
        var command = Phase1AFormattingCatalog.Create(commandId);
        var plan = command.Plan(ProfileRuntime.Current, port);
        if (plan.RequiresPreview)
        {
            return CommandResult.Refused(plan, "This command requires exact-plan preview confirmation, which is not available from the compact Ribbon menu.", RefusalCodes.PreviewRequired);
        }

        return command.Execute(plan, port);
    }

    public static CommandResult Navigate(string commandId)
    {
        var port = new ExcelNavigationAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        var service = NavigationRuntime.Service;
        bool succeeded;
        switch (commandId)
        {
            case "navigate.sheet.previous": succeeded = service.MoveSheet(port, -1, ProfileRuntime.Current.WrapSheetNavigation); break;
            case "navigate.sheet.next": succeeded = service.MoveSheet(port, 1, ProfileRuntime.Current.WrapSheetNavigation); break;
            case "navigate.cell.a1": succeeded = service.Move(port, NavigationTargetKind.A1); break;
            case "navigate.used.first": succeeded = service.Move(port, NavigationTargetKind.UsedFirst); break;
            case "navigate.used.last": succeeded = service.Move(port, NavigationTargetKind.UsedLast); break;
            case "navigate.region.edge.up": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeUp); break;
            case "navigate.region.edge.down": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeDown); break;
            case "navigate.region.edge.left": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeLeft); break;
            case "navigate.region.edge.right": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeRight); break;
            case "navigate.history.back": succeeded = service.Back(port); break;
            case "navigate.history.forward": succeeded = service.Forward(port); break;
            case "navigate.bookmark.next_session": succeeded = service.NextBookmark(port); break;
            case "navigate.bookmark.previous_session": succeeded = service.PreviousBookmark(port); break;
            case "navigate.bookmark.add_session": service.AddBookmark(port); succeeded = true; break;
            case "navigate.bookmark.clear_session": service.ClearBookmarks(); succeeded = true; break;
            default: return CommandResult.Refused(commandId, "The navigation command is not available.", RefusalCodes.CommandUnavailable);
        }

        var location = port.CaptureLocation();
        var plan = new CommandPlan(commandId, CommandImpact.ReadOnly,
            new SelectionContext(location.WorkbookId, location.WorksheetName, location.Address), new string[0], 0, "Read-only navigation.");
        return succeeded ? CommandResult.Success(plan, "Navigation completed without changing workbook content.") :
            CommandResult.Refused(plan, "No valid navigation target is available.", RefusalCodes.SelectionUnsupported);
    }

    private static ExcelSelectionAdapter CreateSelectionAdapter() =>
        new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
}
