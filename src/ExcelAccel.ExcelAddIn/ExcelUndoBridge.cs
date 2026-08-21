using System;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Undo;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Lets Excel's own Ctrl+Z reverse the most recent ExcelAccel change.
///
/// Writing to a workbook through the object model makes Excel discard its native
/// undo stack, which is why this add-in keeps its own session undo. Excel does
/// offer <c>Application.OnUndo</c>, which holds a single custom undo entry: after
/// a change that recorded a receipt, Ctrl+Z runs the macro below instead.
///
/// The limits are Excel's, not ours, and are worth knowing:
/// it is one slot rather than a stack, so a second armed command replaces the
/// first; and any native edit makes Excel record its own undo and discard the
/// custom entry. Ctrl+Z therefore reaches an ExcelAccel change only while that
/// change is the last thing that happened in the workbook. Anything deeper still
/// needs the Undo ExcelAccel command, which walks the full receipt history.
/// </summary>
internal static class ExcelUndoBridge
{
    /// <summary>The macro name Excel invokes when the user chooses Undo.</summary>
    public const string MacroName = "ExcelAccel.UndoLastFromExcel";

    /// <summary>
    /// Offers the last change to Excel's undo slot. Advisory: if Excel will not
    /// accept it, the change still stands and the add-in's own undo still works.
    /// </summary>
    public static void Arm(string commandId)
    {
        // The undo command must not re-arm itself, or Ctrl+Z after an undo would
        // undo the undo.
        if (string.IsNullOrWhiteSpace(commandId) || commandId == UndoLastCommand.Id) return;
        try
        {
            var name = BuiltInCommandRegistry.All.FirstOrDefault(value => value.Id == commandId)?.DisplayName ?? commandId;
            ((dynamic)ExcelDnaUtil.Application).OnUndo("Undo ExcelAccel: " + name, MacroName);
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("ui.on_undo.arm", exception);
        }
    }
}

public static class ExcelUndoCommands
{
    /// <summary>
    /// Invoked by Excel when the user presses Ctrl+Z after an ExcelAccel change.
    /// It runs exactly the same optimistic undo as the ribbon command, so the
    /// receipt is still revalidated against the current cell and refuses rather
    /// than overwriting an edit made since.
    /// </summary>
    [ExcelCommand(Name = ExcelUndoBridge.MacroName, Description = "Reverses the last ExcelAccel property change.")]
    public static void UndoLastFromExcel()
    {
        CallbackBoundary.Run(UndoLastCommand.Id, CommandDispatcher.UndoLastProperty);
    }
}
