using System;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal static class CallbackBoundary
{
    private static readonly ReentrancyGate Gate = new ReentrancyGate();

    public static void Run(string commandId, Func<CommandResult> callback)
    {
        var lease = Gate.TryEnter();
        if (lease is null)
        {
            Show("ExcelAccel is already running a command. Wait for it to finish and try again.", MessageBoxIcon.Information);
            return;
        }

        using (lease)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                var result = callback();
                stopwatch.Stop();
                DiagnosticLog.Info(
                    commandId,
                    result.Succeeded ? "success" : $"refused:{result.RefusalCode}",
                    stopwatch.ElapsedMilliseconds);
                Show(result.Message, result.Succeeded ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (CommandRefusedException exception)
            {
                stopwatch.Stop();
                DiagnosticLog.Info(commandId, $"refused:{exception.RefusalCode}", stopwatch.ElapsedMilliseconds);
                Show(exception.Message, MessageBoxIcon.Warning);
            }
            catch (StateRestoreException exception)
            {
                stopwatch.Stop();
                RuntimeState.Quarantine(commandId);
                DiagnosticLog.Failure(commandId, "STATE_RESTORE_FAILED", exception, stopwatch.ElapsedMilliseconds);
                Show(
                    "ExcelAccel could not fully restore Excel application state. This command is disabled for the rest of the session; save and restart Excel before continuing.",
                    MessageBoxIcon.Error);
            }
            catch (Exception exception)
            {
                stopwatch.Stop();
                DiagnosticLog.Error(commandId, exception, stopwatch.ElapsedMilliseconds);
                Show("ExcelAccel stopped the command safely. No further work was attempted. See the local diagnostic log for the error type.", MessageBoxIcon.Error);
            }
        }
    }

    public static void RunLifecycle(string operationId, Action callback)
    {
        try
        {
            callback();
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error(operationId, exception);
        }
    }

    private static void Show(string message, MessageBoxIcon icon)
    {
        try
        {
            MessageBox.Show(message, "ExcelAccel", MessageBoxButtons.OK, icon);
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("ui.message", exception);
        }
    }
}
