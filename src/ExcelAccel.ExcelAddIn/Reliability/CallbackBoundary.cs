using System;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelAccel.Application.Commands;
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
                    result.Succeeded
                        ? "success"
                        : result.Status == CommandResultStatus.Failed
                            ? $"failed:{result.DiagnosticId}"
                            : $"refused:{result.RefusalCode}",
                    stopwatch.ElapsedMilliseconds);
                var detail = result.Succeeded
                    ? $"Command: {result.CommandId}\n\n{result.Message}"
                    : $"Command: {result.CommandId}\n\nReason: {result.Message}\n\nRemediation: Review the current workbook context and retry. If the refusal persists, export diagnostics.\nCode: {result.RefusalCode}\nDiagnostic ID: {result.DiagnosticId}";
                Show(detail, result.Succeeded
                    ? MessageBoxIcon.Information
                    : result.Status == CommandResultStatus.Failed
                        ? MessageBoxIcon.Error
                        : MessageBoxIcon.Warning);
            }
            catch (CommandRefusedException exception)
            {
                stopwatch.Stop();
                DiagnosticLog.Info(commandId, $"refused:{exception.RefusalCode}", stopwatch.ElapsedMilliseconds);
                Show($"Command: {commandId}\n\nReason: {exception.Message}\n\nRemediation: {exception.Remediation}\nCode: {exception.RefusalCode}", MessageBoxIcon.Warning);
            }
            catch (StateRestoreException exception)
            {
                stopwatch.Stop();
                RuntimeState.Quarantine(commandId);
                var diagnosticId = DiagnosticLog.Failure(commandId, "STATE_RESTORE_FAILED", exception, stopwatch.ElapsedMilliseconds);
                Show(
                    $"ExcelAccel could not fully restore Excel application state. This command is disabled for the rest of the session; save and restart Excel before continuing. Diagnostic ID: {diagnosticId}",
                    MessageBoxIcon.Error);
            }
            catch (Exception exception)
            {
                stopwatch.Stop();
                var diagnosticId = DiagnosticLog.Error(commandId, exception, stopwatch.ElapsedMilliseconds);
                Show($"ExcelAccel stopped the command safely. No further work was attempted. Diagnostic ID: {diagnosticId}", MessageBoxIcon.Error);
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
            var owner = ExcelWindowOwner.TryCreate();
            if (owner is null) MessageBox.Show(message, "ExcelAccel", MessageBoxButtons.OK, icon);
            else MessageBox.Show(owner, message, "ExcelAccel", MessageBoxButtons.OK, icon);
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("ui.message", exception);
        }
    }
}
