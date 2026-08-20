using System;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Reliability;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal static class CallbackBoundary
{
    private static readonly ReentrancyGate Gate = new ReentrancyGate();

    /// <summary>Excel truncates a longer status bar message itself.</summary>
    private const int StatusBarLimit = 255;

    public static void Run(string commandId, Func<CommandResult> callback, bool showResult = true)
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
                            : result.Status == CommandResultStatus.Partial
                                ? $"partial:{result.DiagnosticId}"
                            : $"refused:{result.RefusalCode}",
                    stopwatch.ElapsedMilliseconds);
                if (result.Succeeded)
                {
                    // A command that worked must not interrupt the user. Success
                    // goes to the Excel status bar, which needs no dismissing,
                    // so keyboard-driven work is never blocked by a dialog.
                    ReportSuccess(result.Message);
                }
                else if (showResult)
                {
                    Show(
                        $"Command: {result.CommandId}\n\nReason: {result.Message}\n\nRemediation: Review the current workbook context and retry. If the refusal persists, export diagnostics.\nCode: {result.RefusalCode}\nDiagnostic ID: {result.DiagnosticId}",
                        result.Status == CommandResultStatus.Failed || result.Status == CommandResultStatus.Partial
                            ? MessageBoxIcon.Error
                            : MessageBoxIcon.Warning);
                }
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

    /// <summary>
    /// Writes a successful outcome to the Excel status bar. It is advisory only:
    /// if Excel will not accept it, the command still succeeded and nothing is
    /// surfaced to the user.
    /// </summary>
    private static void ReportSuccess(string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;
        try
        {
            var text = message.Replace("\r", " ").Replace("\n", " ").Trim();
            if (text.Length > StatusBarLimit) text = text.Substring(0, StatusBarLimit - 1) + "\u2026";
            ((dynamic)ExcelDnaUtil.Application).StatusBar = "ExcelAccel: " + text;
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("ui.status_bar", exception);
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
