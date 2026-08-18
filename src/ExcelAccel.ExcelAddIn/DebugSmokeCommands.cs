#if DEBUG
using System;
using ExcelDna.Integration;
using ExcelAccel.Core.Reliability;
using ExcelAccel.ExcelAddIn.Interop;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

public static class DebugSmokeCommands
{
    [ExcelCommand(
        Name = "ExcelAccel.Smoke.ApplyCurrencyFormat",
        Description = "Debug-only integration hook; not compiled into Release builds.")]
    public static void ApplyCurrencyFormat()
    {
        try
        {
            var result = CommandDispatcher.ApplyCurrencyFormat();
            DiagnosticLog.Info("smoke.format.number.currency", result.Succeeded ? "success" : "refused");
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("smoke.format.number.currency", exception);
        }
    }

    [ExcelCommand(
        Name = "ExcelAccel.Smoke.ThrowInsideStateGuard",
        Description = "Debug-only state-restoration fault hook; not compiled into Release builds.")]
    public static void ThrowInsideStateGuard()
    {
        try
        {
            object application = ExcelDnaUtil.Application;
            try
            {
                ApplicationStateGuard.Run(
                    new ExcelApplicationStateAdapter(application),
                    ApplicationStateChangeSet.PropertyMutation(),
                    () => throw new InvalidOperationException("Injected smoke-test failure."));
            }
            catch (InvalidOperationException)
            {
                DiagnosticLog.Info("smoke.state.restore", "expected_failure_contained");
            }
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("smoke.state.restore", exception);
        }
    }
}
#endif
