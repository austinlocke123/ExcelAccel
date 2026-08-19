#if DEBUG
using System;
using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Reliability;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;

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

    [ExcelCommand(Name = "ExcelAccel.Smoke.ApplyFontColorCycle", Description = "Debug-only profile formatting hook.")]
    public static void ApplyFontColorCycle()
    {
        try
        {
            var result = CommandDispatcher.ApplyProfileFormatting("format.font_color.cycle");
            DiagnosticLog.Info("smoke.format.font_color.cycle", result.Succeeded ? "success" : "refused");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.format.font_color.cycle", exception); }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.NavigateA1", Description = "Debug-only navigation hook.")]
    public static void NavigateA1()
    {
        try
        {
            var result = CommandDispatcher.Navigate("navigate.cell.a1");
            DiagnosticLog.Info("smoke.navigate.a1", result.Succeeded ? "success" : "refused");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.navigate.a1", exception); }
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

    [ExcelCommand(
        Name = "ExcelAccel.Smoke.ApplyCurrencyFormatAfterInterveningChange",
        Description = "Debug-only stale-property hook; not compiled into Release builds.")]
    public static void ApplyCurrencyFormatAfterInterveningChange()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var command = new ApplyCurrencyFormatCommand();
            var plan = command.Plan(port.CaptureSelection());
            port.SetNumberFormat("0.00");
            var result = command.Execute(plan, port);
            DiagnosticLog.Info(
                "smoke.format.number.currency.stale",
                result.RefusalCode ?? (result.Succeeded ? "unexpected_success" : "refused_without_code"));
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("smoke.format.number.currency.stale", exception);
        }
    }
}
#endif
