#if DEBUG
using System;
using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Reliability;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using System.Windows.Forms;

namespace ExcelAccel.ExcelAddIn;

public static class DebugSmokeCommands
{
    [ExcelCommand(Name = "ExcelAccel.Smoke.OpenAndCloseCommandSearch", Description = "Debug-only command-search UI lifecycle hook.")]
    public static void OpenAndCloseCommandSearch()
    {
        try
        {
            var result = CommandSearchRuntime.Open();
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
            System.Windows.Forms.Application.DoEvents();
            CommandSearchRuntime.Reset();
            DiagnosticLog.Info("smoke.command.search", "opened_and_closed");
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("smoke.command.search", exception);
            throw;
        }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.OpenAndCloseStyleLibrary", Description = "Debug-only style-library UI lifecycle hook.")]
    public static void OpenAndCloseStyleLibrary()
    {
        try
        {
            var result = StyleLibraryRuntime.Open();
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
            System.Windows.Forms.Application.DoEvents();
            StyleLibraryRuntime.Reset();
            DiagnosticLog.Info("smoke.style.library", "opened_and_closed");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.style.library", exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.ApplyMajorHeaderStyle", Description = "Debug-only built-in style hook.")]
    public static void ApplyMajorHeaderStyle()
    {
        try
        {
            var result = CommandDispatcher.ApplyStyle("major_header", requireBuiltIn: true);
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
            DiagnosticLog.Info("smoke.style.major_header", "success");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.style.major_header", exception); throw; }
    }

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

    [ExcelCommand(Name = "ExcelAccel.Smoke.UndoLastProperty", Description = "Debug-only optimistic undo hook.")]
    public static void UndoLastProperty()
    {
        try
        {
            var result = CommandDispatcher.UndoLastProperty();
            DiagnosticLog.Info("smoke.undo.property", result.Succeeded ? "success" : "refused");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.undo.property", exception); }
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
