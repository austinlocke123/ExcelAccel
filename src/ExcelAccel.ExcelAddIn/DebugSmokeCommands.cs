#if DEBUG
using System;
using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Reliability;
using ExcelAccel.Core.Commands;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.DataCleaning;
using ExcelAccel.Application.Auditing;
using System.Windows.Forms;

namespace ExcelAccel.ExcelAddIn;

public static class DebugSmokeCommands
{
    [ExcelCommand(Name = "ExcelAccel.Smoke.DirectPrecedents", Description = "Debug-only direct-precedent capture hook.")]
    public static string DirectPrecedents()
    {
        try
        {
            var port = new ExcelReferenceSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var result = new DirectPrecedentCoordinator().Execute(port);
            var summary = result.Status + "|" + result.Precedents.Count + "|" +
                result.UnresolvedEdgeCount + "|" + result.ExternalEdgeCount + "|" +
                (result.Precedents.Count == 0 ? "none" : result.Precedents[0].Classification.ToString());
            DiagnosticLog.Info("smoke.audit.precedents.direct", summary);
            return summary;
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.audit.precedents.direct", exception); throw; }
    }

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

    [ExcelCommand(Name = "ExcelAccel.Smoke.FormulaCopyDown", Description = "Debug-only transactional formula block hook.")]
    public static void FormulaCopyDown()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var command = new FormulaBlockCommand(FormulaCommandCatalog.GetRequired("formula.copy.down"));
            var plan = command.PlanCopy(port.CaptureFormulaBlock(), FormulaCopyDirection.Down);
            var result = command.Execute(plan, port, null, UndoRuntime.Store);
            DiagnosticLog.Info("smoke.formula.copy.down.result", result.Status + ":" + result.Message);
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
            DiagnosticLog.Info("smoke.formula.copy.down", "success");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.formula.copy.down", exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.FormulaTranspose", Description = "Debug-only off-selection formula transpose hook.")]
    public static void FormulaTranspose()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var destination = port.CaptureFormulaBlock();
            var sourceContext = new SelectionContext(destination.Selection.Context.WorkbookId,
                destination.Selection.Context.WorksheetName, "B20:C21");
            var source = port.CaptureFormulaBlock(sourceContext);
            var descriptor = new CommandDescriptor("formula.transpose", 1, "Formula Transpose", CommandImpact.Medium,
                new[] { "formula", "value" }, true, "smoke", "CAP-FORM-001",
                CommandContextRequirement.Selection, PreviewPolicy.Threshold, UndoPolicy.SessionPropertyReceipt,
                changedPropertyPolicy: ChangedPropertyPolicy.DeclaredSubset);
            var plan = new FormulaAdvancedCommand(descriptor).PlanTranspose(source, destination);
            var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, UndoRuntime.Store);
            DiagnosticLog.Info("smoke.formula.transpose.result", result.Status + ":" + result.Message);
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.formula.transpose", exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.DataCleaning", Description = "Debug-only transactional data-cleaning hook.")]
    public static void DataCleaning()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var trimDescriptor = DataCleaningCommandCatalog.GetRequired("clean.text.trim_outer");
            var trimPlan = new DataCleaningCommand(trimDescriptor).PlanTrimOuter(port.CaptureFormulaBlock());
            var trimResult = new FormulaBlockCommand(trimDescriptor).Execute(trimPlan, port, trimPlan.CommandPlan.PlanHash, UndoRuntime.Store);
            if (!trimResult.Succeeded) throw new InvalidOperationException(trimResult.Message);
            var zeroDescriptor = DataCleaningCommandCatalog.GetRequired("clean.display.zero_to_blank");
            var zeroPlan = new DataCleaningCommand(zeroDescriptor).PlanDisplayConversion(port.CaptureFormulaBlock(), DisplayValueConversion.ZeroToBlank);
            DiagnosticLog.Info("smoke.data.cleaning.zero_plan",
                zeroPlan.ChangedCount + ":" + zeroPlan.Before.Contents[0, 2].Kind + ":" + zeroPlan.Before.Contents[0, 2].InvariantValue + ":" + zeroPlan.After[0, 2].Kind);
            var zeroResult = new FormulaBlockCommand(zeroDescriptor).Execute(zeroPlan, port, zeroPlan.CommandPlan.PlanHash, UndoRuntime.Store);
            DiagnosticLog.Info("smoke.data.cleaning.result", trimResult.Status + ":" + zeroResult.Status);
            if (!zeroResult.Succeeded) throw new InvalidOperationException(zeroResult.Message);
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.data.cleaning", exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.SelectNumericHardcodes", Description = "Debug-only deterministic selection hook.")]
    public static void SelectNumericHardcodes()
    {
        try
        {
            var result = CommandDispatcher.ApplySelectionCommand("selection.select.numeric_hardcodes");
            DiagnosticLog.Info("smoke.selection.numeric_hardcodes", result.Status + ":" + result.Message);
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.selection.numeric_hardcodes", exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.TypedDataConversions", Description = "Debug-only typed conversion hook.")]
    public static void TypedDataConversions()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            ExecuteDataPlan("clean.convert.text_to_number", command =>
                command.PlanTextToNumber(port.CaptureFormulaBlock(), TextNumberConversionOptions.InvariantFinancial), port);
            ExecuteDataPlan("clean.convert.date_normalize", command =>
                command.PlanNormalizeDateText(port.CaptureFormulaBlock(), new[] { "yyyy-MM-dd", "yyyy/MM/dd", "yyyyMMdd" }, "yyyy-MM-dd"), port);
            ExecuteDataPlan("clean.convert.number_to_text", command =>
                command.PlanNumberToText(port.CaptureFormulaBlock(), "0.################"), port);
            DiagnosticLog.Info("smoke.data.typed_conversions", "success");
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.data.typed_conversions", exception); throw; }
    }

    private static void ExecuteDataPlan(string commandId, Func<DataCleaningCommand, FormulaBlockPlan> planFactory,
        IFormulaBlockPort port)
    {
        var descriptor = DataCleaningCommandCatalog.GetRequired(commandId);
        var plan = planFactory(new DataCleaningCommand(descriptor));
        var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, UndoRuntime.Store);
        if (!result.Succeeded) throw new InvalidOperationException(commandId + ": " + result.Message);
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.PasteValues", Description = "Debug-only values-only paste hook.")]
    public static void PasteValues()
    {
        ExecuteAdvanced("paste.values_only", (port, destination, descriptor) =>
        {
            var source = port.CaptureFormulaBlock(new SelectionContext(destination.Selection.Context.WorkbookId,
                destination.Selection.Context.WorksheetName, "A60:B60"));
            return new FormulaAdvancedCommand(descriptor).PlanPasteValues(source, destination);
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.FormulaFromAbove", Description = "Debug-only formula-from-above hook.")]
    public static void FormulaFromAbove()
    {
        ExecuteAdvanced("fill.formula_from_above", (port, destination, descriptor) =>
        {
            var source = port.CaptureFormulaBlock(new SelectionContext(destination.Selection.Context.WorkbookId,
                destination.Selection.Context.WorksheetName, "A70:B70"));
            return new FormulaAdvancedCommand(descriptor).PlanFormulaFromAbove(source, destination);
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.ValueFromAbove", Description = "Debug-only value-from-above hook.")]
    public static void ValueFromAbove()
    {
        ExecuteAdvanced("fill.value_from_above", (port, destination, descriptor) =>
        {
            var source = port.CaptureFormulaBlock(new SelectionContext(destination.Selection.Context.WorkbookId,
                destination.Selection.Context.WorksheetName, "A70:B70"));
            return new FormulaAdvancedCommand(descriptor).PlanValueFromAbove(source, destination);
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.NumericSequence", Description = "Debug-only numeric-sequence hook.")]
    public static void NumericSequence() => ExecuteAdvanced("fill.numeric_sequence", (port, destination, descriptor) =>
        new FormulaAdvancedCommand(descriptor).PlanNumericSequence(destination, 1, 2, SequenceFillDirection.Right));

    [ExcelCommand(Name = "ExcelAccel.Smoke.DateSequence", Description = "Debug-only date-sequence hook.")]
    public static void DateSequence() => ExecuteAdvanced("fill.date_sequence", (port, destination, descriptor) =>
        new FormulaAdvancedCommand(descriptor).PlanDateSequence(destination, new DateTime(2026, 8, 19), 7,
            SequenceFillDirection.Right, port.CaptureDateSystem()));

    private static void ExecuteAdvanced(string commandId,
        Func<ExcelSelectionAdapter, FormulaBlockSnapshot, CommandDescriptor, FormulaBlockPlan> planFactory)
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var descriptor = FormulaCommandCatalog.GetRequired(commandId);
            var plan = planFactory(port, port.CaptureFormulaBlock(), descriptor);
            var result = new FormulaBlockCommand(descriptor).Execute(plan, port, plan.CommandPlan.PlanHash, UndoRuntime.Store);
            DiagnosticLog.Info("smoke." + commandId, result.Status + ":" + result.Message);
            if (!result.Succeeded) throw new InvalidOperationException(commandId + ": " + result.Message);
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke." + commandId, exception); throw; }
    }

    [ExcelCommand(Name = "ExcelAccel.Smoke.PasteFormats", Description = "Debug-only formats-only paste hook.")]
    public static void PasteFormats()
    {
        try
        {
            var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var destination = port.CaptureFormatBlock();
            var source = port.CaptureFormatBlock(new SelectionContext(destination.Selection.Context.WorkbookId,
                destination.Selection.Context.WorksheetName, "A80"));
            var descriptor = FormulaCommandCatalog.GetRequired("paste.formats_only");
            var command = new ExcelAccel.Application.Formatting.FormatPasteCommand(descriptor);
            var plan = command.Plan(source, destination);
            var result = command.Execute(plan, port, plan.CommandPlan.PlanHash, UndoRuntime.Store);
            DiagnosticLog.Info("smoke.paste.formats_only", result.Status + ":" + result.Message);
            if (!result.Succeeded) throw new InvalidOperationException(result.Message);
        }
        catch (Exception exception) { DiagnosticLog.Error("smoke.paste.formats_only", exception); throw; }
    }
}
#endif
