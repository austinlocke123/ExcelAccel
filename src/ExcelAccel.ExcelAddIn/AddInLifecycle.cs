using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelAccel.Application.Reliability;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

public sealed class AddInLifecycle : IExcelAddIn
{
    public void AutoOpen()
    {
        CallbackBoundary.RunLifecycle("addin.open", () =>
        {
            RuntimeState.Start();
            DiagnosticLog.Info("addin.open", RuntimeState.IsSafeMode ? "safe_mode" : "normal");
        });
    }

    /// <summary>
    /// Tears the add-in down. Excel calls this when the user disables or removes
    /// the add-in while Excel keeps running; on process exit the runtime's own
    /// exit handlers clear the marker instead.
    /// </summary>
    public void AutoClose() => Shutdown();

    /// <summary>
    /// The teardown sequence, exposed so the Debug smoke can exercise it. Excel
    /// does not call <see cref="AutoClose"/> during a COM-automated quit, which
    /// is why this path went unexercised for 290 sessions.
    /// </summary>
    internal static void Shutdown()
    {
        CallbackBoundary.RunLifecycle("addin.close", () =>
        {
            RuntimeState.BeginShutdown();

            // Every reset runs even if an earlier one throws, and StopCleanly runs
            // regardless. Previously one failing reset skipped the rest and the
            // marker deletion with it, which left a stale *.running file and put
            // the user's next Excel session into safe mode, refusing every
            // mutation command for a reason that looked unrelated.
            var failures = LifecycleTeardown.Run(
                Resets(),
                RuntimeState.StopCleanly,
                (name, exception) => DiagnosticLog.Error("addin.close." + name, exception));
            DiagnosticLog.Info("addin.close", failures == 0 ? "normal" : $"partial:{failures}");
        });
    }

    private static IEnumerable<KeyValuePair<string, Action>> Resets()
    {
        yield return Reset("command_search", CommandSearchRuntime.Reset);
        yield return Reset("style_library", StyleLibraryRuntime.Reset);
        yield return Reset("profile", ProfileRuntime.Reset);
        yield return Reset("navigation", NavigationRuntime.Reset);
        yield return Reset("undo", UndoRuntime.Reset);
        yield return Reset("formula_source", FormulaSourceRuntime.Reset);
        yield return Reset("precedent_view", PrecedentViewRuntime.Reset);
        yield return Reset("dependent_view", DependentViewRuntime.Reset);
        yield return Reset("trace_views", TraceViewRuntimes.Reset);
        yield return Reset("inspector_view", InspectorViewRuntime.Reset);
        yield return Reset("model_check", ModelCheckRuntime.Reset);
        yield return Reset("name_inventory", NameInventoryViewRuntime.Reset);
    }

    private static KeyValuePair<string, Action> Reset(string name, Action action) =>
        new KeyValuePair<string, Action>(name, action);
}
