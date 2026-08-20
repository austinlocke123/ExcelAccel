using System;
using ExcelDna.Integration;
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

    public void AutoClose()
    {
        CallbackBoundary.RunLifecycle("addin.close", () =>
        {
            CommandSearchRuntime.Reset();
            StyleLibraryRuntime.Reset();
            ProfileRuntime.Reset();
            NavigationRuntime.Reset();
            UndoRuntime.Reset();
            FormulaSourceRuntime.Reset();
            PrecedentViewRuntime.Reset();
            DependentViewRuntime.Reset();
            TraceViewRuntimes.Reset();
            DiagnosticLog.Info("addin.close", "normal");
            RuntimeState.StopCleanly();
        });
    }
}
