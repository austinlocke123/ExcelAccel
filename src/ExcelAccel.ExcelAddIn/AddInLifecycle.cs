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
            ProfileRuntime.Reset();
            NavigationRuntime.Reset();
            DiagnosticLog.Info("addin.close", "normal");
            RuntimeState.StopCleanly();
        });
    }
}
