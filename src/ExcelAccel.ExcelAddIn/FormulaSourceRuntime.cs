using System;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.ExcelInterop;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn;

internal static class FormulaSourceRuntime
{
    private static readonly TimeSpan MaximumAge = TimeSpan.FromMinutes(30);
    public static FormulaSourceStore Store { get; } = new FormulaSourceStore();

    public static CommandResult Capture()
    {
        var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, Reliability.RuntimeState.VerifyExcelThread);
        var snapshot = port.CaptureFormulaBlock();
        Store.Capture(snapshot, DateTimeOffset.UtcNow);
        return CommandResult.Success("formula.source.capture",
            $"Captured {snapshot.Contents.RowCount:N0}×{snapshot.Contents.ColumnCount:N0} internal formula source for this Excel session.");
    }

    public static bool TryGet(out FormulaBlockSnapshot? snapshot, out string reason) =>
        Store.TryGet(DateTimeOffset.UtcNow, MaximumAge, out snapshot, out reason);

    public static void Reset() => Store.Clear();
}
