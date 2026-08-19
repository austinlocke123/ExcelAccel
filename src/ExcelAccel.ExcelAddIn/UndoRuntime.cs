using ExcelAccel.Application.Undo;

namespace ExcelAccel.ExcelAddIn;

internal static class UndoRuntime
{
    public static SessionUndoStore Store { get; } = new SessionUndoStore();
    public static void Reset() => Store.ClearAll();
}
