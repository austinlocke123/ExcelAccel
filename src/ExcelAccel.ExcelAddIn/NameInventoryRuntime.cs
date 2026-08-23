using System;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Names;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns the named-range inventory view and the snapshot behind it.
/// </summary>
/// <remarks>
/// The view itself is the shared <see cref="TraceViewRuntime"/>, so no second
/// read-only window exists to drift from the auditing ones.
///
/// The captured inventory is retained so search and export work from the same
/// snapshot the user is looking at. Re-reading the workbook for either would let
/// results change under a filter the user did not touch.
/// </remarks>
internal static class NameInventoryViewRuntime
{
    private static readonly object Sync = new object();

    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        NamesCommandCatalog.InventoryId,
        "Read-only inventory of defined names. This view never changes the workbook, and never opens an external source.",
        CommandDispatcher.NavigateToTraceTarget);

    private static NameInventoryResult? _captured;

    public static bool IsOpen => Runtime.IsOpen;

    public static NameInventoryResult? Captured
    {
        get { lock (Sync) return _captured; }
    }

    public static CommandResult Present(NameInventoryResult result, IWorkbookPresencePort presence)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        lock (Sync) _captured = result;
        return Runtime.Present(result.Presentation, result.WorkbookId, presence);
    }

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset()
    {
        Runtime.Reset();
        lock (Sync) _captured = null;
    }
}
