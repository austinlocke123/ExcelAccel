using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// The two auditing views. Each owns its own window and lifecycle, and both run
/// through the one shared <see cref="TraceViewRuntime"/>.
/// </summary>
internal static class PrecedentViewRuntime
{
    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        AuditingCommandCatalog.DirectPrecedentsId,
        "Read-only direct precedents of one formula cell. This view never changes the workbook.");

    public static bool IsOpen => Runtime.IsOpen;

    public static CommandResult Present(DirectPrecedentReport report, string workbookId, IWorkbookPresencePort presence) =>
        Runtime.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset() => Runtime.Reset();
}

internal static class DependentViewRuntime
{
    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        AuditingCommandCatalog.DirectDependentsId,
        "Read-only direct dependents of one target within one worksheet. This view never changes the workbook.");

    public static bool IsOpen => Runtime.IsOpen;

    public static CommandResult Present(DirectDependentReport report, string workbookId, IWorkbookPresencePort presence) =>
        Runtime.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset() => Runtime.Reset();
}
