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
        "Read-only direct precedents of one formula cell. This view never changes the workbook.",
        CommandDispatcher.NavigateToTraceTarget);

    public static bool IsOpen => Runtime.IsOpen;

    public static CommandResult Present(DirectPrecedentReport report, string workbookId, IWorkbookPresencePort presence) =>
        Runtime.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset() => Runtime.Reset();
}

internal static class TraceViewRuntimes
{
    private static readonly TraceViewRuntime Indirect = new TraceViewRuntime(
        AuditingCommandCatalog.IndirectPrecedentsId,
        "Read-only indirect trace within explicit depth and node caps. This view never changes the workbook.",
        CommandDispatcher.NavigateToTraceTarget);

    public static bool IsOpen => Indirect.IsOpen;

    public static CommandResult Present(IndirectTraceReport report, string workbookId, IWorkbookPresencePort presence) =>
        Indirect.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Indirect.RevalidateSource();

    public static void Reset() => Indirect.Reset();
}

internal static class InspectorViewRuntime
{
    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        AuditingCommandCatalog.InspectFormulaId,
        "Read-only structure of one formula. This view never changes the workbook and never evaluates anything.",
        CommandDispatcher.NavigateToTraceTarget);

    public static bool IsOpen => Runtime.IsOpen;

    public static CommandResult Present(FormulaInspectorReport report, string workbookId, IWorkbookPresencePort presence) =>
        Runtime.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset() => Runtime.Reset();
}

internal static class DependentViewRuntime
{
    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        AuditingCommandCatalog.DirectDependentsId,
        "Read-only direct dependents of one target within one worksheet. This view never changes the workbook.",
        CommandDispatcher.NavigateToTraceTarget);

    public static bool IsOpen => Runtime.IsOpen;

    public static CommandResult Present(DirectDependentReport report, string workbookId, IWorkbookPresencePort presence) =>
        Runtime.Present(report.ToPresentation(), workbookId, presence);

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset() => Runtime.Reset();
}
