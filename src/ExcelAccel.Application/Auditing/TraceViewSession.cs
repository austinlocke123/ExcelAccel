using System;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Application.Auditing;

public enum TraceViewAction
{
    /// <summary>The presented result is still backed by an open workbook.</summary>
    Keep,

    /// <summary>The workbook state could not be verified; keep the result under an explicit notice.</summary>
    Warn,

    /// <summary>The source workbook is gone; the result has been discarded.</summary>
    Discard,
}

/// <summary>
/// The lifecycle state of one presented trace result: what is on screen, which
/// workbook it came from, and whether it may still be shown.
///
/// This deliberately lives outside the host view. The view itself is WinForms
/// over COM and cannot be reached from the test project, so keeping the decision
/// here is what makes the lifecycle testable at all; the view is left as a
/// renderer that does what the session decides.
/// </summary>
public sealed class TraceViewSession
{
    public const string UnverifiedNotice =
        "The source workbook state could not be verified; this result is a point-in-time capture.";

    private readonly string _commandId;
    private IWorkbookPresencePort? _presence;
    private bool _revalidating;

    public TraceViewSession(string commandId)
    {
        _commandId = !string.IsNullOrWhiteSpace(commandId)
            ? commandId
            : throw new ArgumentException("A command ID is required.", nameof(commandId));
    }

    public string CommandId => _commandId;

    public TraceResultPresentation? Presentation { get; private set; }

    public string WorkbookId { get; private set; } = string.Empty;

    /// <summary>Set only while the workbook state could not be verified.</summary>
    public string? Notice { get; private set; }

    /// <summary>The last probe failure, for the host to log. Cleared on each probe.</summary>
    public Exception? LastProbeError { get; private set; }

    public bool HasResult => Presentation is not null;

    public void Present(TraceResultPresentation presentation, string workbookId, IWorkbookPresencePort presence)
    {
        Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
        _presence = presence ?? throw new ArgumentNullException(nameof(presence));
        WorkbookId = workbookId ?? string.Empty;
        Notice = null;
        LastProbeError = null;
    }

    /// <summary>
    /// Re-probes the source workbook. A closed workbook discards the result so a
    /// stale trace can never stay on screen; an unverifiable one is kept under an
    /// explicit notice rather than being claimed as live.
    /// </summary>
    public TraceViewAction Revalidate()
    {
        if (_presence is null || !HasResult || string.IsNullOrEmpty(WorkbookId) || _revalidating)
        {
            return TraceViewAction.Keep;
        }

        WorkbookPresence presence;
        LastProbeError = null;
        _revalidating = true;
        try
        {
            presence = _presence.Probe(WorkbookId);
        }
        catch (Exception exception)
        {
            LastProbeError = exception;
            presence = WorkbookPresence.Unknown;
        }
        finally
        {
            _revalidating = false;
        }

        switch (presence)
        {
            case WorkbookPresence.Closed:
                Clear();
                return TraceViewAction.Discard;
            case WorkbookPresence.Unknown:
                Notice = UnverifiedNotice;
                return TraceViewAction.Warn;
            default:
                Notice = null;
                return TraceViewAction.Keep;
        }
    }

    public void Clear()
    {
        Presentation = null;
        _presence = null;
        WorkbookId = string.Empty;
        Notice = null;
    }
}
