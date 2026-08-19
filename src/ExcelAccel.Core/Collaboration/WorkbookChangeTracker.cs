using System;

namespace ExcelAccel.Core.Collaboration;

public sealed class WorkbookChangeTracker
{
    private readonly object _sync = new object();
    private readonly string _workbookIdentity;
    private long _revision;
    private bool _remoteChangeInProgress;
    private bool _remoteChangeObserved;
    private bool _workbookClosed;

    public WorkbookChangeTracker(string workbookIdentity)
    {
        if (string.IsNullOrWhiteSpace(workbookIdentity))
        {
            throw new ArgumentException("A workbook identity is required.", nameof(workbookIdentity));
        }

        _workbookIdentity = workbookIdentity;
    }

    public void Record(WorkbookChangeKind changeKind)
    {
        if (!Enum.IsDefined(typeof(WorkbookChangeKind), changeKind))
        {
            throw new ArgumentOutOfRangeException(nameof(changeKind));
        }

        lock (_sync)
        {
            ThrowIfClosed();
            _revision = checked(_revision + 1);

            if (changeKind == WorkbookChangeKind.RemoteChangeBeginning)
            {
                _remoteChangeInProgress = true;
                _remoteChangeObserved = true;
            }
            else if (changeKind == WorkbookChangeKind.RemoteChangeCompleted)
            {
                _remoteChangeInProgress = false;
                _remoteChangeObserved = true;
            }
        }
    }

    public void Close()
    {
        lock (_sync)
        {
            if (_workbookClosed)
            {
                return;
            }

            _revision = checked(_revision + 1);
            _workbookClosed = true;
            _remoteChangeInProgress = false;
        }
    }

    public WorkbookConcurrencyStamp Capture(
        string preconditionFingerprint,
        WorkbookCollaborationState detectedState)
    {
        if (preconditionFingerprint is null)
        {
            throw new ArgumentNullException(nameof(preconditionFingerprint));
        }

        if (detectedState is null)
        {
            throw new ArgumentNullException(nameof(detectedState));
        }

        lock (_sync)
        {
            var effectiveCoauthoring = _remoteChangeInProgress
                ? CoauthoringState.RemoteChangeInProgress
                : _remoteChangeObserved
                    ? CoauthoringState.RemoteChangeObserved
                    : detectedState.Coauthoring;

            var effectiveState = new WorkbookCollaborationState(
                detectedState.AutoSave,
                effectiveCoauthoring,
                detectedState.RemoteChangeEventsSupported);

            return new WorkbookConcurrencyStamp(
                _workbookIdentity,
                _revision,
                preconditionFingerprint,
                effectiveState,
                _workbookClosed);
        }
    }

    private void ThrowIfClosed()
    {
        if (_workbookClosed)
        {
            throw new InvalidOperationException("A closed workbook tracker cannot accept new events.");
        }
    }
}
