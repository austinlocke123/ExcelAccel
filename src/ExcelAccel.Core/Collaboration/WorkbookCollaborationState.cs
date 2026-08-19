using System;

namespace ExcelAccel.Core.Collaboration;

public enum AutoSaveState
{
    Unknown,
    OffOrDisabled,
    On
}

public enum CoauthoringState
{
    Unknown,
    NotDetected,
    PotentialModern,
    LegacyShared,
    RemoteChangeObserved,
    RemoteChangeInProgress
}

public enum WorkbookLocationKind
{
    Unknown,
    Unsaved,
    LocalOrSyncedPath,
    CloudUrl
}

public sealed class WorkbookCollaborationProbe
{
    public WorkbookCollaborationProbe(
        bool? autoSaveOn,
        bool? legacyShared,
        WorkbookLocationKind location,
        bool remoteChangeEventsHooked)
    {
        if (!Enum.IsDefined(typeof(WorkbookLocationKind), location))
        {
            throw new ArgumentOutOfRangeException(nameof(location));
        }

        AutoSaveOn = autoSaveOn;
        LegacyShared = legacyShared;
        Location = location;
        RemoteChangeEventsHooked = remoteChangeEventsHooked;
    }

    public bool? AutoSaveOn { get; }

    public bool? LegacyShared { get; }

    public WorkbookLocationKind Location { get; }

    public bool RemoteChangeEventsHooked { get; }
}

public static class WorkbookCollaborationClassifier
{
    public static WorkbookCollaborationState Classify(WorkbookCollaborationProbe probe)
    {
        if (probe is null)
        {
            throw new ArgumentNullException(nameof(probe));
        }

        var autoSave = probe.AutoSaveOn.HasValue
            ? probe.AutoSaveOn.Value ? AutoSaveState.On : AutoSaveState.OffOrDisabled
            : AutoSaveState.Unknown;

        CoauthoringState coauthoring;
        if (!probe.LegacyShared.HasValue)
        {
            coauthoring = CoauthoringState.Unknown;
        }
        else if (probe.LegacyShared.Value)
        {
            coauthoring = CoauthoringState.LegacyShared;
        }
        else if (probe.AutoSaveOn == true || probe.Location == WorkbookLocationKind.CloudUrl)
        {
            coauthoring = CoauthoringState.PotentialModern;
        }
        else if (probe.Location == WorkbookLocationKind.Unsaved)
        {
            coauthoring = CoauthoringState.NotDetected;
        }
        else
        {
            // A local-looking path may be a synchronized cloud path with AutoSave
            // manually off. No qualified read-only property proves exclusivity.
            coauthoring = CoauthoringState.Unknown;
        }

        return new WorkbookCollaborationState(
            autoSave,
            coauthoring,
            probe.RemoteChangeEventsHooked);
    }
}

public sealed class WorkbookCollaborationState
{
    public WorkbookCollaborationState(
        AutoSaveState autoSave,
        CoauthoringState coauthoring,
        bool remoteChangeEventsSupported)
    {
        if (!Enum.IsDefined(typeof(AutoSaveState), autoSave))
        {
            throw new ArgumentOutOfRangeException(nameof(autoSave));
        }

        if (!Enum.IsDefined(typeof(CoauthoringState), coauthoring))
        {
            throw new ArgumentOutOfRangeException(nameof(coauthoring));
        }

        AutoSave = autoSave;
        Coauthoring = coauthoring;
        RemoteChangeEventsSupported = remoteChangeEventsSupported;
    }

    public AutoSaveState AutoSave { get; }

    public CoauthoringState Coauthoring { get; }

    public bool RemoteChangeEventsSupported { get; }

    public static WorkbookCollaborationState Unknown { get; } =
        new WorkbookCollaborationState(AutoSaveState.Unknown, CoauthoringState.Unknown, false);
}
