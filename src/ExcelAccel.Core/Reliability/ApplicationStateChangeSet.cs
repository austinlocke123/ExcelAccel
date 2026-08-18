namespace ExcelAccel.Core.Reliability;

public sealed class ApplicationStateChangeSet
{
    public ApplicationStateChangeSet(bool suppressScreenUpdating, bool suppressEvents)
    {
        SuppressScreenUpdating = suppressScreenUpdating;
        SuppressEvents = suppressEvents;
    }

    public bool SuppressScreenUpdating { get; }

    public bool SuppressEvents { get; }

    public static ApplicationStateChangeSet PropertyMutation() => new ApplicationStateChangeSet(true, true);
}
