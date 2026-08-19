namespace ExcelAccel.Core.Collaboration;

public enum WorkbookChangeKind
{
    LocalEdit,
    RemoteChangeBeginning,
    RemoteChangeCompleted,
    StructureChanged,
    Recalculated,
    SaveCompleted,
    ProtectionChanged,
    ReadOnlyChanged,
    AutoSaveStateChanged
}
