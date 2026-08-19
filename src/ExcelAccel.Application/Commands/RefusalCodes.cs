namespace ExcelAccel.Application.Commands;

public static class RefusalCodes
{
    public const string SelectionUnsupported = "SELECTION_UNSUPPORTED";
    public const string MultiAreaUnsupported = "MULTI_AREA_UNSUPPORTED";
    public const string ProtectedTarget = "PROTECTED_TARGET";
    public const string ReadOnlyWorkbook = "READ_ONLY_WORKBOOK";
    public const string ArrayOrSpillUnsafe = "ARRAY_OR_SPILL_UNSAFE";
    public const string StaleContext = "STALE_CONTEXT";
    public const string ResourceLimit = "RESOURCE_LIMIT";
    public const string ExcelCapabilityMissing = "EXCEL_CAPABILITY_MISSING";
    public const string CommandQuarantined = "COMMAND_QUARANTINED";
}
