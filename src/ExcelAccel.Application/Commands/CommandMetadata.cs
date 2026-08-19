using System;

namespace ExcelAccel.Application.Commands;

[Flags]
public enum CommandContextRequirement
{
    None = 0,
    Workbook = 1,
    Worksheet = 2,
    Selection = 4,
    Window = 8,
    Application = 16,
}

public enum PreviewPolicy
{
    None = 0,
    Threshold = 1,
    Mandatory = 2,
}

public enum UndoPolicy
{
    None = 0,
    SessionPropertyReceipt = 1,
}

public enum ChangedPropertyPolicy
{
    Exact = 0,
    DeclaredSubset = 1,
}

public enum CommandResultStatus
{
    Success = 0,
    Refused = 1,
    Cancelled = 2,
    Failed = 3,
    Partial = 4,
}

public enum InvocationSource
{
    Ribbon = 0,
    Shortcut = 1,
    Search = 2,
    Favorite = 3,
    Pane = 4,
    Test = 5,
}
