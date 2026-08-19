using System.Collections.Generic;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public static class BuiltInCommandRegistry
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor(
            InspectSelectionCommand.Id,
            1,
            "Inspect Selection",
            CommandImpact.ReadOnly,
            new string[0],
            true,
            "Ribbon KeyTips: Alt, X, A, I"),
        new CommandDescriptor(
            ApplyCurrencyFormatCommand.Id,
            1,
            "Currency Format",
            CommandImpact.Low,
            new[] { ApplyCurrencyFormatCommand.ChangedProperty },
            true,
            "Ribbon KeyTips: Alt, X, A, C"),
    };

    public static IReadOnlyList<CommandDescriptor> All => Commands;
}
