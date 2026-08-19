using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Discovery;

public static class DiscoveryCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor("command.search.open", 1, "Search Commands", CommandImpact.ReadOnly,
            new string[0], true, "Ribbon KeyTips: Alt, X, A, Q", "CAP-SEARCH-001",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-SEARCH-001", "AC-SEARCH-002", "AC-SEARCH-003", "AC-SEARCH-004" },
            "Discovery", "Find and run a registered ExcelAccel command without scanning the workbook.",
            new[] { "command palette", "find command" }, "Alt, X, A, Q"),
        new CommandDescriptor("favorite.add", 1, "Add Favorite", CommandImpact.Low,
            new[] { "user_profile_favorites" }, true, "Command Search: Ctrl+D", "CAP-FAV-001",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-FAV-001", "AC-FAV-003" }, "Discovery",
            "Add the selected command and fixed arguments to the local profile.", shortcutLabel: "Ctrl+D"),
        new CommandDescriptor("favorite.remove", 1, "Remove Favorite", CommandImpact.Low,
            new[] { "user_profile_favorites" }, true, "Command Search: Ctrl+Shift+D", "CAP-FAV-001",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-FAV-001", "AC-FAV-003" }, "Discovery",
            "Remove the selected favorite from the local profile.", shortcutLabel: "Ctrl+Shift+D"),
        new CommandDescriptor("favorite.invoke", 1, "Invoke Favorite", CommandImpact.ReadOnly,
            new string[0], true, "Command Search: Enter", "CAP-FAV-001",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-CMD-002", "AC-FAV-002", "AC-FAV-003", "AC-FAV-004" }, "Discovery",
            "Resolve the current command contract and route through its normal lifecycle.", shortcutLabel: "Enter",
            inheritsReferencedPolicy: true),
    };

    public static IEnumerable<CommandDescriptor> All => Commands;
}
