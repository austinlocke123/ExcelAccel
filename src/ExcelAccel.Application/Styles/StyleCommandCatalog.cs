using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Styles;

public static class StyleCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor("style.capture", 1, "Capture Local Style", CommandImpact.Low,
            new[] { "user_profile_styles" }, true, "Ribbon KeyTips: Alt, X, A, Y, C", "CAP-STYLE-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None, UndoPolicy.None, new[] { "AC-STYLE-001", "AC-STYLE-002", "AC-STYLE-003" },
            "Styles", "Capture explicitly selected formatting properties from one cell into the local profile.", shortcutLabel: "Alt, X, A, Y, C"),
        Apply("style.apply", "Apply Local Style", "Ribbon KeyTips: Alt, X, A, Y, L", "Alt, X, A, Y, L"),
        Apply("style.apply_builtin", "Apply Built-in Style", "Ribbon KeyTips: Alt, X, A, Y, B", "Alt, X, A, Y, B"),
        new CommandDescriptor("style.delete_local", 1, "Delete Local Style", CommandImpact.Low,
            new[] { "user_profile_styles" }, true, "Ribbon KeyTips: Alt, X, A, Y, D", "CAP-STYLE-001",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None, new[] { "AC-STYLE-008" },
            "Styles", "Delete one local style without reading or changing a workbook.", shortcutLabel: "Alt, X, A, Y, D"),
    };

    public static IEnumerable<CommandDescriptor> All => Commands;
    public static CommandDescriptor GetRequired(string commandId)
    {
        foreach (var command in Commands) if (command.Id == commandId) return command;
        throw new KeyNotFoundException($"Style command '{commandId}' is not registered.");
    }

    private static CommandDescriptor Apply(string id, string name, string route, string shortcut) =>
        new CommandDescriptor(id, 1, name, CommandImpact.Low, StylePropertyCatalog.Supported, true, route, "CAP-STYLE-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.Threshold, UndoPolicy.SessionPropertyReceipt,
            new[] { "AC-STYLE-004", "AC-STYLE-005", "AC-STYLE-006", "AC-STYLE-007", "AC-REL-005" },
            "Styles", "Apply only the formatting properties declared by a versioned style recipe.", shortcutLabel: shortcut,
            changedPropertyPolicy: ChangedPropertyPolicy.DeclaredSubset);
}
