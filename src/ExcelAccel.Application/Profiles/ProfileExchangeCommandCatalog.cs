using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Profiles;

public static class ProfileExchangeCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor("profile.export", 1, "Export Profile", CommandImpact.Medium,
            new[] { "local_profile_package_file" }, true, "Ribbon KeyTips: Alt, X, A, P, E", "CAP-PROF-002",
            CommandContextRequirement.Application, PreviewPolicy.Mandatory, UndoPolicy.None,
            new[] { "AC-PROF-003", "AC-PROF-004", "AC-PROF-005", "AC-PROF-006" }, "Profiles",
            "Export the declarative local profile to an explicit local package after manifest confirmation.", shortcutLabel: "Alt, X, A, P, E"),
        new CommandDescriptor("profile.import.preview", 1, "Preview Profile Import", CommandImpact.ReadOnly,
            new string[0], true, "Ribbon KeyTips: Alt, X, A, P, V", "CAP-PROF-002",
            CommandContextRequirement.Application, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-PROF-002", "AC-PROF-004", "AC-PROF-007" }, "Profiles",
            "Validate a local profile package and show its deterministic diff without changing settings.", shortcutLabel: "Alt, X, A, P, V"),
        new CommandDescriptor("profile.import.apply", 1, "Import Profile", CommandImpact.Medium,
            new[] { "user_profile" }, true, "Ribbon KeyTips: Alt, X, A, P, I", "CAP-PROF-002",
            CommandContextRequirement.Application, PreviewPolicy.Mandatory, UndoPolicy.None,
            new[] { "AC-PROF-002", "AC-PROF-004", "AC-PROF-007", "AC-PROF-008", "AC-PROF-009" }, "Profiles",
            "Preview, validate, back up, and atomically activate a local profile package.", shortcutLabel: "Alt, X, A, P, I"),
        new CommandDescriptor("bindings.cheat_sheet.export", 1, "Export Shortcut Cheat Sheet", CommandImpact.Medium,
            new[] { "local_binding_cheat_sheet_file" }, true, "Ribbon KeyTips: Alt, X, A, P, B", "CAP-PROF-002",
            CommandContextRequirement.Application, PreviewPolicy.Mandatory, UndoPolicy.None,
            new[] { "AC-KEY-004", "AC-PROF-005", "AC-PROF-006" }, "Profiles",
            "Export current Quick Keys, command names, categories, and conflicts to local CSV or HTML.", shortcutLabel: "Alt, X, A, P, B"),
    };
    public static IEnumerable<CommandDescriptor> All => Commands;
}
