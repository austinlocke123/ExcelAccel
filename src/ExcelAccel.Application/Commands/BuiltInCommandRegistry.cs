using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Navigation;
using ExcelAccel.Application.Undo;
using ExcelAccel.Application.Discovery;
using ExcelAccel.Application.Styles;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.DataCleaning;
using ExcelAccel.Application.SelectionTools;
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
            RibbonRoutes.For(InspectSelectionCommand.Id),
            "CAP-CMD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-CMD-001", "AC-CMD-002" }),
        new CommandDescriptor(
            UndoLastCommand.Id,
            1,
            "Undo Last ExcelAccel Property",
            CommandImpact.Low,
            new[] { "qualified_receipt_property" },
            true,
            RibbonRoutes.For(UndoLastCommand.Id),
            "CAP-REL-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-REL-011", "AC-REL-012" }),
        new CommandDescriptor(
            "support.diagnostics.export",
            1,
            "Export Sanitized Diagnostics",
            CommandImpact.Medium,
            new[] { "local_support_file" },
            false,
            RibbonRoutes.For("support.diagnostics.export"),
            "CAP-REL-001",
            CommandContextRequirement.Application,
            PreviewPolicy.Mandatory,
            UndoPolicy.None,
            new[] { "AC-SEC-002", "AC-SEC-004", "AC-UX-005" }),
        }
        .Concat(Phase1AFormattingCatalog.All)
        .Concat(NavigationCommandCatalog.All)
        .Concat(DiscoveryCommandCatalog.All)
        .Concat(StyleCommandCatalog.All)
        .Concat(ProfileExchangeCommandCatalog.All)
        .Concat(FormulaCommandCatalog.All)
        .Concat(DataCleaningCommandCatalog.All)
        .Concat(SelectionCommandCatalog.All)
        .Concat(AuditingCommandCatalog.All)
        .Concat(ModelCheckCommandCatalog.All)
        .OrderBy(command => command.Id, System.StringComparer.Ordinal)
        .ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    public static CommandDescriptor GetRequired(string commandId)
    {
        if (string.IsNullOrWhiteSpace(commandId))
        {
            throw new System.ArgumentException("A command ID is required.", nameof(commandId));
        }

        foreach (var command in Commands)
        {
            if (string.Equals(command.Id, commandId, System.StringComparison.Ordinal))
            {
                return command;
            }
        }

        throw new System.Collections.Generic.KeyNotFoundException($"Command '{commandId}' is not registered.");
    }
}
