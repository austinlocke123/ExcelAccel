using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.SelectionTools;

public static class SelectionCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        Descriptor("selection.select.formulas", "Select Formulas", "FO", "Select cells containing underlying formulas."),
        Descriptor("selection.select.constants", "Select Constants", "CO", "Select nonblank constant cells and exclude formulas."),
        Descriptor("selection.select.blanks", "Select True Blanks", "BL", "Select truly empty cells; formulas returning empty text are excluded."),
        Descriptor("selection.select.numeric_hardcodes", "Select Numeric Hardcodes", "NH", "Select numeric constants and exclude formulas and numeric text."),
        Descriptor("selection.select.external_formulas", "Select External Formulas", "EX", "Select formulas with parsed external-workbook reference qualifiers."),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();
    public static IReadOnlyList<CommandDescriptor> All => Commands;
    public static CommandDescriptor GetRequired(string id) => Commands.First(value => value.Id == id);
    private static CommandDescriptor Descriptor(string id, string name, string keytip, string description) =>
        new CommandDescriptor(id, 1, name, CommandImpact.ReadOnly, Array.Empty<string>(), false,
            "Ribbon KeyTips: Alt, X, A, L, " + keytip, "CAP-SELECT-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None, UndoPolicy.None, new[] { "AC-SELECT-001", "AC-SELECT-002", "AC-SELECT-003", "AC-SELECT-004" },
            "Selection", description, shortcutLabel: "Alt, X, A, L, " + keytip);
}
