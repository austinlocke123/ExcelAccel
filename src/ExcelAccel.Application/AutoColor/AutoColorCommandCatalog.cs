using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.AutoColor;

public static class AutoColorCommandCatalog
{
    public const string SelectionId = "format.auto_color.selection";
    public const string WorksheetId = "format.auto_color.worksheet";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        Descriptor(
            SelectionId,
            "AutoColor Selection",
            "Recolour the selected cells by what they are: typed numbers, formulas, cross-sheet and external references, and errors. Only the font colour changes, and one undo reverses the whole change.",
            new[] { "autocolor", "auto color", "blue black", "colour inputs" },
            PreviewPolicy.Threshold),
        Descriptor(
            WorksheetId,
            "AutoColor Worksheet",
            "Recolour the active worksheet's used range by what each cell is. Always previewed, and remains unavailable until worksheet-scale qualification passes.",
            new[] { "autocolor sheet", "auto color worksheet" },
            PreviewPolicy.Mandatory),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    public static AutoColorScope ScopeFor(string commandId) =>
        string.Equals(commandId, WorksheetId, StringComparison.Ordinal)
            ? AutoColorScope.Worksheet
            : AutoColorScope.Selection;

    public static AutoColorCommand Create(string commandId) =>
        new AutoColorCommand(
            Commands.SingleOrDefault(value => value.Id == commandId)
                ?? throw new KeyNotFoundException($"AutoColor command '{commandId}' is not implemented."),
            ScopeFor(commandId));

    private static CommandDescriptor Descriptor(
        string id,
        string name,
        string description,
        IEnumerable<string> aliases,
        PreviewPolicy previewPolicy) =>
        new CommandDescriptor(
            id,
            1,
            name,
            CommandImpact.Medium,
            // The whole recolour is one coarse property, which is what lets a
            // single undo reverse it; see FontColorBlock.
            new[] { FontColorBlock.ReceiptPropertyId },
            true,
            RibbonRoutes.For(id),
            "CAP-FMT-002",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            previewPolicy,
            UndoPolicy.SessionPropertyReceipt,
            new[] { "AC-FMT-034", "AC-FMT-035", "AC-FMT-036", "AC-FMT-037", "AC-FMT-046", "AC-REL-005" },
            "Formatting",
            description,
            aliases,
            RibbonRoutes.For(id));
}
