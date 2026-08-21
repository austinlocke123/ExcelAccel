using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Formatting;

public static class Phase1AFormattingCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        Descriptor("format.font_color.cycle", "Font Color Cycle", "font_color", "FC", "AC-FMT-001"),
        Descriptor("format.fill_color.cycle", "Fill Color Cycle", "fill_color", "FI", "AC-FMT-001"),
        Descriptor("format.number.general", "General Number Format", "number_format", "NG", "AC-FMT-002"),
        Descriptor("format.number.currency", "Currency Format", "number_format", "NC", "AC-FMT-002"),
        Descriptor("format.number.percentage", "Percentage Format", "number_format", "NP", "AC-FMT-002"),
        Descriptor("format.number.multiple", "Multiple Format", "number_format", "NM", "AC-FMT-002"),
        Descriptor("format.number.date", "Date Format", "number_format", "NT", "AC-FMT-002"),
        Descriptor("format.number.boolean", "Boolean Format", "number_format", "NB", "AC-FMT-002"),
        Descriptor("format.number.decimals.increase", "Increase Decimals", "number_format", "NI", "AC-FMT-003"),
        Descriptor("format.number.decimals.decrease", "Decrease Decimals", "number_format", "ND", "AC-FMT-003"),
        Descriptor("format.alignment.horizontal.cycle", "Horizontal Alignment Cycle", "horizontal_alignment", "AH", "AC-FMT-001"),
        Descriptor("format.alignment.vertical.cycle", "Vertical Alignment Cycle", "vertical_alignment", "AV", "AC-FMT-001"),
        Descriptor("format.indent.increase", "Increase Indent", "indent_level", "II", "AC-FMT-001"),
        Descriptor("format.indent.decrease", "Decrease Indent", "indent_level", "ID", "AC-FMT-001"),
        Descriptor("format.underline.cycle", "Underline Cycle", "underline", "FU", "AC-FMT-001"),
        Descriptor("format.font_size.cycle", "Font Size Cycle", "font_size", "FS", "AC-FMT-001"),
        Descriptor("format.row_height.cycle", "Row Height Cycle", "row_height", "RH", "AC-FMT-004"),
        Descriptor("format.column_width.cycle", "Column Width Cycle", "column_width", "CW", "AC-FMT-004"),
        Descriptor("format.center_across.apply", "Center Across Selection", "horizontal_alignment", "CA", "AC-FMT-005"),
        Descriptor("format.border.sum_bar.apply", "Apply Sum Bar", "borders", "BS", "AC-FMT-006"),
        Descriptor("format.border.remove", "Remove Borders", "borders", "BR", "AC-FMT-006"),
        Descriptor("format.rows.autofit", "AutoFit Rows", "row_height", "AR", "AC-FMT-007"),
        Descriptor("format.columns.autofit", "AutoFit Columns", "column_width", "AC", "AC-FMT-007"),
        Descriptor("view.gridlines.toggle", "Toggle Gridlines", "gridlines", "VG", "AC-FMT-008"),
        Descriptor("view.zoom.set", "Set Zoom 100%", "zoom", "VZ", "AC-FMT-008"),
        Descriptor("view.panes.freeze", "Freeze Panes at Selection", "freeze_panes", "VF", "AC-FMT-009", CommandImpact.Medium),
        Descriptor("view.panes.unfreeze", "Unfreeze Panes", "freeze_panes", "VU", "AC-FMT-009"),
    };

    public static IEnumerable<CommandDescriptor> All => Commands;

    public static ProfileFormattingCommand Create(string commandId)
    {
        var descriptor = Commands.SingleOrDefault(value => value.Id == commandId)
            ?? throw new KeyNotFoundException($"Formatting command '{commandId}' is not implemented.");
        Func<ProfileDefinition, string, string> resolver;
        Func<string, string, bool>? verifier = null;
        // Set where an empty resolver result means "nothing to change" rather
        // than a broken profile, so the user is told the truth.
        string? noChange = null;
        // Set where an empty resolver result means the named cycle is not
        // configured, so the refusal can name the cycle the user asked for.
        string? unconfigured = null;
        switch (commandId)
        {
            case "format.font_color.cycle":
            case "format.fill_color.cycle":
            case "format.font_size.cycle":
            case "format.row_height.cycle":
            case "format.column_width.cycle":
            case "format.alignment.horizontal.cycle":
            case "format.alignment.vertical.cycle":
            case "format.underline.cycle":
            {
                // These commands name no particular cycle on the ribbon, so they
                // follow whichever cycle the user has placed first in the family.
                var family = descriptor.ChangedProperties[0];
                resolver = (p, c) => ProfileFormattingCommand.Next(p.ResolveFirstCycle(family), c);
                unconfigured = $"No {family.Replace('_', ' ')} cycle is configured in the active profile.";
                break;
            }
            case "format.indent.increase":
                resolver = (_, c) => ChangeInteger(c, 1);
                noChange = "This cell is already at the maximum indent level.";
                break;
            case "format.indent.decrease":
                resolver = (_, c) => ChangeInteger(c, -1);
                noChange = "This cell has no indent left to remove.";
                break;
            case "format.number.decimals.increase":
                resolver = (_, c) => NumberFormatDecimals.Change(c, 1);
                noChange = "This number format cannot take another decimal place.";
                break;
            case "format.number.decimals.decrease":
                resolver = (_, c) => NumberFormatDecimals.Change(c, -1);
                noChange = "This number format has no decimal places left to remove.";
                break;
            case "format.center_across.apply": resolver = (_, __) => "center_across"; break;
            case "format.border.sum_bar.apply": resolver = (_, __) => "sum_bar"; break;
            case "format.border.remove": resolver = (_, __) => "none"; break;
            case "format.rows.autofit": resolver = (_, __) => "autofit"; verifier = (_, observed) => IsPositiveNumber(observed); break;
            case "format.columns.autofit": resolver = (_, __) => "autofit"; verifier = (_, observed) => IsPositiveNumber(observed); break;
            case "view.gridlines.toggle": resolver = (_, c) => string.Equals(c, "true", StringComparison.OrdinalIgnoreCase) ? "false" : "true"; break;
            case "view.zoom.set": resolver = (_, __) => "100"; break;
            case "view.panes.freeze": resolver = (_, __) => "true"; break;
            case "view.panes.unfreeze": resolver = (_, __) => "false"; break;
            default:
            {
                // Number-format commands are named on the ribbon, so each binds to
                // the cycle carrying its own identifier rather than to a slot.
                var cycleId = commandId.Substring("format.number.".Length);
                resolver = (p, c) => ProfileFormattingCommand.Next(p.ResolveCycle("number_format", cycleId), c);
                unconfigured = $"The '{cycleId}' number-format cycle is not configured in the active profile.";
                break;
            }
        }

        return new ProfileFormattingCommand(descriptor, resolver, verifier, noChange, unconfigured);
    }

    private static string ChangeInteger(string current, int delta) =>
        int.TryParse(current, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, Math.Min(15, value + delta)).ToString(CultureInfo.InvariantCulture)
            : string.Empty;

    private static bool IsPositiveNumber(string value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && parsed > 0;

    private static CommandDescriptor Descriptor(string id, string name, string property, string keyboard, string acceptance, CommandImpact impact = CommandImpact.Low) =>
        new CommandDescriptor(id, 1, name, impact, new[] { property }, true, RibbonRoutes.For(id), "CAP-FMT-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            impact == CommandImpact.Medium ? PreviewPolicy.Mandatory : PreviewPolicy.None,
            UndoPolicy.SessionPropertyReceipt, new[] { acceptance, "AC-REL-005" }, "Formatting",
            $"Apply {name.ToLowerInvariant()} using the active profile.", shortcutLabel: RibbonRoutes.For(id));
}
