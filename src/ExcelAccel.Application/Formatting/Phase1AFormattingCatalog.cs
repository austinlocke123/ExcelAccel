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
        Descriptor("format.font_color.cycle", "Font Color Cycle", "font_color", "Alt, X, F, C", "AC-FMT-001"),
        Descriptor("format.fill_color.cycle", "Fill Color Cycle", "fill_color", "Alt, X, F, I", "AC-FMT-001"),
        Descriptor("format.number.general", "General Number Format", "number_format", "Alt, X, N, G", "AC-FMT-002"),
        Descriptor("format.number.percentage", "Percentage Format", "number_format", "Alt, X, N, P", "AC-FMT-002"),
        Descriptor("format.number.multiple", "Multiple Format", "number_format", "Alt, X, N, M", "AC-FMT-002"),
        Descriptor("format.number.date", "Date Format", "number_format", "Alt, X, N, D", "AC-FMT-002"),
        Descriptor("format.number.boolean", "Boolean Format", "number_format", "Alt, X, N, B", "AC-FMT-002"),
        Descriptor("format.number.decimals.increase", "Increase Decimals", "number_format", "Alt, X, N, I", "AC-FMT-003"),
        Descriptor("format.number.decimals.decrease", "Decrease Decimals", "number_format", "Alt, X, N, R", "AC-FMT-003"),
        Descriptor("format.alignment.horizontal.cycle", "Horizontal Alignment Cycle", "horizontal_alignment", "Alt, X, A, H", "AC-FMT-001"),
        Descriptor("format.alignment.vertical.cycle", "Vertical Alignment Cycle", "vertical_alignment", "Alt, X, A, V", "AC-FMT-001"),
        Descriptor("format.indent.increase", "Increase Indent", "indent_level", "Alt, X, A, I", "AC-FMT-001"),
        Descriptor("format.indent.decrease", "Decrease Indent", "indent_level", "Alt, X, A, D", "AC-FMT-001"),
        Descriptor("format.underline.cycle", "Underline Cycle", "underline", "Alt, X, F, U", "AC-FMT-001"),
        Descriptor("format.font_size.cycle", "Font Size Cycle", "font_size", "Alt, X, F, S", "AC-FMT-001"),
        Descriptor("format.row_height.cycle", "Row Height Cycle", "row_height", "Alt, X, R, H", "AC-FMT-004"),
        Descriptor("format.column_width.cycle", "Column Width Cycle", "column_width", "Alt, X, C, W", "AC-FMT-004"),
        Descriptor("format.center_across.apply", "Center Across Selection", "horizontal_alignment", "Alt, X, A, C", "AC-FMT-005"),
        Descriptor("format.border.sum_bar.apply", "Apply Sum Bar", "borders", "Alt, X, B, S", "AC-FMT-006"),
        Descriptor("format.border.remove", "Remove Borders", "borders", "Alt, X, B, R", "AC-FMT-006"),
        Descriptor("format.rows.autofit", "AutoFit Rows", "row_height", "Alt, X, R, A", "AC-FMT-007"),
        Descriptor("format.columns.autofit", "AutoFit Columns", "column_width", "Alt, X, C, A", "AC-FMT-007"),
        Descriptor("view.gridlines.toggle", "Toggle Gridlines", "gridlines", "Alt, X, V, G", "AC-FMT-008"),
        Descriptor("view.zoom.set", "Set Zoom 100%", "zoom", "Alt, X, V, Z", "AC-FMT-008"),
        Descriptor("view.panes.freeze", "Freeze Panes at Selection", "freeze_panes", "Alt, X, V, F", "AC-FMT-009", CommandImpact.Medium),
        Descriptor("view.panes.unfreeze", "Unfreeze Panes", "freeze_panes", "Alt, X, V, U", "AC-FMT-009"),
    };

    public static IEnumerable<CommandDescriptor> All => Commands;

    public static ProfileFormattingCommand Create(string commandId)
    {
        var descriptor = Commands.SingleOrDefault(value => value.Id == commandId)
            ?? throw new KeyNotFoundException($"Formatting command '{commandId}' is not implemented.");
        Func<ProfileDefinition, string, string> resolver;
        Func<string, string, bool>? verifier = null;
        switch (commandId)
        {
            case "format.font_color.cycle": resolver = (p, c) => ProfileFormattingCommand.Next(p.FontColorCycle, c); break;
            case "format.fill_color.cycle": resolver = (p, c) => ProfileFormattingCommand.Next(p.FillColorCycle, c); break;
            case "format.font_size.cycle": resolver = (p, c) => NextNumber(p.FontSizeCycle, c); break;
            case "format.row_height.cycle": resolver = (p, c) => NextNumber(p.RowHeightCycle, c); break;
            case "format.column_width.cycle": resolver = (p, c) => NextNumber(p.ColumnWidthCycle, c); break;
            case "format.alignment.horizontal.cycle": resolver = (p, c) => ProfileFormattingCommand.Next(p.HorizontalAlignmentCycle, c); break;
            case "format.alignment.vertical.cycle": resolver = (p, c) => ProfileFormattingCommand.Next(p.VerticalAlignmentCycle, c); break;
            case "format.underline.cycle": resolver = (p, c) => ProfileFormattingCommand.Next(p.UnderlineCycle, c); break;
            case "format.indent.increase": resolver = (_, c) => ChangeInteger(c, 1); break;
            case "format.indent.decrease": resolver = (_, c) => ChangeInteger(c, -1); break;
            case "format.number.decimals.increase": resolver = (_, c) => NumberFormatDecimals.Change(c, 1); break;
            case "format.number.decimals.decrease": resolver = (_, c) => NumberFormatDecimals.Change(c, -1); break;
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
                var key = commandId.Substring("format.number.".Length);
                resolver = (p, _) => p.NumberFormats.TryGetValue(key, out var value) ? value : string.Empty;
                break;
        }

        return new ProfileFormattingCommand(descriptor, resolver, verifier);
    }

    private static string NextNumber(IReadOnlyList<double> values, string current) =>
        ProfileFormattingCommand.Next(values.Select(value => value.ToString("0.####", CultureInfo.InvariantCulture)).ToArray(), current);

    private static string ChangeInteger(string current, int delta) =>
        int.TryParse(current, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, Math.Min(15, value + delta)).ToString(CultureInfo.InvariantCulture)
            : string.Empty;

    private static bool IsPositiveNumber(string value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && parsed > 0;

    private static CommandDescriptor Descriptor(string id, string name, string property, string keyboard, string acceptance, CommandImpact impact = CommandImpact.Low) =>
        new CommandDescriptor(id, 1, name, impact, new[] { property }, true, $"Ribbon KeyTips: {keyboard}", "CAP-FMT-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            impact == CommandImpact.Medium ? PreviewPolicy.Mandatory : PreviewPolicy.None,
            UndoPolicy.SessionPropertyReceipt, new[] { acceptance, "AC-REL-005" });
}
