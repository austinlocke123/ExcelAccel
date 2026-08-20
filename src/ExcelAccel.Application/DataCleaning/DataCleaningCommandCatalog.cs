using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.DataCleaning;

public static class DataCleaningCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        Descriptor("clean.text.trim_outer", "Trim Outer Whitespace", "TO", PreviewPolicy.Threshold, "Trim only approved leading/trailing Unicode whitespace."),
        Descriptor("clean.text.collapse_whitespace", "Collapse Whitespace", "CW", PreviewPolicy.Threshold, "Trim and collapse approved Unicode whitespace runs to one ordinary space."),
        Descriptor("clean.text.remove_nonprinting", "Remove Nonprinting", "RN", PreviewPolicy.Threshold, "Remove only versioned C0/C1 controls while preserving tabs and newlines."),
        Descriptor("clean.convert.text_to_number", "Invariant Text to Number", "TN", PreviewPolicy.Mandatory, "Convert only complete invariant financial-number strings using the documented fixed grammar."),
        Descriptor("clean.convert.number_to_text", "Number to Invariant Text", "NT", PreviewPolicy.Mandatory, "Convert numeric constants with the explicit invariant 0.################ format."),
        Descriptor("clean.convert.date_normalize", "Normalize Date Text", "DN", PreviewPolicy.Mandatory, "Normalize only yyyy-MM-dd, yyyy/MM/dd, or yyyyMMdd text to yyyy-MM-dd."),
        Descriptor("clean.display.blank_to_zero", "Blanks to Zero", "BZ", PreviewPolicy.Mandatory, "Convert true blank constants to numeric zero."),
        Descriptor("clean.display.zero_to_blank", "Zeros to Blank", "ZB", PreviewPolicy.Mandatory, "Convert numeric zero constants to true blanks; formulas and text zero are skipped."),
        Descriptor("clean.display.blank_to_na_text", "Blanks to N/A", "BN", PreviewPolicy.Mandatory, "Convert true blanks to exact N/A text."),
        Descriptor("clean.display.blank_to_nm_text", "Blanks to NM", "BM", PreviewPolicy.Mandatory, "Convert true blanks to exact NM text."),
        Descriptor("clean.display.blank_to_dash_text", "Blanks to Dash", "BD", PreviewPolicy.Mandatory, "Convert true blanks to exact dash text."),
        Descriptor("clean.display.na_text_to_blank", "N/A to Blanks", "NB", PreviewPolicy.Mandatory, "Convert exact case-sensitive N/A text to blanks."),
        Descriptor("clean.display.nm_text_to_blank", "NM to Blanks", "MB", PreviewPolicy.Mandatory, "Convert exact case-sensitive NM text to blanks."),
        Descriptor("clean.display.dash_text_to_blank", "Dashes to Blanks", "DB", PreviewPolicy.Mandatory, "Convert exact dash text to blanks."),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;
    public static CommandDescriptor GetRequired(string id) => Commands.First(value => value.Id == id);

    private static CommandDescriptor Descriptor(string id, string name, string keytip, PreviewPolicy preview, string description) =>
        new CommandDescriptor(id, 1, name, CommandImpact.Medium, new[] { "value" }, false,
            RibbonRoutes.For(id), "CAP-DATA-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            preview, UndoPolicy.SessionPropertyReceipt, new[] { "AC-DATA-001", "AC-DATA-004", "AC-DATA-017", "AC-DATA-019" },
            "Data Cleaning", description, shortcutLabel: RibbonRoutes.For(id));
}
