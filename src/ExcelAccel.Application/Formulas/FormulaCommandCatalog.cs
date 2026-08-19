using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Formulas;

public static class FormulaCommandCatalog
{
    private static readonly string[] FormulaAcceptance = { "AC-FORM-001", "AC-FORM-002", "AC-FORM-003", "AC-FORM-024" };
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor("formula.source.capture", 1, "Capture Formula Source", CommandImpact.ReadOnly,
            Array.Empty<string>(), false, "Ribbon KeyTips: Alt, X, A, M, SC", "CAP-FORM-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None, UndoPolicy.None, new[] { "AC-FORM-013", "AC-FORM-026" }, "Formula",
            "Capture a bounded internal source snapshot for transpose and later qualified paste commands.",
            new[] { "copy source", "transpose source", "internal clipboard" }, "Alt, X, A, M, SC"),
        Descriptor("formula.transpose", "Transpose Captured Source Here", new[] { "formula", "value" }, false, "TP",
            "Transpose the captured source into the exact selected destination without copying formatting.",
            new[] { "transpose formulas", "transpose values" }, "AC-FORM-013", "AC-FORM-016", ChangedPropertyPolicy.DeclaredSubset, PreviewPolicy.Mandatory),
        Descriptor("paste.formulas_only", "Paste Formulas Only", new[] { "formula" }, false, "PF",
            "Paste only formulas from the captured internal source with exact-shape or whole-multiple repetition.",
            new[] { "formula paste", "paste formulas" }, "AC-FORM-026", "AC-FORM-029"),
        Descriptor("paste.values_only", "Paste Values Only", new[] { "value" }, true, "PV",
            "Paste only underlying values from the captured internal source, with explicit formula-to-value counts.",
            new[] { "value paste", "paste values" }, "AC-FORM-026", "AC-FORM-030", previewPolicy: PreviewPolicy.Mandatory),
        Descriptor("paste.formats_only", "Paste Formats Only", new[] { "cell_format_block_v1" }, true, "PT",
            "Paste only the approved bounded nine-property formatting set from the captured internal source.",
            new[] { "format paste", "paste formats" }, "AC-FORM-027", "AC-FORM-029", previewPolicy: PreviewPolicy.Mandatory),
        Descriptor("formula.copy.down", "Smart Copy Down", new[] { "formula" }, false, "CD", "Translate formulas from the top edge into the selected rows.", new[] { "copy down", "smart copy" }, "AC-FORM-008", "AC-FORM-012"),
        Descriptor("formula.copy.right", "Smart Copy Right", new[] { "formula" }, false, "CR", "Translate formulas from the left edge into the selected columns.", new[] { "copy right", "smart copy" }, "AC-FORM-008", "AC-FORM-012"),
        Descriptor("formula.spacing.rows", "Space Formulas by Rows", new[] { "formula" }, false, "SR",
            "Prompt for an explicit positive interval, then translate top-edge formulas into exact row destinations.",
            new[] { "formula spacing rows", "space rows" }, "AC-FORM-008", "AC-FORM-012"),
        Descriptor("formula.spacing.columns", "Space Formulas by Columns", new[] { "formula" }, false, "SL",
            "Prompt for an explicit positive interval, then translate left-edge formulas into exact column destinations.",
            new[] { "formula spacing columns", "space columns" }, "AC-FORM-008", "AC-FORM-012"),
        Descriptor("fill.formula_from_above", "Fill Formula from Above", new[] { "formula" }, true, "FA",
            "Translate the immediately adjacent formula row into the selected destination.",
            new[] { "formula from above", "fill formulas" }, "AC-FORM-035", "AC-FORM-038"),
        Descriptor("fill.value_from_above", "Fill Value from Above", new[] { "value" }, true, "VA",
            "Copy underlying values from the immediately adjacent row, including calculated formula values.",
            new[] { "value from above", "fill values" }, "AC-FORM-036", "AC-FORM-038", previewPolicy: PreviewPolicy.Mandatory),
        Descriptor("fill.numeric_sequence", "Fill Numeric Sequence", new[] { "value" }, false, "NS",
            "Prompt for explicit invariant start, step, and direction, then fill the exact selection.",
            new[] { "number sequence", "series fill" }, "AC-FORM-031", "AC-FORM-034", previewPolicy: PreviewPolicy.Mandatory),
        Descriptor("fill.date_sequence", "Fill Date Sequence", new[] { "value" }, false, "DS",
            "Prompt for explicit ISO start date, day step, direction, and workbook date-system confirmation.",
            new[] { "date sequence", "date fill" }, "AC-FORM-031", "AC-FORM-034", previewPolicy: PreviewPolicy.Mandatory),
        Descriptor("formula.iferror.toggle", "Toggle IFERROR", new[] { "formula" }, true, "IE", "Add or remove the exact configured top-level IFERROR wrapper.", new[] { "if error", "error wrapper" }, "AC-FORM-017", "AC-FORM-020"),
        Descriptor("formula.sign.reverse", "Reverse Sign", new[] { "formula", "value" }, true, "RS", "Apply or remove canonical formula negation; numeric constants are opt-in.", new[] { "negate", "flip sign" }, "AC-FORM-021", "AC-FORM-024", ChangedPropertyPolicy.DeclaredSubset),
        Descriptor("formula.units.to_thousands", "Convert to Thousands", new[] { "formula", "value" }, true, "UT", "Divide formulas and optionally numeric constants by 1,000.", new[] { "scale thousands", "divide 1000" }, "AC-FORM-004", "AC-FORM-025", ChangedPropertyPolicy.DeclaredSubset),
        Descriptor("formula.units.from_thousands", "Convert from Thousands", new[] { "formula", "value" }, true, "UF", "Multiply formulas and optionally numeric constants by 1,000.", new[] { "unscale thousands", "multiply 1000" }, "AC-FORM-004", "AC-FORM-025", ChangedPropertyPolicy.DeclaredSubset),
        Descriptor("formula.units.to_millions", "Convert to Millions", new[] { "formula", "value" }, true, "UM", "Divide formulas and optionally numeric constants by 1,000,000.", new[] { "scale millions", "divide 1000000" }, "AC-FORM-004", "AC-FORM-025", ChangedPropertyPolicy.DeclaredSubset),
        Descriptor("formula.units.from_millions", "Convert from Millions", new[] { "formula", "value" }, true, "UN", "Multiply formulas and optionally numeric constants by 1,000,000.", new[] { "unscale millions", "multiply 1000000" }, "AC-FORM-004", "AC-FORM-025", ChangedPropertyPolicy.DeclaredSubset),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;
    public static CommandDescriptor GetRequired(string commandId) => Commands.First(value => value.Id == commandId);

    private static CommandDescriptor Descriptor(string id, string name, IEnumerable<string> properties, bool fixedParameters,
        string keytip, string description, IEnumerable<string> aliases, string acceptance1, string acceptance2,
        ChangedPropertyPolicy propertyPolicy = ChangedPropertyPolicy.Exact,
        PreviewPolicy previewPolicy = PreviewPolicy.Threshold) =>
        new CommandDescriptor(id, 1, name, CommandImpact.Medium, properties, fixedParameters,
            "Ribbon KeyTips: Alt, X, A, M, " + keytip, "CAP-FORM-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            previewPolicy, UndoPolicy.SessionPropertyReceipt,
            FormulaAcceptance.Concat(new[] { acceptance1, acceptance2 }), "Formula", description, aliases,
            "Alt, X, A, M, " + keytip, changedPropertyPolicy: propertyPolicy);
}
