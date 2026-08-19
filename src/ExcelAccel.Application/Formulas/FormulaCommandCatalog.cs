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
        Descriptor("formula.copy.down", "Smart Copy Down", new[] { "formula" }, false, "CD", "Translate formulas from the top edge into the selected rows.", new[] { "copy down", "smart copy" }, "AC-FORM-008", "AC-FORM-012"),
        Descriptor("formula.copy.right", "Smart Copy Right", new[] { "formula" }, false, "CR", "Translate formulas from the left edge into the selected columns.", new[] { "copy right", "smart copy" }, "AC-FORM-008", "AC-FORM-012"),
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
        ChangedPropertyPolicy propertyPolicy = ChangedPropertyPolicy.Exact) =>
        new CommandDescriptor(id, 1, name, CommandImpact.Medium, properties, fixedParameters,
            "Ribbon KeyTips: Alt, X, A, M, " + keytip, "CAP-FORM-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.Threshold, UndoPolicy.SessionPropertyReceipt,
            FormulaAcceptance.Concat(new[] { acceptance1, acceptance2 }), "Formula", description, aliases,
            "Alt, X, A, M, " + keytip, changedPropertyPolicy: propertyPolicy);
}
