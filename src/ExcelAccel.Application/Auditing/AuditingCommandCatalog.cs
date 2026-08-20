using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Auditing;

public static class AuditingCommandCatalog
{
    public const string DirectPrecedentsId = "audit.precedents.direct";
    public const string DirectDependentsId = "audit.dependents.direct";
    public const string WorkbookDependentsId = "audit.dependents.workbook";
    public const string IndirectPrecedentsId = "audit.precedents.indirect";
    public const string IndirectDependentsId = "audit.dependents.indirect";
    public const string InspectFormulaId = "audit.formula.inspect";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor(
            DirectPrecedentsId,
            1,
            "Trace Direct Precedents",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.precedents.direct"),
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-AUD-001", "AC-AUD-002", "AC-AUD-003", "AC-AUD-004", "AC-AUD-005" },
            "Auditing",
            "Show the direct precedents of one selected formula cell in a read-only view. Closed external workbooks are listed but never opened, and no Excel trace arrow or workbook annotation is used.",
            new[] { "precedents", "trace precedents", "audit precedents" },
            RibbonRoutes.For("audit.precedents.direct")),
        new CommandDescriptor(
            DirectDependentsId,
            1,
            "Trace Direct Dependents",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.dependents.direct"),
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.Threshold,
            UndoPolicy.None,
            new[] { "AC-AUD-006", "AC-AUD-007", "AC-AUD-008", "AC-AUD-009" },
            "Auditing",
            "Scan the active worksheet for formulas that read the selected cell or range, and show them in a read-only view. The scan is bounded, cancellable, confirmed before a large worksheet, and never widens beyond the declared worksheet scope.",
            new[] { "dependents", "trace dependents", "audit dependents", "precedents reverse" },
            RibbonRoutes.For("audit.dependents.direct")),
        new CommandDescriptor(
            WorkbookDependentsId,
            1,
            "Trace Dependents Across Workbook",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.dependents.workbook"),
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.Mandatory,
            UndoPolicy.None,
            new[] { "AC-AUD-006", "AC-AUD-007", "AC-AUD-008", "AC-AUD-009" },
            "Auditing",
            "Scan every worksheet in the workbook for formulas that read the selection. The sheet inventory is always confirmed before anything is read, a worksheet that cannot be bounded is excluded with a stated reason, and the scan is cancellable.",
            new[] { "workbook dependents", "dependents everywhere" },
            RibbonRoutes.For("audit.dependents.workbook")),
        new CommandDescriptor(
            IndirectPrecedentsId,
            1,
            "Trace Indirect Precedents",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.precedents.indirect"),
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-AUD-010", "AC-AUD-011", "AC-AUD-012", "AC-AUD-013", "AC-AUD-014" },
            "Auditing",
            "Follow the precedent chain upstream from the selected formula cell, breadth-first within explicit depth and node caps. Cycles terminate and are shown as cycle edges; reaching a cap produces an explicit truncated result.",
            new[] { "indirect precedents", "precedent chain", "trace upstream" },
            RibbonRoutes.For("audit.precedents.indirect")),
        new CommandDescriptor(
            IndirectDependentsId,
            1,
            "Trace Indirect Dependents",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.dependents.indirect"),
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.Threshold,
            UndoPolicy.None,
            new[] { "AC-AUD-010", "AC-AUD-011", "AC-AUD-012", "AC-AUD-013", "AC-AUD-014" },
            "Auditing",
            "Follow the dependent chain downstream from the selection within this worksheet, breadth-first within explicit depth and node caps. The worksheet is scanned once, cycles terminate, and reaching a cap produces an explicit truncated result.",
            new[] { "indirect dependents", "dependent chain", "trace downstream" },
            RibbonRoutes.For("audit.dependents.indirect")),
        new CommandDescriptor(
            InspectFormulaId,
            1,
            "Inspect Formula",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For("audit.formula.inspect"),
            "CAP-AUD-002",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-AUD-016", "AC-AUD-017", "AC-AUD-018", "AC-AUD-019" },
            "Auditing",
            "Show the structure of the selected formula as a tree of functions, operators, constants, and references, each with its exact source span. Nothing is evaluated, scored, or explained.",
            new[] { "inspect formula", "formula tree", "formula structure" },
            RibbonRoutes.For("audit.formula.inspect")),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    public static CommandDescriptor GetRequired(string id) => Commands.First(value => value.Id == id);
}
