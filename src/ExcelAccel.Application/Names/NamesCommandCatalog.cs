using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Names;

public static class NamesCommandCatalog
{
    public const string InventoryId = "names.inventory.open";
    public const string NavigateTargetId = "names.navigate_target";
    public const string ExportId = "names.inventory.export";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        ReadOnly(
            InventoryId,
            "Named Range Inventory",
            "List every qualified defined name with its scope, target category, and status in a read-only view. Closed external sources are never opened, and nothing is renamed, repointed, or deleted.",
            new[] { "names", "defined names", "named ranges", "name manager" },
            new[] { "AC-NAME-001", "AC-NAME-002", "AC-NAME-003", "AC-NAME-004", "AC-NAME-005", "AC-NAME-006" }),
        ReadOnly(
            NavigateTargetId,
            "Go to Named Range Target",
            "Select the target of the highlighted name when it resolves to a range in this workbook. Constants, expressions, external and broken names stay inspectable but are not navigable.",
            new[] { "go to name", "navigate name" },
            new[] { "AC-NAME-007", "AC-NAV-005" }),
        ReadOnly(
            ExportId,
            "Export Named Range Inventory",
            "Write the current inventory to a local file after confirming a manifest. Name expressions are excluded by default and nothing is transmitted.",
            new[] { "export names" },
            new[] { "AC-NAME-011", "AC-SEC-004" }),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    private static CommandDescriptor ReadOnly(
        string id,
        string name,
        string description,
        IEnumerable<string> aliases,
        IEnumerable<string> acceptanceIds) =>
        new CommandDescriptor(
            id,
            1,
            name,
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For(id),
            "CAP-NAME-001",
            CommandContextRequirement.Workbook,
            PreviewPolicy.None,
            UndoPolicy.None,
            acceptanceIds,
            "Auditing",
            description,
            aliases,
            RibbonRoutes.For(id));
}
