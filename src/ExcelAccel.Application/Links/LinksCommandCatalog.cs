using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Links;

public static class LinksCommandCatalog
{
    public const string InventoryId = "links.inventory.open";
    public const string NavigateUsageId = "links.navigate_usage";
    public const string ExportId = "links.inventory.export";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        ReadOnly(
            InventoryId,
            "External Link Inventory",
            "Group every qualified external link by source and list where it is used, in a read-only view. The source is never opened, contacted, updated, repointed, or broken, and status describes only what this Excel session can see.",
            new[] { "links", "external links", "edit links", "workbook links" },
            new[] { "AC-LINK-001", "AC-LINK-002", "AC-LINK-003", "AC-LINK-004", "AC-LINK-005", "AC-LINK-006", "AC-LINK-007" }),
        ReadOnly(
            NavigateUsageId,
            "Go to Link Usage",
            "Select the cell that carries the highlighted link usage. The external source is never opened, and usages that are not cells stay listed but non-navigable.",
            new[] { "go to link", "navigate link" },
            new[] { "AC-LINK-008", "AC-LINK-009", "AC-LINK-010" }),
        ReadOnly(
            ExportId,
            "Export External Link Inventory",
            "Write the current link inventory to a local file after confirming a manifest. Source paths are excluded by default and nothing is transmitted.",
            new[] { "export links" },
            new[] { "AC-LINK-011", "AC-SEC-004" }),
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
            "CAP-LINK-001",
            CommandContextRequirement.Workbook,
            PreviewPolicy.None,
            UndoPolicy.None,
            acceptanceIds,
            "Auditing",
            description,
            aliases,
            RibbonRoutes.For(id));
}
