using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Navigation;

public static class NavigationCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        D("navigate.sheet.previous", "Previous Visible Sheet", RibbonRoutes.For("navigate.sheet.previous"), "AC-NAV-001"),
        D("navigate.sheet.next", "Next Visible Sheet", RibbonRoutes.For("navigate.sheet.next"), "AC-NAV-001"),
        D("navigate.cell.a1", "Go to A1", "Alt, X, A, V, A", "AC-NAV-002"),
        D("navigate.used.first", "First Used Cell", RibbonRoutes.For("navigate.used.first"), "AC-NAV-003"),
        D("navigate.used.last", "Last Used Cell", RibbonRoutes.For("navigate.used.last"), "AC-NAV-003"),
        D("navigate.region.edge.up", "Region Edge Up", RibbonRoutes.For("navigate.region.edge.up"), "AC-NAV-004"),
        D("navigate.region.edge.down", "Region Edge Down", RibbonRoutes.For("navigate.region.edge.down"), "AC-NAV-004"),
        D("navigate.region.edge.left", "Region Edge Left", RibbonRoutes.For("navigate.region.edge.left"), "AC-NAV-004"),
        D("navigate.region.edge.right", "Region Edge Right", RibbonRoutes.For("navigate.region.edge.right"), "AC-NAV-004"),
        D("navigate.history.back", "Navigation Back", RibbonRoutes.For("navigate.history.back"), "AC-NAV-005"),
        D("navigate.history.forward", "Navigation Forward", RibbonRoutes.For("navigate.history.forward"), "AC-NAV-005"),
        D("navigate.bookmark.add_session", "Add Session Bookmark", RibbonRoutes.For("navigate.bookmark.add_session"), "AC-NAV-006"),
        D("navigate.bookmark.next_session", "Next Session Bookmark", RibbonRoutes.For("navigate.bookmark.next_session"), "AC-NAV-006"),
        D("navigate.bookmark.previous_session", "Previous Session Bookmark", RibbonRoutes.For("navigate.bookmark.previous_session"), "AC-NAV-006"),
        D("navigate.bookmark.clear_session", "Clear Session Bookmarks", RibbonRoutes.For("navigate.bookmark.clear_session"), "AC-NAV-006"),
    };
    public static IEnumerable<CommandDescriptor> All => Commands;
    private static CommandDescriptor D(string id, string name, string route, string acceptance) =>
        new CommandDescriptor(id, 1, name, CommandImpact.ReadOnly, new string[0], true, route, "CAP-NAV-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet, PreviewPolicy.None, UndoPolicy.None, new[] { acceptance },
            "Navigation", $"Navigate to {name.ToLowerInvariant()} without changing workbook content.", shortcutLabel: route);
}
