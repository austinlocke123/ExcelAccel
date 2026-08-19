using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Navigation;

public static class NavigationCommandCatalog
{
    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        D("navigate.sheet.previous", "Previous Visible Sheet", "Alt, X, A, V, P", "AC-NAV-001"),
        D("navigate.sheet.next", "Next Visible Sheet", "Alt, X, A, V, N", "AC-NAV-001"),
        D("navigate.cell.a1", "Go to A1", "Alt, X, A, V, A", "AC-NAV-002"),
        D("navigate.used.first", "First Used Cell", "Alt, X, A, V, F", "AC-NAV-003"),
        D("navigate.used.last", "Last Used Cell", "Alt, X, A, V, L", "AC-NAV-003"),
        D("navigate.region.edge.up", "Region Edge Up", "Alt, X, A, V, U", "AC-NAV-004"),
        D("navigate.region.edge.down", "Region Edge Down", "Alt, X, A, V, D", "AC-NAV-004"),
        D("navigate.region.edge.left", "Region Edge Left", "Alt, X, A, V, E", "AC-NAV-004"),
        D("navigate.region.edge.right", "Region Edge Right", "Alt, X, A, V, R", "AC-NAV-004"),
        D("navigate.history.back", "Navigation Back", "Alt, X, A, V, B", "AC-NAV-005"),
        D("navigate.history.forward", "Navigation Forward", "Alt, X, A, V, O", "AC-NAV-005"),
        D("navigate.bookmark.add_session", "Add Session Bookmark", "Alt, X, A, V, M", "AC-NAV-006"),
        D("navigate.bookmark.next_session", "Next Session Bookmark", "Alt, X, A, V, J", "AC-NAV-006"),
        D("navigate.bookmark.previous_session", "Previous Session Bookmark", "Alt, X, A, V, K", "AC-NAV-006"),
        D("navigate.bookmark.clear_session", "Clear Session Bookmarks", "Alt, X, A, V, C", "AC-NAV-006"),
    };
    public static IEnumerable<CommandDescriptor> All => Commands;
    private static CommandDescriptor D(string id, string name, string route, string acceptance) =>
        new CommandDescriptor(id, 1, name, CommandImpact.ReadOnly, new string[0], true, route, "CAP-NAV-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet, PreviewPolicy.None, UndoPolicy.None, new[] { acceptance });
}
