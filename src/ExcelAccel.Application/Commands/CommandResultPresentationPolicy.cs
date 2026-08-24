using System;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Links;
using ExcelAccel.Application.Names;

namespace ExcelAccel.Application.Commands;

/// <summary>
/// Identifies audit-menu commands whose result surface always presents both
/// success and refusal. Other audit commands can refuse before a view exists and
/// must leave callback-boundary failure reporting enabled.
/// </summary>
public static class CommandResultPresentationPolicy
{
    public static bool PresentsOwnAuditResult(string commandId)
    {
        if (string.IsNullOrWhiteSpace(commandId)) return false;
        return AuditingCommandCatalog.All.Any(command => command.Id == commandId)
            || string.Equals(commandId, NamesCommandCatalog.InventoryId, StringComparison.Ordinal)
            || string.Equals(commandId, LinksCommandCatalog.InventoryId, StringComparison.Ordinal);
    }
}
