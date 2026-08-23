using System;
using ExcelAccel.Application.Profiles;

namespace ExcelAccel.Application.Formatting;

/// <summary>
/// Decides whether a ribbon control bound to a cycle should be shown.
/// </summary>
/// <remarks>
/// The ribbon XML is static, so a button for a deleted cycle cannot remove
/// itself; it has to be hidden through <c>getVisible</c>. Keeping the decision
/// here rather than in the ribbon callback means it can be tested without a host,
/// and means the callback stays a one-liner.
///
/// Only controls bound to one particular cycle are hideable. The decimals
/// commands operate on any number format, including formats belonging to no
/// cycle, so they are always visible.
/// </remarks>
public static class CycleVisibility
{
    private const string NumberPrefix = "format.number.";

    public static bool IsVisible(ProfileDefinition profile, string commandId)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (string.IsNullOrWhiteSpace(commandId))
        {
            return true;
        }

        if (CycleCommandFactory.IsCycleCommand(commandId))
        {
            return true;
        }

        if (commandId.StartsWith(NumberPrefix, StringComparison.Ordinal))
        {
            var cycleId = commandId.Substring(NumberPrefix.Length);
            if (cycleId.StartsWith("decimals.", StringComparison.Ordinal))
            {
                return true;
            }

            return profile.Cycles.TryGet("number_format", cycleId, out _);
        }

        var family = FamilyFor(commandId);
        return family is null || profile.Cycles[family].Count > 0;
    }

    /// <summary>
    /// The family a slot-zero cycle command follows, or null when the command is
    /// not bound to a cycle at all.
    /// </summary>
    private static string? FamilyFor(string commandId)
    {
        switch (commandId)
        {
            case "format.font_color.cycle": return "font_color";
            case "format.fill_color.cycle": return "fill_color";
            case "format.font_size.cycle": return "font_size";
            case "format.row_height.cycle": return "row_height";
            case "format.column_width.cycle": return "column_width";
            case "format.alignment.horizontal.cycle": return "horizontal_alignment";
            case "format.alignment.vertical.cycle": return "vertical_alignment";
            case "format.underline.cycle": return "underline";
            default: return null;
        }
    }
}
