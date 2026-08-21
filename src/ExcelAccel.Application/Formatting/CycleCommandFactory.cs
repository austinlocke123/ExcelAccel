using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Formatting;

/// <summary>
/// Builds command descriptors for the cycles a profile defines, so a cycle the
/// user invented after install is still reachable.
/// </summary>
/// <remarks>
/// The ribbon is a static XML string, so a cycle created after install cannot be
/// given a purpose-built button. Command Search can carry it, because it indexes
/// any descriptor sequence rather than only the static registry.
///
/// Only cycles that no built-in command already covers get a descriptor. A
/// built-in command exists for every number-format cycle the default profile
/// ships and for the first cycle in each property family, and duplicating those
/// would put two entries with different names in the search results for one
/// behaviour.
/// </remarks>
public static class CycleCommandFactory
{
    public const string IdPrefix = "format.cycle.";

    /// <summary>
    /// The route these commands honestly have. They are not on the ribbon, so
    /// advertising an Alt sequence would be a lie.
    /// </summary>
    public const string Route = "Search Commands, then the cycle name";

    public static IReadOnlyList<CommandDescriptor> Descriptors(ProfileDefinition profile)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));

        var covered = new HashSet<string>(StringComparer.Ordinal);
        foreach (var command in Phase1AFormattingCatalog.All)
        {
            if (command.Id.StartsWith("format.number.", StringComparison.Ordinal))
            {
                covered.Add("number_format/" + command.Id.Substring("format.number.".Length));
            }
        }

        var descriptors = new List<CommandDescriptor>();
        foreach (var family in profile.Cycles.Families)
        {
            var cycles = profile.Cycles[family];
            for (var index = 0; index < cycles.Count; index++)
            {
                var cycle = cycles[index];
                var isBuiltIn = covered.Contains(family + "/" + cycle.CycleId)
                    || (index == 0 && !string.Equals(family, "number_format", StringComparison.Ordinal));
                if (isBuiltIn)
                {
                    continue;
                }

                descriptors.Add(Describe(family, cycle));
            }
        }

        return descriptors;
    }

    public static bool IsCycleCommand(string commandId) =>
        !string.IsNullOrEmpty(commandId) && commandId.StartsWith(IdPrefix, StringComparison.Ordinal);

    /// <summary>
    /// Resolves a generated id back to a runnable command. A cycle deleted since
    /// the id was stored refuses by name rather than silently doing nothing.
    /// </summary>
    public static ProfileFormattingCommand Create(ProfileDefinition profile, string commandId)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (!TrySplit(commandId, out var family, out var cycleId))
        {
            throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                $"'{commandId}' is not a cycle command.",
                "Use Search Commands to find the cycle you want.");
        }

        if (!profile.Cycles.TryGet(family, cycleId, out var cycle))
        {
            throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                $"The '{cycleId}' cycle is no longer configured in the active profile.",
                "Add the cycle in Settings, or choose another from Search Commands.");
        }

        return new ProfileFormattingCommand(
            Describe(family, cycle),
            (p, current) => ProfileFormattingCommand.Next(p.ResolveCycle(family, cycleId), current),
            unconfiguredMessage: $"The '{cycleId}' cycle is not configured in the active profile.");
    }

    public static string IdFor(string family, string cycleId) => IdPrefix + family + "." + cycleId;

    private static bool TrySplit(string commandId, out string family, out string cycleId)
    {
        family = string.Empty;
        cycleId = string.Empty;
        if (!IsCycleCommand(commandId))
        {
            return false;
        }

        var remainder = commandId.Substring(IdPrefix.Length);
        // Family identifiers are a known closed set, and a cycle id may itself
        // contain a dot, so match the family rather than splitting on the last one.
        var match = CycleFamilyCatalog.Supported
            .Where(candidate => remainder.StartsWith(candidate + ".", StringComparison.Ordinal))
            .OrderByDescending(candidate => candidate.Length)
            .FirstOrDefault();
        if (match is null)
        {
            return false;
        }

        family = match;
        cycleId = remainder.Substring(match.Length + 1);
        return cycleId.Length > 0;
    }

    private static CommandDescriptor Describe(string family, ProfileCycle cycle) =>
        new CommandDescriptor(
            IdFor(family, cycle.CycleId),
            1,
            cycle.DisplayName,
            CommandImpact.Low,
            new[] { family },
            true,
            Route,
            "CAP-FMT-004",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.SessionPropertyReceipt,
            new[] { "AC-FMT-026", "AC-FMT-029", "AC-REL-005" },
            "Formatting",
            $"Cycle {family.Replace('_', ' ')} through the '{cycle.DisplayName}' cycle from the active profile.",
            new[] { "cycle", cycle.CycleId, family.Replace('_', ' ') },
            Route);
}
