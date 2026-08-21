using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Application.Profiles;

/// <summary>
/// Pure add, remove, rename, and reorder operations over <see cref="ProfileCycles"/>.
/// </summary>
/// <remarks>
/// This lives in the application layer rather than in the settings dialog so it
/// can be tested without a host. The dialog is expected to be a thin shell over
/// these operations, and every one of them returns a new collection rather than
/// mutating, so a rejected edit leaves the caller's profile untouched.
///
/// Deletion is the only way a cycle stops existing. There is deliberately no
/// "clear the entries" operation, because an empty cycle is exactly the phantom
/// slot AC-FMT-039 forbids, and <see cref="ProfileCycle"/> refuses to construct
/// one anyway.
/// </remarks>
public static class ProfileCycleEditor
{
    public static ProfileCycles Add(ProfileCycles cycles, ProfileCycle cycle)
    {
        if (cycles is null) throw new ArgumentNullException(nameof(cycles));
        if (cycle is null) throw new ArgumentNullException(nameof(cycle));

        var existing = cycles[cycle.Family];
        if (existing.Any(value => string.Equals(value.CycleId, cycle.CycleId, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException(
                $"Family '{cycle.Family}' already has a cycle named '{cycle.CycleId}'.", nameof(cycle));
        }

        if (existing.Count >= ProfileCycles.MaximumCyclesPerFamily)
        {
            throw new ArgumentException(
                $"Family '{cycle.Family}' already holds {ProfileCycles.MaximumCyclesPerFamily} cycles, which is the limit. Remove one first.",
                nameof(cycle));
        }

        return Rebuild(cycles, cycle.Family, existing.Concat(new[] { cycle }).ToArray());
    }

    /// <summary>
    /// Removes a cycle. Removing the last cycle in a family removes the family,
    /// so nothing is left behind for a command to find and refuse on.
    /// </summary>
    public static ProfileCycles Remove(ProfileCycles cycles, string family, string cycleId)
    {
        if (cycles is null) throw new ArgumentNullException(nameof(cycles));

        var existing = cycles[family];
        var remaining = existing
            .Where(value => !string.Equals(value.CycleId, cycleId, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (remaining.Length == existing.Count)
        {
            throw new ArgumentException($"Family '{family}' has no cycle named '{cycleId}'.", nameof(cycleId));
        }

        return Rebuild(cycles, family, remaining);
    }

    public static ProfileCycles Rename(ProfileCycles cycles, string family, string cycleId, string displayName)
    {
        if (cycles is null) throw new ArgumentNullException(nameof(cycles));

        var existing = Require(cycles, family, cycleId);
        var renamed = new ProfileCycle(family, existing.CycleId, displayName, existing.Entries);
        return Replace(cycles, family, existing, renamed);
    }

    public static ProfileCycles SetEntries(ProfileCycles cycles, string family, string cycleId, IEnumerable<string> entries)
    {
        if (cycles is null) throw new ArgumentNullException(nameof(cycles));

        var existing = Require(cycles, family, cycleId);
        var updated = new ProfileCycle(family, existing.CycleId, existing.DisplayName, entries);
        return Replace(cycles, family, existing, updated);
    }

    /// <summary>
    /// Moves a cycle within its family's slot order. Slot order is user data and
    /// is load-bearing: commands that name no particular cycle follow whichever
    /// one sits first.
    /// </summary>
    public static ProfileCycles Move(ProfileCycles cycles, string family, string cycleId, int offset)
    {
        if (cycles is null) throw new ArgumentNullException(nameof(cycles));

        var existing = cycles[family].ToList();
        var index = existing.FindIndex(value => string.Equals(value.CycleId, cycleId, StringComparison.OrdinalIgnoreCase));
        if (index < 0)
        {
            throw new ArgumentException($"Family '{family}' has no cycle named '{cycleId}'.", nameof(cycleId));
        }

        var target = index + offset;
        if (target < 0 || target >= existing.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(offset), "The cycle cannot move outside its family.");
        }

        var moved = existing[index];
        existing.RemoveAt(index);
        existing.Insert(target, moved);
        return Rebuild(cycles, family, existing);
    }

    private static ProfileCycle Require(ProfileCycles cycles, string family, string cycleId) =>
        cycles.TryGet(family, cycleId, out var cycle)
            ? cycle
            : throw new ArgumentException($"Family '{family}' has no cycle named '{cycleId}'.", nameof(cycleId));

    private static ProfileCycles Replace(ProfileCycles cycles, string family, ProfileCycle from, ProfileCycle to) =>
        Rebuild(cycles, family, cycles[family].Select(value => ReferenceEquals(value, from) ? to : value).ToArray());

    private static ProfileCycles Rebuild(ProfileCycles cycles, string family, IReadOnlyList<ProfileCycle> replacement)
    {
        var families = cycles.Families
            .Where(name => !string.Equals(name, family, StringComparison.Ordinal))
            .Select(name => new KeyValuePair<string, IEnumerable<ProfileCycle>>(name, cycles[name]))
            .ToList();
        if (replacement.Count > 0)
        {
            families.Add(new KeyValuePair<string, IEnumerable<ProfileCycle>>(family, replacement));
        }

        return new ProfileCycles(families);
    }
}
