using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Styles;

namespace ExcelAccel.Application.Profiles;

/// <summary>
/// The families a cycle may belong to. This is the alphabet, never the contents:
/// no format string, colour, or cycle order is written in code, per AC-FMT-045.
/// Family identifiers are the same property identifiers commands already declare
/// in <c>ChangedProperties</c>, so no command-to-family lookup table exists.
/// </summary>
public static class CycleFamilyCatalog
{
    public static readonly IReadOnlyList<string> Supported = new[]
    {
        "column_width",
        "fill_color",
        "font_color",
        "font_size",
        "horizontal_alignment",
        "number_format",
        "row_height",
        "underline",
        "vertical_alignment",
    };

    public static bool IsSupported(string family) =>
        family is not null && Supported.Contains(family, StringComparer.Ordinal);

    public static bool IsColorFamily(string family) =>
        string.Equals(family, "font_color", StringComparison.Ordinal)
        || string.Equals(family, "fill_color", StringComparison.Ordinal);

    /// <summary>
    /// Normalizes one cycle entry for storage. Style-shaped families delegate to
    /// <see cref="StylePropertyCatalog"/> rather than growing a second validator;
    /// only the dimension families and the colour-reference sigil are handled here.
    /// </summary>
    public static string Normalize(string family, string entry)
    {
        if (!IsSupported(family))
        {
            throw new ArgumentException($"Cycle family '{family}' is not supported.", nameof(family));
        }

        var candidate = entry ?? string.Empty;
        if (candidate.Length == 0 || candidate.Length > 256)
        {
            throw new ArgumentException("Cycle entries must be bounded nonempty text.", nameof(entry));
        }

        if (IsColorFamily(family) && ColorReference.IsReference(candidate))
        {
            if (!ColorReference.TryResolveKey(candidate, out _))
            {
                throw new ArgumentException($"'{candidate}' is not a known AutoColor category reference.", nameof(entry));
            }

            return candidate.ToLowerInvariant();
        }

        switch (family)
        {
            case "row_height":
                return Dimension(candidate, 3, 409, "Row height");
            case "column_width":
                return Dimension(candidate, 1, 255, "Column width");
            default:
                return StylePropertyCatalog.Normalize(family, candidate);
        }
    }

    private static string Dimension(string candidate, double minimum, double maximum, string label)
    {
        if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            || double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum)
        {
            throw new ArgumentException($"{label} must be from {minimum} through {maximum}.", nameof(candidate));
        }

        return value.ToString("0.####", CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// A colour cycle entry that names an AutoColor category rather than a literal
/// colour, so the entry tracks the category when the user recolours it. Literal
/// entries deliberately do not track; see <c>docs/commands/AUTOCOLOR.md</c>.
/// </summary>
public static class ColorReference
{
    public const char Sigil = '@';

    private static readonly IReadOnlyDictionary<string, string> Tokens =
        new SortedDictionary<string, string>(StringComparer.Ordinal)
        {
            ["@cross_sheet"] = "cross_sheet_formula",
            ["@error"] = "error",
            ["@external"] = "external_formula",
            ["@hardcode"] = "numeric_hardcode",
            ["@same_sheet"] = "same_sheet_formula",
            ["@text"] = "text",
        };

    public static IReadOnlyList<string> All => Tokens.Keys.ToArray();

    public static bool IsReference(string entry) =>
        !string.IsNullOrEmpty(entry) && entry[0] == Sigil;

    public static bool TryResolveKey(string entry, out string categoryKey)
    {
        categoryKey = string.Empty;
        if (string.IsNullOrEmpty(entry))
        {
            return false;
        }

        return Tokens.TryGetValue(entry.ToLowerInvariant(), out categoryKey!);
    }
}

/// <summary>
/// One named, ordered cycle. A cycle with no entries cannot be constructed, which
/// is how AC-FMT-039 is met: an unconfigured slot is unrepresentable rather than
/// filtered out at invocation time.
/// </summary>
public sealed class ProfileCycle
{
    public const int MaximumEntries = 32;

    public ProfileCycle(string family, string cycleId, string displayName, IEnumerable<string> entries)
    {
        Family = CycleFamilyCatalog.IsSupported(family)
            ? family
            : throw new ArgumentException($"Cycle family '{family}' is not supported.", nameof(family));
        CycleId = RequireToken(cycleId, nameof(cycleId)).ToLowerInvariant();
        DisplayName = RequireToken(displayName, nameof(displayName));
        if (DisplayName.Length > 64)
        {
            throw new ArgumentException("A cycle display name may be at most 64 characters.", nameof(displayName));
        }

        var normalized = (entries ?? throw new ArgumentNullException(nameof(entries)))
            .Select(entry => CycleFamilyCatalog.Normalize(Family, entry))
            .ToArray();
        if (normalized.Length == 0)
        {
            throw new ArgumentException(
                $"Cycle '{CycleId}' has no entries. Delete the cycle instead of leaving it empty.",
                nameof(entries));
        }

        if (normalized.Length > MaximumEntries)
        {
            throw new ArgumentException($"A cycle may contain at most {MaximumEntries} entries.", nameof(entries));
        }

        if (normalized.Distinct(StringComparer.Ordinal).Count() != normalized.Length)
        {
            throw new ArgumentException($"Cycle '{CycleId}' repeats an entry.", nameof(entries));
        }

        Entries = normalized;
    }

    public string Family { get; }
    public string CycleId { get; }
    public string DisplayName { get; }

    /// <summary>Raw stored entries; colour references are not resolved here.</summary>
    public IReadOnlyList<string> Entries { get; }

    /// <summary>
    /// Derives a display name from an identifier mechanically, so no human-readable
    /// cycle name is hard-coded in the product.
    /// </summary>
    public static string TitleFrom(string identifier)
    {
        var words = (identifier ?? string.Empty)
            .Split(new[] { '_', '-', '.' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(word => word.Length == 1
                ? word.ToUpperInvariant()
                : char.ToUpperInvariant(word[0]) + word.Substring(1))
            .ToArray();
        return words.Length == 0 ? "Cycle" : string.Join(" ", words);
    }

    private static string RequireToken(string value, string parameterName) =>
        string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException("A nonempty token is required.", parameterName)
            : value.Trim();
}

/// <summary>
/// Every configured cycle, grouped by family. A family with no cycles is absent
/// rather than empty, so an unconfigured family cannot round-trip as a phantom.
/// </summary>
public sealed class ProfileCycles
{
    public const int MaximumCyclesPerFamily = 8;

    private readonly SortedDictionary<string, IReadOnlyList<ProfileCycle>> _families;

    public ProfileCycles(IEnumerable<KeyValuePair<string, IEnumerable<ProfileCycle>>> families)
    {
        _families = new SortedDictionary<string, IReadOnlyList<ProfileCycle>>(StringComparer.Ordinal);
        foreach (var family in families ?? throw new ArgumentNullException(nameof(families)))
        {
            if (!CycleFamilyCatalog.IsSupported(family.Key))
            {
                throw new ArgumentException($"Cycle family '{family.Key}' is not supported.", nameof(families));
            }

            if (_families.ContainsKey(family.Key))
            {
                throw new ArgumentException($"Cycle family '{family.Key}' appears more than once.", nameof(families));
            }

            var cycles = (family.Value ?? Array.Empty<ProfileCycle>()).ToArray();
            if (cycles.Length == 0)
            {
                throw new ArgumentException(
                    $"Cycle family '{family.Key}' has no cycles. Omit the family instead of leaving it empty.",
                    nameof(families));
            }

            if (cycles.Length > MaximumCyclesPerFamily)
            {
                throw new ArgumentException(
                    $"Cycle family '{family.Key}' may contain at most {MaximumCyclesPerFamily} cycles.",
                    nameof(families));
            }

            if (cycles.Any(cycle => !string.Equals(cycle.Family, family.Key, StringComparison.Ordinal)))
            {
                throw new ArgumentException($"A cycle is filed under the wrong family '{family.Key}'.", nameof(families));
            }

            if (cycles.Select(cycle => cycle.CycleId).Distinct(StringComparer.Ordinal).Count() != cycles.Length)
            {
                throw new ArgumentException($"Cycle identifiers must be unique within family '{family.Key}'.", nameof(families));
            }

            _families.Add(family.Key, cycles);
        }
    }

    /// <summary>Configured families only, ordinal ordered.</summary>
    public IReadOnlyList<string> Families => _families.Keys.ToArray();

    /// <summary>Cycles in user-controlled slot order, which is never re-sorted.</summary>
    public IReadOnlyList<ProfileCycle> this[string family] =>
        _families.TryGetValue(family ?? string.Empty, out var cycles)
            ? cycles
            : Array.Empty<ProfileCycle>();

    public IReadOnlyList<ProfileCycle> All =>
        _families.SelectMany(family => family.Value).ToArray();

    public bool TryGet(string family, string cycleId, out ProfileCycle cycle)
    {
        cycle = this[family].FirstOrDefault(candidate =>
            string.Equals(candidate.CycleId, cycleId, StringComparison.OrdinalIgnoreCase))!;
        return cycle is not null;
    }
}
