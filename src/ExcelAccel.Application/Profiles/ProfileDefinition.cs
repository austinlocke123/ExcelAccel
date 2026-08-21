using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Styles;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Profiles;

public sealed class ProfileDefinition
{
    public const int CurrentSchemaVersion = 6;
    public const int MaximumBindings = 512;
    public const int MaximumFavorites = 128;
    public const int MaximumLocalStyles = 64;

    public ProfileDefinition(
        int schemaVersion,
        string profileId,
        ProfileCycles cycles,
        IEnumerable<KeyValuePair<string, string>> autoColorColors,
        IEnumerable<QuickKeyBinding> quickKeys,
        IEnumerable<FavoriteDefinition> favorites,
        IEnumerable<StyleRecipe> localStyles,
        long immediatePreviewCellLimit,
        bool wrapSheetNavigation,
        string formulaIfErrorFallback = "0")
    {
        if (schemaVersion != CurrentSchemaVersion)
        {
            throw new ArgumentOutOfRangeException(nameof(schemaVersion), "The profile schema is not supported.");
        }

        SchemaVersion = schemaVersion;
        ProfileId = RequireToken(profileId, nameof(profileId));
        Cycles = cycles ?? throw new ArgumentNullException(nameof(cycles));
        AutoColorColors = NormalizeColorMap(autoColorColors, nameof(autoColorColors));

        var bindings = (quickKeys ?? throw new ArgumentNullException(nameof(quickKeys))).ToArray();
        if (bindings.Length > MaximumBindings)
        {
            throw new ArgumentException($"Profiles may contain at most {MaximumBindings} Quick Key bindings.", nameof(quickKeys));
        }

        QuickKeys = bindings;
        var normalizedFavorites = (favorites ?? throw new ArgumentNullException(nameof(favorites)))
            .OrderBy(value => value.FavoriteId, StringComparer.Ordinal)
            .ToArray();
        if (normalizedFavorites.Length > MaximumFavorites)
            throw new ArgumentException($"Profiles may contain at most {MaximumFavorites} favorites.", nameof(favorites));
        if (normalizedFavorites.Select(value => value.FavoriteId).Distinct(StringComparer.Ordinal).Count() != normalizedFavorites.Length)
            throw new ArgumentException("Favorite IDs must be unique.", nameof(favorites));
        Favorites = normalizedFavorites;
        var normalizedStyles = (localStyles ?? throw new ArgumentNullException(nameof(localStyles)))
            .OrderBy(value => value.StyleId, StringComparer.Ordinal).ToArray();
        if (normalizedStyles.Length > MaximumLocalStyles) throw new ArgumentException($"Profiles may contain at most {MaximumLocalStyles} local styles.", nameof(localStyles));
        if (normalizedStyles.Any(value => value.Origin != StyleOrigin.Local)) throw new ArgumentException("The user profile may contain local styles only.", nameof(localStyles));
        if (normalizedStyles.Select(value => value.StyleId).Distinct(StringComparer.Ordinal).Count() != normalizedStyles.Length) throw new ArgumentException("Local style IDs must be unique.", nameof(localStyles));
        LocalStyles = normalizedStyles;
        if (immediatePreviewCellLimit < 1 || immediatePreviewCellLimit > 1_000_000)
        {
            throw new ArgumentOutOfRangeException(nameof(immediatePreviewCellLimit));
        }

        ImmediatePreviewCellLimit = immediatePreviewCellLimit;
        WrapSheetNavigation = wrapSheetNavigation;
        FormulaIfErrorFallback = ValidateFormulaFallback(formulaIfErrorFallback);
    }

    public int SchemaVersion { get; }
    public string ProfileId { get; }
    public ProfileCycles Cycles { get; }
    public IReadOnlyDictionary<string, string> AutoColorColors { get; }
    public IReadOnlyList<QuickKeyBinding> QuickKeys { get; }
    public IReadOnlyList<FavoriteDefinition> Favorites { get; }
    public IReadOnlyList<StyleRecipe> LocalStyles { get; }
    public long ImmediatePreviewCellLimit { get; }
    public bool WrapSheetNavigation { get; }
    public string FormulaIfErrorFallback { get; }

    /// <summary>
    /// Resolves a named cycle to the values a command applies, substituting each
    /// AutoColor category reference for that category's current colour.
    /// </summary>
    /// <remarks>
    /// Entries that resolve to a value already present are dropped, keeping the
    /// first occurrence. Two categories may legitimately share a colour — the
    /// default palette paints both error and external red, and both same-sheet
    /// and text black. Without collapsing, the stateless advance would match the
    /// earlier index every time and the cycle would oscillate between two values
    /// with the rest unreachable. Give a category its own colour and its entry
    /// reappears automatically.
    /// </remarks>
    public IReadOnlyList<string> ResolveCycle(string family, string cycleId) =>
        Cycles.TryGet(family, cycleId, out var cycle) ? Resolve(cycle) : Array.Empty<string>();

    /// <summary>
    /// Resolves the first configured cycle in a family. Commands whose ribbon label
    /// names no particular cycle follow whichever cycle the user has placed first.
    /// </summary>
    public IReadOnlyList<string> ResolveFirstCycle(string family)
    {
        var cycles = Cycles[family];
        return cycles.Count == 0 ? Array.Empty<string>() : Resolve(cycles[0]);
    }

    private IReadOnlyList<string> Resolve(ProfileCycle cycle)
    {
        var resolved = new List<string>(cycle.Entries.Count);
        foreach (var entry in cycle.Entries)
        {
            var value = entry;
            if (ColorReference.TryResolveKey(entry, out var categoryKey)
                && AutoColorColors.TryGetValue(categoryKey, out var color))
            {
                value = color;
            }

            if (!resolved.Contains(value, StringComparer.OrdinalIgnoreCase))
            {
                resolved.Add(value);
            }
        }

        return resolved;
    }

    public ProfileDefinition WithCycles(ProfileCycles cycles) => new ProfileDefinition(
        CurrentSchemaVersion, ProfileId, cycles, AutoColorColors, QuickKeys, Favorites, LocalStyles,
        ImmediatePreviewCellLimit, WrapSheetNavigation, FormulaIfErrorFallback);

    public ProfileDefinition WithFavorites(IEnumerable<FavoriteDefinition> favorites) => new ProfileDefinition(
        CurrentSchemaVersion, ProfileId, Cycles, AutoColorColors, QuickKeys, favorites, LocalStyles,
        ImmediatePreviewCellLimit, WrapSheetNavigation, FormulaIfErrorFallback);

    public ProfileDefinition WithLocalStyles(IEnumerable<StyleRecipe> localStyles) => new ProfileDefinition(
        CurrentSchemaVersion, ProfileId, Cycles, AutoColorColors, QuickKeys, Favorites, localStyles,
        ImmediatePreviewCellLimit, WrapSheetNavigation, FormulaIfErrorFallback);

    private static string ValidateFormulaFallback(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 1024 || value[0] == '=')
            throw new ArgumentException("The IFERROR fallback must be a 1-1,024 character expression without a leading '='.", nameof(value));
        var result = new FormulaParser().Parse("=" + value, FormulaParseOptions.DefaultA1);
        if (!result.IsSuccess || result.Document!.Disposition == FormulaCoverageDisposition.InspectOnly)
            throw new ArgumentException("The IFERROR fallback is outside the qualified A1 formula subset.", nameof(value));
        return value;
    }

    private static IReadOnlyDictionary<string, string> NormalizeColorMap(IEnumerable<KeyValuePair<string, string>> values, string parameterName)
    {
        var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var item in values ?? throw new ArgumentNullException(parameterName))
        {
            var key = RequireToken(item.Key, parameterName).ToLowerInvariant();
            var color = NormalizeColor(item.Value, parameterName);
            if (result.ContainsKey(key)) throw new ArgumentException("AutoColor categories must be unique.", parameterName);
            result.Add(key, color);
        }
        var required = new[] { "text", "numeric_hardcode", "same_sheet_formula", "cross_sheet_formula", "external_formula", "error" };
        if (required.Any(key => !result.ContainsKey(key))) throw new ArgumentException("The AutoColor recipe is incomplete.", parameterName);
        return result;
    }

    private static string NormalizeColor(string value, string parameterName)
    {
        var color = value?.ToUpperInvariant() ?? string.Empty;
        if (color.Length != 7 || color[0] != '#' || !color.Skip(1).All(IsHex))
        {
            throw new ArgumentException("An AutoColor category requires a #RRGGBB value.", parameterName);
        }

        return color;
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9') || (value >= 'A' && value <= 'F');

    private static string RequireToken(string value, string parameterName) =>
        string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException("A nonempty invariant token is required.", parameterName)
            : value.Trim();
}

public sealed class QuickKeyBinding
{
    public QuickKeyBinding(string commandId, string sequence)
    {
        CommandId = string.IsNullOrWhiteSpace(commandId)
            ? throw new ArgumentException("A command ID is required.", nameof(commandId))
            : commandId.Trim();
        Sequence = string.IsNullOrWhiteSpace(sequence)
            ? throw new ArgumentException("A key sequence is required.", nameof(sequence))
            : sequence.Trim();
    }

    public string CommandId { get; }
    public string Sequence { get; }
}
