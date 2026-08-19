using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Styles;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Profiles;

public sealed class ProfileDefinition
{
    public const int CurrentSchemaVersion = 5;
    public const int MaximumBindings = 512;
    public const int MaximumFavorites = 128;
    public const int MaximumLocalStyles = 64;

    public ProfileDefinition(
        int schemaVersion,
        string profileId,
        IEnumerable<string> fontColorCycle,
        IEnumerable<string> fillColorCycle,
        IEnumerable<double> fontSizeCycle,
        IEnumerable<string> horizontalAlignmentCycle,
        IEnumerable<string> verticalAlignmentCycle,
        IEnumerable<string> underlineCycle,
        IEnumerable<double> rowHeightCycle,
        IEnumerable<double> columnWidthCycle,
        IEnumerable<KeyValuePair<string, string>> autoColorColors,
        IEnumerable<KeyValuePair<string, string>> numberFormats,
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
        FontColorCycle = NormalizeColors(fontColorCycle, nameof(fontColorCycle));
        FillColorCycle = NormalizeColors(fillColorCycle, nameof(fillColorCycle));
        FontSizeCycle = (fontSizeCycle ?? throw new ArgumentNullException(nameof(fontSizeCycle))).ToArray();
        if (FontSizeCycle.Count == 0 || FontSizeCycle.Any(value => value < 6 || value > 72 || double.IsNaN(value) || double.IsInfinity(value)))
        {
            throw new ArgumentException("Font-size cycles require values from 6 through 72 points.", nameof(fontSizeCycle));
        }

        HorizontalAlignmentCycle = NormalizeTokens(horizontalAlignmentCycle, nameof(horizontalAlignmentCycle));
        VerticalAlignmentCycle = NormalizeTokens(verticalAlignmentCycle, nameof(verticalAlignmentCycle));
        UnderlineCycle = NormalizeTokens(underlineCycle, nameof(underlineCycle));
        RowHeightCycle = NormalizeDimensions(rowHeightCycle, 3, 409, nameof(rowHeightCycle));
        ColumnWidthCycle = NormalizeDimensions(columnWidthCycle, 1, 255, nameof(columnWidthCycle));
        AutoColorColors = NormalizeColorMap(autoColorColors, nameof(autoColorColors));

        var formats = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var format in numberFormats ?? throw new ArgumentNullException(nameof(numberFormats)))
        {
            var key = RequireToken(format.Key, nameof(numberFormats));
            if (string.IsNullOrWhiteSpace(format.Value) || formats.ContainsKey(key))
            {
                throw new ArgumentException("Number-format entries require unique keys and nonempty values.", nameof(numberFormats));
            }
            formats.Add(key, format.Value);
        }

        NumberFormats = formats;
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
    public IReadOnlyList<string> FontColorCycle { get; }
    public IReadOnlyList<string> FillColorCycle { get; }
    public IReadOnlyList<double> FontSizeCycle { get; }
    public IReadOnlyList<string> HorizontalAlignmentCycle { get; }
    public IReadOnlyList<string> VerticalAlignmentCycle { get; }
    public IReadOnlyList<string> UnderlineCycle { get; }
    public IReadOnlyList<double> RowHeightCycle { get; }
    public IReadOnlyList<double> ColumnWidthCycle { get; }
    public IReadOnlyDictionary<string, string> AutoColorColors { get; }
    public IReadOnlyDictionary<string, string> NumberFormats { get; }
    public IReadOnlyList<QuickKeyBinding> QuickKeys { get; }
    public IReadOnlyList<FavoriteDefinition> Favorites { get; }
    public IReadOnlyList<StyleRecipe> LocalStyles { get; }
    public long ImmediatePreviewCellLimit { get; }
    public bool WrapSheetNavigation { get; }
    public string FormulaIfErrorFallback { get; }

    public ProfileDefinition WithFavorites(IEnumerable<FavoriteDefinition> favorites) => new ProfileDefinition(
        CurrentSchemaVersion, ProfileId, FontColorCycle, FillColorCycle, FontSizeCycle,
        HorizontalAlignmentCycle, VerticalAlignmentCycle, UnderlineCycle, RowHeightCycle,
        ColumnWidthCycle, AutoColorColors, NumberFormats, QuickKeys, favorites, LocalStyles,
        ImmediatePreviewCellLimit, WrapSheetNavigation, FormulaIfErrorFallback);

    public ProfileDefinition WithLocalStyles(IEnumerable<StyleRecipe> localStyles) => new ProfileDefinition(
        CurrentSchemaVersion, ProfileId, FontColorCycle, FillColorCycle, FontSizeCycle,
        HorizontalAlignmentCycle, VerticalAlignmentCycle, UnderlineCycle, RowHeightCycle,
        ColumnWidthCycle, AutoColorColors, NumberFormats, QuickKeys, Favorites, localStyles,
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

    private static IReadOnlyList<string> NormalizeColors(IEnumerable<string> values, string parameterName)
    {
        var colors = (values ?? throw new ArgumentNullException(parameterName))
            .Select(value => value?.ToUpperInvariant() ?? string.Empty)
            .ToArray();
        if (colors.Length == 0 || colors.Any(value => value.Length != 7 || value[0] != '#' || !value.Skip(1).All(IsHex)))
        {
            throw new ArgumentException("Color cycles require one or more #RRGGBB values.", parameterName);
        }

        return colors;
    }

    private static IReadOnlyDictionary<string, string> NormalizeColorMap(IEnumerable<KeyValuePair<string, string>> values, string parameterName)
    {
        var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var item in values ?? throw new ArgumentNullException(parameterName))
        {
            var key = RequireToken(item.Key, parameterName).ToLowerInvariant();
            var color = NormalizeColors(new[] { item.Value }, parameterName)[0];
            if (result.ContainsKey(key)) throw new ArgumentException("AutoColor categories must be unique.", parameterName);
            result.Add(key, color);
        }
        var required = new[] { "text", "numeric_hardcode", "same_sheet_formula", "cross_sheet_formula", "external_formula", "error" };
        if (required.Any(key => !result.ContainsKey(key))) throw new ArgumentException("The AutoColor recipe is incomplete.", parameterName);
        return result;
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9') || (value >= 'A' && value <= 'F');

    private static IReadOnlyList<string> NormalizeTokens(IEnumerable<string> values, string parameterName)
    {
        var result = (values ?? throw new ArgumentNullException(parameterName))
            .Select(value => RequireToken(value, parameterName).ToLowerInvariant())
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        return result.Length == 0
            ? throw new ArgumentException("A formatting cycle cannot be empty.", parameterName)
            : result;
    }

    private static IReadOnlyList<double> NormalizeDimensions(IEnumerable<double> values, double minimum, double maximum, string parameterName)
    {
        var result = (values ?? throw new ArgumentNullException(parameterName)).ToArray();
        if (result.Length == 0 || result.Any(value => value < minimum || value > maximum || double.IsNaN(value) || double.IsInfinity(value)))
        {
            throw new ArgumentException($"Dimension cycles require values from {minimum} through {maximum}.", parameterName);
        }

        return result;
    }

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
