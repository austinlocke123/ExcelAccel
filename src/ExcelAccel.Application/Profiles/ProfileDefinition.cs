using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Application.Profiles;

public sealed class ProfileDefinition
{
    public const int CurrentSchemaVersion = 1;
    public const int MaximumBindings = 512;

    public ProfileDefinition(
        int schemaVersion,
        string profileId,
        IEnumerable<string> fontColorCycle,
        IEnumerable<string> fillColorCycle,
        IEnumerable<double> fontSizeCycle,
        IEnumerable<KeyValuePair<string, string>> numberFormats,
        IEnumerable<QuickKeyBinding> quickKeys,
        long immediatePreviewCellLimit,
        bool wrapSheetNavigation)
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
        if (immediatePreviewCellLimit < 1 || immediatePreviewCellLimit > 1_000_000)
        {
            throw new ArgumentOutOfRangeException(nameof(immediatePreviewCellLimit));
        }

        ImmediatePreviewCellLimit = immediatePreviewCellLimit;
        WrapSheetNavigation = wrapSheetNavigation;
    }

    public int SchemaVersion { get; }
    public string ProfileId { get; }
    public IReadOnlyList<string> FontColorCycle { get; }
    public IReadOnlyList<string> FillColorCycle { get; }
    public IReadOnlyList<double> FontSizeCycle { get; }
    public IReadOnlyDictionary<string, string> NumberFormats { get; }
    public IReadOnlyList<QuickKeyBinding> QuickKeys { get; }
    public long ImmediatePreviewCellLimit { get; }
    public bool WrapSheetNavigation { get; }

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
