using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAccel.Application.Styles;

public enum StyleOrigin { BuiltIn = 0, Local = 1, OrganizationLocked = 2 }
public enum UnsupportedStylePropertyPolicy { Refuse = 0, Skip = 1 }

public sealed class StyleRecipe
{
    public const int CurrentVersion = 1;
    public StyleRecipe(string styleId, int version, string displayName, StyleOrigin origin,
        UnsupportedStylePropertyPolicy unsupportedPropertyPolicy, IEnumerable<KeyValuePair<string, string>> properties)
    {
        StyleId = RequireText(styleId, nameof(styleId), 128);
        if (version != CurrentVersion) throw new ArgumentOutOfRangeException(nameof(version), "The style recipe version is unsupported.");
        Version = version;
        DisplayName = RequireText(displayName, nameof(displayName), 128);
        Origin = origin;
        UnsupportedPropertyPolicy = unsupportedPropertyPolicy;
        var normalized = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in properties ?? throw new ArgumentNullException(nameof(properties)))
        {
            var id = RequireText(property.Key, nameof(properties), 64).ToLowerInvariant();
            if (normalized.ContainsKey(id)) throw new ArgumentException($"Duplicate style property '{id}'.", nameof(properties));
            normalized.Add(id, StylePropertyCatalog.Normalize(id, property.Value));
        }
        if (normalized.Count < 1 || normalized.Count > StylePropertyCatalog.MaximumProperties)
            throw new ArgumentException($"A style requires 1 through {StylePropertyCatalog.MaximumProperties} supported properties.", nameof(properties));
        Properties = normalized;
    }

    public string StyleId { get; }
    public int Version { get; }
    public string DisplayName { get; }
    public StyleOrigin Origin { get; }
    public UnsupportedStylePropertyPolicy UnsupportedPropertyPolicy { get; }
    public IReadOnlyDictionary<string, string> Properties { get; }
    public bool IsDeletable => Origin == StyleOrigin.Local;

    private static string RequireText(string value, string parameterName, int maximum)
    {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A nonempty value is required.", parameterName);
        var result = value.Trim();
        if (result.Length > maximum || result.Any(char.IsControl)) throw new ArgumentException($"The value must contain at most {maximum} non-control characters.", parameterName);
        return result;
    }
}

public static class StylePropertyCatalog
{
    public const int MaximumProperties = 12;
    public static readonly IReadOnlyList<string> Supported = new[]
    {
        "borders", "fill_color", "font_bold", "font_color", "font_italic", "font_name", "font_size",
        "horizontal_alignment", "indent_level", "number_format", "underline", "vertical_alignment",
    };

    public static string Normalize(string propertyId, string value)
    {
        if (!Supported.Contains(propertyId, StringComparer.Ordinal))
            throw new ArgumentException($"Style property '{propertyId}' is not supported.", nameof(propertyId));
        var candidate = value ?? string.Empty;
        if (candidate.Length > 256 || candidate.Any(character => char.IsControl(character) && character != '\t'))
            throw new ArgumentException("Style property values must be bounded non-control text.", nameof(value));
        switch (propertyId)
        {
            case "font_color":
            case "fill_color":
                var color = candidate.ToUpperInvariant();
                if (color.Length != 7 || color[0] != '#' || !color.Skip(1).All(IsHex)) throw new ArgumentException("A #RRGGBB color is required.", nameof(value));
                return color;
            case "font_bold":
            case "font_italic":
                if (!bool.TryParse(candidate, out var boolean)) throw new ArgumentException("A Boolean style value is required.", nameof(value));
                return boolean.ToString().ToLowerInvariant();
            case "font_size":
                if (!double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) || size < 6 || size > 72)
                    throw new ArgumentException("Font size must be from 6 through 72 points.", nameof(value));
                return size.ToString("0.####", CultureInfo.InvariantCulture);
            case "indent_level":
                if (!int.TryParse(candidate, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indent) || indent < 0 || indent > 15)
                    throw new ArgumentException("Indent level must be from 0 through 15.", nameof(value));
                return indent.ToString(CultureInfo.InvariantCulture);
            case "horizontal_alignment": return RequireToken(candidate, new[] { "general", "left", "center", "right", "center_across" });
            case "vertical_alignment": return RequireToken(candidate, new[] { "bottom", "center", "top" });
            case "underline": return RequireToken(candidate, new[] { "none", "single", "double" });
            case "borders": return RequireToken(candidate, new[] { "none", "sum_bar" });
            case "font_name":
                if (string.IsNullOrWhiteSpace(candidate) || candidate.Length > 64) throw new ArgumentException("A bounded font name is required.", nameof(value));
                return candidate.Trim();
            case "number_format":
                if (string.IsNullOrWhiteSpace(candidate)) throw new ArgumentException("A nonempty number format is required.", nameof(value));
                return candidate;
            default: throw new ArgumentException($"Style property '{propertyId}' is not supported.", nameof(propertyId));
        }
    }

    private static string RequireToken(string value, IEnumerable<string> allowed)
    {
        var token = value.Trim().ToLowerInvariant();
        if (!allowed.Contains(token, StringComparer.Ordinal)) throw new ArgumentException($"Unsupported style token '{value}'.", nameof(value));
        return token;
    }
    private static bool IsHex(char value) => (value >= '0' && value <= '9') || (value >= 'A' && value <= 'F');
}
