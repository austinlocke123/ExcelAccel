using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Application.Styles;

public static class BuiltInStyleCatalog
{
    private static readonly IReadOnlyList<StyleRecipe> Recipes = new[]
    {
        Recipe("major_header", "Major Header", P("font_bold", "true"), P("font_size", "14"), P("font_color", "#FFFFFF"), P("fill_color", "#1F4E78")),
        Recipe("minor_header", "Minor Header", P("font_bold", "true"), P("font_color", "#1F1F1F"), P("fill_color", "#D9EAF7")),
        Recipe("date_header", "Date Header", P("font_bold", "true"), P("horizontal_alignment", "center"), P("number_format", "mmm-yy")),
        Recipe("assumption", "Assumption", P("font_color", "#0000FF"), P("fill_color", "#FFF2CC")),
        Recipe("formula", "Formula", P("font_color", "#000000")),
        Recipe("linked_formula", "Linked Formula", P("font_color", "#008000")),
        Recipe("output", "Output", P("font_bold", "true"), P("borders", "sum_bar")),
        Recipe("warning", "Warning", P("font_bold", "true"), P("font_color", "#FF0000"), P("fill_color", "#F4CCCC")),
        Recipe("total", "Total", P("font_bold", "true"), P("borders", "sum_bar")),
    };

    public static IReadOnlyList<StyleRecipe> All => Recipes;
    public static StyleRecipe GetRequired(string styleId) => Recipes.SingleOrDefault(value => string.Equals(value.StyleId, styleId, StringComparison.Ordinal))
        ?? throw new KeyNotFoundException($"Built-in style '{styleId}' is not registered.");

    private static StyleRecipe Recipe(string id, string name, params KeyValuePair<string, string>[] properties) =>
        new StyleRecipe(id, StyleRecipe.CurrentVersion, name, StyleOrigin.BuiltIn, UnsupportedStylePropertyPolicy.Refuse, properties);
    private static KeyValuePair<string, string> P(string id, string value) => new KeyValuePair<string, string>(id, value);
}
