using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Application.Styles;

public static class StyleLibrary
{
    public static IReadOnlyList<StyleRecipe> AddOrReplace(IReadOnlyList<StyleRecipe> localStyles, StyleRecipe recipe, bool overwrite)
    {
        if (localStyles is null) throw new ArgumentNullException(nameof(localStyles));
        if (recipe is null) throw new ArgumentNullException(nameof(recipe));
        if (recipe.Origin != StyleOrigin.Local) throw new InvalidOperationException("Only local styles can be written to the user profile.");
        if (BuiltInStyleCatalog.All.Any(value => string.Equals(value.StyleId, recipe.StyleId, StringComparison.Ordinal)))
            throw new InvalidOperationException("A local style cannot replace a built-in style ID.");
        var existing = localStyles.FirstOrDefault(value => string.Equals(value.StyleId, recipe.StyleId, StringComparison.Ordinal));
        if (existing is not null && !overwrite) throw new InvalidOperationException($"Local style '{recipe.StyleId}' already exists.");
        return localStyles.Where(value => !string.Equals(value.StyleId, recipe.StyleId, StringComparison.Ordinal))
            .Concat(new[] { recipe }).OrderBy(value => value.StyleId, StringComparer.Ordinal).ToArray();
    }

    public static IReadOnlyList<StyleRecipe> Delete(IReadOnlyList<StyleRecipe> localStyles, string styleId)
    {
        if (localStyles is null) throw new ArgumentNullException(nameof(localStyles));
        if (string.IsNullOrWhiteSpace(styleId)) throw new ArgumentException("A style ID is required.", nameof(styleId));
        if (BuiltInStyleCatalog.All.Any(value => string.Equals(value.StyleId, styleId, StringComparison.Ordinal)))
            throw new InvalidOperationException("Built-in styles cannot be deleted.");
        return localStyles.Where(value => !string.Equals(value.StyleId, styleId, StringComparison.Ordinal)).ToArray();
    }

    public static IReadOnlyList<StyleRecipe> Effective(IReadOnlyList<StyleRecipe> localStyles) =>
        BuiltInStyleCatalog.All.Concat(localStyles ?? throw new ArgumentNullException(nameof(localStyles)))
            .OrderBy(value => value.DisplayName, StringComparer.OrdinalIgnoreCase).ThenBy(value => value.StyleId, StringComparer.Ordinal).ToArray();
}
