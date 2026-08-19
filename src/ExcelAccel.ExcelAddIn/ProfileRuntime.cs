using System;
using System.IO;
using ExcelAccel.Application.Profiles;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.Persistence.Profiles;
using ExcelAccel.Application.Styles;

namespace ExcelAccel.ExcelAddIn;

internal static class ProfileRuntime
{
    private static readonly object Sync = new object();
    private static ProfileDefinition? _current;

    public static ProfileDefinition Current
    {
        get
        {
            lock (Sync)
            {
                return _current ?? (_current = Load());
            }
        }
    }

    public static void Reset()
    {
        lock (Sync) _current = null;
    }

    public static bool AddFavorite(FavoriteDefinition favorite)
    {
        lock (Sync)
        {
            var current = _current ?? (_current = Load());
            var favorites = FavoriteCatalog.Add(current.Favorites, favorite);
            if (ReferenceEquals(favorites, current.Favorites)) return false;
            var updated = current.WithFavorites(favorites);
            new ProfileStore().SaveAtomic(ProfilePath, updated);
            _current = updated;
            return true;
        }
    }

    public static bool RemoveFavorite(string favoriteId)
    {
        lock (Sync)
        {
            var current = _current ?? (_current = Load());
            var favorites = FavoriteCatalog.Remove(current.Favorites, favoriteId);
            if (favorites.Count == current.Favorites.Count) return false;
            var updated = current.WithFavorites(favorites);
            new ProfileStore().SaveAtomic(ProfilePath, updated);
            _current = updated;
            return true;
        }
    }

    public static void SaveLocalStyle(StyleRecipe recipe, bool overwrite)
    {
        lock (Sync)
        {
            var current = _current ?? (_current = Load());
            var styles = StyleLibrary.AddOrReplace(current.LocalStyles, recipe, overwrite);
            var updated = current.WithLocalStyles(styles);
            new ProfileStore().SaveAtomic(ProfilePath, updated);
            _current = updated;
        }
    }

    public static bool DeleteLocalStyle(string styleId)
    {
        lock (Sync)
        {
            var current = _current ?? (_current = Load());
            var styles = StyleLibrary.Delete(current.LocalStyles, styleId);
            if (styles.Count == current.LocalStyles.Count) return false;
            var updated = current.WithLocalStyles(styles);
            new ProfileStore().SaveAtomic(ProfilePath, updated);
            _current = updated;
            return true;
        }
    }

    private static ProfileDefinition Load()
    {
        var store = new ProfileStore();
        try
        {
            return store.LoadOrDefault(ProfilePath);
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException || exception is ArgumentException)
        {
            DiagnosticLog.Failure("profile.load", "invalid_profile_default_used", exception);
            return store.LoadDefault();
        }
    }

    private static string ProfilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAccel", "profile.json");
}
