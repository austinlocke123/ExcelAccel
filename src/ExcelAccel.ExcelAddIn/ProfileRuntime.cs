using System;
using System.IO;
using ExcelAccel.Application.Profiles;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.Persistence.Profiles;

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

    private static ProfileDefinition Load()
    {
        var store = new ProfileStore();
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAccel", "profile.json");
        try
        {
            return store.LoadOrDefault(path);
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException || exception is ArgumentException)
        {
            DiagnosticLog.Failure("profile.load", "invalid_profile_default_used", exception);
            return store.LoadDefault();
        }
    }
}
