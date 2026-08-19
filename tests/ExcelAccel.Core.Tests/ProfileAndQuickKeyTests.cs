using System;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.QuickKeys;
using ExcelAccel.Persistence.Profiles;
using System.Collections.Generic;
using ExcelAccel.Application.Styles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ProfileAndQuickKeyTests
{
    [Fact]
    public void EmbeddedDefaultProfileIsValidAndReferencesRegisteredCommands()
    {
        var profile = new ProfileStore().LoadDefault();

        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, profile.SchemaVersion);
        Assert.Empty(QuickKeyValidator.Validate(profile.QuickKeys));
        Assert.All(profile.QuickKeys, binding =>
            Assert.Contains(BuiltInCommandRegistry.All, command => command.Id == binding.CommandId));
    }

    [Fact]
    public void ProfileSerializationIsDeterministicAndRoundTrips()
    {
        var store = new ProfileStore();
        var profile = store.LoadDefault();

        var first = store.Serialize(profile);
        var second = store.Serialize(store.Parse(first));

        Assert.Equal(first, second);
    }

    [Fact]
    public void AtomicSavePreservesPriorProfileAsBackup()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "profile.json");
        var store = new ProfileStore();
        var profile = store.LoadDefault();

        store.SaveAtomic(path, profile);
        store.SaveAtomic(path, profile);

        Assert.Equal(profile.ProfileId, store.Load(path).ProfileId);
        Assert.True(File.Exists(path + ".bak"));
        Assert.Empty(Directory.GetFiles(sandbox.Path, "*.tmp"));
    }

    [Fact]
    public void UnknownProfileFieldsAreRefused()
    {
        var store = new ProfileStore();
        var json = store.Serialize(store.LoadDefault()).Replace(
            "\"profile_id\":",
            "\"unexpected\": true,\n  \"profile_id\":",
            StringComparison.Ordinal);

        Assert.Throws<InvalidDataException>(() => store.Parse(json));
    }

    [Fact]
    public void FavoritesRoundTripThroughAtomicProfileStorageWithoutChangingOtherSettings()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "profile.json");
        var store = new ProfileStore();
        var original = store.LoadDefault();
        var favorite = new FavoriteDefinition("favorite.currency", "format.number.currency", 1,
            new[] { new KeyValuePair<string, string>("style", "accounting") });

        store.SaveAtomic(path, original.WithFavorites(FavoriteCatalog.Add(original.Favorites, favorite)));
        var loaded = store.Load(path);

        Assert.Single(loaded.Favorites);
        Assert.Equal("accounting", loaded.Favorites[0].Arguments["style"]);
        Assert.Equal(original.NumberFormats, loaded.NumberFormats);
        Assert.Equal(original.QuickKeys.Select(value => value.Sequence).OrderBy(value => value, StringComparer.Ordinal),
            loaded.QuickKeys.Select(value => value.Sequence).OrderBy(value => value, StringComparer.Ordinal));
    }

    [Fact]
    public void LocalStyleRoundTripsDeterministicallyThroughProfileStorage()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "profile.json");
        var store = new ProfileStore();
        var original = store.LoadDefault();
        var style = new StyleRecipe("local.review", 1, "Review", StyleOrigin.Local, UnsupportedStylePropertyPolicy.Refuse,
            new[] { new KeyValuePair<string, string>("font_bold", "true"), new KeyValuePair<string, string>("fill_color", "#FFF2CC") });

        store.SaveAtomic(path, original.WithLocalStyles(StyleLibrary.AddOrReplace(original.LocalStyles, style, overwrite: false)));
        var loaded = store.Load(path);

        Assert.Single(loaded.LocalStyles);
        Assert.Equal("#FFF2CC", loaded.LocalStyles[0].Properties["fill_color"]);
        Assert.Equal(store.Serialize(loaded), store.Serialize(store.Parse(store.Serialize(loaded))));
    }

    [Fact]
    public void QuickKeyValidationFindsReservedDuplicateAndPrefixConflicts()
    {
        var conflicts = QuickKeyValidator.Validate(new[]
        {
            new QuickKeyBinding("one", "Ctrl+C"),
            new QuickKeyBinding("two", "Ctrl+Shift+K"),
            new QuickKeyBinding("three", "Ctrl+Shift+K, C"),
            new QuickKeyBinding("four", "Ctrl+Shift+K, C"),
        });

        Assert.Contains(conflicts, conflict => conflict.Reason.Contains("reserved", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(conflicts, conflict => conflict.Reason.Contains("assigned more than once", StringComparison.Ordinal));
        Assert.Contains(conflicts, conflict => conflict.Reason.Contains("prefix", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void QuickKeyEngineNeverInterceptsEditModeAndSupportsTimeoutAndEscape()
    {
        var binding = new QuickKeyBinding("test.command", "Ctrl+Shift+K, C");
        var engine = new QuickKeyEngine(new[] { binding }, TimeSpan.FromSeconds(1));
        var now = DateTimeOffset.Parse("2026-08-19T00:00:00Z");

        Assert.Equal(QuickKeyOutcome.NotHandled, engine.ProcessStroke("Ctrl+Shift+K", now, excelEditMode: true).Outcome);
        Assert.Equal(QuickKeyOutcome.AwaitingNextStroke, engine.ProcessStroke("Ctrl+Shift+K", now, excelEditMode: false).Outcome);
        Assert.Equal(QuickKeyOutcome.Cancelled, engine.ProcessStroke("Escape", now, excelEditMode: false).Outcome);
        Assert.Equal(QuickKeyOutcome.AwaitingNextStroke, engine.ProcessStroke("Ctrl+Shift+K", now, excelEditMode: false).Outcome);
        Assert.Equal(QuickKeyOutcome.TimedOut, engine.ProcessStroke("C", now.AddSeconds(2), excelEditMode: false).Outcome);
        Assert.Equal(QuickKeyOutcome.AwaitingNextStroke, engine.ProcessStroke("Ctrl+Shift+K", now.AddSeconds(3), excelEditMode: false).Outcome);
        var invoked = engine.ProcessStroke("C", now.AddSeconds(3.5), excelEditMode: false);
        Assert.Equal(QuickKeyOutcome.Invoke, invoked.Outcome);
        Assert.Equal("test.command", invoked.CommandId);
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excelaccel-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
