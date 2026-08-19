using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Discovery;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DiscoveryTests
{
    [Fact]
    public void SearchRanksAllMetadataDeterministicallyAndIncludesAvailability()
    {
        var commands = new[]
        {
            Descriptor("format.currency", "Currency Format", "Formatting", "Apply an accounting display.", new[] { "money" }, "Ctrl+M"),
            Descriptor("inspect.money", "Inspect Selection", "Audit", "Find money cells.", Array.Empty<string>(), "Ctrl+I"),
        };
        var index = new CommandSearchIndex(commands);

        var first = index.Search("money", command => command.Id == "format.currency"
            ? CanExecuteResult.Permit()
            : CanExecuteResult.Refuse("NO_SELECTION", "Unavailable.", "Select a range."));
        var second = index.Search("  MONEY  ", _ => CanExecuteResult.Permit());

        Assert.Equal(new[] { "format.currency", "inspect.money" }, first.Select(value => value.Command.Id));
        Assert.Equal(first.Select(value => value.Command.Id), second.Select(value => value.Command.Id));
        Assert.True(first[0].Availability.Allowed);
        Assert.False(first[1].Availability.Allowed);
        Assert.Equal("Select a range.", first[1].Availability.Remediation);
    }

    [Fact]
    public void SearchUsesOnlyBoundedRegistryMetadataWithinImmediateBudget()
    {
        var commands = Enumerable.Range(0, CommandSearchIndex.MaximumRegistrySize)
            .Select(index => Descriptor($"command.{index:D4}", $"Command {index:D4}", "Test", "Deterministic local metadata.",
                new[] { $"alias-{index:D4}" }, $"Alt,{index:D4}"))
            .ToArray();
        var indexUnderTest = new CommandSearchIndex(commands);
        _ = indexUnderTest.Search("alias-2047", _ => CanExecuteResult.Permit());
        var watch = Stopwatch.StartNew();

        var results = indexUnderTest.Search("alias-2047", _ => CanExecuteResult.Permit());

        watch.Stop();
        Assert.Single(results);
        Assert.Equal("command.2047", results[0].Command.Id);
        Assert.True(watch.ElapsedMilliseconds < 100, $"Search took {watch.ElapsedMilliseconds} ms.");
    }

    [Fact]
    public void FavoriteAddAndRemoveAreIdempotent()
    {
        var favorite = new FavoriteDefinition("favorite.currency", "format.currency", 1);
        var first = FavoriteCatalog.Add(Array.Empty<FavoriteDefinition>(), favorite);
        var duplicate = FavoriteCatalog.Add(first, new FavoriteDefinition("favorite.currency", "format.currency", 1));

        Assert.Same(first, duplicate);
        Assert.Empty(FavoriteCatalog.Remove(FavoriteCatalog.Remove(duplicate, favorite.FavoriteId), favorite.FavoriteId));
        Assert.Throws<InvalidOperationException>(() => FavoriteCatalog.Add(first,
            new FavoriteDefinition("favorite.currency", "different.command", 1)));
    }

    [Fact]
    public void MissingAndIncompatibleFavoritesRemainVisibleWithRemediation()
    {
        var command = Descriptor("format.currency", "Currency Format", "Formatting", "Format numbers.", Array.Empty<string>(), "Ctrl+M");
        var missing = FavoriteCatalog.Resolve(new FavoriteDefinition("missing", "missing.command", 1), new[] { command }, _ => CanExecuteResult.Permit());
        var incompatible = FavoriteCatalog.Resolve(new FavoriteDefinition("old", command.Id, 2), new[] { command }, _ => CanExecuteResult.Permit());

        Assert.Equal(FavoriteResolutionStatus.MissingCommand, missing.Status);
        Assert.False(string.IsNullOrWhiteSpace(missing.Remediation));
        Assert.Equal(FavoriteResolutionStatus.IncompatibleVersion, incompatible.Status);
        Assert.False(incompatible.CanInvoke);
    }

    [Fact]
    public void FavoriteInvocationCanOnlyUseNormalRouterAndCurrentAvailability()
    {
        var command = Descriptor("format.currency", "Currency Format", "Formatting", "Format numbers.", Array.Empty<string>(), "Ctrl+M");
        var favorite = new FavoriteDefinition("currency", command.Id, command.ContractVersion,
            new[] { new KeyValuePair<string, string>("style", "accounting") });
        string? routedCommand = null;
        InvocationSource? routedSource = null;

        var result = FavoriteCatalog.Invoke(favorite, new[] { command }, _ => CanExecuteResult.Permit(),
            (commandId, arguments, source) =>
            {
                routedCommand = commandId;
                routedSource = source;
                Assert.Equal("accounting", arguments["style"]);
                return CommandResult.Success(commandId, "routed");
            });

        Assert.True(result.Succeeded);
        Assert.Equal(command.Id, routedCommand);
        Assert.Equal(InvocationSource.Favorite, routedSource);

        var refused = FavoriteCatalog.Invoke(favorite, new[] { command },
            _ => CanExecuteResult.Refuse("LOCKED", "Locked.", "Unlock it."),
            (_, __, ___) => throw new InvalidOperationException("The router must not be called."));
        Assert.Equal(CommandResultStatus.Refused, refused.Status);
    }

    [Fact]
    public void ProfileV2MigratesToCurrentSchemaWithEmptyPhase1BCollections()
    {
        var store = new ProfileStore();
        var current = store.Serialize(store.LoadDefault());
        var v2 = current.Replace("\"schema_version\": 5", "\"schema_version\": 2", StringComparison.Ordinal)
            .Replace("  \"favorites\": []," + Environment.NewLine, string.Empty, StringComparison.Ordinal)
            .Replace("  \"local_styles\": []," + Environment.NewLine, string.Empty, StringComparison.Ordinal);

        var migrated = store.Parse(v2);

        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, migrated.SchemaVersion);
        Assert.Empty(migrated.Favorites);
        Assert.Empty(migrated.LocalStyles);
        Assert.Contains("\"favorites\": []", store.Serialize(migrated), StringComparison.Ordinal);
    }

    [Fact]
    public void ProfileV3MigratesFavoritesAndAddsEmptyLocalStyles()
    {
        var store = new ProfileStore();
        var profile = store.LoadDefault().WithFavorites(new[] { new FavoriteDefinition("favorite.currency", "format.number.currency", 1) });
        var v3 = store.Serialize(profile)
            .Replace("\"schema_version\": 5", "\"schema_version\": 3", StringComparison.Ordinal)
            .Replace("  \"local_styles\": []," + Environment.NewLine, string.Empty, StringComparison.Ordinal);

        var migrated = store.Parse(v3);

        Assert.Single(migrated.Favorites);
        Assert.Empty(migrated.LocalStyles);
        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, migrated.SchemaVersion);
    }

    [Fact]
    public void ProfileV4MigratesWithQualifiedDefaultIfErrorFallback()
    {
        var store = new ProfileStore();
        var v4 = store.Serialize(store.LoadDefault())
            .Replace("\"schema_version\": 5", "\"schema_version\": 4", StringComparison.Ordinal)
            .Replace("  \"formula_iferror_fallback\": \"0\"" + Environment.NewLine, string.Empty, StringComparison.Ordinal)
            .Replace("  \"wrap_sheet_navigation\": true," + Environment.NewLine,
                "  \"wrap_sheet_navigation\": true" + Environment.NewLine, StringComparison.Ordinal);

        var migrated = store.Parse(v4);

        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, migrated.SchemaVersion);
        Assert.Equal("0", migrated.FormulaIfErrorFallback);
        Assert.Contains("\"formula_iferror_fallback\": \"0\"", store.Serialize(migrated), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("=0")]
    [InlineData("NamedFallback")]
    [InlineData("Table1[Fallback]")]
    public void ProfileRejectsUnsafeIfErrorFallback(string fallback)
    {
        var store = new ProfileStore();
        var json = store.Serialize(store.LoadDefault())
            .Replace("\"formula_iferror_fallback\": \"0\"", "\"formula_iferror_fallback\": \"" + fallback + "\"", StringComparison.Ordinal);

        Assert.Throws<System.IO.InvalidDataException>(() => store.Parse(json));
    }

    private static CommandDescriptor Descriptor(string id, string name, string category, string description,
        IEnumerable<string> aliases, string shortcut) =>
        new CommandDescriptor(id, 1, name, CommandImpact.Low, new[] { "test_property" }, true, shortcut,
            "CAP-TEST", CommandContextRequirement.Selection, PreviewPolicy.None, UndoPolicy.None,
            new[] { "AC-TEST" }, category, description, aliases, shortcut);
}
