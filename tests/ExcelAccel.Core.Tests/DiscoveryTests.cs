using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

    /// <summary>
    /// The legacy shape is a checked-in fixture rather than mutated serializer
    /// output: v6 no longer emits any of the keys these migrations exercise, so
    /// deriving the old shape from the new one is no longer possible.
    /// </summary>
    private static string LegacyProfileV5() => File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", "profile-v5.json"));

    [Fact]
    public void ProfileV5MigratesNumberFormatsToOneEntryCyclesAndLiftsPropertyCycles()
    {
        var store = new ProfileStore();

        var migrated = store.Parse(LegacyProfileV5());

        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, migrated.SchemaVersion);

        // Every v5 number format becomes a one-entry cycle under its own name.
        var numberFormats = migrated.Cycles["number_format"];
        Assert.Equal(6, numberFormats.Count);
        var currency = numberFormats.Single(cycle => cycle.CycleId == "currency");
        Assert.Equal(new[] { "$#,##0.00;($#,##0.00);-" }, currency.Entries);
        Assert.Equal("Currency", currency.DisplayName);

        // Each property family keeps its whole ordered list under one cycle.
        var fontColor = Assert.Single(migrated.Cycles["font_color"]);
        Assert.Equal(new[] { "#000000", "#0000FF", "#008000", "#FF0000" }, fontColor.Entries);
        Assert.Equal(new[] { "8", "9", "10", "11", "12", "14" },
            Assert.Single(migrated.Cycles["font_size"]).Entries);
        Assert.Equal(new[] { "8.43", "10", "12", "15", "20" },
            Assert.Single(migrated.Cycles["column_width"]).Entries);
    }

    [Fact]
    public void ProfileV5CannotCarryACyclesObjectAndV6CannotOmitOne()
    {
        var store = new ProfileStore();
        var withCycles = LegacyProfileV5().Replace(
            "\"profile_id\": \"excelaccel.default.v5\",",
            "\"profile_id\": \"excelaccel.default.v5\",\n  \"cycles\": {},",
            StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => store.Parse(withCycles));

        var v6WithoutCycles = LegacyProfileV5().Replace(
            "\"schema_version\": 5", "\"schema_version\": 6", StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => store.Parse(v6WithoutCycles));
    }

    [Fact]
    public void SerializedProfileEmitsNoLegacyCycleFieldAndNoNullMember()
    {
        var store = new ProfileStore();

        var serialized = store.Serialize(store.LoadDefault());

        Assert.DoesNotContain("number_formats", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("font_color_cycle", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("column_width_cycle", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain(": null", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public void UnconfiguredFamilyIsAbsentRatherThanEmpty()
    {
        var store = new ProfileStore();
        var profile = store.LoadDefault();

        var trimmed = profile.WithCycles(new ProfileCycles(
            profile.Cycles.Families
                .Where(family => family != "underline")
                .Select(family => new KeyValuePair<string, IEnumerable<ProfileCycle>>(
                    family, profile.Cycles[family]))));

        var serialized = store.Serialize(trimmed);
        Assert.DoesNotContain("underline", serialized, StringComparison.Ordinal);
        Assert.Empty(store.Parse(serialized).Cycles["underline"]);
    }

    [Fact]
    public void AnEmptyCycleAndAnEmptyFamilyAreBothUnrepresentable()
    {
        Assert.Throws<ArgumentException>(() =>
            new ProfileCycle("font_size", "empty", "Empty", Array.Empty<string>()));

        Assert.Throws<ArgumentException>(() => new ProfileCycles(new[]
        {
            new KeyValuePair<string, IEnumerable<ProfileCycle>>("font_size", Array.Empty<ProfileCycle>()),
        }));
    }

    [Fact]
    public void NinthCycleInAFamilyIsRefusedNamingTheLimit()
    {
        var cycles = Enumerable.Range(0, 9)
            .Select(index => new ProfileCycle("number_format", "c" + index, "C" + index, new[] { "0.0" + index }))
            .ToArray();

        var error = Assert.Throws<ArgumentException>(() => new ProfileCycles(new[]
        {
            new KeyValuePair<string, IEnumerable<ProfileCycle>>("number_format", cycles),
        }));
        Assert.Contains("8", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ColorReferencesTrackTheCategoryWhileLiteralsStayPinned()
    {
        var store = new ProfileStore();
        var profile = store.LoadDefault();

        var mixed = profile.WithCycles(new ProfileCycles(new[]
        {
            new KeyValuePair<string, IEnumerable<ProfileCycle>>("font_color", new[]
            {
                new ProfileCycle("font_color", "standard", "Font Color", new[] { "@hardcode", "#123456" }),
            }),
        }));

        Assert.Equal(new[] { "#0000FF", "#123456" }, mixed.ResolveFirstCycle("font_color"));

        var recolored = new ProfileDefinition(
            ProfileDefinition.CurrentSchemaVersion,
            mixed.ProfileId,
            mixed.Cycles,
            mixed.AutoColorColors
                .Select(entry => entry.Key == "numeric_hardcode"
                    ? new KeyValuePair<string, string>(entry.Key, "#000080")
                    : entry)
                .ToArray(),
            mixed.QuickKeys, mixed.Favorites, mixed.LocalStyles,
            mixed.ImmediatePreviewCellLimit, mixed.WrapSheetNavigation, mixed.FormulaIfErrorFallback);

        Assert.Equal(new[] { "#000080", "#123456" }, recolored.ResolveFirstCycle("font_color"));
    }

    /// <summary>
    /// Two categories may share a colour, and the default palette does exactly
    /// that. Without collapsing, the stateless advance would match the earlier
    /// index forever and the cycle would oscillate between two values.
    /// </summary>
    [Fact]
    public void DefaultFontColorCycleCollapsesCategoriesThatShareAColour()
    {
        var resolved = new ProfileStore().LoadDefault().ResolveFirstCycle("font_color");

        Assert.Equal(new[] { "#000000", "#0000FF", "#008000", "#FF0000" }, resolved);
    }

    [Fact]
    public void DefaultCurrencyCycleWalksDollarEuroPoundAtZeroAndTwoDecimals()
    {
        var entries = new ProfileStore().LoadDefault().ResolveCycle("number_format", "currency");

        Assert.Equal(6, entries.Count);
        Assert.StartsWith("$#,##0_)", entries[0], StringComparison.Ordinal);
        Assert.StartsWith("$#,##0.00_)", entries[1], StringComparison.Ordinal);
        Assert.StartsWith("\u20ac#,##0_)", entries[2], StringComparison.Ordinal);
        Assert.StartsWith("\u20ac#,##0.00_)", entries[3], StringComparison.Ordinal);
        Assert.StartsWith("\u00a3#,##0_)", entries[4], StringComparison.Ordinal);
        Assert.StartsWith("\u00a3#,##0.00_)", entries[5], StringComparison.Ordinal);
        Assert.All(entries, entry => Assert.Contains("_);(", entry, StringComparison.Ordinal));
    }

    [Fact]
    public void ProfileV2MigratesToCurrentSchemaWithEmptyPhase1BCollections()
    {
        var store = new ProfileStore();
        var v2 = LegacyProfileV5().Replace("\"schema_version\": 5", "\"schema_version\": 2", StringComparison.Ordinal)
            .Replace("  \"favorites\": [],\n", string.Empty, StringComparison.Ordinal)
            .Replace("  \"local_styles\": [],\n", string.Empty, StringComparison.Ordinal);

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
        var v3 = LegacyProfileV5()
            .Replace("\"schema_version\": 5", "\"schema_version\": 3", StringComparison.Ordinal)
            .Replace("\"favorites\": []",
                "\"favorites\": [ { \"favorite_id\": \"favorite.currency\", \"command_id\": \"format.number.currency\", \"contract_version\": 1, \"arguments\": {} } ]",
                StringComparison.Ordinal)
            .Replace("  \"local_styles\": [],\n", string.Empty, StringComparison.Ordinal);

        var migrated = store.Parse(v3);

        Assert.Single(migrated.Favorites);
        Assert.Empty(migrated.LocalStyles);
        Assert.Equal(ProfileDefinition.CurrentSchemaVersion, migrated.SchemaVersion);
    }

    [Fact]
    public void ProfileV4MigratesWithQualifiedDefaultIfErrorFallback()
    {
        var store = new ProfileStore();
        var v4 = LegacyProfileV5()
            .Replace("\"schema_version\": 5", "\"schema_version\": 4", StringComparison.Ordinal)
            .Replace(",\n  \"formula_iferror_fallback\": \"0\"", string.Empty, StringComparison.Ordinal);

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
