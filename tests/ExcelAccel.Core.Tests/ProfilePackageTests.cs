using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.Styles;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ProfilePackageTests
{
    [Fact]
    public void ExportIsDeterministicPreviewedAndRoundTripsThroughStrictImport()
    {
        using var sandbox = new TemporaryDirectory();
        var store = new ProfileStore();
        var service = new ProfilePackageService(store);
        var profile = ModifiedProfile(store);
        var destination = Path.Combine(sandbox.Path, "profile.excelaccel.json");
        var first = service.PlanExport(profile, destination, overwrite: false);
        var second = service.PlanExport(profile, Path.Combine(sandbox.Path, "other.json"), overwrite: false);

        Assert.Equal(first.PackageSha256, second.PackageSha256);
        Assert.Equal(first.PackageBytes, second.PackageBytes);
        Assert.Throws<InvalidOperationException>(() => service.Export(first, "wrong"));
        service.Export(first, first.PlanHash);
        var import = service.PreviewImport(destination, store.LoadDefault(), BuiltInCommandRegistry.All);

        Assert.Equal(0, import.Manifest.AssetCount);
        Assert.Equal(1, import.Diff.FavoritesAdded);
        Assert.Equal(1, import.Diff.StylesAdded);
        Assert.Equal(profile.ProfileId, import.ImportedProfile.ProfileId);
    }

    [Fact]
    public void ExportRefusesExistingDestinationWithoutOverwriteAndDetectsPreviewDrift()
    {
        using var sandbox = new TemporaryDirectory();
        var service = new ProfilePackageService();
        var profile = new ProfileStore().LoadDefault();
        var destination = Path.Combine(sandbox.Path, "profile.json");
        File.WriteAllText(destination, "prior");
        Assert.Throws<IOException>(() => service.PlanExport(profile, destination, overwrite: false));
        var plan = service.PlanExport(profile, destination, overwrite: true);
        File.WriteAllText(destination, "changed-after-preview");

        Assert.Throws<IOException>(() => service.Export(plan, plan.PlanHash));
        Assert.Equal("changed-after-preview", File.ReadAllText(destination));
    }

    [Fact]
    public void ImportApplyRequiresExactSourceAndCurrentProfileAndCreatesBoundedBackup()
    {
        using var sandbox = new TemporaryDirectory();
        var store = new ProfileStore();
        var service = new ProfilePackageService(store);
        var current = store.LoadDefault();
        var imported = ModifiedProfile(store);
        var package = Path.Combine(sandbox.Path, "package.json");
        var export = service.PlanExport(imported, package, overwrite: false);
        service.Export(export, export.PlanHash);
        var plan = service.PreviewImport(package, current, BuiltInCommandRegistry.All);
        var profilePath = Path.Combine(sandbox.Path, "active", "profile.json");
        store.SaveAtomic(profilePath, current);

        var applied = service.ApplyImport(plan, plan.PlanHash, profilePath, current, BuiltInCommandRegistry.All);

        Assert.Single(applied.Favorites);
        Assert.Single(store.Load(profilePath).LocalStyles);
        Assert.True(File.Exists(profilePath + ".import.bak"));
        Assert.Equal(current.ProfileId, store.Load(profilePath + ".import.bak").ProfileId);
    }

    [Fact]
    public void ImportRefusesTamperedPayloadAssetsAndUnknownCommandReferences()
    {
        using var sandbox = new TemporaryDirectory();
        var store = new ProfileStore();
        var service = new ProfilePackageService(store);
        var current = store.LoadDefault();
        var path = Path.Combine(sandbox.Path, "package.json");
        var export = service.PlanExport(current, path, overwrite: false);
        service.Export(export, export.PlanHash);
        var original = File.ReadAllText(path);

        File.WriteAllText(path, original.Replace("\"asset_count\": 0", "\"asset_count\": 1", StringComparison.Ordinal));
        Assert.Throws<InvalidDataException>(() => service.PreviewImport(path, current, BuiltInCommandRegistry.All));
        File.WriteAllText(path, original.Replace(export.Manifest.PayloadSha256, new string('0', 64), StringComparison.Ordinal));
        Assert.Throws<InvalidDataException>(() => service.PreviewImport(path, current, BuiltInCommandRegistry.All));

        var unknown = current.WithFavorites(new[] { new FavoriteDefinition("missing", "missing.command", 1) });
        var unknownPath = Path.Combine(sandbox.Path, "unknown.json");
        var unknownExport = service.PlanExport(unknown, unknownPath, overwrite: false);
        service.Export(unknownExport, unknownExport.PlanHash);
        Assert.Throws<InvalidDataException>(() => service.PreviewImport(unknownPath, current, BuiltInCommandRegistry.All));
    }

    [Fact]
    public void ImportRefusesSourceOrActiveProfileDriftAfterPreview()
    {
        using var sandbox = new TemporaryDirectory();
        var store = new ProfileStore();
        var service = new ProfilePackageService(store);
        var current = store.LoadDefault();
        var path = Path.Combine(sandbox.Path, "package.json");
        var export = service.PlanExport(ModifiedProfile(store), path, overwrite: false);
        service.Export(export, export.PlanHash);
        var plan = service.PreviewImport(path, current, BuiltInCommandRegistry.All);
        File.AppendAllText(path, " ");
        Assert.Throws<InvalidDataException>(() => service.ApplyImport(plan, plan.PlanHash,
            Path.Combine(sandbox.Path, "profile.json"), current, BuiltInCommandRegistry.All));

        File.WriteAllBytes(path, export.PackageBytes);
        var changedCurrent = current.WithFavorites(new[] { new FavoriteDefinition("favorite.currency", "format.number.currency", 1) });
        Assert.Throws<InvalidOperationException>(() => service.ApplyImport(plan, plan.PlanHash,
            Path.Combine(sandbox.Path, "profile.json"), changedCurrent, BuiltInCommandRegistry.All));
    }

    [Fact]
    public void ExportRefusesUnqualifiedFixedFavoriteArguments()
    {
        var store = new ProfileStore();
        var profile = store.LoadDefault().WithFavorites(new[]
        {
            new FavoriteDefinition("fixed", "format.number.currency", 1,
                new[] { new KeyValuePair<string, string>("path", "C:\\Workbook.xlsx") })
        });
        Assert.Throws<InvalidDataException>(() => new ProfilePackageService(store).PlanExport(profile, "profile.json", false));
    }

    [Fact]
    public void BindingCheatSheetIsDeterministicPreviewedAndLocalOnly()
    {
        using var sandbox = new TemporaryDirectory();
        var profile = new ProfileStore().LoadDefault();
        var exporter = new BindingCheatSheetExporter();
        var csvPath = Path.Combine(sandbox.Path, "bindings.csv");
        var first = exporter.Plan(profile.QuickKeys, BuiltInCommandRegistry.All, csvPath, BindingExportFormat.Csv, false);
        var second = exporter.Plan(profile.QuickKeys.Reverse(), BuiltInCommandRegistry.All, Path.Combine(sandbox.Path, "other.csv"), BindingExportFormat.Csv, false);

        Assert.Equal(first.Content, second.Content);
        Assert.Throws<InvalidOperationException>(() => exporter.Export(first, "wrong"));
        exporter.Export(first, first.PlanHash);
        var text = File.ReadAllText(csvPath);
        Assert.Contains("sequence,command_id,command_name,category,conflict", text, StringComparison.Ordinal);
        Assert.DoesNotContain("http", text, StringComparison.OrdinalIgnoreCase);

        var html = exporter.Plan(profile.QuickKeys, BuiltInCommandRegistry.All, Path.Combine(sandbox.Path, "bindings.html"), BindingExportFormat.Html, false);
        exporter.Export(html, html.PlanHash);
        Assert.Contains("<!doctype html>", File.ReadAllText(html.DestinationPath), StringComparison.OrdinalIgnoreCase);
    }

    private static ProfileDefinition ModifiedProfile(ProfileStore store)
    {
        var profile = store.LoadDefault().WithFavorites(new[] { new FavoriteDefinition("favorite.currency", "format.number.currency", 1) });
        var style = new StyleRecipe("local.review", 1, "Review", StyleOrigin.Local, UnsupportedStylePropertyPolicy.Refuse,
            new[] { new KeyValuePair<string, string>("fill_color", "#FFF2CC") });
        return profile.WithLocalStyles(new[] { style });
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory() { Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excelaccel-tests", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(Path); }
        public string Path { get; }
        public void Dispose() { if (Directory.Exists(Path)) Directory.Delete(Path, true); }
    }
}
