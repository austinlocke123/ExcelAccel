using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using Newtonsoft.Json;

namespace ExcelAccel.Persistence.Profiles;

public sealed class ProfilePackageManifest
{
    public ProfilePackageManifest(int profileSchemaVersion, int quickKeyCount, int favoriteCount, int localStyleCount,
        int assetCount, long payloadBytes, string payloadSha256)
    {
        ProfileSchemaVersion = profileSchemaVersion; QuickKeyCount = quickKeyCount; FavoriteCount = favoriteCount;
        LocalStyleCount = localStyleCount; AssetCount = assetCount; PayloadBytes = payloadBytes; PayloadSha256 = payloadSha256;
    }
    public int ProfileSchemaVersion { get; }
    public int QuickKeyCount { get; }
    public int FavoriteCount { get; }
    public int LocalStyleCount { get; }
    public int AssetCount { get; }
    public long PayloadBytes { get; }
    public string PayloadSha256 { get; }
    public override string ToString() =>
        $"Profile schema: {ProfileSchemaVersion}\nQuick Keys: {QuickKeyCount}\nFavorites: {FavoriteCount}\nLocal styles: {LocalStyleCount}\nAssets: {AssetCount}\nPayload bytes: {PayloadBytes}\nSHA-256: {PayloadSha256}";
}

public sealed class ProfileExportPlan
{
    public ProfileExportPlan(string destinationPath, bool overwrite, string existingDestinationSha256,
        byte[] packageBytes, string packageSha256, ProfilePackageManifest manifest, string planHash)
    {
        DestinationPath = destinationPath; Overwrite = overwrite; ExistingDestinationSha256 = existingDestinationSha256;
        PackageBytes = packageBytes; PackageSha256 = packageSha256; Manifest = manifest; PlanHash = planHash;
    }
    public string DestinationPath { get; }
    public bool Overwrite { get; }
    public string ExistingDestinationSha256 { get; }
    public byte[] PackageBytes { get; }
    public string PackageSha256 { get; }
    public ProfilePackageManifest Manifest { get; }
    public string PlanHash { get; }
}

public sealed class ProfileImportDiff
{
    public ProfileImportDiff(int quickKeysBefore, int quickKeysAfter, int favoritesAdded, int favoritesReplaced,
        int favoritesRemoved, int stylesAdded, int stylesReplaced, int stylesRemoved)
    {
        QuickKeysBefore = quickKeysBefore; QuickKeysAfter = quickKeysAfter; FavoritesAdded = favoritesAdded;
        FavoritesReplaced = favoritesReplaced; FavoritesRemoved = favoritesRemoved; StylesAdded = stylesAdded;
        StylesReplaced = stylesReplaced; StylesRemoved = stylesRemoved;
    }
    public int QuickKeysBefore { get; }
    public int QuickKeysAfter { get; }
    public int FavoritesAdded { get; }
    public int FavoritesReplaced { get; }
    public int FavoritesRemoved { get; }
    public int StylesAdded { get; }
    public int StylesReplaced { get; }
    public int StylesRemoved { get; }
    public override string ToString() =>
        $"Quick Keys: {QuickKeysBefore} → {QuickKeysAfter}\nFavorites: +{FavoritesAdded}, ~{FavoritesReplaced}, -{FavoritesRemoved}\nLocal styles: +{StylesAdded}, ~{StylesReplaced}, -{StylesRemoved}";
}

public sealed class ProfileImportPlan
{
    public ProfileImportPlan(string sourcePath, string sourceSha256, string currentProfileSha256,
        ProfileDefinition importedProfile, ProfilePackageManifest manifest, ProfileImportDiff diff, string planHash)
    {
        SourcePath = sourcePath; SourceSha256 = sourceSha256; CurrentProfileSha256 = currentProfileSha256;
        ImportedProfile = importedProfile; Manifest = manifest; Diff = diff; PlanHash = planHash;
    }
    public string SourcePath { get; }
    public string SourceSha256 { get; }
    public string CurrentProfileSha256 { get; }
    public ProfileDefinition ImportedProfile { get; }
    public ProfilePackageManifest Manifest { get; }
    public ProfileImportDiff Diff { get; }
    public string PlanHash { get; }
}

public sealed class ProfilePackageService
{
    public const int PackageSchemaVersion = 1;
    public const long MaximumPackageBytes = 2_097_152;
    private readonly ProfileStore _store;
    public ProfilePackageService(ProfileStore? store = null) => _store = store ?? new ProfileStore();

    public ProfileExportPlan PlanExport(ProfileDefinition profile, string destinationPath, bool overwrite)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (profile.Favorites.Any(value => value.Arguments.Count != 0))
            throw new InvalidDataException("Profile export refuses fixed favorite arguments until a typed non-workbook-content binder is qualified.");
        var destination = RequirePath(destinationPath);
        var existingHash = File.Exists(destination) ? HashFile(destination) : string.Empty;
        if (!overwrite && existingHash.Length != 0) throw new IOException("The export destination already exists; explicit overwrite is required.");
        var profileJson = _store.Serialize(profile);
        var payload = Encoding.UTF8.GetBytes(profileJson);
        var payloadHash = Hash(payload);
        var dto = new PackageDto
        {
            SchemaVersion = PackageSchemaVersion,
            Product = "ExcelAccel",
            ProfileSchemaVersion = profile.SchemaVersion,
            QuickKeyCount = profile.QuickKeys.Count,
            FavoriteCount = profile.Favorites.Count,
            LocalStyleCount = profile.LocalStyles.Count,
            AssetCount = 0,
            PayloadBytes = payload.Length,
            PayloadSha256 = payloadHash,
            ProfileJsonBase64 = Convert.ToBase64String(payload),
        };
        var packageBytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(dto, Formatting.Indented) + Environment.NewLine);
        if (packageBytes.Length > MaximumPackageBytes) throw new InvalidDataException($"The profile package exceeds {MaximumPackageBytes} bytes.");
        var manifest = Manifest(dto);
        var packageHash = Hash(packageBytes);
        var planHash = Hash(Encoding.UTF8.GetBytes($"export\n{destination}\n{overwrite}\n{existingHash}\n{packageHash}"));
        return new ProfileExportPlan(destination, overwrite, existingHash, packageBytes, packageHash, manifest, planHash);
    }

    public void Export(ProfileExportPlan plan, string confirmedPlanHash)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (!string.Equals(plan.PlanHash, confirmedPlanHash, StringComparison.Ordinal)) throw new InvalidOperationException("The exact export plan was not confirmed.");
        var existingHash = File.Exists(plan.DestinationPath) ? HashFile(plan.DestinationPath) : string.Empty;
        if (!string.Equals(existingHash, plan.ExistingDestinationSha256, StringComparison.Ordinal)) throw new IOException("The export destination changed after preview.");
        if (!plan.Overwrite && existingHash.Length != 0) throw new IOException("The export destination now exists.");
        if (!string.Equals(Hash(plan.PackageBytes), plan.PackageSha256, StringComparison.Ordinal)) throw new InvalidDataException("The planned package bytes changed.");
        _ = ParsePackage(plan.PackageBytes);

        var directory = Path.GetDirectoryName(plan.DestinationPath) ?? throw new ArgumentException("The destination requires a parent directory.");
        Directory.CreateDirectory(directory);
        var temporary = Path.Combine(directory, $".{Path.GetFileName(plan.DestinationPath)}.{Guid.NewGuid():N}.tmp");
        var backup = Path.Combine(directory, $".{Path.GetFileName(plan.DestinationPath)}.{Guid.NewGuid():N}.bak");
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(plan.PackageBytes, 0, plan.PackageBytes.Length); stream.Flush(); }
            _ = ParsePackage(File.ReadAllBytes(temporary));
            if (File.Exists(plan.DestinationPath)) File.Replace(temporary, plan.DestinationPath, backup, true);
            else File.Move(temporary, plan.DestinationPath);
            if (!string.Equals(Hash(File.ReadAllBytes(plan.DestinationPath)), plan.PackageSha256, StringComparison.Ordinal))
            {
                if (File.Exists(backup)) File.Replace(backup, plan.DestinationPath, null, true);
                else if (File.Exists(plan.DestinationPath)) File.Delete(plan.DestinationPath);
                throw new IOException("Export post-write verification failed; the prior destination was restored where available.");
            }
            if (File.Exists(backup)) File.Delete(backup);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            if (File.Exists(backup)) File.Delete(backup);
        }
    }

    public ProfileImportPlan PreviewImport(string sourcePath, ProfileDefinition currentProfile,
        IReadOnlyList<CommandDescriptor> registry)
    {
        if (currentProfile is null) throw new ArgumentNullException(nameof(currentProfile));
        if (registry is null) throw new ArgumentNullException(nameof(registry));
        var source = RequirePath(sourcePath);
        var file = new FileInfo(source);
        if (!file.Exists) throw new FileNotFoundException("The profile package does not exist.", source);
        if (file.Length > MaximumPackageBytes) throw new InvalidDataException($"The profile package exceeds {MaximumPackageBytes} bytes.");
        var bytes = File.ReadAllBytes(source);
        var parsed = ParsePackage(bytes);
        ValidateCommandReferences(parsed.Profile, registry);
        var currentHash = Hash(Encoding.UTF8.GetBytes(_store.Serialize(currentProfile)));
        var sourceHash = Hash(bytes);
        var diff = Diff(currentProfile, parsed.Profile);
        var planHash = Hash(Encoding.UTF8.GetBytes($"import\n{source}\n{sourceHash}\n{currentHash}\n{parsed.Manifest.PayloadSha256}"));
        return new ProfileImportPlan(source, sourceHash, currentHash, parsed.Profile, parsed.Manifest, diff, planHash);
    }

    public ProfileDefinition ApplyImport(ProfileImportPlan plan, string confirmedPlanHash, string profilePath,
        ProfileDefinition currentProfile, IReadOnlyList<CommandDescriptor> registry)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (!string.Equals(plan.PlanHash, confirmedPlanHash, StringComparison.Ordinal)) throw new InvalidOperationException("The exact import plan was not confirmed.");
        var currentHash = Hash(Encoding.UTF8.GetBytes(_store.Serialize(currentProfile)));
        if (!string.Equals(currentHash, plan.CurrentProfileSha256, StringComparison.Ordinal)) throw new InvalidOperationException("The active profile changed after import preview.");
        var bytes = File.ReadAllBytes(plan.SourcePath);
        if (!string.Equals(Hash(bytes), plan.SourceSha256, StringComparison.Ordinal)) throw new InvalidDataException("The profile package changed after preview.");
        var parsed = ParsePackage(bytes);
        ValidateCommandReferences(parsed.Profile, registry);
        var target = RequirePath(profilePath);
        WriteBoundedBackup(target + ".import.bak", _store.Serialize(currentProfile));
        _store.SaveAtomic(target, parsed.Profile);
        return parsed.Profile;
    }

    private ParsedPackage ParsePackage(byte[] bytes)
    {
        if (bytes is null || bytes.Length == 0 || bytes.Length > MaximumPackageBytes) throw new InvalidDataException("The profile package is empty or exceeds its size limit.");
        PackageDto dto;
        try
        {
            dto = JsonConvert.DeserializeObject<PackageDto>(Encoding.UTF8.GetString(bytes), new JsonSerializerSettings
            { MissingMemberHandling = MissingMemberHandling.Error, DateParseHandling = DateParseHandling.None, MaxDepth = 16 })
                ?? throw new InvalidDataException("The profile package did not contain an object.");
        }
        catch (JsonException exception) { throw new InvalidDataException("The profile package JSON is invalid or contains unsupported fields.", exception); }
        if (dto.SchemaVersion != PackageSchemaVersion || dto.Product != "ExcelAccel" || dto.AssetCount != 0)
            throw new InvalidDataException("The package schema/product/assets are unsupported.");
        byte[] payload;
        try { payload = Convert.FromBase64String(dto.ProfileJsonBase64); }
        catch (FormatException exception) { throw new InvalidDataException("The profile payload is not valid base64.", exception); }
        if (payload.Length != dto.PayloadBytes || !string.Equals(Hash(payload), dto.PayloadSha256, StringComparison.Ordinal))
            throw new InvalidDataException("The profile payload length or SHA-256 does not match the manifest.");
        var profile = _store.Parse(Encoding.UTF8.GetString(payload));
        if (profile.SchemaVersion != dto.ProfileSchemaVersion || profile.QuickKeys.Count != dto.QuickKeyCount ||
            profile.Favorites.Count != dto.FavoriteCount || profile.LocalStyles.Count != dto.LocalStyleCount)
            throw new InvalidDataException("The profile payload counts do not match the manifest.");
        return new ParsedPackage(profile, Manifest(dto));
    }

    private static void ValidateCommandReferences(ProfileDefinition profile, IReadOnlyList<CommandDescriptor> registry)
    {
        foreach (var commandId in profile.QuickKeys.Select(value => value.CommandId))
            if (!registry.Any(value => value.Id == commandId)) throw new InvalidDataException($"Quick Key references unknown command '{commandId}'.");
        foreach (var favorite in profile.Favorites)
        {
            var command = registry.FirstOrDefault(value => value.Id == favorite.CommandId);
            if (command is null) throw new InvalidDataException($"Favorite references unknown command '{favorite.CommandId}'.");
            if (command.ContractVersion != favorite.ContractVersion) throw new InvalidDataException($"Favorite '{favorite.FavoriteId}' has an incompatible command version.");
        }
    }

    private static ProfileImportDiff Diff(ProfileDefinition before, ProfileDefinition after)
    {
        var beforeFavorites = before.Favorites.ToDictionary(value => value.FavoriteId, StringComparer.Ordinal);
        var afterFavorites = after.Favorites.ToDictionary(value => value.FavoriteId, StringComparer.Ordinal);
        var beforeStyles = before.LocalStyles.ToDictionary(value => value.StyleId, StringComparer.Ordinal);
        var afterStyles = after.LocalStyles.ToDictionary(value => value.StyleId, StringComparer.Ordinal);
        return new ProfileImportDiff(before.QuickKeys.Count, after.QuickKeys.Count,
            afterFavorites.Keys.Count(value => !beforeFavorites.ContainsKey(value)),
            afterFavorites.Keys.Count(value => beforeFavorites.TryGetValue(value, out var prior) && !prior.SemanticallyEquals(afterFavorites[value])),
            beforeFavorites.Keys.Count(value => !afterFavorites.ContainsKey(value)),
            afterStyles.Keys.Count(value => !beforeStyles.ContainsKey(value)),
            afterStyles.Keys.Count(value => beforeStyles.TryGetValue(value, out var prior) && !StyleEquals(prior, afterStyles[value])),
            beforeStyles.Keys.Count(value => !afterStyles.ContainsKey(value)));
    }

    private static ProfilePackageManifest Manifest(PackageDto dto) => new ProfilePackageManifest(dto.ProfileSchemaVersion,
        dto.QuickKeyCount, dto.FavoriteCount, dto.LocalStyleCount, dto.AssetCount, dto.PayloadBytes, dto.PayloadSha256);
    private static string RequirePath(string path) => string.IsNullOrWhiteSpace(path) ? throw new ArgumentException("A local file path is required.", nameof(path)) : Path.GetFullPath(path);
    private static string Hash(byte[] bytes)
    {
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(bytes).Select(value => value.ToString("X2")));
    }
    private static string HashFile(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(stream).Select(value => value.ToString("X2")));
    }
    private static bool StyleEquals(Application.Styles.StyleRecipe first, Application.Styles.StyleRecipe second) =>
        first.Version == second.Version && string.Equals(first.DisplayName, second.DisplayName, StringComparison.Ordinal) &&
        first.UnsupportedPropertyPolicy == second.UnsupportedPropertyPolicy && first.Properties.Count == second.Properties.Count &&
        first.Properties.All(value => second.Properties.TryGetValue(value.Key, out var candidate) && string.Equals(value.Value, candidate, StringComparison.Ordinal));
    private static void WriteBoundedBackup(string path, string contents)
    {
        var directory = Path.GetDirectoryName(path) ?? throw new ArgumentException("The backup path requires a parent directory.");
        Directory.CreateDirectory(directory);
        var temporary = path + ".tmp";
        try
        {
            File.WriteAllText(temporary, contents, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temporary, path, null, true); else File.Move(temporary, path);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private sealed class ParsedPackage
    {
        public ParsedPackage(ProfileDefinition profile, ProfilePackageManifest manifest) { Profile = profile; Manifest = manifest; }
        public ProfileDefinition Profile { get; }
        public ProfilePackageManifest Manifest { get; }
    }

    private sealed class PackageDto
    {
        [JsonProperty("schema_version", Order = 1, Required = Required.Always)] public int SchemaVersion { get; set; }
        [JsonProperty("product", Order = 2, Required = Required.Always)] public string Product { get; set; } = string.Empty;
        [JsonProperty("profile_schema_version", Order = 3, Required = Required.Always)] public int ProfileSchemaVersion { get; set; }
        [JsonProperty("quick_key_count", Order = 4, Required = Required.Always)] public int QuickKeyCount { get; set; }
        [JsonProperty("favorite_count", Order = 5, Required = Required.Always)] public int FavoriteCount { get; set; }
        [JsonProperty("local_style_count", Order = 6, Required = Required.Always)] public int LocalStyleCount { get; set; }
        [JsonProperty("asset_count", Order = 7, Required = Required.Always)] public int AssetCount { get; set; }
        [JsonProperty("payload_bytes", Order = 8, Required = Required.Always)] public long PayloadBytes { get; set; }
        [JsonProperty("payload_sha256", Order = 9, Required = Required.Always)] public string PayloadSha256 { get; set; } = string.Empty;
        [JsonProperty("profile_json_base64", Order = 10, Required = Required.Always)] public string ProfileJsonBase64 { get; set; } = string.Empty;
    }
}
