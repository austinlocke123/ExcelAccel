using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelAccel.Application.Profiles;
using Newtonsoft.Json;

namespace ExcelAccel.Persistence.Profiles;

public sealed class ProfileStore
{
    public const long MaximumProfileBytes = 1_048_576;
    private const string DefaultResourceSuffix = "Defaults.default-profile.json";
    private readonly object _sync = new object();

    public ProfileDefinition Load(string path)
    {
        var resolvedPath = RequirePath(path);
        var file = new FileInfo(resolvedPath);
        if (!file.Exists)
        {
            throw new FileNotFoundException("The profile file does not exist.", resolvedPath);
        }

        if (file.Length > MaximumProfileBytes)
        {
            throw new InvalidDataException($"The profile exceeds the {MaximumProfileBytes}-byte limit.");
        }

        return Parse(File.ReadAllText(resolvedPath, Encoding.UTF8));
    }

    public ProfileDefinition LoadOrDefault(string path) =>
        File.Exists(RequirePath(path)) ? Load(path) : LoadDefault();

    public ProfileDefinition LoadDefault()
    {
        var assembly = typeof(ProfileStore).GetTypeInfo().Assembly;
        var resourceName = assembly.GetManifestResourceNames()
            .SingleOrDefault(name => name.EndsWith(DefaultResourceSuffix, StringComparison.Ordinal));
        if (resourceName is null)
        {
            throw new InvalidOperationException("The embedded default profile is missing.");
        }

        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException("The embedded default profile could not be opened.");
        using var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, leaveOpen: false);
        return Parse(reader.ReadToEnd());
    }

    public void SaveAtomic(string path, ProfileDefinition profile)
    {
        if (profile is null)
        {
            throw new ArgumentNullException(nameof(profile));
        }

        var resolvedPath = RequirePath(path);
        var directory = Path.GetDirectoryName(resolvedPath)
            ?? throw new ArgumentException("The profile path requires a parent directory.", nameof(path));
        var json = Serialize(profile);
        if (Encoding.UTF8.GetByteCount(json) > MaximumProfileBytes)
        {
            throw new InvalidDataException($"The serialized profile exceeds the {MaximumProfileBytes}-byte limit.");
        }

        lock (_sync)
        {
            Directory.CreateDirectory(directory);
            var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(resolvedPath)}.{Guid.NewGuid():N}.tmp");
            var backupPath = resolvedPath + ".bak";
            try
            {
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(json);
                    writer.Flush();
                    stream.Flush();
                }

                _ = Load(temporaryPath);
                if (File.Exists(resolvedPath))
                {
                    File.Replace(temporaryPath, resolvedPath, backupPath, ignoreMetadataErrors: true);
                }
                else
                {
                    File.Move(temporaryPath, resolvedPath);
                }
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }
    }

    public string Serialize(ProfileDefinition profile)
    {
        if (profile is null)
        {
            throw new ArgumentNullException(nameof(profile));
        }

        return JsonConvert.SerializeObject(ProfileDto.From(profile), Formatting.Indented) + Environment.NewLine;
    }

    public ProfileDefinition Parse(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            throw new InvalidDataException("The profile JSON is empty.");
        }

        if (Encoding.UTF8.GetByteCount(json) > MaximumProfileBytes)
        {
            throw new InvalidDataException($"The profile exceeds the {MaximumProfileBytes}-byte limit.");
        }

        try
        {
            var dto = JsonConvert.DeserializeObject<ProfileDto>(json, new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Error,
                DateParseHandling = DateParseHandling.None,
                MaxDepth = 32,
            }) ?? throw new InvalidDataException("The profile JSON did not contain an object.");
            return dto.ToDefinition();
        }
        catch (JsonException exception)
        {
            throw new InvalidDataException("The profile JSON is invalid or contains unsupported fields.", exception);
        }
    }

    private static string RequirePath(string path) =>
        string.IsNullOrWhiteSpace(path)
            ? throw new ArgumentException("A profile path is required.", nameof(path))
            : Path.GetFullPath(path);

    private sealed class ProfileDto
    {
        [JsonProperty("schema_version", Order = 1, Required = Required.Always)]
        public int SchemaVersion { get; set; }

        [JsonProperty("profile_id", Order = 2, Required = Required.Always)]
        public string ProfileId { get; set; } = string.Empty;

        [JsonProperty("font_color_cycle", Order = 3, Required = Required.Always)]
        public string[] FontColorCycle { get; set; } = Array.Empty<string>();

        [JsonProperty("fill_color_cycle", Order = 4, Required = Required.Always)]
        public string[] FillColorCycle { get; set; } = Array.Empty<string>();

        [JsonProperty("font_size_cycle", Order = 5, Required = Required.Always)]
        public double[] FontSizeCycle { get; set; } = Array.Empty<double>();

        [JsonProperty("number_formats", Order = 6, Required = Required.Always)]
        public SortedDictionary<string, string> NumberFormats { get; set; } = new SortedDictionary<string, string>(StringComparer.Ordinal);

        [JsonProperty("quick_keys", Order = 7, Required = Required.Always)]
        public QuickKeyDto[] QuickKeys { get; set; } = Array.Empty<QuickKeyDto>();

        [JsonProperty("immediate_preview_cell_limit", Order = 8, Required = Required.Always)]
        public long ImmediatePreviewCellLimit { get; set; }

        [JsonProperty("wrap_sheet_navigation", Order = 9, Required = Required.Always)]
        public bool WrapSheetNavigation { get; set; }

        public ProfileDefinition ToDefinition() => new ProfileDefinition(
            SchemaVersion,
            ProfileId,
            FontColorCycle,
            FillColorCycle,
            FontSizeCycle,
            NumberFormats,
            QuickKeys.Select(value => new QuickKeyBinding(value.CommandId, value.Sequence)),
            ImmediatePreviewCellLimit,
            WrapSheetNavigation);

        public static ProfileDto From(ProfileDefinition profile) => new ProfileDto
        {
            SchemaVersion = profile.SchemaVersion,
            ProfileId = profile.ProfileId,
            FontColorCycle = profile.FontColorCycle.ToArray(),
            FillColorCycle = profile.FillColorCycle.ToArray(),
            FontSizeCycle = profile.FontSizeCycle.ToArray(),
            NumberFormats = CopyFormats(profile.NumberFormats),
            QuickKeys = profile.QuickKeys
                .OrderBy(value => value.CommandId, StringComparer.Ordinal)
                .Select(value => new QuickKeyDto { CommandId = value.CommandId, Sequence = value.Sequence })
                .ToArray(),
            ImmediatePreviewCellLimit = profile.ImmediatePreviewCellLimit,
            WrapSheetNavigation = profile.WrapSheetNavigation,
        };

        private static SortedDictionary<string, string> CopyFormats(IReadOnlyDictionary<string, string> source)
        {
            var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
            foreach (var item in source)
            {
                result.Add(item.Key, item.Value);
            }

            return result;
        }
    }

    private sealed class QuickKeyDto
    {
        [JsonProperty("command_id", Order = 1, Required = Required.Always)]
        public string CommandId { get; set; } = string.Empty;

        [JsonProperty("sequence", Order = 2, Required = Required.Always)]
        public string Sequence { get; set; } = string.Empty;
    }
}
