using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.Styles;
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
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception exception) when (exception is ArgumentException || exception is InvalidOperationException)
        {
            throw new InvalidDataException("The profile contains invalid or incompatible settings.", exception);
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

        [JsonProperty("cycles", Order = 3, Required = Required.Default)]
        public SortedDictionary<string, CycleDto[]>? Cycles { get; set; }

        // Schema v2-v5 shape. These members are read so existing profiles keep
        // loading, and are never written: MissingMemberHandling.Error would refuse
        // an old profile the moment they were removed. NullValueHandling.Ignore is
        // load-bearing rather than tidiness -- without it every saved v6 profile
        // would carry "font_color_cycle": null and fail the unknown-field check on
        // the next read.
        [JsonProperty("font_color_cycle", Order = 101, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public string[]? FontColorCycle { get; set; }

        [JsonProperty("fill_color_cycle", Order = 102, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public string[]? FillColorCycle { get; set; }

        [JsonProperty("font_size_cycle", Order = 103, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public double[]? FontSizeCycle { get; set; }

        [JsonProperty("number_formats", Order = 104, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public SortedDictionary<string, string>? NumberFormats { get; set; }

        [JsonProperty("horizontal_alignment_cycle", Order = 105, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public string[]? HorizontalAlignmentCycle { get; set; }

        [JsonProperty("vertical_alignment_cycle", Order = 106, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public string[]? VerticalAlignmentCycle { get; set; }

        [JsonProperty("underline_cycle", Order = 107, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public string[]? UnderlineCycle { get; set; }

        [JsonProperty("row_height_cycle", Order = 108, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public double[]? RowHeightCycle { get; set; }

        [JsonProperty("column_width_cycle", Order = 109, Required = Required.Default, NullValueHandling = NullValueHandling.Ignore)]
        public double[]? ColumnWidthCycle { get; set; }

        [JsonProperty("auto_color_colors", Order = 12, Required = Required.Always)]
        public SortedDictionary<string, string> AutoColorColors { get; set; } = new SortedDictionary<string, string>(StringComparer.Ordinal);

        [JsonProperty("quick_keys", Order = 13, Required = Required.Always)]
        public QuickKeyDto[] QuickKeys { get; set; } = Array.Empty<QuickKeyDto>();

        [JsonProperty("favorites", Order = 14, Required = Required.Default)]
        public FavoriteDto[]? Favorites { get; set; }

        [JsonProperty("local_styles", Order = 15, Required = Required.Default)]
        public StyleDto[]? LocalStyles { get; set; }

        [JsonProperty("immediate_preview_cell_limit", Order = 16, Required = Required.Always)]
        public long ImmediatePreviewCellLimit { get; set; }

        [JsonProperty("wrap_sheet_navigation", Order = 17, Required = Required.Always)]
        public bool WrapSheetNavigation { get; set; }

        [JsonProperty("formula_iferror_fallback", Order = 18, Required = Required.Default)]
        public string? FormulaIfErrorFallback { get; set; }

        public ProfileDefinition ToDefinition()
        {
            if (SchemaVersion < 2 || SchemaVersion > ProfileDefinition.CurrentSchemaVersion)
                throw new InvalidDataException($"Profile schema {SchemaVersion} is not supported.");
            if (SchemaVersion >= 3 && Favorites is null) throw new InvalidDataException("Schema v3+ profiles require a favorites array.");
            if (SchemaVersion >= 4 && LocalStyles is null) throw new InvalidDataException("Schema v4 profiles require a local_styles array.");
            if (SchemaVersion >= 5 && FormulaIfErrorFallback is null) throw new InvalidDataException("Schema v5 profiles require formula_iferror_fallback.");
            if (SchemaVersion >= 6 && Cycles is null) throw new InvalidDataException("Schema v6 profiles require a cycles object.");
            if (SchemaVersion < 6 && Cycles is not null) throw new InvalidDataException("Schema v2-v5 profiles cannot carry a cycles object.");
            return new ProfileDefinition(
            ProfileDefinition.CurrentSchemaVersion,
            ProfileId,
            Cycles is null ? LiftLegacyCycles() : ReadCycles(Cycles),
            AutoColorColors,
            QuickKeys.Select(value => new QuickKeyBinding(value.CommandId, value.Sequence)),
            (Favorites ?? Array.Empty<FavoriteDto>()).Select(value => value.ToDefinition()),
            (LocalStyles ?? Array.Empty<StyleDto>()).Select(value => value.ToDefinition()),
            ImmediatePreviewCellLimit,
            WrapSheetNavigation,
            FormulaIfErrorFallback ?? "0");
        }

        private const string LegacyCycleId = "standard";

        /// <summary>
        /// Schema v2-v5 held exactly one cycle per property family, plus one number
        /// format per family name. Each property cycle keeps its whole ordered list
        /// under a single "standard" cycle, and each number format becomes a
        /// one-entry cycle keyed by the name it already had, so no setting is lost
        /// and every existing format.number.* command keeps resolving (AC-FMT-031).
        /// </summary>
        private ProfileCycles LiftLegacyCycles()
        {
            var formats = NumberFormats
                ?? throw new InvalidDataException("Schema v2-v5 profiles require number_formats.");
            if (formats.Count > ProfileCycles.MaximumCyclesPerFamily)
                throw new InvalidDataException(
                    $"This profile defines {formats.Count} number formats, more than the {ProfileCycles.MaximumCyclesPerFamily} cycles one family may hold. Remove some and retry.");

            return new ProfileCycles(new[]
            {
                Family("column_width", ColumnWidthCycle),
                Family("fill_color", FillColorCycle),
                Family("font_color", FontColorCycle),
                Family("font_size", FontSizeCycle),
                Family("horizontal_alignment", HorizontalAlignmentCycle),
                new KeyValuePair<string, IEnumerable<ProfileCycle>>("number_format", formats
                    .Select(entry => new ProfileCycle(
                        "number_format", entry.Key, ProfileCycle.TitleFrom(entry.Key), new[] { entry.Value }))
                    .ToArray()),
                Family("row_height", RowHeightCycle),
                Family("underline", UnderlineCycle),
                Family("vertical_alignment", VerticalAlignmentCycle),
            });
        }

        private static KeyValuePair<string, IEnumerable<ProfileCycle>> Family(string family, string[]? entries) =>
            new KeyValuePair<string, IEnumerable<ProfileCycle>>(family, new[]
            {
                new ProfileCycle(family, LegacyCycleId, ProfileCycle.TitleFrom(family),
                    entries ?? throw new InvalidDataException($"Schema v2-v5 profiles require {family}_cycle.")),
            });

        private static KeyValuePair<string, IEnumerable<ProfileCycle>> Family(string family, double[]? values) =>
            Family(family, values?
                .Select(value => value.ToString("0.####", CultureInfo.InvariantCulture))
                .ToArray());

        private static ProfileCycles ReadCycles(SortedDictionary<string, CycleDto[]> source) =>
            new ProfileCycles(source.Select(family =>
                new KeyValuePair<string, IEnumerable<ProfileCycle>>(
                    family.Key,
                    (family.Value ?? Array.Empty<CycleDto>())
                        .Select(value => new ProfileCycle(family.Key, value.CycleId, value.DisplayName, value.Entries))
                        .ToArray())));

        public static ProfileDto From(ProfileDefinition profile) => new ProfileDto
        {
            SchemaVersion = profile.SchemaVersion,
            ProfileId = profile.ProfileId,
            Cycles = CopyCycles(profile.Cycles),
            AutoColorColors = CopyFormats(profile.AutoColorColors),
            QuickKeys = profile.QuickKeys
                .OrderBy(value => value.CommandId, StringComparer.Ordinal)
                .Select(value => new QuickKeyDto { CommandId = value.CommandId, Sequence = value.Sequence })
                .ToArray(),
            Favorites = profile.Favorites
                .OrderBy(value => value.FavoriteId, StringComparer.Ordinal)
                .Select(FavoriteDto.From)
                .ToArray(),
            LocalStyles = profile.LocalStyles
                .OrderBy(value => value.StyleId, StringComparer.Ordinal)
                .Select(StyleDto.From)
                .ToArray(),
            ImmediatePreviewCellLimit = profile.ImmediatePreviewCellLimit,
            WrapSheetNavigation = profile.WrapSheetNavigation,
            FormulaIfErrorFallback = profile.FormulaIfErrorFallback,
        };

        private static SortedDictionary<string, CycleDto[]> CopyCycles(ProfileCycles cycles)
        {
            var result = new SortedDictionary<string, CycleDto[]>(StringComparer.Ordinal);
            foreach (var family in cycles.Families)
            {
                result.Add(family, cycles[family].Select(CycleDto.From).ToArray());
            }

            return result;
        }

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

    private sealed class CycleDto
    {
        [JsonProperty("cycle_id", Order = 1, Required = Required.Always)]
        public string CycleId { get; set; } = string.Empty;

        [JsonProperty("display_name", Order = 2, Required = Required.Always)]
        public string DisplayName { get; set; } = string.Empty;

        [JsonProperty("entries", Order = 3, Required = Required.Always)]
        public string[] Entries { get; set; } = Array.Empty<string>();

        public static CycleDto From(ProfileCycle cycle) => new CycleDto
        {
            CycleId = cycle.CycleId,
            DisplayName = cycle.DisplayName,
            Entries = cycle.Entries.ToArray(),
        };
    }

    private sealed class QuickKeyDto
    {
        [JsonProperty("command_id", Order = 1, Required = Required.Always)]
        public string CommandId { get; set; } = string.Empty;

        [JsonProperty("sequence", Order = 2, Required = Required.Always)]
        public string Sequence { get; set; } = string.Empty;
    }

    private sealed class FavoriteDto
    {
        [JsonProperty("favorite_id", Order = 1, Required = Required.Always)]
        public string FavoriteId { get; set; } = string.Empty;

        [JsonProperty("command_id", Order = 2, Required = Required.Always)]
        public string CommandId { get; set; } = string.Empty;

        [JsonProperty("contract_version", Order = 3, Required = Required.Always)]
        public int ContractVersion { get; set; }

        [JsonProperty("arguments", Order = 4, Required = Required.Always)]
        public SortedDictionary<string, string> Arguments { get; set; } = new SortedDictionary<string, string>(StringComparer.Ordinal);

        public FavoriteDefinition ToDefinition() => new FavoriteDefinition(FavoriteId, CommandId, ContractVersion, Arguments);

        public static FavoriteDto From(FavoriteDefinition favorite) => new FavoriteDto
        {
            FavoriteId = favorite.FavoriteId,
            CommandId = favorite.CommandId,
            ContractVersion = favorite.ContractVersion,
            Arguments = CopyArguments(favorite.Arguments),
        };

        private static SortedDictionary<string, string> CopyArguments(IReadOnlyDictionary<string, string> source)
        {
            var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
            foreach (var item in source) result.Add(item.Key, item.Value);
            return result;
        }
    }

    private sealed class StyleDto
    {
        [JsonProperty("style_id", Order = 1, Required = Required.Always)]
        public string StyleId { get; set; } = string.Empty;
        [JsonProperty("version", Order = 2, Required = Required.Always)]
        public int Version { get; set; }
        [JsonProperty("display_name", Order = 3, Required = Required.Always)]
        public string DisplayName { get; set; } = string.Empty;
        [JsonProperty("unsupported_property_policy", Order = 4, Required = Required.Always)]
        public string UnsupportedPropertyPolicy { get; set; } = string.Empty;
        [JsonProperty("properties", Order = 5, Required = Required.Always)]
        public SortedDictionary<string, string> Properties { get; set; } = new SortedDictionary<string, string>(StringComparer.Ordinal);

        public StyleRecipe ToDefinition()
        {
            if (!Enum.TryParse(UnsupportedPropertyPolicy, true, out UnsupportedStylePropertyPolicy policy))
                throw new InvalidDataException($"Unsupported style property policy '{UnsupportedPropertyPolicy}'.");
            return new StyleRecipe(StyleId, Version, DisplayName, StyleOrigin.Local, policy, Properties);
        }

        public static StyleDto From(StyleRecipe recipe) => new StyleDto
        {
            StyleId = recipe.StyleId,
            Version = recipe.Version,
            DisplayName = recipe.DisplayName,
            UnsupportedPropertyPolicy = recipe.UnsupportedPropertyPolicy.ToString().ToLowerInvariant(),
            Properties = CopyArguments(recipe.Properties),
        };

        private static SortedDictionary<string, string> CopyArguments(IReadOnlyDictionary<string, string> source)
        {
            var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
            foreach (var item in source) result.Add(item.Key, item.Value);
            return result;
        }
    }
}
