using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAccel.Persistence.ModelCheck;

/// <summary>
/// One suppressed finding. It stores the rule identity and the normalized
/// fingerprint only — never a formula, a value, or any other workbook content.
/// </summary>
public sealed class ModelCheckIgnoreEntry
{
    public ModelCheckIgnoreEntry(string ruleId, int ruleVersion, string fingerprint, string scopeNote = "")
    {
        RuleId = !string.IsNullOrWhiteSpace(ruleId) ? ruleId : throw new ArgumentException("A rule ID is required.", nameof(ruleId));
        RuleVersion = ruleVersion >= 1 ? ruleVersion : throw new ArgumentOutOfRangeException(nameof(ruleVersion));
        Fingerprint = IsFingerprint(fingerprint)
            ? fingerprint
            : throw new ArgumentException("A fingerprint must be a 64-character hex digest.", nameof(fingerprint));
        ScopeNote = scopeNote ?? string.Empty;
    }

    public string RuleId { get; }

    public int RuleVersion { get; }

    public string Fingerprint { get; }

    /// <summary>A short human note about what was ignored. Never workbook content.</summary>
    public string ScopeNote { get; }

    public static bool IsFingerprint(string? value) =>
        value is not null && value.Length == 64 && value.All(character =>
            (character >= '0' && character <= '9') || (character >= 'a' && character <= 'f'));
}

/// <summary>
/// The local ignore set, held in its own atomic file beside the profile.
///
/// It is deliberately separate from the profile document: adding it there would
/// bump the shipped profile schema and ripple through profile export and import.
/// The atomicity, locality, and explicit-export properties the contract requires
/// are provided here, and the ignore set is portable only through the same
/// deliberate export or import action.
/// </summary>
public sealed class ModelCheckIgnoreStore
{
    public const int MaximumEntries = 2_048;
    public const int SchemaVersion = 1;

    private static readonly object Sync = new object();

    public IReadOnlyList<ModelCheckIgnoreEntry> Load(string path)
    {
        var resolved = RequirePath(path);
        if (!File.Exists(resolved)) return Array.Empty<ModelCheckIgnoreEntry>();
        var entries = new List<ModelCheckIgnoreEntry>();
        foreach (var line in File.ReadAllLines(resolved, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line) || line[0] == '#') continue;
            var parts = line.Split('\t');
            if (parts.Length < 3) continue;
            if (!int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var version)) continue;
            if (!ModelCheckIgnoreEntry.IsFingerprint(parts[2])) continue;
            entries.Add(new ModelCheckIgnoreEntry(parts[0], version, parts[2], parts.Length > 3 ? parts[3] : string.Empty));
            if (entries.Count >= MaximumEntries) break;
        }

        return entries;
    }

    /// <summary>
    /// Writes the ignore set through temp, validate, then replace, so a failure
    /// leaves the existing file intact.
    /// </summary>
    public void SaveAtomic(string path, IEnumerable<ModelCheckIgnoreEntry> entries)
    {
        if (entries is null) throw new ArgumentNullException(nameof(entries));
        var resolved = RequirePath(path);
        var ordered = entries
            .GroupBy(entry => entry.Fingerprint, StringComparer.Ordinal)
            .Select(group => group.First())
            .OrderBy(entry => entry.RuleId, StringComparer.Ordinal)
            .ThenBy(entry => entry.Fingerprint, StringComparer.Ordinal)
            .ToArray();
        if (ordered.Length > MaximumEntries)
        {
            throw new InvalidDataException($"The ignore set is limited to {MaximumEntries:N0} entries.");
        }

        var directory = Path.GetDirectoryName(resolved)
            ?? throw new ArgumentException("The ignore-set path requires a parent directory.", nameof(path));
        var builder = new StringBuilder();
        builder.Append("# ExcelAccel Model Check local ignores, schema ")
            .Append(SchemaVersion.ToString(CultureInfo.InvariantCulture))
            .Append('\n');
        builder.Append("# rule_id\trule_version\tfingerprint\tscope_note\n");
        foreach (var entry in ordered)
        {
            builder.Append(entry.RuleId).Append('\t')
                .Append(entry.RuleVersion.ToString(CultureInfo.InvariantCulture)).Append('\t')
                .Append(entry.Fingerprint).Append('\t')
                .Append(entry.ScopeNote.Replace('\t', ' ').Replace('\n', ' ')).Append('\n');
        }

        lock (Sync)
        {
            Directory.CreateDirectory(directory);
            var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(resolved)}.{Guid.NewGuid():N}.tmp");
            try
            {
                File.WriteAllText(temporaryPath, builder.ToString(), new UTF8Encoding(false));

                // Validate what was written before it replaces anything.
                var reloaded = Load(temporaryPath);
                if (reloaded.Count != ordered.Length)
                {
                    throw new InvalidDataException("The ignore set did not round-trip; the existing file was left unchanged.");
                }

                if (File.Exists(resolved)) File.Replace(temporaryPath, resolved, null);
                else File.Move(temporaryPath, resolved);
            }
            finally
            {
                TryDelete(temporaryPath);
            }
        }
    }

    private static string RequirePath(string path) =>
        !string.IsNullOrWhiteSpace(path) ? path : throw new ArgumentException("A path is required.", nameof(path));

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
