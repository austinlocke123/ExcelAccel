using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAccel.Core.Names;

namespace ExcelAccel.Persistence.Names;

/// <summary>What the export will contain, shown before anything is written.</summary>
public sealed class NameExportManifest
{
    public NameExportManifest(string destination, int nameCount, bool includeExpressions, string coverage)
    {
        Destination = destination;
        NameCount = nameCount;
        IncludeExpressions = includeExpressions;
        Coverage = coverage;
    }

    public string Destination { get; }
    public int NameCount { get; }

    /// <summary>
    /// Off by default. A name's target can carry a file path or a business term,
    /// so the raw expression leaves the machine only on an explicit choice.
    /// </summary>
    public bool IncludeExpressions { get; }

    public string Coverage { get; }

    public IReadOnlyList<string> IncludedFields => NameInventorySearch.ExportFields(IncludeExpressions);

    public override string ToString() =>
        string.Format(
            CultureInfo.InvariantCulture,
            "{0}\n{1:N0} name(s)\nFields: {2}\nExpressions: {3}\n{4}",
            Destination,
            NameCount,
            string.Join(", ", IncludedFields),
            IncludeExpressions ? "included" : "excluded",
            Coverage);
}

/// <summary>
/// Writes a name inventory to a local file. Nothing is transmitted, and the file
/// is written whole or not at all.
/// </summary>
public sealed class NameInventoryExporter
{
    public const int MaximumNames = 5_000;

    public NameExportManifest Plan(string destination, NameInventory inventory, bool includeExpressions)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        return new NameExportManifest(
            RequirePath(destination),
            inventory.Entries.Count,
            includeExpressions,
            NameInventorySearch.CoverageStatement(inventory));
    }

    public void Export(NameExportManifest manifest, NameInventory inventory)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        if (inventory.Entries.Count > MaximumNames)
        {
            throw new InvalidDataException($"The export is limited to {MaximumNames:N0} names.");
        }

        var resolved = RequirePath(manifest.Destination);
        var directory = Path.GetDirectoryName(resolved)
            ?? throw new ArgumentException("The export path requires a parent directory.", nameof(manifest));

        var builder = new StringBuilder();
        builder.Append("# ").Append(Quote(manifest.Coverage)).Append('\n');
        builder.Append(string.Join(",", manifest.IncludedFields)).Append('\n');
        foreach (var entry in inventory.Entries)
        {
            var fields = NameInventorySearch.ExportRow(entry, manifest.IncludeExpressions);
            builder.Append(string.Join(",", fields.Select(Quote))).Append('\n');
        }

        // Written through a temporary file in the same directory so a failure
        // cannot leave a half-written export that looks complete.
        var temporary = Path.Combine(directory, "." + Path.GetFileName(resolved) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            File.WriteAllText(temporary, builder.ToString(), new UTF8Encoding(false));
            if (File.Exists(resolved))
            {
                File.Replace(temporary, resolved, null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(temporary, resolved);
            }
        }
        finally
        {
            if (File.Exists(temporary))
            {
                try { File.Delete(temporary); } catch (IOException) { }
            }
        }
    }

    private static string Quote(string value)
    {
        var text = value ?? string.Empty;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    private static string RequirePath(string destination)
    {
        if (string.IsNullOrWhiteSpace(destination))
        {
            throw new ArgumentException("An export destination is required.", nameof(destination));
        }

        return Path.GetFullPath(destination);
    }
}
