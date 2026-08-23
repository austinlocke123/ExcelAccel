using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAccel.Core.Links;

namespace ExcelAccel.Persistence.Links;

public sealed class LinkExportManifest
{
    public LinkExportManifest(string destination, int sourceCount, int usageCount, bool includePaths, string coverage)
    {
        Destination = destination;
        SourceCount = sourceCount;
        UsageCount = usageCount;
        IncludePaths = includePaths;
        Coverage = coverage;
    }

    public string Destination { get; }
    public int SourceCount { get; }
    public int UsageCount { get; }

    /// <summary>
    /// Off by default. A link path routinely names a drive, a client, or a deal,
    /// so the full path leaves the machine only on an explicit choice.
    /// </summary>
    public bool IncludePaths { get; }

    public string Coverage { get; }

    public IReadOnlyList<string> IncludedFields => LinkInventorySearch.ExportFields(IncludePaths);

    public override string ToString() =>
        string.Format(
            CultureInfo.InvariantCulture,
            "{0}\n{1:N0} source(s), {2:N0} usage(s)\nFields: {3}\nSource paths: {4}\n{5}",
            Destination,
            SourceCount,
            UsageCount,
            string.Join(", ", IncludedFields),
            IncludePaths ? "included" : "redacted to file name",
            Coverage);
}

/// <summary>
/// Writes a link inventory to a local file. Nothing is transmitted, and the file
/// is written whole or not at all.
/// </summary>
public sealed class LinkInventoryExporter
{
    public const int MaximumUsages = 5_000;

    public LinkExportManifest Plan(string destination, LinkInventory inventory, bool includePaths)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        return new LinkExportManifest(
            RequirePath(destination),
            inventory.Sources.Count,
            inventory.UsageCount,
            includePaths,
            LinkInventorySearch.CoverageStatement(inventory));
    }

    public void Export(LinkExportManifest manifest, LinkInventory inventory)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        if (inventory.UsageCount > MaximumUsages)
        {
            throw new InvalidDataException($"The export is limited to {MaximumUsages:N0} usages.");
        }

        var resolved = RequirePath(manifest.Destination);
        var directory = Path.GetDirectoryName(resolved)
            ?? throw new ArgumentException("The export path requires a parent directory.", nameof(manifest));

        var builder = new StringBuilder();
        builder.Append("# ").Append(Quote(manifest.Coverage)).Append('\n');
        builder.Append(string.Join(",", manifest.IncludedFields)).Append('\n');
        foreach (var source in inventory.Sources)
        {
            if (source.Usages.Count == 0)
            {
                builder.Append(string.Join(",", new[]
                {
                    manifest.IncludePaths ? source.DisplaySource : source.Token,
                    source.Status.ToString(),
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    "No usage found in a scanned category.",
                    "false",
                }.Select(Quote))).Append('\n');
                continue;
            }

            foreach (var usage in source.Usages)
            {
                builder
                    .Append(string.Join(",", LinkInventorySearch.ExportRow(source, usage, manifest.IncludePaths).Select(Quote)))
                    .Append('\n');
            }
        }

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

    private static string Quote(string value) =>
        "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";

    private static string RequirePath(string destination)
    {
        if (string.IsNullOrWhiteSpace(destination))
        {
            throw new ArgumentException("An export destination is required.", nameof(destination));
        }

        return Path.GetFullPath(destination);
    }
}
