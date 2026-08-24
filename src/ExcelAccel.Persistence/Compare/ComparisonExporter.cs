using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAccel.Core.Compare;

namespace ExcelAccel.Persistence.Compare;

public sealed class ComparisonExportManifest
{
    public ComparisonExportManifest(
        string destination,
        string sourceIdentity,
        string targetIdentity,
        int differenceCount,
        int comparedCells,
        bool includeContent,
        string coverage)
    {
        Destination = destination;
        SourceIdentity = sourceIdentity;
        TargetIdentity = targetIdentity;
        DifferenceCount = differenceCount;
        ComparedCells = comparedCells;
        IncludeContent = includeContent;
        Coverage = coverage;
    }

    public string Destination { get; }
    public string SourceIdentity { get; }
    public string TargetIdentity { get; }
    public int DifferenceCount { get; }
    public int ComparedCells { get; }

    /// <summary>
    /// Off by default. The differing cells are the ones most likely to carry the
    /// numbers and formulas a model is about.
    /// </summary>
    public bool IncludeContent { get; }

    public string Coverage { get; }

    public IReadOnlyList<string> IncludedFields => ComparisonPresentation.ExportFields(IncludeContent);

    public override string ToString() =>
        string.Format(
            CultureInfo.InvariantCulture,
            "{0}\n\nSource: {1}\nTarget: {2}\n{3:N0} difference(s) across {4:N0} compared cell(s)\nFields: {5}\nCell contents: {6}\n{7}",
            Destination,
            SourceIdentity,
            TargetIdentity,
            DifferenceCount,
            ComparedCells,
            string.Join(", ", IncludedFields),
            IncludeContent ? "included" : "excluded",
            Coverage);
}

/// <summary>
/// Writes comparison differences to a local file, whole or not at all. Nothing is
/// transmitted and neither compared workbook is touched.
/// </summary>
public sealed class ComparisonExporter
{
    public ComparisonExportManifest Plan(string destination, ComparisonResult result, bool includeContent)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        return new ComparisonExportManifest(
            RequirePath(destination),
            result.Source.Identity,
            result.Target.Identity,
            result.Differences.Count,
            result.ComparedCells,
            includeContent,
            ComparisonPresentation.CoverageStatement(result));
    }

    public void Export(ComparisonExportManifest manifest, ComparisonResult result)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        if (result is null) throw new ArgumentNullException(nameof(result));

        var resolved = RequirePath(manifest.Destination);
        var directory = Path.GetDirectoryName(resolved)
            ?? throw new ArgumentException("The export path requires a parent directory.", nameof(manifest));

        var builder = new StringBuilder();
        builder.Append("# ").Append(Quote(manifest.SourceIdentity + " vs " + manifest.TargetIdentity)).Append('\n');
        builder.Append("# ").Append(Quote(manifest.Coverage)).Append('\n');
        builder.Append(string.Join(",", manifest.IncludedFields)).Append('\n');
        foreach (var difference in result.Differences)
        {
            builder
                .Append(string.Join(",", ComparisonPresentation.ExportRow(difference, manifest.IncludeContent).Select(Quote)))
                .Append('\n');
        }

        // Through a temporary file in the same directory, so a failure leaves any
        // existing export exactly as it was rather than half rewritten.
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
