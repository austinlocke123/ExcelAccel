using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAccel.Persistence.ModelCheck;

/// <summary>One finding as the exporter sees it, already redacted by the caller.</summary>
public sealed class ExportableFinding
{
    public ExportableFinding(
        string ruleId,
        int ruleVersion,
        string severity,
        string worksheetName,
        string address,
        string statement,
        string coverage,
        string fingerprint,
        IReadOnlyList<string> evidence)
    {
        RuleId = ruleId;
        RuleVersion = ruleVersion;
        Severity = severity;
        WorksheetName = worksheetName;
        Address = address;
        Statement = statement;
        Coverage = coverage;
        Fingerprint = fingerprint;
        Evidence = evidence ?? Array.Empty<string>();
    }

    public string RuleId { get; }

    public int RuleVersion { get; }

    public string Severity { get; }

    public string WorksheetName { get; }

    public string Address { get; }

    public string Statement { get; }

    public string Coverage { get; }

    public string Fingerprint { get; }

    /// <summary>Rule evidence. Excluded unless the user opts in.</summary>
    public IReadOnlyList<string> Evidence { get; }
}

/// <summary>
/// The manifest a user confirms before anything is written. It states exactly
/// which workbook-derived fields the file will contain.
/// </summary>
public sealed class ExportManifest
{
    public ExportManifest(string destination, int findingCount, bool includeEvidence)
    {
        Destination = destination;
        FindingCount = findingCount;
        IncludeEvidence = includeEvidence;
    }

    public string Destination { get; }

    public int FindingCount { get; }

    /// <summary>Off by default. Evidence can quote formula structure.</summary>
    public bool IncludeEvidence { get; }

    public IReadOnlyList<string> IncludedFields => IncludeEvidence
        ? new[] { "rule_id", "rule_version", "severity", "worksheet", "address", "statement", "coverage", "fingerprint", "evidence" }
        : new[] { "rule_id", "rule_version", "severity", "worksheet", "address", "statement", "coverage", "fingerprint" };

    public IReadOnlyList<string> ExcludedFields => IncludeEvidence
        ? new[] { "cell formulas", "cell values" }
        : new[] { "cell formulas", "cell values", "rule evidence" };

    public IReadOnlyList<string> Lines => new[]
    {
        "Destination: " + Destination,
        "Findings: " + FindingCount.ToString("N0", CultureInfo.InvariantCulture),
        "Included fields: " + string.Join(", ", IncludedFields),
        "Excluded: " + string.Join(", ", ExcludedFields),
        "Nothing is transmitted; the file is written locally.",
    };
}

/// <summary>
/// Writes findings to a local file through temp, validate, then replace. It
/// never transmits, and a failure leaves any existing destination untouched.
/// </summary>
public sealed class ModelCheckResultExporter
{
    public const int MaximumFindings = 5_000;

    private static readonly object Sync = new object();

    public ExportManifest Plan(string destination, IReadOnlyList<ExportableFinding> findings, bool includeEvidence)
    {
        if (findings is null) throw new ArgumentNullException(nameof(findings));
        return new ExportManifest(RequirePath(destination), findings.Count, includeEvidence);
    }

    public void Export(ExportManifest manifest, IReadOnlyList<ExportableFinding> findings)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        if (findings is null) throw new ArgumentNullException(nameof(findings));
        if (findings.Count > MaximumFindings)
        {
            throw new InvalidDataException($"The export is limited to {MaximumFindings:N0} findings.");
        }

        var resolved = RequirePath(manifest.Destination);
        var directory = Path.GetDirectoryName(resolved)
            ?? throw new ArgumentException("The export path requires a parent directory.", nameof(manifest));

        var builder = new StringBuilder();
        builder.Append(string.Join(",", manifest.IncludedFields)).Append('\n');
        foreach (var finding in findings
            .OrderBy(value => value.RuleId, StringComparer.Ordinal)
            .ThenBy(value => value.WorksheetName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(value => value.Address, StringComparer.OrdinalIgnoreCase))
        {
            var fields = new List<string>
            {
                finding.RuleId,
                finding.RuleVersion.ToString(CultureInfo.InvariantCulture),
                finding.Severity,
                finding.WorksheetName,
                finding.Address,
                finding.Statement,
                finding.Coverage,
                finding.Fingerprint,
            };
            if (manifest.IncludeEvidence) fields.Add(string.Join(" | ", finding.Evidence));
            builder.Append(string.Join(",", fields.Select(Quote))).Append('\n');
        }

        lock (Sync)
        {
            Directory.CreateDirectory(directory);
            var temporaryPath = Path.Combine(directory, $".{Path.GetFileName(resolved)}.{Guid.NewGuid():N}.tmp");
            try
            {
                File.WriteAllText(temporaryPath, builder.ToString(), new UTF8Encoding(false));

                // Validate the written file before it replaces anything.
                var lines = File.ReadAllLines(temporaryPath, Encoding.UTF8);
                if (lines.Length != findings.Count + 1)
                {
                    throw new InvalidDataException("The export did not round-trip; the destination was left unchanged.");
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

    private static string Quote(string value)
    {
        var text = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    private static string RequirePath(string path) =>
        !string.IsNullOrWhiteSpace(path) ? path : throw new ArgumentException("A destination is required.", nameof(path));

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
