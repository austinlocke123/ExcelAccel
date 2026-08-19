using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.QuickKeys;

namespace ExcelAccel.Persistence.Profiles;

public enum BindingExportFormat { Csv = 0, Html = 1 }

public sealed class BindingExportPlan
{
    public BindingExportPlan(string destinationPath, BindingExportFormat format, bool overwrite, string existingSha256,
        byte[] content, int bindingCount, int conflictCount, string planHash)
    {
        DestinationPath = destinationPath; Format = format; Overwrite = overwrite; ExistingSha256 = existingSha256;
        Content = content; BindingCount = bindingCount; ConflictCount = conflictCount; PlanHash = planHash;
    }
    public string DestinationPath { get; }
    public BindingExportFormat Format { get; }
    public bool Overwrite { get; }
    public string ExistingSha256 { get; }
    public byte[] Content { get; }
    public int BindingCount { get; }
    public int ConflictCount { get; }
    public string PlanHash { get; }
    public string Manifest => $"Destination: {DestinationPath}\nFormat: {Format}\nBindings: {BindingCount}\nConflicts: {ConflictCount}\nBytes: {Content.Length}";
}

public sealed class BindingCheatSheetExporter
{
    public BindingExportPlan Plan(IEnumerable<QuickKeyBinding> bindings, IReadOnlyList<CommandDescriptor> registry,
        string destinationPath, BindingExportFormat format, bool overwrite)
    {
        var normalized = (bindings ?? throw new ArgumentNullException(nameof(bindings)))
            .OrderBy(value => QuickKeyValidator.Normalize(value.Sequence), StringComparer.Ordinal).ThenBy(value => value.CommandId, StringComparer.Ordinal).ToArray();
        var commands = registry ?? throw new ArgumentNullException(nameof(registry));
        foreach (var binding in normalized)
            if (!commands.Any(value => value.Id == binding.CommandId)) throw new InvalidDataException($"Binding references unknown command '{binding.CommandId}'.");
        var destination = string.IsNullOrWhiteSpace(destinationPath) ? throw new ArgumentException("A destination is required.", nameof(destinationPath)) : Path.GetFullPath(destinationPath);
        var existing = File.Exists(destination) ? HashFile(destination) : string.Empty;
        if (!overwrite && existing.Length != 0) throw new IOException("The binding export destination exists; explicit overwrite is required.");
        var conflicts = QuickKeyValidator.Validate(normalized);
        var content = Encoding.UTF8.GetBytes(format == BindingExportFormat.Csv
            ? Csv(normalized, commands, conflicts)
            : Html(normalized, commands, conflicts));
        var contentHash = Hash(content);
        var planHash = Hash(Encoding.UTF8.GetBytes($"bindings\n{destination}\n{format}\n{overwrite}\n{existing}\n{contentHash}"));
        return new BindingExportPlan(destination, format, overwrite, existing, content, normalized.Length, conflicts.Count, planHash);
    }

    public void Export(BindingExportPlan plan, string confirmedPlanHash)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (!string.Equals(plan.PlanHash, confirmedPlanHash, StringComparison.Ordinal)) throw new InvalidOperationException("The exact binding export plan was not confirmed.");
        var existing = File.Exists(plan.DestinationPath) ? HashFile(plan.DestinationPath) : string.Empty;
        if (!string.Equals(existing, plan.ExistingSha256, StringComparison.Ordinal)) throw new IOException("The binding export destination changed after preview.");
        var directory = Path.GetDirectoryName(plan.DestinationPath) ?? throw new ArgumentException("The destination requires a parent directory.");
        Directory.CreateDirectory(directory);
        var temporary = Path.Combine(directory, $".{Path.GetFileName(plan.DestinationPath)}.{Guid.NewGuid():N}.tmp");
        var backup = Path.Combine(directory, $".{Path.GetFileName(plan.DestinationPath)}.{Guid.NewGuid():N}.bak");
        try
        {
            File.WriteAllBytes(temporary, plan.Content);
            if (File.Exists(plan.DestinationPath)) File.Replace(temporary, plan.DestinationPath, backup, true); else File.Move(temporary, plan.DestinationPath);
            if (!string.Equals(Hash(File.ReadAllBytes(plan.DestinationPath)), Hash(plan.Content), StringComparison.Ordinal))
            {
                if (File.Exists(backup)) File.Replace(backup, plan.DestinationPath, null, true); else if (File.Exists(plan.DestinationPath)) File.Delete(plan.DestinationPath);
                throw new IOException("Binding export verification failed; the prior destination was restored where available.");
            }
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); if (File.Exists(backup)) File.Delete(backup); }
    }

    private static string Csv(IEnumerable<QuickKeyBinding> bindings, IReadOnlyList<CommandDescriptor> registry, IReadOnlyList<QuickKeyConflict> conflicts)
    {
        var builder = new StringBuilder("sequence,command_id,command_name,category,conflict\r\n");
        foreach (var binding in bindings)
        {
            var command = registry.First(value => value.Id == binding.CommandId);
            var conflict = string.Join("; ", conflicts.Where(value => value.Sequence.IndexOf(QuickKeyValidator.Normalize(binding.Sequence), StringComparison.OrdinalIgnoreCase) >= 0).Select(value => value.Reason));
            builder.Append(Q(QuickKeyValidator.Normalize(binding.Sequence))).Append(',').Append(Q(command.Id)).Append(',').Append(Q(command.DisplayName)).Append(',').Append(Q(command.Category)).Append(',').Append(Q(conflict)).Append("\r\n");
        }
        return builder.ToString();
    }

    private static string Html(IEnumerable<QuickKeyBinding> bindings, IReadOnlyList<CommandDescriptor> registry, IReadOnlyList<QuickKeyConflict> conflicts)
    {
        var builder = new StringBuilder("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>ExcelAccel bindings</title></head><body><h1>ExcelAccel bindings</h1><table><thead><tr><th>Sequence</th><th>Command</th><th>Category</th><th>Conflict</th></tr></thead><tbody>");
        foreach (var binding in bindings)
        {
            var command = registry.First(value => value.Id == binding.CommandId);
            var conflict = string.Join("; ", conflicts.Where(value => value.Sequence.IndexOf(QuickKeyValidator.Normalize(binding.Sequence), StringComparison.OrdinalIgnoreCase) >= 0).Select(value => value.Reason));
            builder.Append("<tr><td>").Append(WebUtility.HtmlEncode(QuickKeyValidator.Normalize(binding.Sequence))).Append("</td><td>")
                .Append(WebUtility.HtmlEncode(command.DisplayName)).Append("</td><td>").Append(WebUtility.HtmlEncode(command.Category))
                .Append("</td><td>").Append(WebUtility.HtmlEncode(conflict)).Append("</td></tr>");
        }
        return builder.Append("</tbody></table></body></html>\n").ToString();
    }
    private static string Q(string value) => "\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";
    private static string Hash(byte[] bytes) { using var sha = SHA256.Create(); return string.Concat(sha.ComputeHash(bytes).Select(value => value.ToString("X2"))); }
    private static string HashFile(string path) { using var stream = File.OpenRead(path); using var sha = SHA256.Create(); return string.Concat(sha.ComputeHash(stream).Select(value => value.ToString("X2"))); }
}
