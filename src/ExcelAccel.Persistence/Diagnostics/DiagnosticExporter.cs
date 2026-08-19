using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAccel.Persistence.Diagnostics;

public sealed class DiagnosticExportPlan
{
    internal DiagnosticExportPlan(string destinationPath, int byteCount, string contentHash, string planHash)
    {
        DestinationPath = destinationPath;
        ByteCount = byteCount;
        ContentHash = contentHash;
        PlanHash = planHash;
    }
    public string DestinationPath { get; }
    public int ByteCount { get; }
    public string ContentHash { get; }
    public string PlanHash { get; }
    public string Manifest => $"One local sanitized diagnostic log ({ByteCount:N0} bytes); no workbook formulas, values, names, paths, images, or automatic transmission.";
}

public sealed class DiagnosticExporter
{
    public const int MaximumExportBytes = 5 * 1024 * 1024;

    public DiagnosticExportPlan Plan(string destinationPath, byte[] sanitizedLog)
    {
        if (sanitizedLog is null) throw new ArgumentNullException(nameof(sanitizedLog));
        if (sanitizedLog.Length > MaximumExportBytes) throw new InvalidDataException("The diagnostic export exceeds its size bound.");
        var destination = RequireLocalPath(destinationPath);
        var contentHash = Hash(sanitizedLog);
        var canonical = destination + "\n" + sanitizedLog.Length + "\n" + contentHash;
        return new DiagnosticExportPlan(destination, sanitizedLog.Length, contentHash, Hash(Encoding.UTF8.GetBytes(canonical)));
    }

    public void Export(DiagnosticExportPlan plan, string confirmedPlanHash, byte[] sanitizedLog)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (!string.Equals(plan.PlanHash, confirmedPlanHash, StringComparison.Ordinal)) throw new InvalidOperationException("Exact diagnostic manifest confirmation is required.");
        if (sanitizedLog is null || sanitizedLog.Length != plan.ByteCount || !string.Equals(Hash(sanitizedLog), plan.ContentHash, StringComparison.Ordinal))
            throw new InvalidOperationException("The diagnostic content changed after preview.");
        var directory = Path.GetDirectoryName(plan.DestinationPath) ?? throw new InvalidOperationException("A destination directory is required.");
        Directory.CreateDirectory(directory);
        var temporary = Path.Combine(directory, "." + Path.GetFileName(plan.DestinationPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        var backup = plan.DestinationPath + ".backup";
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                stream.Write(sanitizedLog, 0, sanitizedLog.Length);
                stream.Flush();
            }
            if (!string.Equals(Hash(File.ReadAllBytes(temporary)), plan.ContentHash, StringComparison.Ordinal)) throw new IOException("Temporary diagnostic verification failed.");
            if (File.Exists(plan.DestinationPath)) File.Replace(temporary, plan.DestinationPath, backup, true);
            else File.Move(temporary, plan.DestinationPath);
            if (!string.Equals(Hash(File.ReadAllBytes(plan.DestinationPath)), plan.ContentHash, StringComparison.Ordinal)) throw new IOException("Final diagnostic verification failed.");
            if (File.Exists(backup)) File.Delete(backup);
        }
        catch
        {
            if (File.Exists(backup))
            {
                if (File.Exists(plan.DestinationPath)) File.Replace(backup, plan.DestinationPath, null, true);
                else File.Move(backup, plan.DestinationPath);
            }
            throw;
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private static string RequireLocalPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination path is required.", nameof(path));
        var full = Path.GetFullPath(path);
        if (new Uri(full).IsUnc) throw new ArgumentException("Diagnostic export supports local destinations only.", nameof(path));
        return full;
    }

    private static string Hash(byte[] value)
    {
        using var sha = SHA256.Create();
        var bytes = sha.ComputeHash(value);
        var builder = new StringBuilder(bytes.Length * 2);
        foreach (var item in bytes) builder.Append(item.ToString("x2"));
        return builder.ToString();
    }
}
