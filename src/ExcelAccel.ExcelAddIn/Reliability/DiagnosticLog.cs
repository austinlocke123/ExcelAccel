using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal static class DiagnosticLog
{
    private static readonly object Sync = new object();

    public static void Info(string operationId, string outcome, long? elapsedMilliseconds = null) =>
        Write("INFO", operationId, outcome, null, elapsedMilliseconds);

    public static void Error(string operationId, Exception exception, long? elapsedMilliseconds = null) =>
        Write("ERROR", operationId, "failed", exception.GetType().FullName, elapsedMilliseconds);

    private static void Write(string level, string operationId, string outcome, string? errorType, long? elapsedMilliseconds)
    {
        try
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelAccel",
                "logs");
            Directory.CreateDirectory(directory);

            var line = new StringBuilder()
                .Append(DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture)).Append('\t')
                .Append(level).Append('\t')
                .Append(Sanitize(operationId)).Append('\t')
                .Append(Sanitize(outcome)).Append('\t')
                .Append(elapsedMilliseconds?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append('\t')
                .Append(Sanitize(errorType ?? string.Empty))
                .AppendLine()
                .ToString();

            lock (Sync)
            {
                File.AppendAllText(Path.Combine(directory, "excelaccel.log"), line, Encoding.UTF8);
            }
        }
        catch
        {
            // Diagnostics are best-effort and must never destabilize Excel.
        }
    }

    private static string Sanitize(string value) =>
        value.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
}
