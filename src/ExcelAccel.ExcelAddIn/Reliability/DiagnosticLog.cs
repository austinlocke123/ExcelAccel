using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal static class DiagnosticLog
{
    private static readonly object Sync = new object();

    private const long MaximumLogBytes = 1_048_576;

    public static void Info(string operationId, string outcome, long? elapsedMilliseconds = null) =>
        Write("INFO", string.Empty, operationId, outcome, null, elapsedMilliseconds);

    public static string Error(string operationId, Exception exception, long? elapsedMilliseconds = null)
    {
        var diagnosticId = Guid.NewGuid().ToString("N");
        Write("ERROR", diagnosticId, operationId, "failed", exception.GetType().FullName, elapsedMilliseconds);
        return diagnosticId;
    }

    public static string Failure(string operationId, string failureCode, Exception exception, long? elapsedMilliseconds = null)
    {
        var diagnosticId = Guid.NewGuid().ToString("N");
        Write("ERROR", diagnosticId, operationId, failureCode, exception.GetType().FullName, elapsedMilliseconds);
        return diagnosticId;
    }

    public static string LogPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAccel", "logs", "excelaccel.log");

    private static void Write(string level, string diagnosticId, string operationId, string outcome, string? errorType, long? elapsedMilliseconds)
    {
        try
        {
            var directory = Path.GetDirectoryName(LogPath) ?? throw new InvalidOperationException();
            Directory.CreateDirectory(directory);

            var line = new StringBuilder()
                .Append(DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture)).Append('\t')
                .Append(level).Append('\t')
                .Append(Sanitize(diagnosticId)).Append('\t')
                .Append(Sanitize(operationId)).Append('\t')
                .Append(Sanitize(outcome)).Append('\t')
                .Append(elapsedMilliseconds?.ToString(CultureInfo.InvariantCulture) ?? string.Empty).Append('\t')
                .Append(Sanitize(errorType ?? string.Empty))
                .AppendLine()
                .ToString();

            lock (Sync)
            {
                if (File.Exists(LogPath) && new FileInfo(LogPath).Length + Encoding.UTF8.GetByteCount(line) > MaximumLogBytes)
                {
                    var prior = LogPath + ".previous";
                    if (File.Exists(prior)) File.Delete(prior);
                    File.Move(LogPath, prior);
                }
                File.AppendAllText(LogPath, line, Encoding.UTF8);
            }
        }
        catch
        {
            // Diagnostics are best-effort and must never destabilize Excel.
        }
    }

    private static string Sanitize(string value)
    {
        var safe = value.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
        return safe.Length <= 256 ? safe : safe.Substring(0, 256);
    }
}
