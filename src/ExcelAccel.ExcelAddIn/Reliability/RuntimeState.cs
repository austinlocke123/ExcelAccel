using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal static class RuntimeState
{
    private static readonly int ProcessId = Process.GetCurrentProcess().Id;
    private static int _excelThreadId;

    private static string SessionDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelAccel",
        "sessions");

    private static string MarkerPath => Path.Combine(SessionDirectory, $"{ProcessId}.running");

    public static bool IsSafeMode { get; private set; }

    public static void Start()
    {
        _excelThreadId = Thread.CurrentThread.ManagedThreadId;
        Directory.CreateDirectory(SessionDirectory);
        IsSafeMode = Directory.EnumerateFiles(SessionDirectory, "*.running")
            .Any(IsStaleMarker);
        File.WriteAllText(MarkerPath, DateTimeOffset.UtcNow.ToString("O"));
        AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
        AppDomain.CurrentDomain.DomainUnload += OnDomainUnload;
    }

    public static void StopCleanly()
    {
        AppDomain.CurrentDomain.ProcessExit -= OnProcessExit;
        AppDomain.CurrentDomain.DomainUnload -= OnDomainUnload;
        TryDelete(MarkerPath);
        IsSafeMode = false;
    }

    public static void VerifyExcelThread()
    {
        if (_excelThreadId == 0 || Thread.CurrentThread.ManagedThreadId != _excelThreadId)
        {
            throw new CommandRefusedException("Excel object access was refused because the callback is not on the owning Excel thread.");
        }
    }

    private static bool IsStaleMarker(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        if (!int.TryParse(name, out var processId))
        {
            TryDelete(path);
            return true;
        }

        if (processId == ProcessId)
        {
            return true;
        }

        try
        {
            using (Process.GetProcessById(processId))
            {
                return false;
            }
        }
        catch (ArgumentException)
        {
            TryDelete(path);
            return true;
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static void OnProcessExit(object? sender, EventArgs eventArgs) => TryDelete(MarkerPath);

    private static void OnDomainUnload(object? sender, EventArgs eventArgs) => TryDelete(MarkerPath);
}
