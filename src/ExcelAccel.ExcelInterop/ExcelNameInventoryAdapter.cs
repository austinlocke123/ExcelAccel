using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Names;
using ExcelAccel.Core.Names;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Reads defined-name metadata from the active workbook.
/// </summary>
/// <remarks>
/// Read-only by construction: it touches <c>Names</c> collections and nothing
/// else, never opens a workbook, and never resolves a name that points somewhere
/// closed. A name whose properties cannot be read is still emitted, with an empty
/// expression, so it appears as unresolved rather than vanishing from a list the
/// user believes is complete.
/// </remarks>
public sealed class ExcelNameInventoryAdapter : INameInventoryPort
{
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;

    public ExcelNameInventoryAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
    }

    public string CaptureWorkbookId()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureWorkbookIdOnce);
    }

    public IReadOnlyList<NameRecord> CaptureNames()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureNamesOnce);
    }

    private string CaptureWorkbookIdOnce()
    {
        object? applicationObject = null;
        object? workbookObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported,
                    "An active workbook is required.",
                    "Open a workbook and retry.");
            }

            return (string)((dynamic)workbookObject).Name;
        }
        finally
        {
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private IReadOnlyList<NameRecord> CaptureNamesOnce()
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? namesObject = null;
        object? worksheetsObject = null;
        var records = new List<NameRecord>();
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported,
                    "An active workbook is required.",
                    "Open a workbook and retry.");
            }

            namesObject = ((dynamic)workbookObject).Names;
            Collect(namesObject, NameScopeKind.Workbook, null, records);

            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            int worksheetCount = ((dynamic)worksheetsObject).Count;
            for (var index = 1; index <= worksheetCount && records.Count < NameInventoryBuilder.MaximumNames; index++)
            {
                object? worksheetObject = null;
                object? sheetNamesObject = null;
                try
                {
                    worksheetObject = ((dynamic)worksheetsObject)[index];
                    var worksheetName = (string)((dynamic)worksheetObject).Name;
                    sheetNamesObject = ((dynamic)worksheetObject).Names;
                    Collect(sheetNamesObject, NameScopeKind.Worksheet, worksheetName, records);
                }
                finally
                {
                    ComRelease.Owned(sheetNamesObject);
                    ComRelease.Owned(worksheetObject);
                }
            }

            return records;
        }
        finally
        {
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(namesObject);
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private static void Collect(object? namesObject, NameScopeKind scope, string? worksheetName, List<NameRecord> records)
    {
        if (namesObject is null)
        {
            return;
        }

        int count;
        try { count = ((dynamic)namesObject).Count; }
        catch { return; }

        for (var index = 1; index <= count && records.Count < NameInventoryBuilder.MaximumNames; index++)
        {
            object? nameObject = null;
            try
            {
                nameObject = ((dynamic)namesObject)[index];
                var record = TryRead(nameObject, scope, worksheetName);
                if (record is not null)
                {
                    records.Add(record);
                }
            }
            catch
            {
                // A single unreadable name must not abandon the inventory.
            }
            finally
            {
                ComRelease.Owned(nameObject);
            }
        }
    }

    private static NameRecord? TryRead(object nameObject, NameScopeKind scope, string? worksheetName)
    {
        string displayName;
        try { displayName = (string)((dynamic)nameObject).Name; }
        catch { return null; }
        if (string.IsNullOrWhiteSpace(displayName))
        {
            return null;
        }

        // A worksheet-scoped name reports as "Sheet1!Local"; the inventory holds
        // the scope separately, so the qualifier is stripped from the display.
        var localName = displayName;
        if (scope == NameScopeKind.Worksheet)
        {
            var separator = localName.LastIndexOf('!');
            if (separator >= 0 && separator + 1 < localName.Length)
            {
                localName = localName.Substring(separator + 1);
            }
        }

        var refersTo = string.Empty;
        try { refersTo = (string)((dynamic)nameObject).RefersTo ?? string.Empty; }
        catch { refersTo = string.Empty; }

        var visible = true;
        try { visible = (bool)((dynamic)nameObject).Visible; }
        catch { visible = true; }

        // Excel exposes no built-in flag; ReservedNames in the core layer owns
        // that recognition so it can be tested without Excel.
        return new NameRecord(localName, scope, worksheetName, refersTo, visible, ReservedNames.IsReserved(localName));
    }

}
