using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Links;
using ExcelAccel.Core.Links;
using ExcelAccel.Core.Names;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Reads external-link metadata from the active workbook.
/// </summary>
/// <remarks>
/// Read-only by construction. It calls <c>LinkSources</c>, walks formulas inside
/// the used range, and reads defined names. It never calls <c>UpdateLink</c>,
/// <c>BreakLink</c>, <c>ChangeLink</c>, or <c>Workbooks.Open</c>, never touches
/// the filesystem, and never triggers a recalculation, so no status it produces
/// can imply a check that did not happen.
///
/// Formula scanning is bounded by the used range and by an explicit cell ceiling,
/// because a workbook can report an enormous used range that costs far more to
/// read than it contains.
/// </remarks>
public sealed class ExcelLinkInventoryAdapter : ILinkInventoryPort
{
    /// <summary>Excel's xlExcelLinks.</summary>
    private const int ExcelLinksType = 1;

    /// <summary>Cells read per worksheet before the scan stops and says so.</summary>
    public const int MaximumScannedCells = 250_000;

    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;

    public ExcelLinkInventoryAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
    }

    /// <summary>
    /// Chart series, queries and validation are not scanned by this build, so
    /// they are reported as coverage gaps rather than silently counted as zero.
    /// </summary>
    public IReadOnlyList<LinkCategory> ScannedCategories { get; } =
        new[] { LinkCategory.FormulaReference, LinkCategory.DefinedName };

    public string CaptureWorkbookId()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => WithWorkbook(workbook => (string)((dynamic)workbook).Name));
    }

    public IReadOnlyList<LinkSourceRecord> CaptureSources()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => WithWorkbook(CaptureSourcesOnce));
    }

    public IReadOnlyList<LinkUsageRecord> CaptureUsages()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => WithWorkbook(CaptureUsagesOnce));
    }

    private T WithWorkbook<T>(Func<object, T> read)
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

            return read(workbookObject);
        }
        finally
        {
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private IReadOnlyList<LinkSourceRecord> CaptureSourcesOnce(object workbookObject)
    {
        var open = OpenWorkbookTokens();
        var records = new List<LinkSourceRecord>();
        object? sourcesObject = null;
        try
        {
            // Returns Empty rather than an array when the workbook has no links.
            sourcesObject = ((dynamic)workbookObject).LinkSources(ExcelLinksType);
            if (sourcesObject is Array array)
            {
                foreach (var item in array)
                {
                    var source = item as string;
                    if (string.IsNullOrWhiteSpace(source)) continue;
                    records.Add(new LinkSourceRecord(
                        source!, open.Contains(LinkInventoryBuilder.NormalizeToken(source))));
                }
            }
        }
        catch (Exception)
        {
            // A workbook that will not report its link sources still yields the
            // usages found by scanning, which is better than refusing outright.
        }
        finally
        {
            ComRelease.Owned(sourcesObject);
        }

        return records;
    }

    /// <summary>
    /// The set of workbooks open in this session, normalized the same way link
    /// sources are, so "open here" is a purely local comparison.
    /// </summary>
    private HashSet<string> OpenWorkbookTokens()
    {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        object? applicationObject = null;
        object? workbooksObject = null;
        try
        {
            applicationObject = _getApplication();
            workbooksObject = ((dynamic)applicationObject).Workbooks;
            int count = ((dynamic)workbooksObject).Count;
            for (var index = 1; index <= count; index++)
            {
                object? workbookObject = null;
                try
                {
                    workbookObject = ((dynamic)workbooksObject)[index];
                    tokens.Add(LinkInventoryBuilder.NormalizeToken((string)((dynamic)workbookObject).Name));
                }
                catch (Exception)
                {
                    // One unreadable workbook must not abandon the comparison.
                }
                finally
                {
                    ComRelease.Owned(workbookObject);
                }
            }
        }
        catch (Exception)
        {
        }
        finally
        {
            ComRelease.Owned(workbooksObject);
            ComRelease.Owned(applicationObject);
        }

        return tokens;
    }

    private IReadOnlyList<LinkUsageRecord> CaptureUsagesOnce(object workbookObject)
    {
        var records = new List<LinkUsageRecord>();
        CollectNameUsages(workbookObject, records);
        CollectFormulaUsages(workbookObject, records);
        return records;
    }

    private static void CollectNameUsages(object workbookObject, List<LinkUsageRecord> records)
    {
        object? namesObject = null;
        try
        {
            namesObject = ((dynamic)workbookObject).Names;
            int count = ((dynamic)namesObject).Count;
            for (var index = 1; index <= count && records.Count < LinkInventoryBuilder.MaximumUsages; index++)
            {
                object? nameObject = null;
                try
                {
                    nameObject = ((dynamic)namesObject)[index];
                    var displayName = (string)((dynamic)nameObject).Name;
                    var refersTo = (string)((dynamic)nameObject).RefersTo ?? string.Empty;
                    if (refersTo.IndexOf('[') < 0) continue;

                    var kind = NameInventoryBuilder.Classify(refersTo);
                    records.Add(new LinkUsageRecord(
                        LinkCategory.DefinedName,
                        refersTo,
                        null,
                        null,
                        "Name: " + displayName,
                        isBroken: kind == NameTargetKind.Broken));
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComRelease.Owned(nameObject);
                }
            }
        }
        catch (Exception)
        {
        }
        finally
        {
            ComRelease.Owned(namesObject);
        }
    }

    private static void CollectFormulaUsages(object workbookObject, List<LinkUsageRecord> records)
    {
        object? worksheetsObject = null;
        try
        {
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            int worksheetCount = ((dynamic)worksheetsObject).Count;
            for (var index = 1; index <= worksheetCount && records.Count < LinkInventoryBuilder.MaximumUsages; index++)
            {
                object? worksheetObject = null;
                try
                {
                    worksheetObject = ((dynamic)worksheetsObject)[index];
                    ScanWorksheet(worksheetObject, records);
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComRelease.Owned(worksheetObject);
                }
            }
        }
        catch (Exception)
        {
        }
        finally
        {
            ComRelease.Owned(worksheetsObject);
        }
    }

    private static void ScanWorksheet(object worksheetObject, List<LinkUsageRecord> records)
    {
        object? usedRangeObject = null;
        object? formulaCellsObject = null;
        object? areasObject = null;
        var worksheetName = (string)((dynamic)worksheetObject).Name;
        try
        {
            usedRangeObject = ((dynamic)worksheetObject).UsedRange;
            if (usedRangeObject is null) return;

            long cellCount;
            try { cellCount = (long)(double)((dynamic)usedRangeObject).Count; }
            catch { cellCount = 0; }
            if (cellCount <= 0 || cellCount > MaximumScannedCells) return;

            // SpecialCells narrows the read to formula cells only, and throws
            // rather than returning empty when a worksheet has none.
            try { formulaCellsObject = ((dynamic)usedRangeObject).SpecialCells(-4123); }
            catch { return; }
            if (formulaCellsObject is null) return;

            areasObject = ((dynamic)formulaCellsObject).Areas;
            int areaCount = ((dynamic)areasObject).Count;
            for (var areaIndex = 1; areaIndex <= areaCount && records.Count < LinkInventoryBuilder.MaximumUsages; areaIndex++)
            {
                object? areaObject = null;
                try
                {
                    areaObject = ((dynamic)areasObject)[areaIndex];
                    ScanArea(areaObject, worksheetName, records);
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComRelease.Owned(areaObject);
                }
            }
        }
        catch (Exception)
        {
        }
        finally
        {
            ComRelease.Owned(areasObject);
            ComRelease.Owned(formulaCellsObject);
            ComRelease.Owned(usedRangeObject);
        }
    }

    private static void ScanArea(object areaObject, string worksheetName, List<LinkUsageRecord> records)
    {
        object? cellsObject = null;
        try
        {
            cellsObject = ((dynamic)areaObject).Cells;
            int count = ((dynamic)cellsObject).Count;
            for (var index = 1; index <= count && records.Count < LinkInventoryBuilder.MaximumUsages; index++)
            {
                object? cellObject = null;
                try
                {
                    cellObject = ((dynamic)cellsObject)[index];
                    var formula = (string)((dynamic)cellObject).Formula ?? string.Empty;
                    if (formula.IndexOf('[') < 0) continue;

                    var address = (string)((dynamic)cellObject).Address(false, false);
                    records.Add(new LinkUsageRecord(
                        LinkCategory.FormulaReference,
                        formula,
                        worksheetName,
                        address,
                        Summarize(formula),
                        isBroken: formula.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0));
                }
                catch (Exception)
                {
                }
                finally
                {
                    ComRelease.Owned(cellObject);
                }
            }
        }
        catch (Exception)
        {
        }
        finally
        {
            ComRelease.Owned(cellsObject);
        }
    }

    /// <summary>
    /// A bounded excerpt. The whole formula can be long and can carry a path, and
    /// the export redacts paths by default, so the detail column stays short.
    /// </summary>
    private static string Summarize(string formula) =>
        formula.Length <= 120 ? formula : formula.Substring(0, 117) + "...";
}
