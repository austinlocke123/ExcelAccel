using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// One worksheet's place in a workbook-scope plan: either an included read plan,
/// or an explicit exclusion with the reason it could not be planned.
/// </summary>
public sealed class WorksheetScanEntry
{
    internal WorksheetScanEntry(string worksheetName, DependentScanRegion? region, string? exclusionCode, string? exclusionMessage)
    {
        WorksheetName = worksheetName;
        Region = region;
        ExclusionCode = exclusionCode;
        ExclusionMessage = exclusionMessage;
    }

    public string WorksheetName { get; }

    public DependentScanRegion? Region { get; }

    public string? ExclusionCode { get; }

    public string? ExclusionMessage { get; }

    public bool IsIncluded => Region is not null;

    public long CellCount => Region?.CellCount ?? 0;

    public int BlockCount => Region?.BlockCount ?? 0;
}

/// <summary>
/// A bounded read plan across the worksheets of one workbook.
///
/// Every ceiling is applied here, in pure code. A worksheet that cannot be
/// planned within the per-worksheet bounds is excluded with a stated reason
/// rather than failing the whole workbook, and an excluded worksheet is a
/// coverage gap that blocks any completeness claim. The workbook as a whole is
/// refused only when the plan cannot be bounded at all.
/// </summary>
public sealed class WorkbookScanPlan
{
    /// <summary>Largest worksheet count a workbook-scope scan will plan.</summary>
    public const int MaximumWorksheets = 64;

    /// <summary>Largest aggregate cell count across every included worksheet.</summary>
    public const long MaximumWorkbookCells = 1_000_000;

    private WorkbookScanPlan(string workbookId, IReadOnlyList<WorksheetScanEntry> entries)
    {
        WorkbookId = workbookId;
        Entries = entries;
    }

    public string WorkbookId { get; }

    /// <summary>Every worksheet considered, in canonical order, included or not.</summary>
    public IReadOnlyList<WorksheetScanEntry> Entries { get; }

    public IReadOnlyList<WorksheetScanEntry> Included =>
        Entries.Where(entry => entry.IsIncluded).ToArray();

    public IReadOnlyList<WorksheetScanEntry> Excluded =>
        Entries.Where(entry => !entry.IsIncluded).ToArray();

    public long TotalCellCount => Entries.Sum(entry => entry.CellCount);

    public int TotalBlockCount => Entries.Sum(entry => entry.BlockCount);

    /// <summary>An excluded worksheet cannot be read, so completeness is impossible.</summary>
    public bool CanClaimCompleteness => Excluded.Count == 0;

    /// <summary>
    /// Builds the plan from each worksheet's reported used region. The reported
    /// regions are untrusted input, exactly as they are for a single worksheet.
    /// </summary>
    public static bool TryCreate(
        string workbookId,
        IEnumerable<UsedRegionBounds> worksheets,
        out WorkbookScanPlan? plan,
        out string? refusalCode,
        out string? message)
    {
        if (string.IsNullOrWhiteSpace(workbookId)) throw new ArgumentException("A workbook identity is required.", nameof(workbookId));
        if (worksheets is null) throw new ArgumentNullException(nameof(worksheets));
        plan = null;
        refusalCode = null;
        message = null;

        var ordered = worksheets
            .OrderBy(sheet => sheet.WorksheetName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (ordered.Length > MaximumWorksheets)
        {
            refusalCode = AuditRefusalCodes.ScanRegionTooLarge;
            message = $"The workbook has {ordered.Length:N0} worksheets, which exceeds the {MaximumWorksheets:N0}-worksheet ceiling for a workbook scan. Scan a worksheet at a time.";
            return false;
        }

        var entries = new List<WorksheetScanEntry>();
        long total = 0;
        foreach (var sheet in ordered)
        {
            if (DependentScanRegion.TryCreate(sheet, out var region, out var code, out var sheetMessage))
            {
                total = checked(total + region!.CellCount);
                entries.Add(new WorksheetScanEntry(sheet.WorksheetName, region, null, null));
            }
            else
            {
                entries.Add(new WorksheetScanEntry(sheet.WorksheetName, null, code, sheetMessage));
            }
        }

        var included = entries.Count(entry => entry.IsIncluded);
        if (included == 0 && entries.Count != 0)
        {
            // Every worksheet was excluded, so there is nothing to read. Report
            // the first reason rather than returning an empty plan that would
            // read as "no findings".
            var first = entries[0];
            refusalCode = first.ExclusionCode;
            message = first.ExclusionMessage;
            return false;
        }

        if (total > MaximumWorkbookCells)
        {
            refusalCode = AuditRefusalCodes.ScanRegionTooLarge;
            message = $"The workbook's included worksheets cover {total:N0} cells, which exceeds the {MaximumWorkbookCells:N0}-cell workbook scan ceiling. Scan a worksheet at a time.";
            return false;
        }

        plan = new WorkbookScanPlan(workbookId, entries.AsReadOnly());
        return true;
    }

    /// <summary>
    /// The inventory a user confirms before a workbook scan reads anything: which
    /// worksheets are included, which are excluded and why, and the workload.
    /// </summary>
    public IReadOnlyList<string> InventoryLines()
    {
        var lines = new List<string>
        {
            "Workbook: " + WorkbookId,
            "Worksheets included: " + Included.Count.ToString("N0", CultureInfo.InvariantCulture),
            "Worksheets excluded: " + Excluded.Count.ToString("N0", CultureInfo.InvariantCulture),
            "Cells to read: " + TotalCellCount.ToString("N0", CultureInfo.InvariantCulture),
            "Bounded blocks: " + TotalBlockCount.ToString("N0", CultureInfo.InvariantCulture),
        };
        foreach (var entry in Included)
        {
            lines.Add("  include " + entry.WorksheetName + ": " + entry.CellCount.ToString("N0", CultureInfo.InvariantCulture) + " cells");
        }

        foreach (var entry in Excluded)
        {
            lines.Add("  EXCLUDE " + entry.WorksheetName + ": " + entry.ExclusionCode);
        }

        if (!CanClaimCompleteness)
        {
            lines.Add("An excluded worksheet cannot be read, so this scan cannot claim completeness.");
        }

        return lines.AsReadOnly();
    }
}
