using System;
using System.Collections.Generic;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Names;

namespace ExcelAccel.Application.Names;

/// <summary>
/// Reads defined-name metadata. Nothing here opens a workbook or contacts an
/// external source; a name pointing at a closed workbook is reported from local
/// metadata alone.
/// </summary>
public interface INameInventoryPort
{
    string CaptureWorkbookId();
    IReadOnlyList<NameRecord> CaptureNames();
}

public sealed class NameInventoryRequest
{
    public NameInventoryRequest(bool includeHidden, bool includeBuiltIn, string? query = null)
    {
        IncludeHidden = includeHidden;
        IncludeBuiltIn = includeBuiltIn;
        Query = query;
    }

    public bool IncludeHidden { get; }
    public bool IncludeBuiltIn { get; }
    public string? Query { get; }
}

public sealed class NameInventoryResult
{
    public NameInventoryResult(NameInventory inventory, string workbookId, TraceResultPresentation presentation)
    {
        Inventory = inventory;
        WorkbookId = workbookId;
        Presentation = presentation;
    }

    public NameInventory Inventory { get; }
    public string WorkbookId { get; }
    public TraceResultPresentation Presentation { get; }
}

/// <summary>
/// Builds a name inventory and projects it into the shared trace presentation.
/// The inventory is captured once; searching filters that snapshot rather than
/// re-reading the workbook, so typing cannot trigger a scan.
/// </summary>
public sealed class NameInventoryCoordinator
{
    public const string CommandId = "names.inventory.open";

    public NameInventoryResult Open(INameInventoryPort port, NameInventoryRequest request)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        if (request is null) throw new ArgumentNullException(nameof(request));

        var workbookId = port.CaptureWorkbookId();
        if (string.IsNullOrWhiteSpace(workbookId))
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                "An open workbook is required to inventory its names.",
                "Open the workbook and retry.");
        }

        var records = port.CaptureNames()
            ?? throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                "The workbook's defined names could not be read.",
                "Retry once the workbook is idle.");

        var inventory = NameInventoryBuilder.Build(records, request.IncludeHidden, request.IncludeBuiltIn);
        return Present(inventory, workbookId, request.Query);
    }

    /// <summary>
    /// Re-filters an inventory already captured. Deliberately takes no port: a
    /// search must never reach Excel.
    /// </summary>
    public NameInventoryResult Search(NameInventoryResult captured, string? query)
    {
        if (captured is null) throw new ArgumentNullException(nameof(captured));
        return Present(captured.Inventory, captured.WorkbookId, query);
    }

    private static NameInventoryResult Present(NameInventory inventory, string workbookId, string? query)
    {
        var filtered = string.IsNullOrWhiteSpace(query)
            ? inventory.Entries
            : NameInventorySearch.Filter(inventory, query);
        return new NameInventoryResult(
            inventory,
            workbookId,
            NameInventoryPresentation.Create(inventory, workbookId, filtered));
    }
}
