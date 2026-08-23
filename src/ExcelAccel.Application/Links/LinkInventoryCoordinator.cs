using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Links;

namespace ExcelAccel.Application.Links;

/// <summary>
/// Reads link metadata. Nothing here opens, updates, repoints, breaks, or
/// recalculates a link, and nothing contacts a source.
/// </summary>
public interface ILinkInventoryPort
{
    string CaptureWorkbookId();

    /// <summary>Sources Excel itself reports, and whether each is open here.</summary>
    IReadOnlyList<LinkSourceRecord> CaptureSources();

    /// <summary>Usages found in the categories this build scans.</summary>
    IReadOnlyList<LinkUsageRecord> CaptureUsages();

    /// <summary>The categories the capture actually covered, so gaps are honest.</summary>
    IReadOnlyList<LinkCategory> ScannedCategories { get; }
}

public sealed class LinkInventoryResult
{
    public LinkInventoryResult(LinkInventory inventory, string workbookId, TraceResultPresentation presentation)
    {
        Inventory = inventory;
        WorkbookId = workbookId;
        Presentation = presentation;
    }

    public LinkInventory Inventory { get; }
    public string WorkbookId { get; }
    public TraceResultPresentation Presentation { get; }
}

public sealed class LinkInventoryCoordinator
{
    public const string CommandId = "links.inventory.open";

    public LinkInventoryResult Open(ILinkInventoryPort port)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));

        var workbookId = port.CaptureWorkbookId();
        if (string.IsNullOrWhiteSpace(workbookId))
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                "An open workbook is required to inventory its links.",
                "Open the workbook and retry.");
        }

        var sources = port.CaptureSources() ?? Array.Empty<LinkSourceRecord>();
        var usages = port.CaptureUsages() ?? Array.Empty<LinkUsageRecord>();
        var inventory = LinkInventoryBuilder.Build(sources, usages, port.ScannedCategories);
        return Present(inventory, workbookId, null);
    }

    /// <summary>
    /// Re-filters an inventory already captured. Takes no port, because a search
    /// must never reach Excel (AC-LINK-007).
    /// </summary>
    public LinkInventoryResult Search(LinkInventoryResult captured, string? query)
    {
        if (captured is null) throw new ArgumentNullException(nameof(captured));
        return Present(captured.Inventory, captured.WorkbookId, query);
    }

    private static LinkInventoryResult Present(LinkInventory inventory, string workbookId, string? query)
    {
        var filtered = string.IsNullOrWhiteSpace(query)
            ? inventory.Sources
            : LinkInventorySearch.Filter(inventory, query);
        return new LinkInventoryResult(
            inventory,
            workbookId,
            LinkInventoryPresentation.Create(inventory, workbookId, filtered));
    }
}
