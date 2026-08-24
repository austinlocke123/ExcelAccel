using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Compare;

namespace ExcelAccel.Application.Compare;

/// <summary>
/// Captures the two sides of a comparison. Nothing here opens, saves, closes, or
/// recalculates a workbook; both sides must already be open.
/// </summary>
public interface IComparisonPort
{
    /// <summary>The range the user has captured as the source, if any.</summary>
    ComparisonBlock? CaptureSource();

    /// <summary>The current selection as the target side.</summary>
    ComparisonBlock CaptureTarget();
}

public sealed class ComparisonSession
{
    public ComparisonSession(ComparisonResult result, TraceResultPresentation presentation)
    {
        Result = result;
        Presentation = presentation;
    }

    public ComparisonResult Result { get; }
    public TraceResultPresentation Presentation { get; }
}

/// <summary>
/// Runs a same-shape comparison and projects it for the shared read-only view.
/// </summary>
/// <remarks>
/// The source side is captured first by an explicit command, exactly like the
/// existing paste-source capture, so a comparison always names two sides the user
/// chose rather than inferring one.
/// </remarks>
public sealed class ComparisonCoordinator
{
    public ComparisonSession Compare(IComparisonPort port, ComparisonCategories? categories = null)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));

        var source = port.CaptureSource();
        if (source is null)
        {
            throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                "Capture a comparison source first, then select the range to compare it against.",
                "Use Capture Comparison Source on the first range.");
        }

        var target = port.CaptureTarget();
        if (string.Equals(source.Identity, target.Identity, StringComparison.OrdinalIgnoreCase))
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                "The source and target are the same range, so there is nothing to compare.",
                "Select a different range as the target.");
        }

        try
        {
            var result = SameShapeComparer.Compare(source, target, categories);
            return new ComparisonSession(result, ComparisonPresentation.Create(result));
        }
        catch (ShapeMismatchException exception)
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported, exception.Message, "Choose ranges with equal dimensions.");
        }
        catch (ComparisonBoundException exception)
        {
            throw new CommandRefusedException(
                RefusalCodes.ResourceLimit, exception.Message, "Compare a smaller range.");
        }
    }

    /// <summary>
    /// Re-projects a captured result against the other side. Takes no port,
    /// because switching sides must not re-read either workbook.
    /// </summary>
    public ComparisonSession ShowTargetSide(ComparisonSession captured)
    {
        if (captured is null) throw new ArgumentNullException(nameof(captured));
        return new ComparisonSession(
            captured.Result, ComparisonPresentation.Create(captured.Result, navigateTarget: true));
    }
}

public static class CompareCommandCatalog
{
    public const string CaptureSourceId = "compare.source.capture";
    public const string RangesId = "compare.ranges.same_shape";
    public const string NavigateTargetId = "compare.result.navigate_target";
    public const string ExportId = "compare.results.export";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        ReadOnly(
            CaptureSourceId,
            "Capture Comparison Source",
            "Remember the selected range as the source side of a comparison. Nothing is read from the workbook until the comparison itself runs.",
            new[] { "compare source", "capture compare" },
            new[] { "AC-CMP-001" }),
        ReadOnly(
            RangesId,
            "Compare With Captured Source",
            "Compare the selection against the captured source, position by position. Both sides must already be open, neither is opened, saved, recalculated, or changed, and unequal shapes are refused rather than aligned.",
            new[] { "compare ranges", "diff ranges", "compare" },
            new[] { "AC-CMP-001", "AC-CMP-002", "AC-CMP-003", "AC-CMP-004", "AC-CMP-005", "AC-CMP-006", "AC-CMP-007", "AC-CMP-008", "AC-CMP-014" }),
        ReadOnly(
            NavigateTargetId,
            "Show Comparison Target Side",
            "Re-list the same differences against the target side, so a result can be navigated on either workbook. Neither side is re-read.",
            new[] { "compare target", "other side" },
            new[] { "AC-CMP-016", "AC-NAV-005" }),
        ReadOnly(
            ExportId,
            "Export Comparison Results",
            "Write the current differences to a local file after confirming a manifest. Cell contents are excluded by default and nothing is transmitted.",
            new[] { "export comparison", "export diff" },
            new[] { "AC-CMP-017", "AC-CMP-018", "AC-CMP-019", "AC-SEC-004" }),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    private static CommandDescriptor ReadOnly(
        string id,
        string name,
        string description,
        IEnumerable<string> aliases,
        IEnumerable<string> acceptanceIds) =>
        new CommandDescriptor(
            id,
            1,
            name,
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            RibbonRoutes.For(id),
            "CAP-CMP-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            acceptanceIds,
            "Auditing",
            description,
            aliases,
            RibbonRoutes.For(id));
}
