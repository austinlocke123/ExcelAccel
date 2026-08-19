using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.AutoColor;

public enum AutoColorCategory
{
    Text,
    NumericHardcode,
    SameSheetFormula,
    CrossSheetFormula,
    ExternalFormula,
    Error,
    Unsupported,
}

public enum CellScalarKind { Empty, Text, Number, Boolean, Error, Unsupported }
public enum AutoColorScope { Selection, Worksheet }

public sealed class AutoColorCellSnapshot
{
    public AutoColorCellSnapshot(string address, CellScalarKind scalarKind, string formula, string fontColor)
    {
        Address = string.IsNullOrWhiteSpace(address) ? throw new ArgumentException("A cell address is required.", nameof(address)) : address;
        ScalarKind = scalarKind;
        Formula = formula ?? string.Empty;
        FontColor = fontColor ?? string.Empty;
    }
    public string Address { get; }
    public CellScalarKind ScalarKind { get; }
    public string Formula { get; }
    public string FontColor { get; }
}

public sealed class AutoColorChange
{
    public AutoColorChange(string address, AutoColorCategory category, string beforeColor, string afterColor)
    {
        Address = address;
        Category = category;
        BeforeColor = beforeColor;
        AfterColor = afterColor;
    }
    public string Address { get; }
    public AutoColorCategory Category { get; }
    public string BeforeColor { get; }
    public string AfterColor { get; }
}

public sealed class AutoColorPlan
{
    public AutoColorPlan(SelectionContext context, AutoColorScope scope, IEnumerable<AutoColorChange> changes,
        IEnumerable<KeyValuePair<AutoColorCategory, int>> counts, int unsupportedCount, bool requiresPreview, string fingerprint)
    {
        Context = context;
        Scope = scope;
        Changes = changes.OrderBy(value => value.Address, StringComparer.Ordinal).ToArray();
        Counts = counts.ToDictionary(value => value.Key, value => value.Value);
        UnsupportedCount = unsupportedCount;
        RequiresPreview = requiresPreview;
        Fingerprint = fingerprint;
    }
    public SelectionContext Context { get; }
    public AutoColorScope Scope { get; }
    public IReadOnlyList<AutoColorChange> Changes { get; }
    public IReadOnlyDictionary<AutoColorCategory, int> Counts { get; }
    public int UnsupportedCount { get; }
    public bool RequiresPreview { get; }
    public string Fingerprint { get; }
}

public static class AutoColorPlanner
{
    public const int MaximumPlannableCells = 250_000;
    private static readonly Regex ExternalReference = new Regex(@"\[[^\]]+\]", RegexOptions.CultureInvariant);
    private static readonly Regex SheetReference = new Regex(@"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!", RegexOptions.CultureInvariant);

    public static AutoColorCategory Classify(AutoColorCellSnapshot cell)
    {
        if (cell.ScalarKind == CellScalarKind.Error) return AutoColorCategory.Error;
        if (!string.IsNullOrWhiteSpace(cell.Formula))
        {
            if (ExternalReference.IsMatch(cell.Formula)) return AutoColorCategory.ExternalFormula;
            if (SheetReference.IsMatch(cell.Formula)) return AutoColorCategory.CrossSheetFormula;
            return AutoColorCategory.SameSheetFormula;
        }
        if (cell.ScalarKind == CellScalarKind.Text) return AutoColorCategory.Text;
        if (cell.ScalarKind == CellScalarKind.Number) return AutoColorCategory.NumericHardcode;
        return AutoColorCategory.Unsupported;
    }

    public static AutoColorPlan Plan(ProfileDefinition profile, SelectionSnapshot selection,
        IEnumerable<AutoColorCellSnapshot> cells, AutoColorScope scope)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        var ordered = (cells ?? throw new ArgumentNullException(nameof(cells))).OrderBy(value => value.Address, StringComparer.Ordinal).ToArray();
        if (ordered.Length > MaximumPlannableCells || ordered.Length != selection.CellCount)
            throw new CommandRefusedException(RefusalCodes.ResourceLimit, "AutoColor scope exceeds the qualified planning bound or snapshot count.", "Select a smaller stable scope.");
        if (selection.Safety.AreaCount != 1 || selection.Safety.HasMergedCells || selection.Safety.WorksheetProtected || selection.Safety.WorkbookReadOnly ||
            !selection.Safety.DynamicArraySpillCheckSupported || selection.Safety.HasLegacyArray || selection.Safety.HasDynamicArraySpill)
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "AutoColor requires one editable, unmerged, array/spill-safe scope.", "Choose a supported scope.");

        var counts = Enum.GetValues(typeof(AutoColorCategory)).Cast<AutoColorCategory>().ToDictionary(value => value, _ => 0);
        var changes = new List<AutoColorChange>();
        foreach (var cell in ordered)
        {
            var category = Classify(cell);
            counts[category]++;
            if (category == AutoColorCategory.Unsupported) continue;
            var key = CategoryKey(category);
            var desired = profile.AutoColorColors[key];
            if (!string.Equals(cell.FontColor, desired, StringComparison.OrdinalIgnoreCase))
                changes.Add(new AutoColorChange(cell.Address, category, cell.FontColor, desired));
        }

        var fingerprintSource = string.Join("\n", ordered.Select(value => value.Address + "\0" + value.ScalarKind + "\0" + value.Formula + "\0" + value.FontColor));
        return new AutoColorPlan(selection.Context, scope, changes, counts, counts[AutoColorCategory.Unsupported],
            scope == AutoColorScope.Worksheet || ordered.Length > profile.ImmediatePreviewCellLimit,
            PreconditionFingerprint.Create(fingerprintSource));
    }

    public static CanExecuteResult ExecutionGate() => CanExecuteResult.Refuse(
        "PERFORMANCE_QUALIFICATION_REQUIRED",
        "AutoColor execution remains unavailable until AC-P0-006 qualification and rollback evidence pass.",
        "Run and approve the Qualification performance profile before enabling this command.");

    private static string CategoryKey(AutoColorCategory category)
    {
        switch (category)
        {
            case AutoColorCategory.Text: return "text";
            case AutoColorCategory.NumericHardcode: return "numeric_hardcode";
            case AutoColorCategory.SameSheetFormula: return "same_sheet_formula";
            case AutoColorCategory.CrossSheetFormula: return "cross_sheet_formula";
            case AutoColorCategory.ExternalFormula: return "external_formula";
            case AutoColorCategory.Error: return "error";
            default: throw new ArgumentOutOfRangeException(nameof(category));
        }
    }
}
