using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;
using ExcelAccel.Core.ModelCheck;

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
    /// <summary>
    /// Classifies one cell. Precedence is Error, numeric hardcode, external,
    /// cross-sheet, same-sheet, text; the first match wins.
    /// </summary>
    /// <remarks>
    /// A hardcode outranks external and cross-sheet deliberately: the point of
    /// the colouring is to make typed numbers findable, and a number buried in an
    /// external reference is exactly the kind that hides. No allowlist applies,
    /// so <c>=A1*2</c> is a hardcode. That is a deliberate divergence from Model
    /// Check, whose embedded-constant rule allowlists 0, 1, -1, 2, 100, 12 and
    /// 365 because its job is to raise findings worth a person's attention. A
    /// colour map is worse for a missed hardcode than an over-coloured one, so
    /// the two rules must not be unified. See docs/commands/AUTOCOLOR.md.
    /// </remarks>
    public static AutoColorCategory Classify(AutoColorCellSnapshot cell, string originWorksheet)
    {
        if (cell is null) throw new ArgumentNullException(nameof(cell));
        if (cell.ScalarKind == CellScalarKind.Error) return AutoColorCategory.Error;

        if (string.IsNullOrWhiteSpace(cell.Formula))
        {
            switch (cell.ScalarKind)
            {
                case CellScalarKind.Number: return AutoColorCategory.NumericHardcode;
                case CellScalarKind.Text: return AutoColorCategory.Text;
                default: return AutoColorCategory.Unsupported;
            }
        }

        var parsed = new FormulaParser().Parse(cell.Formula, FormulaParseOptions.DefaultA1);
        if (!parsed.IsSuccess || parsed.Document is null)
        {
            // A formula outside qualified parser coverage cannot be classified
            // honestly, and guessing would either hide a hardcode or invent one.
            // Unsupported cells are counted and left exactly as they are.
            return AutoColorCategory.Unsupported;
        }

        // ReadEmbeddedLiterals also returns empty on a parse failure, which is
        // why success is checked above rather than relying on an empty result.
        if (FormulaShape.ReadEmbeddedLiterals(cell.Formula).Count > 0)
        {
            return AutoColorCategory.NumericHardcode;
        }

        var references = parsed.Document.References;
        if (references.Any(reference => reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0))
        {
            return AutoColorCategory.ExternalFormula;
        }

        if (references.Any(reference => reference.Qualifier is not null && !IsOrigin(reference.Qualifier, originWorksheet)))
        {
            return AutoColorCategory.CrossSheetFormula;
        }

        return AutoColorCategory.SameSheetFormula;
    }

    /// <summary>
    /// A sheet-qualified reference to the sheet the formula lives on is
    /// same-sheet: <c>=Sheet1!A1</c> written on Sheet1 is not cross-sheet.
    /// </summary>
    private static bool IsOrigin(string qualifier, string originWorksheet)
    {
        var name = qualifier.Trim();
        if (name.Length >= 2 && name[0] == '\'' && name[name.Length - 1] == '\'')
        {
            name = name.Substring(1, name.Length - 2).Replace("''", "'");
        }

        return string.Equals(name, originWorksheet ?? string.Empty, StringComparison.OrdinalIgnoreCase);
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
            var category = Classify(cell, selection.Context.WorksheetName);
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

    /// <summary>
    /// Selection scope is bounded by the selection the user made and is
    /// permitted. Worksheet scope recolours thousands of cells in one action and
    /// stays refused until it has a transactional adapter, rollback and
    /// fault-injection evidence, and a worksheet-scale preview.
    /// </summary>
    public static CanExecuteResult ExecutionGate(AutoColorScope scope) =>
        scope == AutoColorScope.Selection
            ? CanExecuteResult.Permit()
            : CanExecuteResult.Refuse(
                "PERFORMANCE_QUALIFICATION_REQUIRED",
                "AutoColor over a worksheet remains unavailable until AC-P0-006 qualification and rollback evidence pass.",
                "Run AutoColor over a selection, or approve the Qualification performance profile.");

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
