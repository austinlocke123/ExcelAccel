using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Compare;

/// <summary>What a difference is about.</summary>
public enum DifferenceCategory
{
    Formula,
    Constant,
    DisplayedValue,
    NumberFormat,

    /// <summary>
    /// One side holds something the other does not, such as a formula against a
    /// constant. Reported as its own category rather than forced into either.
    /// </summary>
    TypeMismatch,

    /// <summary>
    /// The pair could not be examined. Never treated as equal; see AC-CMP-005.
    /// </summary>
    Unsupported,
}

/// <summary>
/// How two formulas differ, when both are formulas.
/// </summary>
public enum FormulaDifferenceKind
{
    /// <summary>Byte-identical.</summary>
    Identical,

    /// <summary>Different text, and the parser covers both, so the difference is real.</summary>
    Structure,

    /// <summary>
    /// Different text that normalizes to the same relative shape, which is what a
    /// copied formula looks like from a different anchor.
    /// </summary>
    EquivalentShape,

    /// <summary>
    /// At least one side is outside qualified parser coverage. The texts differ,
    /// but nothing is claimed about why.
    /// </summary>
    OutsideCoverage,
}

/// <summary>One cell as captured from one side, before any comparison.</summary>
public sealed class ComparisonCell
{
    public ComparisonCell(string address, string formula, string value, string numberFormat, string displayedValue = "")
    {
        Address = string.IsNullOrWhiteSpace(address)
            ? throw new ArgumentException("A cell address is required.", nameof(address))
            : address.Trim().ToUpperInvariant();
        Formula = formula ?? string.Empty;
        Value = value ?? string.Empty;
        NumberFormat = numberFormat ?? string.Empty;
        DisplayedValue = displayedValue ?? string.Empty;
    }

    public string Address { get; }

    /// <summary>Empty when the cell holds no formula.</summary>
    public string Formula { get; }

    public string Value { get; }
    public string NumberFormat { get; }
    public string DisplayedValue { get; }

    public bool IsFormula => Formula.Length > 0 && Formula[0] == '=';
}

/// <summary>One rectangular side of a comparison.</summary>
public sealed class ComparisonBlock
{
    public ComparisonBlock(string workbookId, string worksheetName, string address, int rowCount, int columnCount, IEnumerable<ComparisonCell> cells)
    {
        WorkbookId = Require(workbookId, nameof(workbookId));
        WorksheetName = Require(worksheetName, nameof(worksheetName));
        Address = Require(address, nameof(address));
        if (rowCount < 1) throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (columnCount < 1) throw new ArgumentOutOfRangeException(nameof(columnCount));
        RowCount = rowCount;
        ColumnCount = columnCount;
        Cells = (cells ?? throw new ArgumentNullException(nameof(cells))).ToArray();
        if (Cells.Count != rowCount * columnCount)
        {
            throw new ArgumentException(
                $"A {rowCount}x{columnCount} block requires {rowCount * columnCount} cells, not {Cells.Count}.", nameof(cells));
        }
    }

    public string WorkbookId { get; }
    public string WorksheetName { get; }
    public string Address { get; }
    public int RowCount { get; }
    public int ColumnCount { get; }

    /// <summary>Row-major, so index <c>row * ColumnCount + column</c>.</summary>
    public IReadOnlyList<ComparisonCell> Cells { get; }

    public string Identity => WorkbookId + "!" + WorksheetName + "!" + Address;

    private static string Require(string value, string name) =>
        string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("A value is required.", name) : value.Trim();
}

public sealed class CellDifference
{
    public CellDifference(
        int row,
        int column,
        DifferenceCategory category,
        ComparisonCell source,
        ComparisonCell target,
        FormulaDifferenceKind formulaKind,
        string statement)
    {
        Row = row;
        Column = column;
        Category = category;
        Source = source;
        Target = target;
        FormulaKind = formulaKind;
        Statement = statement;
    }

    /// <summary>Zero-based position within the compared rectangle.</summary>
    public int Row { get; }

    public int Column { get; }
    public DifferenceCategory Category { get; }
    public ComparisonCell Source { get; }
    public ComparisonCell Target { get; }
    public FormulaDifferenceKind FormulaKind { get; }

    /// <summary>A short description carrying no raw content of its own.</summary>
    public string Statement { get; }
}

public sealed class ComparisonResult
{
    public ComparisonResult(
        ComparisonBlock source,
        ComparisonBlock target,
        IReadOnlyList<CellDifference> differences,
        IReadOnlyDictionary<DifferenceCategory, int> counts,
        int comparedCells,
        int excludedByBound,
        IReadOnlyList<string> coverageGaps,
        string fingerprint)
    {
        Source = source;
        Target = target;
        Differences = differences;
        Counts = counts;
        ComparedCells = comparedCells;
        ExcludedByBound = excludedByBound;
        CoverageGaps = coverageGaps;
        Fingerprint = fingerprint;
    }

    public ComparisonBlock Source { get; }
    public ComparisonBlock Target { get; }
    public IReadOnlyList<CellDifference> Differences { get; }
    public IReadOnlyDictionary<DifferenceCategory, int> Counts { get; }
    public int ComparedCells { get; }

    /// <summary>Differences dropped because the result set hit its ceiling.</summary>
    public int ExcludedByBound { get; }

    public IReadOnlyList<string> CoverageGaps { get; }

    /// <summary>Covers both sides, so any edit to either invalidates the result.</summary>
    public string Fingerprint { get; }

    public bool IsComplete => ExcludedByBound == 0 && CoverageGaps.Count == 0;
}

/// <summary>Which categories a comparison examines.</summary>
public sealed class ComparisonCategories
{
    public ComparisonCategories(bool formulas = true, bool constants = true, bool numberFormats = false, bool displayedValues = false)
    {
        Formulas = formulas;
        Constants = constants;
        NumberFormats = numberFormats;
        DisplayedValues = displayedValues;
    }

    public bool Formulas { get; }
    public bool Constants { get; }
    public bool NumberFormats { get; }
    public bool DisplayedValues { get; }

    public static ComparisonCategories Default => new ComparisonCategories();

    /// <summary>
    /// The categories deliberately not examined, named so a result can never
    /// imply they were checked and found equal.
    /// </summary>
    public IReadOnlyList<string> Unexamined()
    {
        var unexamined = new List<string>();
        if (!Formulas) unexamined.Add("category_not_compared:formula");
        if (!Constants) unexamined.Add("category_not_compared:constant");
        if (!NumberFormats) unexamined.Add("category_not_compared:number_format");
        if (!DisplayedValues) unexamined.Add("category_not_compared:displayed_value");
        return unexamined;
    }
}

public static class SameShapeComparer
{
    public const int MaximumCells = 250_000;
    public const int MaximumDifferences = 5_000;

    /// <summary>
    /// Compares two blocks position by position. Shape mismatch is refused rather
    /// than aligned: there is no row matching, no shift detection, and no
    /// truncation to a common size.
    /// </summary>
    public static ComparisonResult Compare(ComparisonBlock source, ComparisonBlock target, ComparisonCategories? categories = null)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (source.RowCount != target.RowCount || source.ColumnCount != target.ColumnCount)
        {
            throw new ShapeMismatchException(
                $"The sources are {source.RowCount}x{source.ColumnCount} and {target.RowCount}x{target.ColumnCount}. "
                + "Same-shape comparison does not align or truncate; choose ranges of equal dimensions.");
        }

        if (source.Cells.Count > MaximumCells)
        {
            throw new ComparisonBoundException(
                $"The comparison covers {source.Cells.Count:N0} cells, beyond the qualified bound of {MaximumCells:N0}.");
        }

        var selected = categories ?? ComparisonCategories.Default;
        var counts = Enum.GetValues(typeof(DifferenceCategory)).Cast<DifferenceCategory>().ToDictionary(value => value, _ => 0);
        var differences = new List<CellDifference>();
        var gaps = new SortedSet<string>(StringComparer.Ordinal);
        var excluded = 0;

        for (var index = 0; index < source.Cells.Count; index++)
        {
            var left = source.Cells[index];
            var right = target.Cells[index];
            var difference = CompareCell(index / source.ColumnCount, index % source.ColumnCount, left, right, selected, gaps);
            if (difference is null) continue;

            counts[difference.Category]++;
            if (differences.Count < MaximumDifferences) differences.Add(difference);
            else excluded++;
        }

        foreach (var unexamined in selected.Unexamined()) gaps.Add(unexamined);
        if (excluded > 0) gaps.Add("difference_count_exceeded_qualified_bound");

        return new ComparisonResult(
            source, target, differences, counts, source.Cells.Count, excluded, gaps.ToArray(), Fingerprint(source, target));
    }

    private static CellDifference? CompareCell(
        int row,
        int column,
        ComparisonCell source,
        ComparisonCell target,
        ComparisonCategories categories,
        SortedSet<string> gaps)
    {
        if (source.IsFormula != target.IsFormula)
        {
            return new CellDifference(
                row, column, DifferenceCategory.TypeMismatch, source, target, FormulaDifferenceKind.Identical,
                source.IsFormula ? "Formula on the source, constant on the target." : "Constant on the source, formula on the target.");
        }

        if (source.IsFormula)
        {
            if (categories.Formulas)
            {
                var kind = CompareFormulas(source.Formula, target.Formula);
                if (kind == FormulaDifferenceKind.OutsideCoverage)
                {
                    gaps.Add("formula_outside_parser_coverage");
                }

                if (kind != FormulaDifferenceKind.Identical)
                {
                    return new CellDifference(row, column, DifferenceCategory.Formula, source, target, kind, StatementFor(kind));
                }
            }
        }
        else if (categories.Constants && !string.Equals(source.Value, target.Value, StringComparison.Ordinal))
        {
            return new CellDifference(
                row, column, DifferenceCategory.Constant, source, target, FormulaDifferenceKind.Identical,
                "The stored values differ.");
        }

        if (categories.NumberFormats && !string.Equals(source.NumberFormat, target.NumberFormat, StringComparison.Ordinal))
        {
            return new CellDifference(
                row, column, DifferenceCategory.NumberFormat, source, target, FormulaDifferenceKind.Identical,
                "The number formats differ.");
        }

        if (categories.DisplayedValues && !string.Equals(source.DisplayedValue, target.DisplayedValue, StringComparison.Ordinal))
        {
            return new CellDifference(
                row, column, DifferenceCategory.DisplayedValue, source, target, FormulaDifferenceKind.Identical,
                "The displayed values differ.");
        }

        return null;
    }

    /// <summary>
    /// Distinguishes identical text, a real structural difference, and two
    /// formulas that are the same shape from different anchors. Anything the
    /// parser does not cover is reported as such rather than guessed at.
    /// </summary>
    public static FormulaDifferenceKind CompareFormulas(string sourceFormula, string targetFormula)
    {
        if (string.Equals(sourceFormula, targetFormula, StringComparison.Ordinal))
        {
            return FormulaDifferenceKind.Identical;
        }

        var parser = new FormulaParser();
        var left = parser.Parse(sourceFormula, FormulaParseOptions.DefaultA1);
        var right = parser.Parse(targetFormula, FormulaParseOptions.DefaultA1);
        if (!left.IsSuccess || !right.IsSuccess || left.Document is null || right.Document is null)
        {
            return FormulaDifferenceKind.OutsideCoverage;
        }

        return SameTokenShape(left.Document, right.Document)
            ? FormulaDifferenceKind.EquivalentShape
            : FormulaDifferenceKind.Structure;
    }

    /// <summary>
    /// Two formulas share a shape when their meaningful tokens agree kind for
    /// kind, with references treated as interchangeable. That is what a formula
    /// copied to another anchor looks like.
    /// </summary>
    private static bool SameTokenShape(FormulaSyntaxDocument left, FormulaSyntaxDocument right)
    {
        var leftTokens = Meaningful(left);
        var rightTokens = Meaningful(right);
        if (leftTokens.Count != rightTokens.Count) return false;

        for (var index = 0; index < leftTokens.Count; index++)
        {
            var a = leftTokens[index];
            var b = rightTokens[index];
            if (a.Kind != b.Kind) return false;
            if (a.Kind == FormulaTokenKind.Reference) continue;
            if (!string.Equals(a.Text, b.Text, StringComparison.OrdinalIgnoreCase)) return false;
        }

        return true;
    }

    private static IReadOnlyList<FormulaToken> Meaningful(FormulaSyntaxDocument document) =>
        document.Tokens
            .Where(token => token.Kind != FormulaTokenKind.Whitespace && token.Kind != FormulaTokenKind.Prefix)
            .ToArray();

    private static string StatementFor(FormulaDifferenceKind kind)
    {
        switch (kind)
        {
            case FormulaDifferenceKind.Structure: return "The formulas differ in structure.";
            case FormulaDifferenceKind.EquivalentShape: return "Same shape, different references.";
            case FormulaDifferenceKind.OutsideCoverage: return "The formulas differ; at least one is outside qualified parser coverage.";
            default: return "The formulas are identical.";
        }
    }

    /// <summary>
    /// Digests both sides, so an edit to either invalidates the result. Hashing
    /// incrementally keeps this usable at the cell bound, where a concatenated
    /// string would not be.
    /// </summary>
    private static string Fingerprint(ComparisonBlock source, ComparisonBlock target)
    {
        using (var hash = SHA256.Create())
        {
            foreach (var block in new[] { source, target })
            {
                Feed(hash, block.Identity + "" + block.RowCount + "" + block.ColumnCount + "\n");
                foreach (var cell in block.Cells)
                {
                    Feed(hash, cell.Address + "" + cell.Formula + "" + cell.Value + "" + cell.NumberFormat + "\n");
                }
            }

            hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            var builder = new StringBuilder(64);
            foreach (var value in hash.Hash!) builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            return builder.ToString();
        }
    }

    private static void Feed(SHA256 hash, string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        hash.TransformBlock(bytes, 0, bytes.Length, null, 0);
    }
}

/// <summary>Raised when two sides are not the same shape. Never aligned around.</summary>
public sealed class ShapeMismatchException : Exception
{
    public ShapeMismatchException(string message) : base(message) { }
}

/// <summary>Raised when a comparison exceeds the qualified cell bound.</summary>
public sealed class ComparisonBoundException : Exception
{
    public ComparisonBoundException(string message) : base(message) { }
}
