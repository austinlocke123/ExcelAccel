using System;

namespace ExcelAccel.Core.Formulas;

public sealed class FormulaParseOptions
{
    public const int HardMaximumLength = 32768;
    public const int HardMaximumTokens = 16384;
    public const int HardMaximumNesting = 256;

    public FormulaParseOptions(
        FormulaDialect dialect,
        int maximumLength = 8192,
        int maximumTokens = 4096,
        int maximumNesting = 64)
    {
        Dialect = dialect ?? throw new ArgumentNullException(nameof(dialect));
        MaximumLength = RequireInRange(maximumLength, HardMaximumLength, nameof(maximumLength));
        MaximumTokens = RequireInRange(maximumTokens, HardMaximumTokens, nameof(maximumTokens));
        MaximumNesting = RequireInRange(maximumNesting, HardMaximumNesting, nameof(maximumNesting));
    }

    public FormulaDialect Dialect { get; }

    public int MaximumLength { get; }

    public int MaximumTokens { get; }

    public int MaximumNesting { get; }

    public static FormulaParseOptions DefaultA1 { get; } = new FormulaParseOptions(FormulaDialect.InvariantA1);

    private static int RequireInRange(int value, int hardMaximum, string parameterName)
    {
        if (value <= 0 || value > hardMaximum)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                $"The limit must be between 1 and {hardMaximum}.");
        }

        return value;
    }
}
