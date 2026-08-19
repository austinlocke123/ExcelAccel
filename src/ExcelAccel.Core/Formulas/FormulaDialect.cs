using System;

namespace ExcelAccel.Core.Formulas;

public enum FormulaNotation
{
    A1,
    R1C1
}

public sealed class FormulaDialect
{
    public FormulaDialect(
        FormulaNotation notation,
        char listSeparator = ',',
        char decimalSeparator = '.')
    {
        if (listSeparator != ',' && listSeparator != ';')
        {
            throw new ArgumentOutOfRangeException(nameof(listSeparator), "Only comma and semicolon list separators are qualified.");
        }

        if (decimalSeparator != '.' && decimalSeparator != ',')
        {
            throw new ArgumentOutOfRangeException(nameof(decimalSeparator), "Only period and comma decimal separators are qualified.");
        }

        if (listSeparator == decimalSeparator)
        {
            throw new ArgumentException("List and decimal separators must differ.");
        }

        Notation = notation;
        ListSeparator = listSeparator;
        DecimalSeparator = decimalSeparator;
    }

    public FormulaNotation Notation { get; }

    public char ListSeparator { get; }

    public char DecimalSeparator { get; }

    public static FormulaDialect InvariantA1 { get; } = new FormulaDialect(FormulaNotation.A1);

    public static FormulaDialect InvariantR1C1 { get; } = new FormulaDialect(FormulaNotation.R1C1);
}
