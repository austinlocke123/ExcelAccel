using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Application.Formatting;

public enum NumberFormatVerdict
{
    /// <summary>Nothing to say about this entry.</summary>
    Accepted,

    /// <summary>
    /// Structurally usable, but Excel is expected to rewrite it on assignment,
    /// which breaks the cycle that contains it.
    /// </summary>
    RewriteExpected,

    /// <summary>Not a usable number format at all.</summary>
    Rejected,
}

public sealed class NumberFormatDiagnostic
{
    public NumberFormatDiagnostic(NumberFormatVerdict verdict, string message, string? suggestion = null)
    {
        Verdict = verdict;
        Message = message ?? string.Empty;
        Suggestion = suggestion;
    }

    public NumberFormatVerdict Verdict { get; }
    public string Message { get; }
    public string? Suggestion { get; }

    public static readonly NumberFormatDiagnostic Accepted =
        new NumberFormatDiagnostic(NumberFormatVerdict.Accepted, string.Empty);
}

/// <summary>
/// Advisory checks on a number-format cycle entry.
/// </summary>
/// <remarks>
/// These are deliberately <b>not</b> enforced in <see cref="ProfileCycle"/>'s
/// constructor. Constructor validation also runs during schema migration, so
/// tightening it would make an existing profile containing a locale-qualified
/// currency format fail to load, and <c>ProfileRuntime</c> falls back to the
/// default profile when a load throws. A validation improvement that silently
/// discards a user's settings is worse than the problem it fixes.
///
/// The settings editor calls these when the user types an entry, which is the
/// point where a warning helps and nothing is lost by refusing.
/// </remarks>
public static class NumberFormatDiagnostics
{
    public const int MaximumLength = 256;
    public const int MaximumSections = 4;

    /// <summary>
    /// Everything decidable from the string alone. The round-trip question needs
    /// a live Excel and is answered by <see cref="EvaluateRoundTrip"/> once a
    /// probe has supplied what Excel stored.
    /// </summary>
    public static NumberFormatDiagnostic Inspect(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return Reject("A number format cannot be empty.");
        }

        var value = candidate!;
        if (value.Length > MaximumLength)
        {
            return Reject($"A number format may be at most {MaximumLength} characters.");
        }

        if (value[0] == '=')
        {
            return Reject("A number format is not a formula and cannot start with '='.");
        }

        if (value.Any(character => char.IsControl(character)))
        {
            return Reject("A number format cannot contain control characters.");
        }

        var structure = InspectStructure(value);
        if (structure is not null)
        {
            return structure;
        }

        var qualified = LocaleQualifiedCurrency(value);
        if (qualified is not null)
        {
            return qualified;
        }

        return NumberFormatDiagnostic.Accepted;
    }

    /// <summary>
    /// Compares what was written against what Excel stored back. A cycle finds its
    /// position by matching the cell's stored format against its entries, so an
    /// entry Excel rewrites can never match itself and its cycle sticks on the
    /// first entry forever.
    /// </summary>
    public static NumberFormatDiagnostic EvaluateRoundTrip(string candidate, string stored)
    {
        if (string.IsNullOrEmpty(candidate))
        {
            return Reject("A number format cannot be empty.");
        }

        if (string.Equals(candidate, stored, StringComparison.Ordinal))
        {
            return NumberFormatDiagnostic.Accepted;
        }

        return new NumberFormatDiagnostic(
            NumberFormatVerdict.RewriteExpected,
            $"Excel stores this format as '{stored}', not as written. A cycle matches the stored format, "
            + "so this entry would never match itself and the cycle stays on its first entry.",
            stored);
    }

    private static NumberFormatDiagnostic? InspectStructure(string value)
    {
        var inQuotes = false;
        var bracketDepth = 0;
        var sections = 1;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (character == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes)
            {
                continue;
            }

            switch (character)
            {
                case '[':
                    bracketDepth++;
                    break;
                case ']':
                    bracketDepth--;
                    if (bracketDepth < 0)
                    {
                        return Reject("A number format has a ']' with no matching '['.");
                    }

                    break;
                case ';':
                    sections++;
                    break;
                case '\\':
                    index++;
                    break;
            }
        }

        if (inQuotes)
        {
            return Reject("A number format has an unterminated quoted string.");
        }

        if (bracketDepth != 0)
        {
            return Reject("A number format has an unclosed '['.");
        }

        if (sections > MaximumSections)
        {
            return Reject($"A number format may have at most {MaximumSections} sections separated by ';'.");
        }

        return null;
    }

    /// <summary>
    /// The measured trap. Excel rewrites <c>[$£-en-GB]</c> to <c>[$£-809]</c> and
    /// <c>[$€-x-euro2]</c> to <c>[$€-2]</c> on assignment. These are exactly what
    /// Excel's own currency dialog produces, so a user reaches for them naturally.
    /// </summary>
    private static NumberFormatDiagnostic? LocaleQualifiedCurrency(string value)
    {
        var index = value.IndexOf("[$", StringComparison.Ordinal);
        if (index < 0)
        {
            return null;
        }

        var close = value.IndexOf(']', index);
        if (close < 0)
        {
            return null;
        }

        var token = value.Substring(index + 2, close - index - 2);
        var dash = token.IndexOf('-');
        if (dash < 0)
        {
            // [$€] with no locale is stored verbatim.
            return null;
        }

        var symbol = token.Substring(0, dash);
        return new NumberFormatDiagnostic(
            NumberFormatVerdict.RewriteExpected,
            $"Excel rewrites the locale-qualified currency token '[${token}]' when it stores this format, so a cycle "
            + "containing it would never match itself and stays on its first entry.",
            ReplaceEveryQualifiedToken(value, symbol));
    }

    /// <summary>
    /// Replaces every locale-qualified token, not just the first. A format
    /// normally carries one per section, and leaving the negative section
    /// qualified would suggest a replacement that is still broken.
    /// </summary>
    private static string? ReplaceEveryQualifiedToken(string value, string symbol)
    {
        if (string.IsNullOrEmpty(symbol))
        {
            return null;
        }

        var result = value;
        while (true)
        {
            var index = result.IndexOf("[$", StringComparison.Ordinal);
            if (index < 0)
            {
                return result;
            }

            var close = result.IndexOf(']', index);
            if (close < 0 || result.IndexOf('-', index, close - index) < 0)
            {
                return result;
            }

            result = result.Remove(index, close - index + 1).Insert(index, symbol);
        }
    }

    private static NumberFormatDiagnostic Reject(string message) =>
        new NumberFormatDiagnostic(NumberFormatVerdict.Rejected, message);
}
