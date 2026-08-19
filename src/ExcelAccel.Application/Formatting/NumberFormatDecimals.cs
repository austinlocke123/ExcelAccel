using System;
using System.Text.RegularExpressions;

namespace ExcelAccel.Application.Formatting;

public static class NumberFormatDecimals
{
    private static readonly Regex NumericPattern = new Regex("(?<whole>[0#?](?:[0#?,]*[0#?])?)(?:\\.(?<decimal>[0#?]+))?", RegexOptions.CultureInvariant);

    public static string Change(string format, int delta)
    {
        if (string.IsNullOrWhiteSpace(format) || (delta != 1 && delta != -1) || format.IndexOfAny(new[] { '[', ']', 'E', 'e', '/' }) >= 0)
        {
            return string.Empty;
        }

        var changed = false;
        var result = NumericPattern.Replace(format, match =>
        {
            var decimals = match.Groups["decimal"].Value;
            if (delta > 0)
            {
                changed = true;
                return match.Groups["whole"].Value + "." + decimals + "0";
            }

            if (decimals.Length == 0)
            {
                return match.Value;
            }

            changed = true;
            return match.Groups["whole"].Value + (decimals.Length == 1 ? string.Empty : "." + decimals.Substring(0, decimals.Length - 1));
        });
        return changed ? result : string.Empty;
    }
}
