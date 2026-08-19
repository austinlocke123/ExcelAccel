using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAccel.Core.Commands;

public static class PreconditionFingerprint
{
    public const int MaximumComponentCount = 4096;
    public const int MaximumTotalCharacters = 1_000_000;

    public static string Create(params string?[] components)
    {
        if (components is null)
        {
            throw new ArgumentNullException(nameof(components));
        }

        if (components.Length > MaximumComponentCount)
        {
            throw new ArgumentOutOfRangeException(nameof(components), "The fingerprint component limit was exceeded.");
        }

        var canonical = new StringBuilder();
        var totalCharacters = 0;
        foreach (var component in components)
        {
            if (component is null)
            {
                canonical.Append("-1:");
            }
            else
            {
                totalCharacters = checked(totalCharacters + component.Length);
                if (totalCharacters > MaximumTotalCharacters)
                {
                    throw new ArgumentOutOfRangeException(nameof(components), "The fingerprint character limit was exceeded.");
                }

                canonical.Append(component.Length.ToString(CultureInfo.InvariantCulture));
                canonical.Append(':');
                canonical.Append(component);
            }

            canonical.Append('|');
        }

        var bytes = Encoding.UTF8.GetBytes(canonical.ToString());
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(bytes);
        var hex = new StringBuilder(hash.Length * 2);
        foreach (var value in hash)
        {
            hex.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        }

        return hex.ToString();
    }
}
