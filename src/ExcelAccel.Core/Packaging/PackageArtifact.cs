using System;
using System.IO;
using System.Linq;

namespace ExcelAccel.Core.Packaging;

public sealed class PackageArtifact
{
    public PackageArtifact(string relativePath, long length, string sha256)
    {
        RelativePath = NormalizeRelativePath(relativePath);
        if (length < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(length));
        }

        if (sha256 is null)
        {
            throw new ArgumentNullException(nameof(sha256));
        }

        string normalizedHash = sha256.Trim().ToUpperInvariant();
        if (normalizedHash.Length != 64 || normalizedHash.Any(character => !IsHex(character)))
        {
            throw new ArgumentException("SHA-256 must contain exactly 64 hexadecimal characters.", nameof(sha256));
        }

        Length = length;
        Sha256 = normalizedHash;
    }

    public string RelativePath { get; }

    public long Length { get; }

    public string Sha256 { get; }

    private static string NormalizeRelativePath(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("A relative artifact path is required.", nameof(value));
        }

        string normalized = value.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        if (Path.IsPathRooted(normalized))
        {
            throw new ArgumentException("Artifact paths must be relative.", nameof(value));
        }

        string[] components = normalized.Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.None);
        if (components.Any(component =>
                string.IsNullOrWhiteSpace(component) ||
                string.Equals(component, ".", StringComparison.Ordinal) ||
                string.Equals(component, "..", StringComparison.Ordinal)))
        {
            throw new ArgumentException("Artifact paths cannot contain empty, current, or parent components.", nameof(value));
        }

        return string.Join(Path.DirectorySeparatorChar.ToString(), components);
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9') ||
        (value >= 'A' && value <= 'F');
}
