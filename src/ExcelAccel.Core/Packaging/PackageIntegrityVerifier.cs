using System;
using System.IO;
using System.Security.Cryptography;

namespace ExcelAccel.Core.Packaging;

public sealed class PackageIntegrityResult
{
    internal PackageIntegrityResult(bool succeeded, string code, string message, string? resolvedPath)
    {
        Succeeded = succeeded;
        Code = code;
        Message = message;
        ResolvedPath = resolvedPath;
    }

    public bool Succeeded { get; }

    public string Code { get; }

    public string Message { get; }

    public string? ResolvedPath { get; }
}

public static class PackageIntegrityVerifier
{
    public const string Verified = "PACKAGE_VERIFIED";
    public const string RootMissing = "PACKAGE_ROOT_MISSING";
    public const string ArtifactMissing = "PACKAGE_ARTIFACT_MISSING";
    public const string ArtifactOutsideRoot = "PACKAGE_PATH_OUTSIDE_ROOT";
    public const string LengthMismatch = "PACKAGE_LENGTH_MISMATCH";
    public const string HashMismatch = "PACKAGE_HASH_MISMATCH";
    public const string ResourceLimit = "PACKAGE_RESOURCE_LIMIT";

    public static PackageIntegrityResult Verify(string packageRoot, PackageArtifact artifact, long maximumArtifactBytes)
    {
        if (string.IsNullOrWhiteSpace(packageRoot))
        {
            throw new ArgumentException("A package root is required.", nameof(packageRoot));
        }
        if (artifact is null)
        {
            throw new ArgumentNullException(nameof(artifact));
        }
        if (maximumArtifactBytes < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maximumArtifactBytes));
        }

        string root = Path.GetFullPath(packageRoot);
        if (!Directory.Exists(root))
        {
            return Failure(RootMissing, "The package root does not exist.");
        }

        string rootPrefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string artifactPath = Path.GetFullPath(Path.Combine(root, artifact.RelativePath));
        if (!artifactPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return Failure(ArtifactOutsideRoot, "The artifact resolves outside the package root.");
        }
        if (!File.Exists(artifactPath))
        {
            return Failure(ArtifactMissing, "The package artifact does not exist.");
        }

        var information = new FileInfo(artifactPath);
        if (information.Length > maximumArtifactBytes || artifact.Length > maximumArtifactBytes)
        {
            return Failure(ResourceLimit, "The package artifact exceeds the configured size limit.");
        }
        if (information.Length != artifact.Length)
        {
            return Failure(LengthMismatch, "The package artifact length does not match its manifest entry.");
        }

        string actualHash;
        using (var stream = new FileStream(artifactPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var sha256 = SHA256.Create())
        {
            actualHash = BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
        }

        if (!string.Equals(actualHash, artifact.Sha256, StringComparison.Ordinal))
        {
            return Failure(HashMismatch, "The package artifact SHA-256 does not match its manifest entry.");
        }

        return new PackageIntegrityResult(true, Verified, "The package artifact length and SHA-256 are valid.", artifactPath);
    }

    private static PackageIntegrityResult Failure(string code, string message) =>
        new PackageIntegrityResult(false, code, message, resolvedPath: null);
}
