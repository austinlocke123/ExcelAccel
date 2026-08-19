using System;
using System.IO;
using System.Security.Cryptography;
using ExcelAccel.Core.Packaging;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class PackageIntegrityTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "excelaccel-package-tests", Guid.NewGuid().ToString("N"));

    [Fact]
    public void VerifyAcceptsExactLengthAndHash()
    {
        var artifact = WriteArtifact("payload/ExcelAccel.xll", new byte[] { 1, 2, 3 });

        var result = PackageIntegrityVerifier.Verify(_root, artifact, maximumArtifactBytes: 1024);

        Assert.True(result.Succeeded);
        Assert.Equal(PackageIntegrityVerifier.Verified, result.Code);
        Assert.Equal(Path.Combine(_root, "payload", "ExcelAccel.xll"), result.ResolvedPath);
    }

    [Fact]
    public void VerifyRejectsChangedContentWithSameLength()
    {
        var artifact = WriteArtifact("payload/ExcelAccel.xll", new byte[] { 1, 2, 3 });
        File.WriteAllBytes(Path.Combine(_root, "payload", "ExcelAccel.xll"), new byte[] { 3, 2, 1 });

        var result = PackageIntegrityVerifier.Verify(_root, artifact, maximumArtifactBytes: 1024);

        Assert.False(result.Succeeded);
        Assert.Equal(PackageIntegrityVerifier.HashMismatch, result.Code);
    }

    [Fact]
    public void VerifyRejectsChangedLengthBeforeHashing()
    {
        var artifact = WriteArtifact("ExcelAccel.xll", new byte[] { 1, 2, 3 });
        File.WriteAllBytes(Path.Combine(_root, "ExcelAccel.xll"), new byte[] { 1, 2, 3, 4 });

        var result = PackageIntegrityVerifier.Verify(_root, artifact, maximumArtifactBytes: 1024);

        Assert.Equal(PackageIntegrityVerifier.LengthMismatch, result.Code);
    }

    [Fact]
    public void VerifyRejectsResourceLimit()
    {
        var artifact = WriteArtifact("ExcelAccel.xll", new byte[] { 1, 2, 3 });

        var result = PackageIntegrityVerifier.Verify(_root, artifact, maximumArtifactBytes: 2);

        Assert.Equal(PackageIntegrityVerifier.ResourceLimit, result.Code);
    }

    [Theory]
    [InlineData("../escape.xll")]
    [InlineData("payload/../escape.xll")]
    [InlineData("/rooted.xll")]
    [InlineData("payload//file.xll")]
    [InlineData("payload/./file.xll")]
    public void ArtifactRejectsUnsafePath(string path)
    {
        Assert.Throws<ArgumentException>(() => new PackageArtifact(path, 1, new string('A', 64)));
    }

    [Theory]
    [InlineData("")]
    [InlineData("A")]
    [InlineData("GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG")]
    public void ArtifactRejectsInvalidHash(string hash)
    {
        Assert.ThrowsAny<ArgumentException>(() => new PackageArtifact("ExcelAccel.xll", 1, hash));
    }

    [Fact]
    public void VerifyReportsMissingRootAndArtifact()
    {
        var artifact = new PackageArtifact("ExcelAccel.xll", 1, new string('A', 64));
        Assert.Equal(
            PackageIntegrityVerifier.RootMissing,
            PackageIntegrityVerifier.Verify(_root, artifact, 1024).Code);

        Directory.CreateDirectory(_root);
        Assert.Equal(
            PackageIntegrityVerifier.ArtifactMissing,
            PackageIntegrityVerifier.Verify(_root, artifact, 1024).Code);
    }

    public void Dispose()
    {
        if (Directory.Exists(_root))
        {
            Directory.Delete(_root, recursive: true);
        }
    }

    private PackageArtifact WriteArtifact(string relativePath, byte[] content)
    {
        string path = Path.Combine(_root, relativePath.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, content);
        using var sha256 = SHA256.Create();
        string hash = BitConverter.ToString(sha256.ComputeHash(content)).Replace("-", string.Empty);
        return new PackageArtifact(relativePath, content.Length, hash);
    }
}
