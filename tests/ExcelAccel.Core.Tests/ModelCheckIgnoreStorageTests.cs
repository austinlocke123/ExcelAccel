using System;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Persistence.ModelCheck;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

/// <summary>
/// Locks in a decision rather than a behaviour: Model Check ignores live in
/// their own atomic file, not in the profile document.
/// </summary>
/// <remarks>
/// The deciding reason is blast radius. `ProfileStore` rejects a profile it
/// cannot parse whole, and `ProfileRuntime` falls back to the embedded default
/// when a load throws, so folding ignores into the profile would mean a damaged
/// ignore list costs the user every setting they have — cycles, colours, quick
/// keys, favorites — to fix a suppression list.
///
/// It is also the wrong shape: an ignore is a fingerprint of a finding in one
/// model, while a profile is user-wide and portable. Carrying ignores into an
/// exported profile would move entries that can never match on the other side.
///
/// If a future change does fold them in, these tests fail, which is the point:
/// it should be a decision, not a drift.
/// </remarks>
public sealed class ModelCheckIgnoreStorageTests
{
    [Fact]
    public void TheProfileCarriesNoIgnoreData()
    {
        var serialized = new ProfileStore().Serialize(new ProfileStore().LoadDefault());

        Assert.DoesNotContain("ignore", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("fingerprint", serialized, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheProfileTypeExposesNoIgnoreMember() =>
        Assert.DoesNotContain(
            typeof(ProfileDefinition).GetProperties(),
            property => property.Name.IndexOf("ignore", StringComparison.OrdinalIgnoreCase) >= 0);

    /// <summary>
    /// The properties the contract needs are provided by the separate store, so
    /// nothing is given up by keeping it separate.
    /// </summary>
    [Fact]
    public void TheIgnoreStoreIsAtomicAndValidatesBeforeReplacing()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "model-check-ignores.tsv");
        var store = new ModelCheckIgnoreStore();
        var first = new ModelCheckIgnoreEntry("check.formula.inconsistent", 1, new string('a', 64), "row 12");
        var second = new ModelCheckIgnoreEntry("check.reference.external", 1, new string('b', 64));

        store.SaveAtomic(path, new[] { first });
        store.SaveAtomic(path, new[] { first, second });

        var loaded = store.Load(path);
        Assert.Equal(2, loaded.Count);
        Assert.Empty(Directory.GetFiles(sandbox.Path, "*.tmp"));
    }

    [Fact]
    public void ANoteNeverCarriesWorkbookContentIntoTheFile()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "model-check-ignores.tsv");
        var store = new ModelCheckIgnoreStore();

        // Tabs and newlines are the file's own separators; a note carrying them
        // would otherwise reshape the row it sits in.
        store.SaveAtomic(path, new[]
        {
            new ModelCheckIgnoreEntry("check.formula.inconsistent", 1, new string('c', 64), "note\twith\nseparators"),
        });

        var lines = File.ReadAllLines(path).Where(line => !line.StartsWith("#", StringComparison.Ordinal)).ToArray();
        var single = Assert.Single(lines);
        Assert.Equal(4, single.Split('\t').Length);
    }

    [Fact]
    public void ADuplicateFingerprintIsStoredOnce()
    {
        using var sandbox = new TemporaryDirectory();
        var path = Path.Combine(sandbox.Path, "model-check-ignores.tsv");
        var fingerprint = new string('d', 64);
        var store = new ModelCheckIgnoreStore();

        store.SaveAtomic(path, new[]
        {
            new ModelCheckIgnoreEntry("check.formula.inconsistent", 1, fingerprint),
            new ModelCheckIgnoreEntry("check.formula.inconsistent", 1, fingerprint, "again"),
        });

        Assert.Single(store.Load(path));
    }

    [Fact]
    public void AMissingFileIsAnEmptySetRatherThanAFailure() =>
        Assert.Empty(new ModelCheckIgnoreStore().Load(
            Path.Combine(Path.GetTempPath(), "excelaccel-no-such-" + Guid.NewGuid().ToString("N") + ".tsv")));

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "excelaccel-ignores-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
        }
    }
}
