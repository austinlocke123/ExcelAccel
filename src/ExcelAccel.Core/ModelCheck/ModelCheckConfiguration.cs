using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// Explicit rule configuration with approved defaults. Every threshold and
/// exclusion a rule uses is stated here, so no rule guesses intent.
/// </summary>
public sealed class ModelCheckConfiguration
{
    public ModelCheckConfiguration(
        int minimumPeerCount = 3,
        IEnumerable<double>? allowedEmbeddedLiterals = null,
        IEnumerable<string>? literalToleratingFunctions = null,
        IEnumerable<string>? ignoredFingerprints = null,
        bool treatBlanksAsPeerBreaks = true,
        int allowlistVersion = 1)
    {
        MinimumPeerCount = minimumPeerCount >= 2
            ? minimumPeerCount
            : throw new ArgumentOutOfRangeException(nameof(minimumPeerCount), "A peer region needs at least two cells.");
        AllowedEmbeddedLiterals = Array.AsReadOnly(
            (allowedEmbeddedLiterals ?? new[] { 0d, 1d, -1d, 2d, 100d, 12d, 365d })
            .Distinct()
            .OrderBy(value => value)
            .ToArray());
        LiteralToleratingFunctions = Array.AsReadOnly(
            (literalToleratingFunctions ?? new[] { "ROUND", "ROUNDUP", "ROUNDDOWN", "OFFSET", "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "LARGE", "SMALL" })
            .Select(value => value.ToUpperInvariant())
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray());
        IgnoredFingerprints = new HashSet<string>(ignoredFingerprints ?? Array.Empty<string>(), StringComparer.Ordinal);
        TreatBlanksAsPeerBreaks = treatBlanksAsPeerBreaks;
        AllowlistVersion = allowlistVersion >= 1 ? allowlistVersion : throw new ArgumentOutOfRangeException(nameof(allowlistVersion));
    }

    /// <summary>Smallest run of same-shaped formulas that counts as a peer region.</summary>
    public int MinimumPeerCount { get; }

    /// <summary>Versioned numeric literals that do not raise an embedded-hardcode finding.</summary>
    public IReadOnlyList<double> AllowedEmbeddedLiterals { get; }

    /// <summary>
    /// Functions whose arguments are structurally expected to be literals, such
    /// as a ROUND digit count. This is an explicit structural exclusion, not an
    /// inference about intent.
    /// </summary>
    public IReadOnlyList<string> LiteralToleratingFunctions { get; }

    /// <summary>Fingerprints suppressed by the local profile.</summary>
    public ISet<string> IgnoredFingerprints { get; }

    public bool TreatBlanksAsPeerBreaks { get; }

    public int AllowlistVersion { get; }

    public static ModelCheckConfiguration Default { get; } = new ModelCheckConfiguration();

    public ModelCheckConfiguration WithIgnoredFingerprints(IEnumerable<string> fingerprints) =>
        new ModelCheckConfiguration(
            MinimumPeerCount,
            AllowedEmbeddedLiterals,
            LiteralToleratingFunctions,
            fingerprints,
            TreatBlanksAsPeerBreaks,
            AllowlistVersion);

    public ModelCheckConfiguration WithAllowedEmbeddedLiterals(IEnumerable<double> literals, int allowlistVersion) =>
        new ModelCheckConfiguration(
            MinimumPeerCount,
            literals,
            LiteralToleratingFunctions,
            IgnoredFingerprints,
            TreatBlanksAsPeerBreaks,
            allowlistVersion);

    /// <summary>Canonical text of the settings that affect findings, for evidence.</summary>
    public string CanonicalDescription =>
        "peers>=" + MinimumPeerCount.ToString(CultureInfo.InvariantCulture) +
        ";allowlist_v" + AllowlistVersion.ToString(CultureInfo.InvariantCulture) +
        ";literals=" + string.Join(",", AllowedEmbeddedLiterals.Select(value => value.ToString("R", CultureInfo.InvariantCulture))) +
        ";blank_breaks=" + (TreatBlanksAsPeerBreaks ? "1" : "0");
}
