using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Core.ModelCheck;

public enum ModelCheckSeverity
{
    Information,
    Advisory,
    Attention,
}

/// <summary>
/// How completely a rule could see what it needed. This is a coverage category,
/// never a probability that a finding is correct.
/// </summary>
public enum RuleCoverage
{
    Exact,
    Partial,
    Unsupported,
}

public enum ModelCheckScopeKind
{
    Selection,
    Worksheet,
    Workbook,
}

/// <summary>One immutable cell as a rule sees it. Captured once per scan.</summary>
public sealed class ModelCheckCell
{
    public ModelCheckCell(
        AuditCellIdentity identity,
        string? formula,
        AuditCellClassification classification,
        string numberFormat,
        string? errorText = null)
    {
        Identity = identity ?? throw new ArgumentNullException(nameof(identity));
        Formula = formula;
        Classification = classification;
        NumberFormat = numberFormat ?? string.Empty;
        ErrorText = errorText;
    }

    public AuditCellIdentity Identity { get; }

    /// <summary>The cell's formula, or null when it holds no formula.</summary>
    public string? Formula { get; }

    public AuditCellClassification Classification { get; }

    public string NumberFormat { get; }

    /// <summary>The Excel error text when the cell holds one, otherwise null.</summary>
    public string? ErrorText { get; }

    public bool IsFormula => !string.IsNullOrEmpty(Formula);

    public bool IsBlank => Classification == AuditCellClassification.Blank;
}

/// <summary>
/// The immutable data one scan runs against. Every enabled rule sees exactly
/// this, so an identical snapshot yields identical findings.
/// </summary>
public sealed class ModelCheckSnapshot
{
    public const int MaximumCells = 250_000;

    private readonly IReadOnlyDictionary<AuditCellIdentity, ModelCheckCell> _byIdentity;

    public ModelCheckSnapshot(ModelCheckScopeKind scope, string workbookId, IEnumerable<ModelCheckCell> cells)
    {
        Scope = scope;
        WorkbookId = !string.IsNullOrWhiteSpace(workbookId)
            ? workbookId
            : throw new ArgumentException("A workbook identity is required.", nameof(workbookId));
        var ordered = (cells ?? throw new ArgumentNullException(nameof(cells)))
            .OrderBy(cell => cell.Identity.WorksheetName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(cell => cell.Identity.Address, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (ordered.Length > MaximumCells)
        {
            throw new ArgumentOutOfRangeException(nameof(cells), $"A scan snapshot is limited to {MaximumCells:N0} cells.");
        }

        var map = new Dictionary<AuditCellIdentity, ModelCheckCell>();
        foreach (var cell in ordered)
        {
            if (map.ContainsKey(cell.Identity)) throw new ArgumentException($"Duplicate snapshot cell '{cell.Identity}'.", nameof(cells));
            map.Add(cell.Identity, cell);
        }

        Cells = Array.AsReadOnly(ordered);
        _byIdentity = new ReadOnlyDictionary<AuditCellIdentity, ModelCheckCell>(map);
    }

    public ModelCheckScopeKind Scope { get; }

    public string WorkbookId { get; }

    /// <summary>Cells in canonical worksheet-then-address order.</summary>
    public IReadOnlyList<ModelCheckCell> Cells { get; }

    public string ScopeLabel => Scope.ToString().ToLowerInvariant();

    public ModelCheckCell? Find(AuditCellIdentity identity) =>
        _byIdentity.TryGetValue(identity ?? throw new ArgumentNullException(nameof(identity)), out var cell) ? cell : null;
}

/// <summary>
/// One finding. It states a rule, a location, and evidence, and suggests
/// inspection. It never declares that a model is wrong and carries no score.
/// </summary>
public sealed class ModelCheckFinding
{
    public ModelCheckFinding(
        string ruleId,
        int ruleVersion,
        ModelCheckSeverity severity,
        AuditCellIdentity target,
        string statement,
        IEnumerable<string> evidence,
        RuleCoverage coverage,
        string fingerprint,
        string? peerContext = null)
    {
        RuleId = !string.IsNullOrWhiteSpace(ruleId) ? ruleId : throw new ArgumentException("A rule ID is required.", nameof(ruleId));
        RuleVersion = ruleVersion >= 1 ? ruleVersion : throw new ArgumentOutOfRangeException(nameof(ruleVersion));
        Severity = severity;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Statement = !string.IsNullOrWhiteSpace(statement) ? statement : throw new ArgumentException("A rule statement is required.", nameof(statement));
        Evidence = Array.AsReadOnly((evidence ?? throw new ArgumentNullException(nameof(evidence))).ToArray());
        Coverage = coverage;
        Fingerprint = !string.IsNullOrWhiteSpace(fingerprint) ? fingerprint : throw new ArgumentException("A fingerprint is required.", nameof(fingerprint));
        PeerContext = peerContext;
    }

    public string RuleId { get; }

    public int RuleVersion { get; }

    public ModelCheckSeverity Severity { get; }

    public AuditCellIdentity Target { get; }

    public string Statement { get; }

    public IReadOnlyList<string> Evidence { get; }

    /// <summary>A coverage category, not a correctness probability.</summary>
    public RuleCoverage Coverage { get; }

    /// <summary>
    /// Deterministic local fingerprint for ignore matching. It is a hash of
    /// normalized rule and location inputs and never contains raw formula or
    /// value content.
    /// </summary>
    public string Fingerprint { get; }

    public string? PeerContext { get; }

    /// <summary>Findings always point at a real cell in the scanned snapshot.</summary>
    public bool IsNavigable => true;

    /// <summary>The canonical sort key that makes a scan's output order stable.</summary>
    public string CanonicalKey =>
        RuleId + "|" + RuleVersion.ToString(CultureInfo.InvariantCulture) + "|" +
        Target.WorksheetName + "|" + SortableAddress(Target.Address) + "|" + Fingerprint;

    private static string SortableAddress(string address)
    {
        if (!AuditAddress.TryParse(address, out var rectangle)) return address;
        return rectangle.FirstColumn.ToString("D5", CultureInfo.InvariantCulture) + ":" +
            rectangle.FirstRow.ToString("D7", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Builds the deterministic fingerprint for a rule at a location. Callers
    /// pass normalized inputs only; raw formulas and values must never be given.
    /// </summary>
    public static string CreateFingerprint(string ruleId, int ruleVersion, AuditCellIdentity target, params string?[] normalizedInputs)
    {
        if (target is null) throw new ArgumentNullException(nameof(target));
        var components = new List<string?>
        {
            ruleId,
            ruleVersion.ToString(CultureInfo.InvariantCulture),
            target.WorkbookId,
            target.WorksheetName,
            target.Address,
        };
        components.AddRange(normalizedInputs ?? Array.Empty<string?>());
        return PreconditionFingerprint.Create(components.ToArray());
    }
}

/// <summary>A rule that could not complete. Never silently dropped.</summary>
public sealed class ModelCheckRuleFailure
{
    public ModelCheckRuleFailure(string ruleId, int ruleVersion, string diagnostic)
    {
        RuleId = ruleId;
        RuleVersion = ruleVersion;
        Diagnostic = diagnostic ?? string.Empty;
    }

    public string RuleId { get; }

    public int RuleVersion { get; }

    /// <summary>A safe diagnostic. It carries no workbook content.</summary>
    public string Diagnostic { get; }
}

/// <summary>Per-rule coverage across the scanned scope.</summary>
public sealed class ModelCheckRuleCoverage
{
    public ModelCheckRuleCoverage(string ruleId, int ruleVersion, RuleCoverage coverage, int evaluatedCellCount, int findingCount)
    {
        RuleId = ruleId;
        RuleVersion = ruleVersion;
        Coverage = coverage;
        EvaluatedCellCount = evaluatedCellCount;
        FindingCount = findingCount;
    }

    public string RuleId { get; }

    public int RuleVersion { get; }

    public RuleCoverage Coverage { get; }

    public int EvaluatedCellCount { get; }

    public int FindingCount { get; }
}
