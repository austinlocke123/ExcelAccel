using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// What a rule may read. It exposes the immutable snapshot and the rule's own
/// configuration, and nothing that could mutate a workbook.
/// </summary>
public sealed class ModelCheckContext
{
    public ModelCheckContext(ModelCheckSnapshot snapshot, ModelCheckConfiguration configuration)
    {
        Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
        Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
    }

    public ModelCheckSnapshot Snapshot { get; }

    public ModelCheckConfiguration Configuration { get; }
}

public interface IModelCheckRule
{
    string RuleId { get; }

    int Version { get; }

    ModelCheckSeverity DefaultSeverity { get; }

    /// <summary>A concise statement of what the rule looks for.</summary>
    string Statement { get; }

    /// <summary>
    /// Evaluates the rule. Must be deterministic for an identical snapshot and
    /// configuration, and must never mutate anything.
    /// </summary>
    ModelCheckRuleResult Evaluate(ModelCheckContext context);
}

public sealed class ModelCheckRuleResult
{
    public ModelCheckRuleResult(IEnumerable<ModelCheckFinding> findings, RuleCoverage coverage, int evaluatedCellCount)
    {
        Findings = Array.AsReadOnly((findings ?? throw new ArgumentNullException(nameof(findings))).ToArray());
        Coverage = coverage;
        EvaluatedCellCount = evaluatedCellCount;
    }

    public IReadOnlyList<ModelCheckFinding> Findings { get; }

    public RuleCoverage Coverage { get; }

    public int EvaluatedCellCount { get; }

    public static ModelCheckRuleResult Unsupported() =>
        new ModelCheckRuleResult(Array.Empty<ModelCheckFinding>(), RuleCoverage.Unsupported, 0);
}

public sealed class ModelCheckScanResult
{
    internal ModelCheckScanResult(
        AuditTraceStatus status,
        ModelCheckSnapshot snapshot,
        IReadOnlyList<string> enabledRuleIds,
        IEnumerable<ModelCheckFinding> findings,
        IEnumerable<ModelCheckRuleCoverage> ruleCoverage,
        IEnumerable<ModelCheckRuleFailure> ruleFailures,
        bool truncated,
        int suppressedFindingCount,
        string? refusalCode,
        string? message)
    {
        Status = status;
        Snapshot = snapshot;
        EnabledRuleIds = enabledRuleIds;
        Findings = Array.AsReadOnly(findings.ToArray());
        RuleCoverage = Array.AsReadOnly(ruleCoverage.ToArray());
        RuleFailures = Array.AsReadOnly(ruleFailures.ToArray());
        Truncated = truncated;
        SuppressedFindingCount = suppressedFindingCount;
        RefusalCode = refusalCode;
        Message = message;
    }

    public AuditTraceStatus Status { get; }

    public ModelCheckSnapshot Snapshot { get; }

    public IReadOnlyList<string> EnabledRuleIds { get; }

    /// <summary>Findings in canonical order.</summary>
    public IReadOnlyList<ModelCheckFinding> Findings { get; }

    public IReadOnlyList<ModelCheckRuleCoverage> RuleCoverage { get; }

    /// <summary>Rules that failed. A failure never silently omits a rule.</summary>
    public IReadOnlyList<ModelCheckRuleFailure> RuleFailures { get; }

    public bool Truncated { get; }

    /// <summary>Findings hidden by an active local ignore in this scan.</summary>
    public int SuppressedFindingCount { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public string ScanScope => Snapshot.ScopeLabel;

    /// <summary>
    /// A scan is complete only when every enabled rule ran to exact coverage
    /// with no failure and no truncation.
    /// </summary>
    public bool CanClaimCompleteness => Status == AuditTraceStatus.Complete;

    /// <summary>A scan that never ran, carrying its categorized reason.</summary>
    public static ModelCheckScanResult Refused(ModelCheckSnapshot snapshot, string code, string message) =>
        new ModelCheckScanResult(
            AuditTraceStatus.Refused,
            snapshot ?? throw new ArgumentNullException(nameof(snapshot)),
            Array.Empty<string>(),
            Array.Empty<ModelCheckFinding>(),
            Array.Empty<ModelCheckRuleCoverage>(),
            Array.Empty<ModelCheckRuleFailure>(),
            false,
            0,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A refusal code is required.", nameof(code)),
            message ?? string.Empty);
}

/// <summary>
/// Runs the enabled rules against one stable snapshot.
///
/// Rules run in a stable order and findings are sorted canonically, so
/// reordering execution cannot change the output. A rule that throws is
/// recorded as a failure and makes the scan incomplete; it is never dropped. A
/// cancelled scan is refused so a partial run is never presented as a completed
/// scan.
/// </summary>
public sealed class ModelCheckEngine
{
    public const int MaximumFindings = 5_000;

    public ModelCheckScanResult Run(
        ModelCheckSnapshot snapshot,
        IEnumerable<IModelCheckRule> rules,
        ModelCheckConfiguration? configuration = null,
        Func<bool>? cancellationRequested = null,
        Action<string>? onRuleStarted = null)
    {
        if (snapshot is null) throw new ArgumentNullException(nameof(snapshot));
        if (rules is null) throw new ArgumentNullException(nameof(rules));
        var effectiveConfiguration = configuration ?? ModelCheckConfiguration.Default;
        var ordered = rules
            .OrderBy(rule => rule.RuleId, StringComparer.Ordinal)
            .ThenBy(rule => rule.Version)
            .ToArray();
        var enabledIds = Array.AsReadOnly(ordered.Select(rule => rule.RuleId).ToArray());
        var context = new ModelCheckContext(snapshot, effectiveConfiguration);

        var findings = new List<ModelCheckFinding>();
        var coverage = new List<ModelCheckRuleCoverage>();
        var failures = new List<ModelCheckRuleFailure>();
        var truncated = false;
        var suppressed = 0;

        foreach (var rule in ordered)
        {
            if (cancellationRequested?.Invoke() == true)
            {
                return new ModelCheckScanResult(
                    AuditTraceStatus.Refused,
                    snapshot,
                    enabledIds,
                    Array.Empty<ModelCheckFinding>(),
                    Array.Empty<ModelCheckRuleCoverage>(),
                    Array.Empty<ModelCheckRuleFailure>(),
                    false,
                    0,
                    ModelCheckRefusalCodes.ScanCancelled,
                    "The scan was cancelled; prior results remain and no partial scan is reported as complete.");
            }

            onRuleStarted?.Invoke(rule.RuleId);
            try
            {
                var result = rule.Evaluate(context);
                var accepted = new List<ModelCheckFinding>();
                foreach (var finding in result.Findings)
                {
                    // A local ignore suppresses only findings whose rule-specific
                    // fingerprint matches exactly, so it can never hide a
                    // different rule or a different location.
                    if (effectiveConfiguration.IgnoredFingerprints.Contains(finding.Fingerprint))
                    {
                        suppressed++;
                        continue;
                    }

                    if (findings.Count + accepted.Count >= MaximumFindings)
                    {
                        truncated = true;
                        break;
                    }

                    accepted.Add(finding);
                }

                findings.AddRange(accepted);
                coverage.Add(new ModelCheckRuleCoverage(rule.RuleId, rule.Version, result.Coverage, result.EvaluatedCellCount, accepted.Count));
            }
            catch (Exception exception)
            {
                // A rule failure is reported with its rule ID and a safe
                // diagnostic. The scan becomes incomplete rather than quietly
                // losing a rule.
                failures.Add(new ModelCheckRuleFailure(rule.RuleId, rule.Version, exception.GetType().Name));
                coverage.Add(new ModelCheckRuleCoverage(rule.RuleId, rule.Version, RuleCoverage.Unsupported, 0, 0));
            }
        }

        findings.Sort((left, right) => string.Compare(left.CanonicalKey, right.CanonicalKey, StringComparison.Ordinal));

        var incomplete = truncated ||
            failures.Count != 0 ||
            coverage.Any(entry => entry.Coverage != ModelCheck.RuleCoverage.Exact);

        return new ModelCheckScanResult(
            incomplete ? AuditTraceStatus.Partial : AuditTraceStatus.Complete,
            snapshot,
            enabledIds,
            findings,
            coverage,
            failures,
            truncated,
            suppressed,
            null,
            incomplete ? "The scan is partial; inspect the reported rule coverage, failures, and truncation." : null);
    }
}

public static class ModelCheckRefusalCodes
{
    public const string ScanCancelled = "CHECK_SCAN_CANCELLED";
    public const string PreviewRequired = "CHECK_PREVIEW_REQUIRED";
    public const string ScopeTooLarge = "CHECK_SCOPE_TOO_LARGE";
    public const string StaleTarget = "CHECK_STALE_TARGET";
}
