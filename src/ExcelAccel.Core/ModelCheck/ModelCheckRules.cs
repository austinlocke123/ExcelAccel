using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// A formula that breaks the shape its neighbours share. The rule reports the
/// peer group, the baseline shape, and the exception; it never says the cell is
/// wrong.
/// </summary>
public sealed class PatternInconsistencyRule : IModelCheckRule
{
    public const string Id = "check.formula.pattern_inconsistency";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Attention;

    public string Statement => "A formula differs from the shape its peer region shares.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;
        var coverage = RuleCoverage.Exact;

        foreach (var region in PeerRegion.Build(context.Snapshot, context.Configuration))
        {
            var formulaCells = region.Cells.Where(cell => cell.IsFormula).ToArray();
            if (formulaCells.Length < context.Configuration.MinimumPeerCount) continue;
            evaluated += formulaCells.Length;

            var shapes = formulaCells.ToDictionary(cell => cell, cell => FormulaShape.TryNormalize(cell.Formula, cell.Identity));
            // A region the parser cannot fully cover can never be called
            // consistent, so it is reported as partial coverage instead.
            if (shapes.Values.Any(shape => shape is null))
            {
                coverage = RuleCoverage.Partial;
                continue;
            }

            var (baseline, baselineCount) = PeerRegion.Majority(shapes.Values);
            if (baseline is null || baselineCount < context.Configuration.MinimumPeerCount) continue;

            foreach (var cell in formulaCells)
            {
                if (string.Equals(shapes[cell], baseline, StringComparison.Ordinal)) continue;
                findings.Add(new ModelCheckFinding(
                    RuleId,
                    Version,
                    DefaultSeverity,
                    cell.Identity,
                    Statement,
                    new[]
                    {
                        "Peer region: " + region.Label,
                        "Peers sharing the baseline shape: " + baselineCount.ToString(CultureInfo.InvariantCulture),
                        "Baseline shape: " + baseline,
                        "This cell's shape: " + shapes[cell],
                    },
                    RuleCoverage.Exact,
                    ModelCheckFinding.CreateFingerprint(RuleId, Version, cell.Identity, region.Label, baseline, shapes[cell]),
                    region.Label));
            }
        }

        return new ModelCheckRuleResult(findings, coverage, evaluated);
    }
}

/// <summary>
/// A constant sitting inside an otherwise formula-consistent region. Reported
/// separately from a numeric literal embedded inside a formula.
/// </summary>
public sealed class ConstantInterruptsRegionRule : IModelCheckRule
{
    public const string Id = "check.constant.interrupts_formula_region";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Attention;

    public string Statement => "A constant interrupts a region of otherwise consistent formulas.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;

        foreach (var region in PeerRegion.Build(context.Snapshot, context.Configuration))
        {
            var formulaCells = region.Cells.Where(cell => cell.IsFormula).ToArray();
            var constants = region.Cells
                .Where(cell => !cell.IsFormula && !cell.IsBlank && cell.Classification != AuditCellClassification.Unknown)
                .ToArray();
            if (formulaCells.Length < context.Configuration.MinimumPeerCount || constants.Length == 0) continue;
            evaluated += region.Cells.Count;

            var shapes = formulaCells.Select(cell => FormulaShape.TryNormalize(cell.Formula, cell.Identity)).ToArray();
            var (baseline, baselineCount) = PeerRegion.Majority(shapes);
            if (baseline is null || baselineCount < context.Configuration.MinimumPeerCount) continue;

            foreach (var constant in constants)
            {
                findings.Add(new ModelCheckFinding(
                    RuleId,
                    Version,
                    DefaultSeverity,
                    constant.Identity,
                    Statement,
                    new[]
                    {
                        "Peer region: " + region.Label,
                        "Surrounding formulas sharing a shape: " + baselineCount.ToString(CultureInfo.InvariantCulture),
                        "Baseline shape: " + baseline,
                        "This cell holds a constant of kind: " + constant.Classification,
                    },
                    RuleCoverage.Exact,
                    ModelCheckFinding.CreateFingerprint(RuleId, Version, constant.Identity, region.Label, baseline),
                    region.Label));
            }
        }

        return new ModelCheckRuleResult(findings, RuleCoverage.Exact, evaluated);
    }
}

/// <summary>
/// A numeric literal written inside a formula. Allowlisted values and literals
/// in structurally literal-taking functions are excluded, both by explicit
/// configuration rather than inference.
/// </summary>
public sealed class EmbeddedNumericConstantRule : IModelCheckRule
{
    public const string Id = "check.formula.embedded_numeric_constant";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Advisory;

    public string Statement => "A formula embeds a numeric constant.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;

        foreach (var cell in context.Snapshot.Cells.Where(value => value.IsFormula))
        {
            evaluated++;
            foreach (var literal in FormulaShape.ReadEmbeddedLiterals(cell.Formula))
            {
                if (context.Configuration.AllowedEmbeddedLiterals.Contains(literal.Value)) continue;
                if (context.Configuration.LiteralToleratingFunctions.Contains(literal.EnclosingFunction)) continue;

                var span = literal.Span.Start.ToString(CultureInfo.InvariantCulture) + "+" +
                    literal.Span.Length.ToString(CultureInfo.InvariantCulture);
                findings.Add(new ModelCheckFinding(
                    RuleId,
                    Version,
                    DefaultSeverity,
                    cell.Identity,
                    Statement,
                    new[]
                    {
                        "Literal: " + literal.SourceText,
                        "Source span: [" + span + "]",
                        "Enclosing function: " + (literal.EnclosingFunction.Length == 0 ? "(none)" : literal.EnclosingFunction),
                        "Allowlist version: " + context.Configuration.AllowlistVersion.ToString(CultureInfo.InvariantCulture),
                    },
                    RuleCoverage.Exact,
                    ModelCheckFinding.CreateFingerprint(
                        RuleId,
                        Version,
                        cell.Identity,
                        literal.Value.ToString("R", CultureInfo.InvariantCulture),
                        span,
                        literal.EnclosingFunction)));
            }
        }

        return new ModelCheckRuleResult(findings, RuleCoverage.Exact, evaluated);
    }
}

/// <summary>
/// Cells holding an Excel error, and formulas carrying a broken reference. The
/// rule reads the captured snapshot and never recalculates.
/// </summary>
public sealed class FormulaErrorRule : IModelCheckRule
{
    public const string Id = "check.formula.error";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Attention;

    public string Statement => "A cell holds an Excel error value or a broken reference.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;

        foreach (var cell in context.Snapshot.Cells)
        {
            evaluated++;
            var brokenReference = cell.IsFormula && cell.Formula!.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0;
            var hasError = cell.Classification == AuditCellClassification.Error || !string.IsNullOrEmpty(cell.ErrorText);
            if (!brokenReference && !hasError) continue;

            var kind = brokenReference ? "#REF!" : (cell.ErrorText ?? "error");
            findings.Add(new ModelCheckFinding(
                RuleId,
                Version,
                DefaultSeverity,
                cell.Identity,
                Statement,
                new[]
                {
                    "Error kind: " + kind,
                    "Source: " + (brokenReference ? "broken reference in the formula" : "captured cell value"),
                    "The workbook was not recalculated for this rule.",
                },
                RuleCoverage.Exact,
                ModelCheckFinding.CreateFingerprint(RuleId, Version, cell.Identity, kind)));
        }

        return new ModelCheckRuleResult(findings, RuleCoverage.Exact, evaluated);
    }
}

/// <summary>
/// Formulas referencing another workbook. The source is never opened, contacted,
/// or refreshed.
/// </summary>
public sealed class ExternalReferenceRule : IModelCheckRule
{
    public const string Id = "check.reference.external";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Advisory;

    public string Statement => "A formula references another workbook.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;
        var parser = new FormulaParser();

        foreach (var cell in context.Snapshot.Cells.Where(value => value.IsFormula))
        {
            evaluated++;
            var parse = parser.Parse(cell.Formula!, new FormulaParseOptions(FormulaDialect.InvariantA1));
            if (!parse.IsSuccess) continue;
            var external = parse.Document!.References
                .Where(reference => reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0)
                .ToArray();
            if (external.Length == 0) continue;

            var qualifiers = external
                .Select(reference => reference.Qualifier!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            findings.Add(new ModelCheckFinding(
                RuleId,
                Version,
                DefaultSeverity,
                cell.Identity,
                Statement,
                new[]
                {
                    "External sources: " + string.Join(", ", qualifiers),
                    "External reference count: " + external.Length.ToString(CultureInfo.InvariantCulture),
                    "No external workbook was opened, contacted, or refreshed.",
                },
                RuleCoverage.Exact,
                ModelCheckFinding.CreateFingerprint(RuleId, Version, cell.Identity, string.Join("|", qualifiers))));
        }

        return new ModelCheckRuleResult(findings, RuleCoverage.Exact, evaluated);
    }
}

/// <summary>
/// Cycles among the formulas captured in the snapshot, found by bounded graph
/// traversal. Iterative-calculation settings are never read or changed, and the
/// workbook is never recalculated.
/// </summary>
public sealed class CircularReferenceRule : IModelCheckRule
{
    public const string Id = "check.reference.circular";
    public const int MaximumGraphEdges = 50_000;

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Attention;

    public string Statement => "A formula participates in a reference cycle.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var analyzer = new DirectPrecedentAnalyzer();
        var edges = new Dictionary<AuditCellIdentity, List<AuditCellIdentity>>();
        var coverage = RuleCoverage.Exact;
        var evaluated = 0;
        var edgeCount = 0;

        foreach (var cell in context.Snapshot.Cells.Where(value => value.IsFormula))
        {
            evaluated++;
            var plan = analyzer.CreateCapturePlan(cell.Identity, cell.Formula!);
            if (plan.NameCandidates.Count != 0) coverage = RuleCoverage.Partial;
            var targets = new List<AuditCellIdentity>();
            foreach (var target in plan.LocalTargets)
            {
                if (edgeCount >= MaximumGraphEdges)
                {
                    coverage = RuleCoverage.Partial;
                    break;
                }

                if (context.Snapshot.Find(target) is null && !SpansSnapshotCell(context.Snapshot, target, out _)) continue;
                targets.Add(target);
                edgeCount++;
            }

            edges[cell.Identity] = targets;
        }

        var findings = new List<ModelCheckFinding>();
        foreach (var cycle in FindCycles(edges))
        {
            var label = string.Join(" -> ", cycle.Select(AuditPresentationLabels.Location));
            foreach (var member in cycle)
            {
                findings.Add(new ModelCheckFinding(
                    RuleId,
                    Version,
                    DefaultSeverity,
                    member,
                    Statement,
                    new[]
                    {
                        "Cycle: " + label,
                        "Cycle length: " + cycle.Count.ToString(CultureInfo.InvariantCulture),
                        "Iterative calculation settings were not read or changed.",
                    },
                    coverage,
                    ModelCheckFinding.CreateFingerprint(RuleId, Version, member, label),
                    label));
            }
        }

        return new ModelCheckRuleResult(findings, coverage, evaluated);
    }

    private static bool SpansSnapshotCell(ModelCheckSnapshot snapshot, AuditCellIdentity target, out AuditCellIdentity? match)
    {
        match = null;
        if (!AuditAddress.TryParse(target.Address, out var rectangle) || rectangle.IsSingleCell) return false;
        foreach (var cell in snapshot.Cells)
        {
            if (!string.Equals(cell.Identity.WorksheetName, target.WorksheetName, StringComparison.OrdinalIgnoreCase)) continue;
            if (!AuditAddress.TryParse(cell.Identity.Address, out var candidate)) continue;
            if (!candidate.Intersects(rectangle)) continue;
            match = cell.Identity;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Iterative depth-first search. Each cycle is reported once, keyed by its
    /// canonical member set, so an identical snapshot yields identical findings.
    /// </summary>
    private static IReadOnlyList<IReadOnlyList<AuditCellIdentity>> FindCycles(
        IReadOnlyDictionary<AuditCellIdentity, List<AuditCellIdentity>> edges)
    {
        var cycles = new List<IReadOnlyList<AuditCellIdentity>>();
        var seenCycleKeys = new HashSet<string>(StringComparer.Ordinal);
        var permanently = new HashSet<AuditCellIdentity>();

        foreach (var start in edges.Keys.OrderBy(key => key.ToString(), StringComparer.OrdinalIgnoreCase))
        {
            if (permanently.Contains(start)) continue;
            var path = new List<AuditCellIdentity>();
            var onPath = new HashSet<AuditCellIdentity>();
            var stack = new Stack<(AuditCellIdentity Node, int NextEdge)>();
            stack.Push((start, 0));
            path.Add(start);
            onPath.Add(start);

            while (stack.Count != 0)
            {
                var (node, nextEdge) = stack.Pop();
                if (!edges.TryGetValue(node, out var targets) || nextEdge >= targets.Count)
                {
                    permanently.Add(node);
                    if (path.Count != 0) path.RemoveAt(path.Count - 1);
                    onPath.Remove(node);
                    continue;
                }

                stack.Push((node, nextEdge + 1));
                var next = targets[nextEdge];
                if (onPath.Contains(next))
                {
                    var startIndex = path.IndexOf(next);
                    if (startIndex >= 0)
                    {
                        var cycle = path.Skip(startIndex).ToArray();
                        var key = string.Join("|", cycle.Select(value => value.ToString()).OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
                        if (seenCycleKeys.Add(key)) cycles.Add(cycle);
                    }

                    continue;
                }

                if (permanently.Contains(next) || !edges.ContainsKey(next)) continue;
                stack.Push((next, 0));
                path.Add(next);
                onPath.Add(next);
            }
        }

        return cycles;
    }
}

/// <summary>
/// A number format that differs from the one its peer region shares. The rule
/// reports the baseline and the exception and never changes formatting.
/// </summary>
public sealed class NumberFormatInconsistencyRule : IModelCheckRule
{
    public const string Id = "check.format.number_inconsistency";

    public string RuleId => Id;

    public int Version => 1;

    public ModelCheckSeverity DefaultSeverity => ModelCheckSeverity.Advisory;

    public string Statement => "A number format differs from the one its peer region shares.";

    public ModelCheckRuleResult Evaluate(ModelCheckContext context)
    {
        var findings = new List<ModelCheckFinding>();
        var evaluated = 0;

        foreach (var region in PeerRegion.Build(context.Snapshot, context.Configuration))
        {
            var cells = region.Cells.Where(cell => !cell.IsBlank).ToArray();
            if (cells.Length < context.Configuration.MinimumPeerCount) continue;
            evaluated += cells.Length;

            var formats = cells.Select(cell => Normalize(cell.NumberFormat)).ToArray();
            var (baseline, baselineCount) = PeerRegion.Majority(formats);
            if (baseline is null || baselineCount < context.Configuration.MinimumPeerCount) continue;

            for (var index = 0; index < cells.Length; index++)
            {
                if (string.Equals(formats[index], baseline, StringComparison.Ordinal)) continue;
                findings.Add(new ModelCheckFinding(
                    RuleId,
                    Version,
                    DefaultSeverity,
                    cells[index].Identity,
                    Statement,
                    new[]
                    {
                        "Peer region: " + region.Label,
                        "Peers sharing the baseline format: " + baselineCount.ToString(CultureInfo.InvariantCulture),
                        "Baseline format: " + baseline,
                        "This cell's format: " + formats[index],
                    },
                    RuleCoverage.Exact,
                    ModelCheckFinding.CreateFingerprint(RuleId, Version, cells[index].Identity, region.Label, baseline, formats[index]),
                    region.Label));
            }
        }

        return new ModelCheckRuleResult(findings, RuleCoverage.Exact, evaluated);
    }

    /// <summary>
    /// Normalized format identity: case and surrounding whitespace do not make
    /// two formats different. Values and formulas never affect equivalence.
    /// </summary>
    private static string Normalize(string numberFormat) =>
        string.IsNullOrWhiteSpace(numberFormat) ? "General" : numberFormat.Trim().ToUpperInvariant();
}

public static class ModelCheckRuleCatalog
{
    public static IReadOnlyList<IModelCheckRule> All { get; } = new IModelCheckRule[]
    {
        new PatternInconsistencyRule(),
        new ConstantInterruptsRegionRule(),
        new EmbeddedNumericConstantRule(),
        new FormulaErrorRule(),
        new ExternalReferenceRule(),
        new CircularReferenceRule(),
        new NumberFormatInconsistencyRule(),
    };

    public static IReadOnlyList<string> AllRuleIds { get; } =
        All.Select(rule => rule.RuleId).OrderBy(value => value, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<IModelCheckRule> Select(IEnumerable<string>? ruleIds)
    {
        if (ruleIds is null) return All;
        var wanted = new HashSet<string>(ruleIds, StringComparer.Ordinal);
        return All.Where(rule => wanted.Contains(rule.RuleId)).ToArray();
    }
}
