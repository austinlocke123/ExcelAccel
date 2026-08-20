using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

public enum DependentScanScopeKind
{
    Worksheet,
    Workbook,
}

/// <summary>
/// The explicitly declared region a dependent scan may read. A scan never
/// widens its own scope; an unsupported scope is refused with a stable code.
/// </summary>
public sealed class DependentScanScope
{
    private DependentScanScope(DependentScanScopeKind kind, string workbookId, string? worksheetName)
    {
        Kind = kind;
        WorkbookId = !string.IsNullOrWhiteSpace(workbookId)
            ? workbookId
            : throw new ArgumentException("A workbook identity is required.", nameof(workbookId));
        WorksheetName = worksheetName;
    }

    public DependentScanScopeKind Kind { get; }

    public string WorkbookId { get; }

    public string? WorksheetName { get; }

    public string Label => Kind == DependentScanScopeKind.Worksheet ? "worksheet" : "workbook";

    public static DependentScanScope Worksheet(string workbookId, string worksheetName) =>
        new DependentScanScope(
            DependentScanScopeKind.Worksheet,
            workbookId,
            !string.IsNullOrWhiteSpace(worksheetName)
                ? worksheetName
                : throw new ArgumentException("A worksheet name is required.", nameof(worksheetName)));

    /// <summary>Every worksheet of one workbook.</summary>
    public static DependentScanScope Workbook(string workbookId) =>
        new DependentScanScope(DependentScanScopeKind.Workbook, workbookId, null);
}

public sealed class AuditFormulaCell
{
    public AuditFormulaCell(AuditCellIdentity identity, string formula)
    {
        Identity = identity ?? throw new ArgumentNullException(nameof(identity));
        Formula = formula ?? throw new ArgumentNullException(nameof(formula));
    }

    public AuditCellIdentity Identity { get; }

    public string Formula { get; }
}

public sealed class DirectDependent
{
    internal DirectDependent(AuditCellIdentity dependent, IEnumerable<AuditReferenceEvidence> evidence)
    {
        Dependent = dependent;
        Evidence = Array.AsReadOnly(evidence.ToArray());
        if (Evidence.Count == 0) throw new ArgumentException("At least one source edge is required.", nameof(evidence));
    }

    public AuditCellIdentity Dependent { get; }

    public IReadOnlyList<AuditReferenceEvidence> Evidence { get; }

    public int Depth => 1;
}

public sealed class DirectDependentResult
{
    internal DirectDependentResult(
        AuditTraceStatus status,
        AuditCellIdentity target,
        DependentScanScope scope,
        IEnumerable<DirectDependent> dependents,
        int scannedFormulaCount,
        int coverageGapCount,
        bool truncated,
        string? limitationCode,
        string? refusalCode,
        string? message)
    {
        Status = status;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Scope = scope ?? throw new ArgumentNullException(nameof(scope));
        Dependents = Array.AsReadOnly((dependents ?? throw new ArgumentNullException(nameof(dependents))).ToArray());
        ScannedFormulaCount = scannedFormulaCount;
        CoverageGapCount = coverageGapCount;
        Truncated = truncated;
        LimitationCode = limitationCode;
        RefusalCode = refusalCode;
        Message = message;
    }

    public AuditTraceStatus Status { get; }

    public AuditCellIdentity Target { get; }

    public DependentScanScope Scope { get; }

    public IReadOnlyList<DirectDependent> Dependents { get; }

    public int ScannedFormulaCount { get; }

    /// <summary>
    /// Formulas inside the declared scope that could not be fully resolved
    /// within qualified parser coverage. Any gap blocks a completeness claim.
    /// </summary>
    public int CoverageGapCount { get; }

    public bool Truncated { get; }

    public string ScanScope => Scope.Label;

    public string? LimitationCode { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public bool CanClaimCompleteness => Status == AuditTraceStatus.Complete;

    public static DirectDependentResult Refused(AuditCellIdentity target, DependentScanScope scope, string code, string message) =>
        new DirectDependentResult(
            AuditTraceStatus.Refused,
            target ?? throw new ArgumentNullException(nameof(target)),
            scope ?? throw new ArgumentNullException(nameof(scope)),
            Array.Empty<DirectDependent>(),
            0,
            0,
            false,
            null,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A refusal code is required.", nameof(code)),
            message ?? string.Empty);
}

/// <summary>
/// A bounded reverse-reference index over the formulas of one declared scope.
/// Each formula is parsed exactly once at build time; queries then intersect a
/// target rectangle against the stored footprints, so repeated queries never
/// re-parse. It reads only the formulas it was given, opens nothing, evaluates
/// nothing, and never widens its scope.
/// </summary>
public sealed class ReverseReferenceIndex
{
    public const int MaximumScannedFormulas = 20_000;

    private readonly DependentScanScope _scope;
    private readonly IReadOnlyList<DependentEntry> _entries;

    private ReverseReferenceIndex(
        DependentScanScope scope,
        IReadOnlyList<DependentEntry> entries,
        int scannedFormulaCount,
        int coverageGapCount,
        bool truncated)
    {
        _scope = scope;
        _entries = entries;
        ScannedFormulaCount = scannedFormulaCount;
        CoverageGapCount = coverageGapCount;
        Truncated = truncated;
    }

    public DependentScanScope Scope => _scope;

    public int ScannedFormulaCount { get; }

    public int CoverageGapCount { get; }

    public bool Truncated { get; }

    public static ReverseReferenceIndex Build(
        DependentScanScope scope,
        IEnumerable<AuditFormulaCell> formulas,
        IEnumerable<AuditNameBinding>? names = null,
        FormulaDialect? dialect = null,
        int externalGapCount = 0)
    {
        if (scope is null) throw new ArgumentNullException(nameof(scope));
        if (formulas is null) throw new ArgumentNullException(nameof(formulas));
        var effectiveDialect = dialect ?? FormulaDialect.InvariantA1;
        var nameArray = (names ?? Enumerable.Empty<AuditNameBinding>()).ToArray();
        var parser = new FormulaParser();
        var entries = new List<DependentEntry>();
        var scanned = 0;
        // A worksheet the plan could not include is a coverage gap in its own
        // right: its formulas were never read.
        var gaps = externalGapCount;
        var truncated = false;

        foreach (var cell in formulas)
        {
            if (cell is null) throw new ArgumentException("Formula cells cannot be null.", nameof(formulas));
            if (scanned >= MaximumScannedFormulas)
            {
                truncated = true;
                break;
            }

            scanned++;
            if (!InScope(scope, cell.Identity))
            {
                // A caller that supplies out-of-scope formulas is a scope
                // expansion. Count it as a gap rather than silently reading it.
                gaps++;
                continue;
            }

            var footprints = new List<ReferenceFootprint>();
            if (!TryCollectFootprints(parser, effectiveDialect, nameArray, cell, footprints)) gaps++;
            if (footprints.Count != 0) entries.Add(new DependentEntry(cell.Identity, footprints));
        }

        return new ReverseReferenceIndex(scope, new ReadOnlyCollection<DependentEntry>(entries), scanned, gaps, truncated);
    }

    public DirectDependentResult FindDirectDependents(AuditCellIdentity target)
    {
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (!InScope(_scope, target))
        {
            return DirectDependentResult.Refused(
                target,
                _scope,
                AuditRefusalCodes.TargetOutsideScope,
                "The target is outside the declared scan scope.");
        }

        if (!AuditAddress.TryParse(target.Address, out var targetRectangle))
        {
            return DirectDependentResult.Refused(
                target,
                _scope,
                AuditRefusalCodes.NotationUnsupported,
                "Dependent scanning is qualified only for a single A1 cell or rectangular range target.");
        }

        var dependents = new List<DirectDependent>();
        foreach (var entry in _entries)
        {
            var evidence = entry.Footprints
                .Where(footprint =>
                    string.Equals(footprint.WorksheetName, target.WorksheetName, StringComparison.OrdinalIgnoreCase) &&
                    footprint.Rectangle.Intersects(targetRectangle))
                .Select(footprint => footprint.Evidence)
                .ToArray();
            if (evidence.Length != 0) dependents.Add(new DirectDependent(entry.Identity, evidence));
        }

        dependents.Sort((left, right) => string.Compare(
            left.Dependent.ToString(), right.Dependent.ToString(), StringComparison.OrdinalIgnoreCase));

        var incomplete = CoverageGapCount > 0 || Truncated;
        return new DirectDependentResult(
            incomplete ? AuditTraceStatus.Partial : AuditTraceStatus.Complete,
            target,
            _scope,
            dependents,
            ScannedFormulaCount,
            CoverageGapCount,
            Truncated,
            Truncated ? AuditRefusalCodes.ScanTruncated : null,
            null,
            incomplete
                ? "Direct dependents are partial; inspect the reported coverage gaps and truncation."
                : null);
    }

    private static bool InScope(DependentScanScope scope, AuditCellIdentity identity)
    {
        if (!string.Equals(scope.WorkbookId, identity.WorkbookId, StringComparison.OrdinalIgnoreCase)) return false;
        return scope.Kind != DependentScanScopeKind.Worksheet ||
            string.Equals(scope.WorksheetName, identity.WorksheetName, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns false when the formula is a coverage gap. Footprints that could
    /// be resolved are still collected, so a gap reduces the completeness claim
    /// without discarding real edges.
    /// </summary>
    private static bool TryCollectFootprints(
        FormulaParser parser,
        FormulaDialect dialect,
        IReadOnlyList<AuditNameBinding> names,
        AuditFormulaCell cell,
        ICollection<ReferenceFootprint> footprints)
    {
        if (dialect.Notation != FormulaNotation.A1) return false;
        if (cell.Formula.Length == 0 || cell.Formula[0] != '=') return false;
        var parse = parser.Parse(cell.Formula, new FormulaParseOptions(dialect));
        if (!parse.IsSuccess) return false;
        var document = parse.Document!;
        var complete = !document.LimitationCodes.Any(CanHideAnInScopeEdge);

        foreach (var reference in document.References)
        {
            var qualifier = UnquoteQualifier(reference.Qualifier);
            if (qualifier is not null && qualifier.IndexOf('[') >= 0)
            {
                // An external reference cannot address anything in this scope.
                continue;
            }

            if (!AuditAddress.TryResolve(reference, out var rectangle))
            {
                complete = false;
                continue;
            }

            footprints.Add(new ReferenceFootprint(
                qualifier ?? cell.Identity.WorksheetName,
                rectangle,
                new AuditReferenceEvidence(
                    reference.SourceText,
                    reference.Span,
                    reference.IsRange ? AuditReferenceKind.Range : AuditReferenceKind.Cell)));
        }

        for (var index = 0; index < document.Tokens.Count; index++)
        {
            if (!AuditNameCandidates.IsNameCandidate(document.Tokens, index)) continue;
            var token = document.Tokens[index];
            var binding = names
                .Where(value => string.Equals(value.Name, token.Text, StringComparison.OrdinalIgnoreCase))
                .Where(value => value.Scope == AuditNameScope.Workbook ||
                    string.Equals(value.WorksheetName, cell.Identity.WorksheetName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(value => value.Scope == AuditNameScope.Worksheet ? 0 : 1)
                .FirstOrDefault();
            if (binding is null || !AuditAddress.TryParse(binding.Target.Address, out var bound))
            {
                complete = false;
                continue;
            }

            footprints.Add(new ReferenceFootprint(
                binding.Target.WorksheetName,
                bound,
                new AuditReferenceEvidence(token.Text, token.Span, AuditReferenceKind.Name)));
        }

        return complete;
    }

    /// <summary>
    /// Whether a coverage limitation could conceal a reference to an in-scope
    /// cell, and therefore has to be counted as a gap.
    ///
    /// A structured reference resolves to real cells but produces no parsed
    /// reference at all, so it is a genuine blind spot. A dynamic-array or
    /// implicit-intersection reference has an extent that cannot be known
    /// without evaluating. An intersection's operands are parsed, but the cells
    /// actually read are only their overlap, so treating the operands as read
    /// would over-report.
    ///
    /// The others cannot hide an in-scope edge. An external reference addresses
    /// another workbook and can never name a cell in this scope. A union's
    /// operands are each parsed and each genuinely read, so recording them is
    /// exact. A defined name is resolved or counted as unresolved by this index
    /// directly, so the parser's note about it adds nothing.
    /// </summary>
    private static bool CanHideAnInScopeEdge(string limitationCode) =>
        limitationCode == FormulaRefusalCodes.StructuredReferenceInspectOnly ||
        limitationCode == FormulaRefusalCodes.DynamicArrayInspectOnly ||
        limitationCode == FormulaRefusalCodes.IntersectionInspectOnly;

    private static string? UnquoteQualifier(string? qualifier)
    {
        if (qualifier is null) return null;
        return qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\''
            ? qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'")
            : qualifier;
    }

    private sealed class DependentEntry
    {
        public DependentEntry(AuditCellIdentity identity, IReadOnlyList<ReferenceFootprint> footprints)
        {
            Identity = identity;
            Footprints = footprints;
        }

        public AuditCellIdentity Identity { get; }

        public IReadOnlyList<ReferenceFootprint> Footprints { get; }
    }

    private sealed class ReferenceFootprint
    {
        public ReferenceFootprint(string worksheetName, AuditRectangle rectangle, AuditReferenceEvidence evidence)
        {
            WorksheetName = worksheetName;
            Rectangle = rectangle;
            Evidence = evidence;
        }

        public string WorksheetName { get; }

        public AuditRectangle Rectangle { get; }

        public AuditReferenceEvidence Evidence { get; }
    }
}
