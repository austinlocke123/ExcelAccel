using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

public enum AuditCellClassification
{
    Formula,
    Value,
    Error,
    Blank,
    Mixed,
    Unknown,
}

public enum AuditReferenceKind
{
    Cell,
    Range,
    Name,
    External,
    Unresolved,
}

public enum AuditTraceStatus
{
    Complete,
    Partial,
    Refused,
}

public enum AuditNameScope
{
    Workbook,
    Worksheet,
}

public sealed class AuditCellIdentity : IEquatable<AuditCellIdentity>
{
    public AuditCellIdentity(string workbookId, string worksheetName, string address)
    {
        WorkbookId = Require(workbookId, nameof(workbookId));
        WorksheetName = Require(worksheetName, nameof(worksheetName));
        Address = Require(address, nameof(address)).ToUpperInvariant();
    }

    public string WorkbookId { get; }
    public string WorksheetName { get; }
    public string Address { get; }

    public bool Equals(AuditCellIdentity? other) => other is not null &&
        StringComparer.OrdinalIgnoreCase.Equals(WorkbookId, other.WorkbookId) &&
        StringComparer.OrdinalIgnoreCase.Equals(WorksheetName, other.WorksheetName) &&
        StringComparer.OrdinalIgnoreCase.Equals(Address, other.Address);

    public override bool Equals(object? obj) => Equals(obj as AuditCellIdentity);

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = StringComparer.OrdinalIgnoreCase.GetHashCode(WorkbookId);
            hash = (hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(WorksheetName);
            return (hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(Address);
        }
    }

    public override string ToString() => $"{WorkbookId}|{WorksheetName}|{Address}";

    private static string Require(string value, string name) =>
        !string.IsNullOrWhiteSpace(value) ? value : throw new ArgumentException("A non-empty value is required.", name);
}

public sealed class AuditNameBinding
{
    public AuditNameBinding(string name, AuditNameScope scope, AuditCellIdentity target, string? worksheetName = null)
    {
        Name = !string.IsNullOrWhiteSpace(name) ? name : throw new ArgumentException("A name is required.", nameof(name));
        Scope = scope;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        WorksheetName = scope == AuditNameScope.Worksheet
            ? (!string.IsNullOrWhiteSpace(worksheetName) ? worksheetName : throw new ArgumentException("Worksheet-scoped names require a worksheet.", nameof(worksheetName)))
            : null;
    }

    public string Name { get; }
    public AuditNameScope Scope { get; }
    public string? WorksheetName { get; }
    public AuditCellIdentity Target { get; }
}

public sealed class ReferenceSnapshotIndex
{
    public const int MaximumEntries = 4096;
    private readonly IReadOnlyDictionary<AuditCellIdentity, AuditCellClassification> _cells;
    private readonly IReadOnlyList<AuditNameBinding> _names;

    public ReferenceSnapshotIndex(
        IEnumerable<KeyValuePair<AuditCellIdentity, AuditCellClassification>> cells,
        IEnumerable<AuditNameBinding>? names = null)
    {
        if (cells is null) throw new ArgumentNullException(nameof(cells));
        var cellMap = new Dictionary<AuditCellIdentity, AuditCellClassification>();
        foreach (var pair in cells)
        {
            if (pair.Key is null) throw new ArgumentException("Snapshot identities cannot be null.", nameof(cells));
            if (cellMap.ContainsKey(pair.Key)) throw new ArgumentException($"Duplicate snapshot identity '{pair.Key}'.", nameof(cells));
            cellMap.Add(pair.Key, pair.Value);
            if (cellMap.Count > MaximumEntries) throw new ArgumentOutOfRangeException(nameof(cells), $"A snapshot index is limited to {MaximumEntries:N0} entries.");
        }

        var nameArray = (names ?? Enumerable.Empty<AuditNameBinding>()).ToArray();
        if (nameArray.Any(name => name is null)) throw new ArgumentException("Name bindings cannot contain null entries.", nameof(names));
        if (nameArray.Length > MaximumEntries) throw new ArgumentOutOfRangeException(nameof(names), $"A snapshot index is limited to {MaximumEntries:N0} names.");
        _cells = new ReadOnlyDictionary<AuditCellIdentity, AuditCellClassification>(cellMap);
        _names = Array.AsReadOnly(nameArray);
    }

    public IReadOnlyDictionary<AuditCellIdentity, AuditCellClassification> Cells => _cells;
    public IReadOnlyList<AuditNameBinding> Names => _names;

    public AuditCellClassification Classify(AuditCellIdentity identity) =>
        _cells.TryGetValue(identity ?? throw new ArgumentNullException(nameof(identity)), out var classification)
            ? classification
            : AuditCellClassification.Unknown;
}

public sealed class FormulaReferenceSnapshot
{
    public FormulaReferenceSnapshot(
        AuditCellIdentity target,
        string formula,
        ReferenceSnapshotIndex index,
        FormulaDialect? dialect = null)
    {
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Formula = formula ?? throw new ArgumentNullException(nameof(formula));
        Index = index ?? throw new ArgumentNullException(nameof(index));
        Dialect = dialect ?? FormulaDialect.InvariantA1;
    }

    public AuditCellIdentity Target { get; }
    public string Formula { get; }
    public ReferenceSnapshotIndex Index { get; }
    public FormulaDialect Dialect { get; }
}

public sealed class DirectPrecedentCapturePlan
{
    internal DirectPrecedentCapturePlan(
        AuditCellIdentity target,
        string formula,
        FormulaDialect dialect,
        IEnumerable<AuditCellIdentity> localTargets,
        IEnumerable<string> nameCandidates)
    {
        Target = target;
        Formula = formula;
        Dialect = dialect;
        LocalTargets = Array.AsReadOnly(localTargets.Distinct().OrderBy(value => value.ToString(), StringComparer.OrdinalIgnoreCase).ToArray());
        NameCandidates = Array.AsReadOnly(nameCandidates.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToArray());
    }

    public AuditCellIdentity Target { get; }
    public string Formula { get; }
    public FormulaDialect Dialect { get; }
    public IReadOnlyList<AuditCellIdentity> LocalTargets { get; }
    public IReadOnlyList<string> NameCandidates { get; }
}

public sealed class AuditReferenceEvidence
{
    public AuditReferenceEvidence(string sourceText, FormulaSourceSpan sourceSpan, AuditReferenceKind kind)
    {
        SourceText = sourceText ?? throw new ArgumentNullException(nameof(sourceText));
        SourceSpan = sourceSpan;
        Kind = kind;
    }

    public string SourceText { get; }
    public FormulaSourceSpan SourceSpan { get; }
    public AuditReferenceKind Kind { get; }
}

public sealed class DirectPrecedent
{
    internal DirectPrecedent(
        string nodeId,
        AuditReferenceKind kind,
        AuditCellIdentity? target,
        string displayTarget,
        AuditCellClassification classification,
        IEnumerable<AuditReferenceEvidence> evidence,
        bool isExternal,
        bool isClosedExternal,
        bool isUnresolved)
    {
        NodeId = !string.IsNullOrWhiteSpace(nodeId) ? nodeId : throw new ArgumentException("A node ID is required.", nameof(nodeId));
        Kind = kind;
        Target = target;
        DisplayTarget = !string.IsNullOrWhiteSpace(displayTarget) ? displayTarget : throw new ArgumentException("A display target is required.", nameof(displayTarget));
        Classification = classification;
        Evidence = Array.AsReadOnly((evidence ?? throw new ArgumentNullException(nameof(evidence))).ToArray());
        if (Evidence.Count == 0) throw new ArgumentException("At least one source edge is required.", nameof(evidence));
        IsExternal = isExternal;
        IsClosedExternal = isClosedExternal;
        IsUnresolved = isUnresolved;
    }

    public string NodeId { get; }
    public AuditReferenceKind Kind { get; }
    public AuditCellIdentity? Target { get; }
    public string DisplayTarget { get; }
    public AuditCellClassification Classification { get; }
    public IReadOnlyList<AuditReferenceEvidence> Evidence { get; }
    public int Depth => 1;
    public bool IsExternal { get; }
    public bool IsClosedExternal { get; }
    public bool IsUnresolved { get; }
}

public sealed class DirectPrecedentResult
{
    internal DirectPrecedentResult(
        AuditTraceStatus status,
        AuditCellIdentity source,
        IEnumerable<DirectPrecedent> precedents,
        FormulaCoverageDisposition coverage,
        string? limitationCode,
        string? refusalCode,
        string? message)
    {
        Status = status;
        Source = source ?? throw new ArgumentNullException(nameof(source));
        Precedents = Array.AsReadOnly((precedents ?? throw new ArgumentNullException(nameof(precedents))).ToArray());
        Coverage = coverage;
        LimitationCode = limitationCode;
        RefusalCode = refusalCode;
        Message = message;
    }

    public AuditTraceStatus Status { get; }
    public AuditCellIdentity Source { get; }
    public IReadOnlyList<DirectPrecedent> Precedents { get; }
    public FormulaCoverageDisposition Coverage { get; }
    public string ScanScope => "target_formula";
    public string? LimitationCode { get; }
    public string? RefusalCode { get; }
    public string? Message { get; }
    public bool CanClaimCompleteness => Status == AuditTraceStatus.Complete;
    public int UnresolvedEdgeCount => Precedents.Count(item => item.IsUnresolved);
    public int ExternalEdgeCount => Precedents.Count(item => item.IsExternal);

    public static DirectPrecedentResult Refused(AuditCellIdentity source, string code, string message) =>
        new DirectPrecedentResult(
            AuditTraceStatus.Refused,
            source ?? throw new ArgumentNullException(nameof(source)),
            Array.Empty<DirectPrecedent>(),
            FormulaCoverageDisposition.Refuse,
            null,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A refusal code is required.", nameof(code)),
            message ?? string.Empty);
}

public static class AuditRefusalCodes
{
    public const string TargetNotFormula = "AUDIT_TARGET_NOT_FORMULA";
    public const string NotationUnsupported = "AUDIT_NOTATION_UNSUPPORTED";
    public const string NameUnresolved = "AUDIT_NAME_UNRESOLVED";
    public const string StaleTarget = "AUDIT_STALE_TARGET";
    public const string ScopeUnsupported = "AUDIT_SCOPE_UNSUPPORTED";
    public const string TargetOutsideScope = "AUDIT_TARGET_OUTSIDE_SCOPE";
    public const string ScanTruncated = "AUDIT_SCAN_TRUNCATED";
    public const string ScanRegionTooLarge = "AUDIT_SCAN_REGION_TOO_LARGE";
    public const string ScanRegionUnsupported = "AUDIT_SCAN_REGION_UNSUPPORTED";
    public const string ScanCancelled = "AUDIT_SCAN_CANCELLED";
    public const string PreviewRequired = "AUDIT_PREVIEW_REQUIRED";
    public const string DepthLimitReached = "AUDIT_DEPTH_LIMIT_REACHED";
}
