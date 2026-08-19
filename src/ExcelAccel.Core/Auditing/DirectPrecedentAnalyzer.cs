using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

public sealed class DirectPrecedentAnalyzer
{
    private readonly FormulaParser _parser = new FormulaParser();

    public DirectPrecedentCapturePlan CreateCapturePlan(
        AuditCellIdentity target,
        string formula,
        FormulaDialect? dialect = null)
    {
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        dialect ??= FormulaDialect.InvariantA1;
        var localTargets = new List<AuditCellIdentity>();
        var names = new List<string>();
        if (dialect.Notation == FormulaNotation.A1 && formula.Length > 0 && formula[0] == '=')
        {
            var parse = _parser.Parse(formula, new FormulaParseOptions(dialect));
            if (parse.IsSuccess)
            {
                var document = parse.Document!;
                foreach (var reference in document.References)
                {
                    var qualifier = UnquoteQualifier(reference.Qualifier);
                    if (qualifier is not null && qualifier.IndexOf('[') >= 0) continue;
                    localTargets.Add(new AuditCellIdentity(target.WorkbookId, qualifier ?? target.WorksheetName, CanonicalAddress(reference)));
                }
                for (var index = 0; index < document.Tokens.Count; index++)
                {
                    var token = document.Tokens[index];
                    if (IsNameCandidate(document.Tokens, index)) names.Add(token.Text);
                }
            }
        }
        return new DirectPrecedentCapturePlan(target, formula, dialect, localTargets, names);
    }

    public DirectPrecedentResult Analyze(FormulaReferenceSnapshot snapshot)
    {
        if (snapshot is null) throw new ArgumentNullException(nameof(snapshot));
        if (snapshot.Formula.Length == 0 || snapshot.Formula[0] != '=')
        {
            return Refused(snapshot, AuditRefusalCodes.TargetNotFormula, "Direct precedents require one formula cell.");
        }

        if (snapshot.Dialect.Notation != FormulaNotation.A1)
        {
            return Refused(snapshot, AuditRefusalCodes.NotationUnsupported, "Direct precedent resolution is qualified only for A1 formulas.");
        }

        var parse = _parser.Parse(snapshot.Formula, new FormulaParseOptions(snapshot.Dialect));
        if (!parse.IsSuccess)
        {
            return new DirectPrecedentResult(
                AuditTraceStatus.Refused,
                snapshot.Target,
                Array.Empty<DirectPrecedent>(),
                FormulaCoverageDisposition.Refuse,
                null,
                parse.RefusalCode,
                parse.Message);
        }

        var document = parse.Document!;
        var accumulators = new Dictionary<string, PrecedentAccumulator>(StringComparer.OrdinalIgnoreCase);
        foreach (var reference in document.References)
        {
            AddReference(snapshot, reference, accumulators);
        }

        AddNames(snapshot, document, accumulators);
        if (document.Disposition == FormulaCoverageDisposition.InspectOnly && accumulators.Count == 0)
        {
            AddUnresolved(
                accumulators,
                document.LimitationCode ?? FormulaRefusalCodes.InvalidSyntax,
                snapshot.Formula,
                new FormulaSourceSpan(0, snapshot.Formula.Length));
        }

        var precedents = accumulators.Values
            .OrderBy(item => item.SortKey, StringComparer.OrdinalIgnoreCase)
            .Select(item => item.ToPrecedent())
            .ToArray();
        var incomplete = document.Disposition == FormulaCoverageDisposition.InspectOnly ||
            precedents.Any(item => item.IsExternal || item.IsUnresolved);
        return new DirectPrecedentResult(
            incomplete ? AuditTraceStatus.Partial : AuditTraceStatus.Complete,
            snapshot.Target,
            precedents,
            document.Disposition,
            document.LimitationCode,
            null,
            incomplete ? "Direct precedents are partial; inspect the reported coverage and unresolved edges." : null);
    }

    private static DirectPrecedentResult Refused(FormulaReferenceSnapshot snapshot, string code, string message) =>
        new DirectPrecedentResult(
            AuditTraceStatus.Refused,
            snapshot.Target,
            Array.Empty<DirectPrecedent>(),
            FormulaCoverageDisposition.Refuse,
            null,
            code,
            message);

    private static void AddReference(
        FormulaReferenceSnapshot snapshot,
        FormulaReference reference,
        IDictionary<string, PrecedentAccumulator> results)
    {
        var address = CanonicalAddress(reference);
        var qualifier = UnquoteQualifier(reference.Qualifier);
        if (qualifier is not null && qualifier.IndexOf('[') >= 0)
        {
            var externalKey = "external|" + qualifier + "|" + address;
            Add(results, externalKey, AuditReferenceKind.External, null, qualifier + "!" + address,
                AuditCellClassification.Unknown, reference.SourceText, reference.Span, true, true, false);
            return;
        }

        var sheet = qualifier ?? snapshot.Target.WorksheetName;
        var target = new AuditCellIdentity(snapshot.Target.WorkbookId, sheet, address);
        var kind = reference.IsRange ? AuditReferenceKind.Range : AuditReferenceKind.Cell;
        var classification = snapshot.Index.Classify(target);
        Add(results, "local|" + target, kind, target, target.ToString(), classification,
            reference.SourceText, reference.Span, false, false, classification == AuditCellClassification.Unknown);
    }

    private static void AddNames(
        FormulaReferenceSnapshot snapshot,
        FormulaSyntaxDocument document,
        IDictionary<string, PrecedentAccumulator> results)
    {
        for (var index = 0; index < document.Tokens.Count; index++)
        {
            var token = document.Tokens[index];
            if (!IsNameCandidate(document.Tokens, index))
            {
                continue;
            }

            var matches = snapshot.Index.Names
                .Where(binding => string.Equals(binding.Name, token.Text, StringComparison.OrdinalIgnoreCase))
                .Where(binding => binding.Scope == AuditNameScope.Workbook ||
                    string.Equals(binding.WorksheetName, snapshot.Target.WorksheetName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(binding => binding.Scope == AuditNameScope.Worksheet ? 0 : 1)
                .ToArray();
            var selected = matches.FirstOrDefault();
            if (selected is null)
            {
                AddUnresolved(results, AuditRefusalCodes.NameUnresolved + "|" + token.Text, token.Text, token.Span);
                continue;
            }

            var target = selected.Target;
            var classification = snapshot.Index.Classify(target);
            Add(results, "local|" + target, AuditReferenceKind.Name, target, target.ToString(),
                classification, token.Text, token.Span, false, false, classification == AuditCellClassification.Unknown);
        }
    }

    private static FormulaToken? NextSignificant(IReadOnlyList<FormulaToken> tokens, int start)
    {
        for (var index = start; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != FormulaTokenKind.Whitespace) return tokens[index];
        }
        return null;
    }

    private static bool IsNameCandidate(IReadOnlyList<FormulaToken> tokens, int index)
    {
        var token = tokens[index];
        if (token.Kind != FormulaTokenKind.Identifier || IsBoolean(token.Text)) return false;
        var next = NextSignificant(tokens, index + 1);
        if (next?.Kind == FormulaTokenKind.OpenParenthesis) return false;
        return index + 1 >= tokens.Count || tokens[index + 1].Kind != FormulaTokenKind.BracketedIdentifier;
    }

    private static bool IsBoolean(string value) =>
        string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase);

    private static void AddUnresolved(
        IDictionary<string, PrecedentAccumulator> results,
        string key,
        string sourceText,
        FormulaSourceSpan span) =>
        Add(results, "unresolved|" + key, AuditReferenceKind.Unresolved, null, sourceText,
            AuditCellClassification.Unknown, sourceText, span, false, false, true);

    private static void Add(
        IDictionary<string, PrecedentAccumulator> results,
        string key,
        AuditReferenceKind kind,
        AuditCellIdentity? target,
        string displayTarget,
        AuditCellClassification classification,
        string sourceText,
        FormulaSourceSpan span,
        bool external,
        bool closedExternal,
        bool unresolved)
    {
        if (!results.TryGetValue(key, out var item))
        {
            item = new PrecedentAccumulator(key, kind, target, displayTarget, classification, external, closedExternal, unresolved);
            results.Add(key, item);
        }
        item.Evidence.Add(new AuditReferenceEvidence(sourceText, span, kind));
    }

    private static string CanonicalAddress(FormulaReference reference)
    {
        var first = reference.First;
        if (reference.Second is null) return Endpoint(first);
        var second = reference.Second;
        var minRow = Math.Min(first.Row.Value, second.Row.Value);
        var maxRow = Math.Max(first.Row.Value, second.Row.Value);
        var minColumn = Math.Min(first.Column.Value, second.Column.Value);
        var maxColumn = Math.Max(first.Column.Value, second.Column.Value);
        return Cell(minRow, minColumn) + ":" + Cell(maxRow, maxColumn);
    }

    private static string Endpoint(FormulaReferenceEndpoint endpoint) => Cell(endpoint.Row.Value, endpoint.Column.Value);

    private static string Cell(int row, int column)
    {
        var name = string.Empty;
        for (var remaining = column; remaining > 0; remaining /= 26)
        {
            var zeroBased = (remaining - 1) % 26;
            name = (char)('A' + zeroBased) + name;
        }
        return name + row.ToString(CultureInfo.InvariantCulture);
    }

    private static string? UnquoteQualifier(string? qualifier)
    {
        if (qualifier is null) return null;
        return qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\''
            ? qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'")
            : qualifier;
    }

    private sealed class PrecedentAccumulator
    {
        public PrecedentAccumulator(
            string sortKey,
            AuditReferenceKind kind,
            AuditCellIdentity? target,
            string displayTarget,
            AuditCellClassification classification,
            bool external,
            bool closedExternal,
            bool unresolved)
        {
            SortKey = sortKey;
            Kind = kind;
            Target = target;
            DisplayTarget = displayTarget;
            Classification = classification;
            IsExternal = external;
            IsClosedExternal = closedExternal;
            IsUnresolved = unresolved;
        }

        public string SortKey { get; }
        public AuditReferenceKind Kind { get; }
        public AuditCellIdentity? Target { get; }
        public string DisplayTarget { get; }
        public AuditCellClassification Classification { get; }
        public bool IsExternal { get; }
        public bool IsClosedExternal { get; }
        public bool IsUnresolved { get; }
        public List<AuditReferenceEvidence> Evidence { get; } = new List<AuditReferenceEvidence>();

        public DirectPrecedent ToPrecedent() => new DirectPrecedent(
            SortKey, Kind, Target, DisplayTarget, Classification, Evidence,
            IsExternal, IsClosedExternal, IsUnresolved);
    }
}
