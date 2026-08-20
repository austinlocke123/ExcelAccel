using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// The single definition of how auditing results are worded and formatted for
/// display. Every trace presentation reads its labels from here so that
/// precedent, dependent, and later traversal views cannot describe the same
/// state in different words.
/// </summary>
public static class AuditPresentationLabels
{
    public static string Status(AuditTraceStatus status) => status switch
    {
        AuditTraceStatus.Complete => "Complete",
        AuditTraceStatus.Partial => "Partial",
        AuditTraceStatus.Refused => "Refused",
        _ => "Unknown",
    };

    public static string Coverage(FormulaCoverageDisposition coverage) => coverage switch
    {
        FormulaCoverageDisposition.Transform => "Fully parsed",
        FormulaCoverageDisposition.RoundTrip => "Fully parsed (round-trip only)",
        FormulaCoverageDisposition.InspectOnly => "Parser coverage gap (inspect only)",
        FormulaCoverageDisposition.Refuse => "Refused",
        _ => "Unknown",
    };

    public static string Kind(AuditReferenceKind kind) => kind switch
    {
        AuditReferenceKind.Cell => "Cell",
        AuditReferenceKind.Range => "Range",
        AuditReferenceKind.Name => "Name",
        AuditReferenceKind.External => "External",
        AuditReferenceKind.Unresolved => "Unresolved",
        _ => "Unknown",
    };

    public static string Classification(AuditCellClassification classification) => classification switch
    {
        AuditCellClassification.Formula => "Formula",
        AuditCellClassification.Value => "Value",
        AuditCellClassification.Error => "Error",
        AuditCellClassification.Blank => "Blank",
        AuditCellClassification.Mixed => "Mixed",
        _ => "Not captured",
    };

    /// <summary>Worksheet-qualified display form, for example "Model!D10".</summary>
    public static string Location(AuditCellIdentity identity) =>
        (identity ?? throw new ArgumentNullException(nameof(identity))).WorksheetName + "!" + identity.Address;

    public static string Count(int value) => value.ToString("N0", CultureInfo.InvariantCulture);

    public static string Count(long value) => value.ToString("N0", CultureInfo.InvariantCulture);

    /// <summary>Source text with its span, for example "A1 [1+2]".</summary>
    public static string Evidence(AuditReferenceEvidence evidence)
    {
        if (evidence is null) throw new ArgumentNullException(nameof(evidence));
        return evidence.SourceText + " [" +
            evidence.SourceSpan.Start.ToString(CultureInfo.InvariantCulture) + "+" +
            evidence.SourceSpan.Length.ToString(CultureInfo.InvariantCulture) + "]";
    }

    public static string EvidenceList(IEnumerable<AuditReferenceEvidence> evidence) =>
        string.Join("; ", (evidence ?? throw new ArgumentNullException(nameof(evidence))).Select(Evidence));

    public static string CompletenessLine(bool canClaimCompleteness) =>
        "Completeness: " + (canClaimCompleteness ? "claimed" : "not claimed");
}
