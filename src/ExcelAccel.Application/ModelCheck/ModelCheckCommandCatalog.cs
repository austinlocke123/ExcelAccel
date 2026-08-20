using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.ModelCheck;

public static class ModelCheckCommandCatalog
{
    public const string RunSelectionId = "model_check.run.selection";
    public const string RunWorksheetId = "model_check.run.worksheet";
    public const string RescanId = "model_check.rescan";
    public const string IgnoreLocalId = "model_check.finding.ignore_local";
    public const string UnignoreLocalId = "model_check.finding.unignore_local";
    public const string ExportId = "model_check.results.export";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        ReadOnly(
            RunSelectionId,
            "Model Check Selection",
            "MS",
            "Run the enabled Model Check rules over the current selection and show the findings in a read-only view.",
            PreviewPolicy.None,
            new[] { "AC-CHECK-001", "AC-CHECK-002", "AC-CHECK-003", "AC-CHECK-004", "AC-CHECK-005", "AC-CHECK-006" }),
        ReadOnly(
            RunWorksheetId,
            "Model Check Worksheet",
            "MW",
            "Run the enabled Model Check rules over the active worksheet's used region. A large worksheet is confirmed before anything is read.",
            PreviewPolicy.Threshold,
            new[] { "AC-CHECK-001", "AC-CHECK-002", "AC-CHECK-006", "AC-CHECK-007" }),
        ReadOnly(
            RescanId,
            "Model Check Rescan",
            "MR",
            "Repeat the exact prior scope and rule configuration against a newly captured snapshot. Prior findings are never relabelled as current.",
            PreviewPolicy.None,
            new[] { "AC-CHECK-034" }),
        new CommandDescriptor(
            IgnoreLocalId,
            1,
            "Ignore Finding Locally",
            CommandImpact.Low,
            new[] { "local_model_check_ignores" },
            false,
            "Ribbon KeyTips: Alt, X, A, K, MI",
            "CAP-CHECK-001",
            CommandContextRequirement.Application,
            PreviewPolicy.Mandatory,
            UndoPolicy.None,
            new[] { "AC-CHECK-030", "AC-CHECK-031", "AC-CHECK-032", "AC-CHECK-033" },
            "Model Check",
            "Suppress the selected finding locally by its normalized fingerprint. No formula or value content is stored, and a rescan is required for it to take effect.",
            new[] { "ignore finding", "suppress finding" },
            "Alt, X, A, K, MI"),
        new CommandDescriptor(
            UnignoreLocalId,
            1,
            "Manage Local Ignores",
            CommandImpact.Low,
            new[] { "local_model_check_ignores" },
            false,
            "Ribbon KeyTips: Alt, X, A, K, MU",
            "CAP-CHECK-001",
            CommandContextRequirement.Application,
            PreviewPolicy.Mandatory,
            UndoPolicy.None,
            new[] { "AC-CHECK-031", "AC-CHECK-032", "AC-CHECK-033" },
            "Model Check",
            "Show the active local ignores and remove selected entries. A rescan is required for a removal to take effect.",
            new[] { "unignore finding", "manage ignores" },
            "Alt, X, A, K, MU"),
        new CommandDescriptor(
            ExportId,
            1,
            "Export Model Check Results",
            CommandImpact.Medium,
            new[] { "local_export_file" },
            false,
            "Ribbon KeyTips: Alt, X, A, K, ME",
            "CAP-CHECK-001",
            CommandContextRequirement.Application,
            PreviewPolicy.Mandatory,
            UndoPolicy.None,
            new[] { "AC-CHECK-035", "AC-CHECK-036", "AC-CHECK-037", "AC-SEC-004" },
            "Model Check",
            "Write the current findings to a local file after confirming a manifest. Formulas and values are excluded by default and nothing is transmitted.",
            new[] { "export findings", "export results" },
            "Alt, X, A, K, ME"),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    private static CommandDescriptor ReadOnly(
        string id,
        string name,
        string keytip,
        string description,
        PreviewPolicy previewPolicy,
        IEnumerable<string> acceptanceIds) =>
        new CommandDescriptor(
            id,
            1,
            name,
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            "Ribbon KeyTips: Alt, X, A, K, " + keytip,
            "CAP-CHECK-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            previewPolicy,
            UndoPolicy.None,
            acceptanceIds,
            "Model Check",
            description,
            new[] { "model check", "check model", "review model" },
            "Alt, X, A, K, " + keytip);
}
