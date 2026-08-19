using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Auditing;

public static class AuditingCommandCatalog
{
    public const string DirectPrecedentsId = "audit.precedents.direct";

    private static readonly IReadOnlyList<CommandDescriptor> Commands = new[]
    {
        new CommandDescriptor(
            DirectPrecedentsId,
            1,
            "Trace Direct Precedents",
            CommandImpact.ReadOnly,
            Array.Empty<string>(),
            false,
            "Ribbon KeyTips: Alt, X, A, A, PD",
            "CAP-AUD-001",
            CommandContextRequirement.Workbook | CommandContextRequirement.Worksheet | CommandContextRequirement.Selection,
            PreviewPolicy.None,
            UndoPolicy.None,
            new[] { "AC-AUD-001", "AC-AUD-002", "AC-AUD-003", "AC-AUD-004", "AC-AUD-005" },
            "Auditing",
            "Show the direct precedents of one selected formula cell in a read-only view. Closed external workbooks are listed but never opened, and no Excel trace arrow or workbook annotation is used.",
            new[] { "precedents", "trace precedents", "audit precedents" },
            "Alt, X, A, A, PD"),
    }.OrderBy(value => value.Id, StringComparer.Ordinal).ToArray();

    public static IReadOnlyList<CommandDescriptor> All => Commands;

    public static CommandDescriptor GetRequired(string id) => Commands.First(value => value.Id == id);
}
