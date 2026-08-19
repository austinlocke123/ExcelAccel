using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAccel.Application.Commands;

public static class CanonicalPlanSerializer
{
    public const int SchemaVersion = 1;

    public static string Serialize(CommandPlan plan)
    {
        if (plan is null)
        {
            throw new ArgumentNullException(nameof(plan));
        }

        var builder = new StringBuilder(512);
        builder.Append("{\"schema_version\":").Append(SchemaVersion);
        AppendString(builder, "command_id", plan.CommandId);
        builder.Append(",\"contract_version\":").Append(plan.ContractVersion.ToString(CultureInfo.InvariantCulture));
        AppendString(builder, "impact", plan.Impact.ToString().ToLowerInvariant());
        AppendString(builder, "workbook_id", plan.Context.WorkbookId);
        AppendString(builder, "worksheet_name", plan.Context.WorksheetName);
        AppendString(builder, "address", plan.Context.Address);
        builder.Append(",\"changed_properties\":[");
        for (var index = 0; index < plan.ChangedProperties.Count; index++)
        {
            if (index > 0)
            {
                builder.Append(',');
            }

            AppendQuoted(builder, plan.ChangedProperties[index]);
        }

        builder.Append(']');
        builder.Append(",\"affected_cell_count\":").Append(plan.AffectedCellCount.ToString(CultureInfo.InvariantCulture));
        AppendString(builder, "precondition_fingerprint", plan.PreconditionFingerprint);
        builder.Append(",\"requires_preview\":").Append(plan.RequiresPreview ? "true" : "false");
        builder.Append(",\"arguments\":{");
        var argumentIndex = 0;
        foreach (var argument in plan.Arguments)
        {
            if (argumentIndex++ > 0)
            {
                builder.Append(',');
            }

            AppendQuoted(builder, argument.Key);
            builder.Append(':');
            AppendQuoted(builder, argument.Value);
        }

        builder.Append("}}");
        return builder.ToString();
    }

    public static string Hash(CommandPlan plan)
    {
        var bytes = Encoding.UTF8.GetBytes(Serialize(plan));
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(bytes);
        var builder = new StringBuilder(hash.Length * 2);
        foreach (var value in hash)
        {
            builder.Append(value.ToString("x2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static void AppendString(StringBuilder builder, string name, string value)
    {
        builder.Append(',');
        AppendQuoted(builder, name);
        builder.Append(':');
        AppendQuoted(builder, value);
    }

    private static void AppendQuoted(StringBuilder builder, string value)
    {
        builder.Append('"');
        foreach (var character in value)
        {
            switch (character)
            {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\b': builder.Append("\\b"); break;
                case '\f': builder.Append("\\f"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (character < 0x20)
                    {
                        builder.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        builder.Append(character);
                    }
                    break;
            }
        }

        builder.Append('"');
    }
}
