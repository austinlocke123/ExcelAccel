using System;
using System.Linq;

namespace ExcelAccel.Application.Commands;

public static class CommandExecutionGate
{
    public static CanExecuteResult Authorize(
        CommandDescriptor descriptor,
        CommandPlan plan,
        string? confirmedPlanHash = null)
    {
        if (descriptor is null)
        {
            throw new ArgumentNullException(nameof(descriptor));
        }

        if (plan is null)
        {
            throw new ArgumentNullException(nameof(plan));
        }

        var propertiesMatch = descriptor.ChangedPropertyPolicy == ChangedPropertyPolicy.Exact
            ? descriptor.ChangedProperties.SequenceEqual(plan.ChangedProperties, StringComparer.Ordinal)
            : plan.ChangedProperties.Count > 0 && plan.ChangedProperties.All(value => descriptor.ChangedProperties.Contains(value, StringComparer.Ordinal));
        if (!string.Equals(descriptor.Id, plan.CommandId, StringComparison.Ordinal) ||
            descriptor.ContractVersion != plan.ContractVersion ||
            descriptor.Impact != plan.Impact ||
            !propertiesMatch)
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.ContractMismatch,
                "The command plan no longer matches the registered command contract.",
                "Refresh the command and try again.");
        }

        var confirmationRequired = descriptor.PreviewPolicy == PreviewPolicy.Mandatory || plan.RequiresPreview;
        if (confirmationRequired && !string.Equals(plan.PlanHash, confirmedPlanHash, StringComparison.Ordinal))
        {
            return CanExecuteResult.Refuse(
                RefusalCodes.PreviewRequired,
                "This command requires confirmation of the exact current plan.",
                "Review and confirm the refreshed preview before execution.");
        }

        return CanExecuteResult.Permit();
    }
}
