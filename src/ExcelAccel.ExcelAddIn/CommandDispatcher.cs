using ExcelDna.Integration;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Navigation;
using ExcelAccel.Core.Commands;
using ExcelAccel.Application.Undo;
using System;
using System.IO;
using System.Windows.Forms;
using ExcelAccel.Persistence.Diagnostics;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using ExcelAccel.Application.Profiles;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Styles;
using ExcelAccel.Persistence.Profiles;
using ExcelAccel.Application.Formulas;

namespace ExcelAccel.ExcelAddIn;

internal static class CommandDispatcher
{
    public static CommandResult InvokeRegistered(string commandId, IReadOnlyDictionary<string, string>? arguments, InvocationSource source)
    {
        if (string.IsNullOrWhiteSpace(commandId))
            return CommandResult.Refused("command.invoke", "A command ID is required.", RefusalCodes.CommandUnavailable);
        var fixedArguments = arguments ?? new Dictionary<string, string>();
        if (fixedArguments.Count != 0)
            return CommandResult.Refused(commandId, "This command does not yet accept fixed invocation arguments.", RefusalCodes.ContractMismatch);
        DiagnosticLog.Info(commandId, $"invocation_source:{source.ToString().ToLowerInvariant()}");
        if (commandId == InspectSelectionCommand.Id) return InspectSelection();
        if (commandId == ApplyCurrencyFormatCommand.Id) return ApplyCurrencyFormat();
        if (commandId == UndoLastCommand.Id) return UndoLastProperty();
        if (commandId == "support.diagnostics.export") return ExportDiagnostics();
        if (commandId == "command.search.open") return CommandSearchRuntime.Open();
        if (commandId == "profile.export") return ExportProfile();
        if (commandId == "profile.import.preview") return ImportProfile(apply: false);
        if (commandId == "profile.import.apply") return ImportProfile(apply: true);
        if (commandId == "bindings.cheat_sheet.export") return ExportBindingCheatSheet();
        if (commandId.StartsWith("style.", StringComparison.Ordinal)) return StyleLibraryRuntime.Open();
        if (commandId.StartsWith("favorite.", StringComparison.Ordinal))
            return CommandResult.Refused(commandId, "Select a concrete command or favorite in Command Search.", RefusalCodes.CommandUnavailable);
        if (Phase1AFormattingCatalog.All.Any(value => value.Id == commandId)) return ApplyProfileFormatting(commandId);
        if (NavigationCommandCatalog.All.Any(value => value.Id == commandId)) return Navigate(commandId);
        if (FormulaCommandCatalog.All.Any(value => value.Id == commandId)) return ApplyFormulaCommand(commandId);
        return CommandResult.Refused(commandId, "The registered command has no available host dispatcher.", RefusalCodes.CommandUnavailable);
    }

    public static Func<CommandDescriptor, CanExecuteResult> CaptureAvailability()
    {
        SelectionSnapshot? snapshot = null;
        CommandRefusedException? captureFailure = null;
        try { snapshot = CreateSelectionAdapter().CaptureSelection(); }
        catch (CommandRefusedException exception) { captureFailure = exception; }

        return descriptor =>
        {
            if (descriptor.ContextRequirement == CommandContextRequirement.Application ||
                descriptor.Id.StartsWith("favorite.", StringComparison.Ordinal) || descriptor.Id == "command.search.open")
                return CanExecuteResult.Permit();
            if (snapshot is null)
                return CanExecuteResult.Refuse(captureFailure?.RefusalCode ?? RefusalCodes.CommandUnavailable,
                    captureFailure?.Message ?? "Excel context is unavailable.", captureFailure?.Remediation ?? "Open a workbook and select a cell range.");
            if (descriptor.Impact != CommandImpact.ReadOnly &&
                !descriptor.ChangedProperties.Contains("user_profile_favorites") &&
                (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(descriptor.Id)))
                return CanExecuteResult.Refuse(RefusalCodes.CommandQuarantined, "Workbook mutation is disabled in safe mode or quarantine.", "Restart Excel cleanly before retrying.");
            if (descriptor.Id == ApplyCurrencyFormatCommand.Id) return new ApplyCurrencyFormatCommand().CanExecute(snapshot);
            if (descriptor.Impact != CommandImpact.ReadOnly && descriptor.ContextRequirement.HasFlag(CommandContextRequirement.Selection))
            {
                if (snapshot.Safety.AreaCount != 1 || snapshot.Safety.HasMergedCells)
                    return CanExecuteResult.Refuse(RefusalCodes.SelectionUnsupported, "The command requires one unmerged rectangular selection.", "Select one unmerged range.");
                if (snapshot.Safety.WorksheetProtected || snapshot.Safety.WorkbookReadOnly)
                    return CanExecuteResult.Refuse(RefusalCodes.ProtectedTarget, "The target is protected or read-only.", "Use an editable, unprotected target.");
                if (!snapshot.Safety.DynamicArraySpillCheckSupported || snapshot.Safety.HasLegacyArray || snapshot.Safety.HasDynamicArraySpill)
                    return CanExecuteResult.Refuse(RefusalCodes.ArrayOrSpillUnsafe, "The selection intersects an unqualified array or spill state.", "Select cells outside array/spill ranges.");
                if (snapshot.CellCount > ProfileFormattingCommand.MaximumCellCount)
                    return CanExecuteResult.Refuse(RefusalCodes.ResourceLimit, "The selection exceeds the immediate command limit.", "Select a smaller range.");
            }
            if (descriptor.Id == "navigate.history.back" && NavigationRuntime.Session.HistoryCount < 2)
                return CanExecuteResult.Refuse(RefusalCodes.CommandUnavailable, "No prior navigation location is available.", "Navigate first, then retry.");
            if ((descriptor.Id == "navigate.bookmark.next_session" || descriptor.Id == "navigate.bookmark.previous_session") && NavigationRuntime.Session.BookmarkCount == 0)
                return CanExecuteResult.Refuse(RefusalCodes.CommandUnavailable, "No session bookmark is available.", "Add a session bookmark first.");
            return CanExecuteResult.Permit();
        };
    }

    public static CommandResult InspectSelection()
    {
        var port = CreateSelectionAdapter();
        var command = new InspectSelectionCommand();
        var plan = command.Plan(port.CaptureSelection());
        return command.Execute(plan, port);
    }

    public static CommandResult UndoLastProperty()
    {
        var port = CreateSelectionAdapter();
        var workbookId = port.CaptureSelection().Context.WorkbookId;
        return UndoLastCommand.Execute(workbookId, UndoRuntime.Store, port, DateTimeOffset.UtcNow);
    }

    public static CommandResult ExportDiagnostics()
    {
        const string commandId = "support.diagnostics.export";
        using (var dialog = new SaveFileDialog
        {
            Title = "Export ExcelAccel diagnostics",
            Filter = "Text file (*.txt)|*.txt",
            DefaultExt = "txt",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = "excelaccel-support.txt",
        })
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
                return CommandResult.Refused(commandId, "Diagnostic export was cancelled; no file was written.", "USER_CANCELLED");
            var bytes = File.Exists(DiagnosticLog.LogPath) ? File.ReadAllBytes(DiagnosticLog.LogPath) : new byte[0];
            var exporter = new DiagnosticExporter();
            var plan = exporter.Plan(dialog.FileName, bytes);
            var confirmation = owner is null
                ? MessageBox.Show(plan.Manifest + "\n\nCreate this local file?", "ExcelAccel diagnostic manifest", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                : MessageBox.Show(owner, plan.Manifest + "\n\nCreate this local file?", "ExcelAccel diagnostic manifest", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (confirmation != DialogResult.Yes)
                return CommandResult.Refused(commandId, "Diagnostic export was not confirmed; no file was written.", "PREVIEW_NOT_CONFIRMED");
            exporter.Export(plan, plan.PlanHash, bytes);
            return CommandResult.Success(commandId, "The sanitized local diagnostic file was created. Nothing was transmitted.");
        }
    }

    public static CommandResult ExportProfile()
    {
        const string commandId = "profile.export";
        using (var dialog = new SaveFileDialog
        {
            Title = "Export ExcelAccel profile",
            Filter = "ExcelAccel profile (*.excelaccel-profile.json)|*.excelaccel-profile.json|JSON file (*.json)|*.json",
            DefaultExt = "excelaccel-profile.json", AddExtension = true, OverwritePrompt = true,
            FileName = "ExcelAccel-profile.excelaccel-profile.json",
        })
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
                return CommandResult.Refused(commandId, "Profile export was cancelled.", "USER_CANCELLED");
            var service = new ProfilePackageService();
            var overwrite = File.Exists(dialog.FileName);
            var plan = service.PlanExport(ProfileRuntime.Current, dialog.FileName, overwrite);
            var text = plan.Manifest + $"\nDestination: {plan.DestinationPath}\nPackage SHA-256: {plan.PackageSha256}\n\nCreate this exact local package?";
            var confirmation = owner is null ? MessageBox.Show(text, "ExcelAccel profile export", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                : MessageBox.Show(owner, text, "ExcelAccel profile export", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (confirmation != DialogResult.Yes) return CommandResult.Refused(commandId, "Profile export manifest was not confirmed.", "PREVIEW_NOT_CONFIRMED");
            service.Export(plan, plan.PlanHash);
            return CommandResult.Success(commandId, "The local profile package was created and verified. Nothing was launched or transmitted.");
        }
    }

    public static CommandResult ImportProfile(bool apply)
    {
        var commandId = apply ? "profile.import.apply" : "profile.import.preview";
        using (var dialog = new OpenFileDialog
        {
            Title = "Select ExcelAccel profile package",
            Filter = "ExcelAccel profile (*.excelaccel-profile.json;*.json)|*.excelaccel-profile.json;*.json",
            CheckFileExists = true, Multiselect = false,
        })
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
                return CommandResult.Refused(commandId, "Profile import was cancelled.", "USER_CANCELLED");
            var service = new ProfilePackageService();
            var current = ProfileRuntime.Current;
            var plan = service.PreviewImport(dialog.FileName, current, BuiltInCommandRegistry.All);
            var preview = plan.Manifest + "\n\n" + plan.Diff + $"\nSource SHA-256: {plan.SourceSha256}";
            if (!apply) return CommandResult.Success(commandId, "Import preview validated without changing settings:\n\n" + preview);
            var confirmation = owner is null ? MessageBox.Show(preview + "\n\nBack up the current profile and apply this exact import?", "ExcelAccel profile import", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                : MessageBox.Show(owner, preview + "\n\nBack up the current profile and apply this exact import?", "ExcelAccel profile import", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirmation != DialogResult.Yes) return CommandResult.Refused(commandId, "Profile import was not confirmed.", "PREVIEW_NOT_CONFIRMED");
            var imported = service.ApplyImport(plan, plan.PlanHash, ProfileRuntime.ProfilePath, current, BuiltInCommandRegistry.All);
            ProfileRuntime.Activate(imported);
            return CommandResult.Success(commandId, "The validated profile was backed up and atomically activated.");
        }
    }

    public static CommandResult ExportBindingCheatSheet()
    {
        const string commandId = "bindings.cheat_sheet.export";
        using (var dialog = new SaveFileDialog
        {
            Title = "Export ExcelAccel shortcut cheat sheet",
            Filter = "HTML file (*.html)|*.html|CSV file (*.csv)|*.csv",
            DefaultExt = "html", AddExtension = true, OverwritePrompt = true, FileName = "ExcelAccel-shortcuts.html",
        })
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
                return CommandResult.Refused(commandId, "Shortcut export was cancelled.", "USER_CANCELLED");
            var format = dialog.FilterIndex == 2 ? BindingExportFormat.Csv : BindingExportFormat.Html;
            var exporter = new BindingCheatSheetExporter();
            var plan = exporter.Plan(ProfileRuntime.Current.QuickKeys, BuiltInCommandRegistry.All, dialog.FileName, format, File.Exists(dialog.FileName));
            var confirmation = owner is null ? MessageBox.Show(plan.Manifest + "\n\nCreate this exact local file?", "ExcelAccel shortcuts", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                : MessageBox.Show(owner, plan.Manifest + "\n\nCreate this exact local file?", "ExcelAccel shortcuts", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (confirmation != DialogResult.Yes) return CommandResult.Refused(commandId, "Shortcut export was not confirmed.", "PREVIEW_NOT_CONFIRMED");
            exporter.Export(plan, plan.PlanHash);
            return CommandResult.Success(commandId, "The local shortcut cheat sheet was created and verified.");
        }
    }

    public static CommandResult ApplyCurrencyFormat()
    {
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(ApplyCurrencyFormatCommand.Id))
        {
            return CommandResult.Refused(
                ApplyCurrencyFormatCommand.Id,
                "ExcelAccel is in safe mode after an unclean prior session. Restart Excel cleanly before using mutation commands.",
                RefusalCodes.CommandQuarantined);
        }

        var port = CreateSelectionAdapter();
        var command = new ApplyCurrencyFormatCommand();
        var snapshot = port.CaptureSelection();
        var canExecute = command.CanExecute(snapshot);
        if (!canExecute.Allowed)
        {
            return CommandResult.Refused(
                ApplyCurrencyFormatCommand.Id,
                $"{canExecute.Message} {canExecute.Remediation}".Trim(),
                canExecute.RefusalCode);
        }

        var plan = command.Plan(snapshot);
        return command.Execute(plan, port, UndoRuntime.Store);
    }

    public static CommandResult ApplyStyle(string styleId, bool requireBuiltIn = false)
    {
        const string localCommandId = "style.apply";
        const string builtInCommandId = "style.apply_builtin";
        var style = StyleLibrary.Effective(ProfileRuntime.Current.LocalStyles)
            .FirstOrDefault(value => string.Equals(value.StyleId, styleId, StringComparison.Ordinal));
        var commandId = style?.Origin == StyleOrigin.BuiltIn ? builtInCommandId : localCommandId;
        if (style is null || (requireBuiltIn && style.Origin != StyleOrigin.BuiltIn))
            return CommandResult.Refused(commandId, $"Style '{styleId}' is unavailable.", RefusalCodes.CommandUnavailable);
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(commandId))
            return CommandResult.Refused(commandId, "Style mutation is disabled in safe mode or quarantine.", RefusalCodes.CommandQuarantined);
        var port = CreateSelectionAdapter();
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired(commandId));
        var plan = command.Plan(style, port, ProfileRuntime.Current.ImmediatePreviewCellLimit);
        string? confirmation = null;
        if (plan.CommandPlan.RequiresPreview)
        {
            var text = plan.CommandPlan.Summary + "\n\nProperties: " + string.Join(", ", plan.CommandPlan.ChangedProperties) +
                $"\nSkipped: {plan.Skipped.Count}\n\nApply this exact plan?";
            var owner = ExcelWindowOwner.TryCreate();
            var response = owner is null
                ? MessageBox.Show(text, "ExcelAccel style preview", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                : MessageBox.Show(owner, text, "ExcelAccel style preview", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (response != DialogResult.Yes) return CommandResult.Refused(plan.CommandPlan, "Style preview was cancelled.", "USER_CANCELLED");
            confirmation = plan.CommandPlan.PlanHash;
        }
        return command.Execute(plan, port, confirmation, UndoRuntime.Store);
    }

    public static CommandResult CaptureLocalStyle(string displayName, IReadOnlyList<string> propertyIds)
    {
        if (RuntimeState.IsSafeMode)
            return CommandResult.Refused("style.capture", "Local style capture is disabled in safe mode.", RefusalCodes.CommandQuarantined);
        var existing = ProfileRuntime.Current.LocalStyles.FirstOrDefault(value =>
            string.Equals(value.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));
        var overwrite = existing is not null;
        if (overwrite)
        {
            var owner = ExcelWindowOwner.TryCreate();
            var response = owner is null
                ? MessageBox.Show($"Replace local style '{existing!.DisplayName}'?", "ExcelAccel", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                : MessageBox.Show(owner, $"Replace local style '{existing!.DisplayName}'?", "ExcelAccel", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (response != DialogResult.Yes) return CommandResult.Refused("style.capture", "Existing style replacement was cancelled.", "USER_CANCELLED");
        }
        var styleId = existing?.StyleId ?? "local." + Guid.NewGuid().ToString("N");
        var recipe = new StyleCaptureCommand().Capture(styleId, displayName, propertyIds, CreateSelectionAdapter());
        ProfileRuntime.SaveLocalStyle(recipe, overwrite);
        return CommandResult.Success("style.capture", $"Saved local style '{recipe.DisplayName}' with {recipe.Properties.Count} formatting properties.");
    }

    public static CommandResult DeleteLocalStyle(string styleId)
    {
        var removed = ProfileRuntime.DeleteLocalStyle(styleId);
        return CommandResult.Success("style.delete_local", removed ? "The local style was deleted." : "The local style was already absent.");
    }

    public static CommandResult ApplyProfileFormatting(string commandId)
    {
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(commandId))
        {
            return CommandResult.Refused(commandId, "Mutation commands are disabled in safe mode or quarantine.", RefusalCodes.CommandQuarantined);
        }

        var port = CreateSelectionAdapter();
        var command = Phase1AFormattingCatalog.Create(commandId);
        var plan = command.Plan(ProfileRuntime.Current, port);
        if (plan.RequiresPreview)
        {
            return CommandResult.Refused(plan, "This command requires exact-plan preview confirmation, which is not available from the compact Ribbon menu.", RefusalCodes.PreviewRequired);
        }

        return command.Execute(plan, port, receiptSink: UndoRuntime.Store);
    }

    public static CommandResult Navigate(string commandId)
    {
        var port = new ExcelNavigationAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        var service = NavigationRuntime.Service;
        bool succeeded;
        switch (commandId)
        {
            case "navigate.sheet.previous": succeeded = service.MoveSheet(port, -1, ProfileRuntime.Current.WrapSheetNavigation); break;
            case "navigate.sheet.next": succeeded = service.MoveSheet(port, 1, ProfileRuntime.Current.WrapSheetNavigation); break;
            case "navigate.cell.a1": succeeded = service.Move(port, NavigationTargetKind.A1); break;
            case "navigate.used.first": succeeded = service.Move(port, NavigationTargetKind.UsedFirst); break;
            case "navigate.used.last": succeeded = service.Move(port, NavigationTargetKind.UsedLast); break;
            case "navigate.region.edge.up": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeUp); break;
            case "navigate.region.edge.down": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeDown); break;
            case "navigate.region.edge.left": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeLeft); break;
            case "navigate.region.edge.right": succeeded = service.Move(port, NavigationTargetKind.RegionEdgeRight); break;
            case "navigate.history.back": succeeded = service.Back(port); break;
            case "navigate.history.forward": succeeded = service.Forward(port); break;
            case "navigate.bookmark.next_session": succeeded = service.NextBookmark(port); break;
            case "navigate.bookmark.previous_session": succeeded = service.PreviousBookmark(port); break;
            case "navigate.bookmark.add_session": service.AddBookmark(port); succeeded = true; break;
            case "navigate.bookmark.clear_session": service.ClearBookmarks(); succeeded = true; break;
            default: return CommandResult.Refused(commandId, "The navigation command is not available.", RefusalCodes.CommandUnavailable);
        }

        var location = port.CaptureLocation();
        var plan = new CommandPlan(commandId, CommandImpact.ReadOnly,
            new SelectionContext(location.WorkbookId, location.WorksheetName, location.Address), new string[0], 0, "Read-only navigation.");
        return succeeded ? CommandResult.Success(plan, "Navigation completed without changing workbook content.") :
            CommandResult.Refused(plan, "No valid navigation target is available.", RefusalCodes.SelectionUnsupported);
    }

    public static CommandResult ApplyFormulaCommand(string commandId)
    {
        if (RuntimeState.IsSafeMode || RuntimeState.IsQuarantined(commandId))
            return CommandResult.Refused(commandId, "Formula mutation is disabled in safe mode or quarantine.", RefusalCodes.CommandQuarantined);
        var descriptor = FormulaCommandCatalog.GetRequired(commandId);
        var port = CreateSelectionAdapter();
        var command = new FormulaBlockCommand(descriptor);
        var snapshot = port.CaptureFormulaBlock();
        var previewLimit = checked((int)Math.Min(ProfileRuntime.Current.ImmediatePreviewCellLimit, int.MaxValue));
        FormulaBlockPlan plan;
        switch (commandId)
        {
            case "formula.copy.down":
                plan = command.PlanCopy(snapshot, FormulaCopyDirection.Down, previewLimit);
                break;
            case "formula.copy.right":
                plan = command.PlanCopy(snapshot, FormulaCopyDirection.Right, previewLimit);
                break;
            case "formula.iferror.toggle":
                plan = command.PlanIfError(snapshot, ProfileRuntime.Current.FormulaIfErrorFallback, previewLimit);
                break;
            case "formula.sign.reverse":
                plan = command.PlanReverseSign(snapshot, includeNumericConstants: false, previewLimit);
                break;
            case "formula.units.to_thousands":
                plan = command.PlanScale(snapshot, 1000, divide: true, includeNumericConstants: false, previewLimit);
                break;
            case "formula.units.from_thousands":
                plan = command.PlanScale(snapshot, 1000, divide: false, includeNumericConstants: false, previewLimit);
                break;
            case "formula.units.to_millions":
                plan = command.PlanScale(snapshot, 1000000, divide: true, includeNumericConstants: false, previewLimit);
                break;
            case "formula.units.from_millions":
                plan = command.PlanScale(snapshot, 1000000, divide: false, includeNumericConstants: false, previewLimit);
                break;
            default:
                return CommandResult.Refused(commandId, "The formula command has no qualified host route.", RefusalCodes.CommandUnavailable);
        }

        string? confirmation = null;
        if (plan.CommandPlan.RequiresPreview)
        {
            var preview = plan.CommandPlan.Summary + "\n\n" + string.Join("\n", plan.Samples) + "\n\nApply this exact plan?";
            var owner = ExcelWindowOwner.TryCreate();
            var response = owner is null
                ? MessageBox.Show(preview, "ExcelAccel formula preview", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                : MessageBox.Show(owner, preview, "ExcelAccel formula preview", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (response != DialogResult.Yes) return CommandResult.Refused(plan.CommandPlan, "Formula preview was cancelled.", "USER_CANCELLED");
            confirmation = plan.CommandPlan.PlanHash;
        }
        return command.Execute(plan, port, confirmation, UndoRuntime.Store);
    }

    private static ExcelSelectionAdapter CreateSelectionAdapter() =>
        new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
}
