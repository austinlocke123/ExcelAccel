using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Styles;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class StyleRecipeTests
{
    [Fact]
    public void CaptureReadsExactlySelectedSupportedFormattingPropertiesFromOneCell()
    {
        var port = Port(1, ("font_bold", "true"), ("font_color", "#112233"), ("number_format", "0.0%"));
        var recipe = new StyleCaptureCommand().Capture("local.test", "Test", new[] { "font_color", "font_bold" }, port);

        Assert.Equal(new[] { "font_bold", "font_color" }, recipe.Properties.Keys);
        Assert.Equal(new[] { "font_bold", "font_color" }, port.ReadProperties.OrderBy(value => value, StringComparer.Ordinal));
        Assert.DoesNotContain(recipe.Properties.Keys, value => value.Contains("formula", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void CaptureRefusesUnsupportedPropertyWithoutReadingAnything()
    {
        var port = Port(1, ("font_color", "#112233"));

        var refusal = Assert.Throws<CommandRefusedException>(() =>
            new StyleCaptureCommand().Capture("local.bad", "Bad", new[] { "font_color", "formula" }, port));

        Assert.Equal(RefusalCodes.CommandUnavailable, refusal.RefusalCode);
        Assert.Empty(port.ReadProperties);
    }

    [Fact]
    public void BuiltInsAreVersionedValidAndCannotBeDeleted()
    {
        Assert.Equal(9, BuiltInStyleCatalog.All.Count);
        Assert.All(BuiltInStyleCatalog.All, style =>
        {
            Assert.Equal(StyleRecipe.CurrentVersion, style.Version);
            Assert.Equal(StyleOrigin.BuiltIn, style.Origin);
            Assert.False(style.IsDeletable);
            Assert.NotEmpty(style.Properties);
        });
        Assert.Throws<InvalidOperationException>(() => StyleLibrary.Delete(Array.Empty<StyleRecipe>(), "major_header"));
    }

    [Fact]
    public void ApplyUsesOneExecutablePlanAndBatchReceiptUndoRestoresEveryProperty()
    {
        var recipe = Recipe(("font_bold", "true"), ("font_color", "#0000FF"));
        var port = Port(10, ("font_bold", "false"), ("font_color", "#000000"));
        var store = new SessionUndoStore();
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired("style.apply"));
        var plan = command.Plan(recipe, port, 50_000);

        var result = command.Execute(plan, port, receiptSink: store);

        Assert.True(result.Succeeded);
        Assert.Equal(new[] { "font_bold", "font_color" }, plan.CommandPlan.ChangedProperties);
        Assert.Equal("true", port.Values["font_bold"]);
        Assert.Equal("#0000FF", port.Values["font_color"]);
        Assert.Equal(1, store.Count("Book.xlsx"));

        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);
        Assert.True(undo.Succeeded);
        Assert.Equal("false", port.Values["font_bold"]);
        Assert.Equal("#000000", port.Values["font_color"]);
    }

    [Fact]
    public void ApplyFailureRollsBackAllAttemptedPropertiesBeforeReturningFailure()
    {
        var recipe = Recipe(("font_bold", "true"), ("font_color", "#0000FF"));
        var port = Port(10, ("font_bold", "false"), ("font_color", "#000000"));
        port.ThrowAfterWriteProperty = "font_color";
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired("style.apply"));

        var result = command.Execute(command.Plan(recipe, port, 50_000), port);

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.Equal("STYLE_APPLY_ROLLED_BACK", result.DiagnosticId);
        Assert.Equal("false", port.Values["font_bold"]);
        Assert.Equal("#000000", port.Values["font_color"]);
    }

    [Fact]
    public void IncompleteRollbackReportsExactRemainingChangedProperty()
    {
        var recipe = Recipe(("font_bold", "true"), ("font_color", "#0000FF"));
        var port = Port(10, ("font_bold", "false"), ("font_color", "#000000"));
        port.ThrowAfterWriteProperty = "font_color";
        port.RefuseRollbackProperty = "font_bold";
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired("style.apply"));

        var result = command.Execute(command.Plan(recipe, port, 50_000), port);

        Assert.Equal(CommandResultStatus.Partial, result.Status);
        Assert.Equal("STYLE_ROLLBACK_INCOMPLETE", result.DiagnosticId);
        Assert.Contains("font_bold", result.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("font_color", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ThresholdPreviewRequiresTheExactStylePlanHash()
    {
        var recipe = Recipe(("font_bold", "true"));
        var port = Port(10, ("font_bold", "false"));
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired("style.apply"));
        var plan = command.Plan(recipe, port, immediatePreviewCellLimit: 5);

        var refused = command.Execute(plan, port);
        var applied = command.Execute(plan, port, plan.CommandPlan.PlanHash);

        Assert.Equal(RefusalCodes.PreviewRequired, refused.RefusalCode);
        Assert.True(applied.Succeeded);
    }

    [Fact]
    public void BatchUndoCompensatesAWriteThatReportsFailureAfterMutation()
    {
        var recipe = Recipe(("font_bold", "true"), ("font_color", "#0000FF"));
        var port = Port(10, ("font_bold", "false"), ("font_color", "#000000"));
        var store = new SessionUndoStore();
        var command = new StyleApplyCommand(StyleCommandCatalog.GetRequired("style.apply"));
        Assert.True(command.Execute(command.Plan(recipe, port, 50_000), port, receiptSink: store).Succeeded);
        port.ThrowAfterWriteProperty = "font_color";

        var undo = store.TryUndo("Book.xlsx", port, DateTimeOffset.UtcNow);

        Assert.Equal(UndoOutcome.WriteFailed, undo.Outcome);
        Assert.Equal("true", port.Values["font_bold"]);
        Assert.Equal("#0000FF", port.Values["font_color"]);
    }

    private static StyleRecipe Recipe(params (string Id, string Value)[] properties) =>
        new StyleRecipe("local.test", 1, "Test", StyleOrigin.Local, UnsupportedStylePropertyPolicy.Refuse,
            properties.Select(value => new KeyValuePair<string, string>(value.Id, value.Value)));

    private static FakeStylePort Port(long cells, params (string Id, string Value)[] properties) =>
        new FakeStylePort(cells, properties.ToDictionary(value => value.Id, value => value.Value, StringComparer.Ordinal));

    private sealed class FakeStylePort : IStylePort
    {
        private readonly SelectionSnapshot _snapshot;
        public FakeStylePort(long cells, Dictionary<string, string> values)
        {
            Values = values;
            _snapshot = new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", cells == 1 ? "A1" : $"A1:A{cells}"), cells,
                false, "General", new SelectionSafetyState(1, false, false, false, false, false, true));
        }
        public Dictionary<string, string> Values { get; }
        public List<string> ReadProperties { get; } = new List<string>();
        public string? ThrowAfterWriteProperty { get; set; }
        public string? RefuseRollbackProperty { get; set; }
        public SelectionSnapshot CaptureSelection() => _snapshot;
        public void SetNumberFormat(string formatCode) => Values["number_format"] = formatCode;
        public string ReadFormattingProperty(string propertyId)
        {
            ReadProperties.Add(propertyId);
            return Values.TryGetValue(propertyId, out var value) ? value : throw new CommandRefusedException(RefusalCodes.CommandUnavailable, "Unsupported.", "Remove it.");
        }
        public void WriteFormattingProperty(string propertyId, string invariantValue)
        {
            if (propertyId == RefuseRollbackProperty && Values.TryGetValue(propertyId, out var current) && invariantValue != current)
            {
                if ((propertyId == "font_bold" && invariantValue == "false") || (propertyId == "font_color" && invariantValue == "#000000"))
                    throw new InvalidOperationException("Injected rollback failure.");
            }
            Values[propertyId] = invariantValue;
            if (propertyId == ThrowAfterWriteProperty)
            {
                ThrowAfterWriteProperty = null;
                throw new InvalidOperationException("Injected write failure after mutation.");
            }
        }
        public bool TryRead(SelectionContext target, string propertyId, out string value)
        { value = string.Empty; if (!target.Equals(_snapshot.Context) || !Values.TryGetValue(propertyId, out var found)) return false; value = found; return true; }
        public bool TryWrite(SelectionContext target, string propertyId, string value)
        { if (!target.Equals(_snapshot.Context)) return false; try { WriteFormattingProperty(propertyId, value); return true; } catch { return false; } }
    }
}
