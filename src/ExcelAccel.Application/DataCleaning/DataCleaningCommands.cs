using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.DataCleaning;

public enum DisplayValueConversion
{
    BlankToZero,
    ZeroToBlank,
    BlankToNaText,
    BlankToNmText,
    BlankToDashText,
    NaTextToBlank,
    NmTextToBlank,
    DashTextToBlank,
}

public sealed class TextNumberConversionOptions
{
    public TextNumberConversionOptions(char decimalSeparator, char thousandsSeparator,
        bool allowLeadingSign, bool allowParentheses, IEnumerable<string> currencySymbols, bool allowPercent)
    {
        if (decimalSeparator == thousandsSeparator) throw new ArgumentException("Decimal and thousands separators must differ.");
        if (char.IsDigit(decimalSeparator) || char.IsDigit(thousandsSeparator)) throw new ArgumentException("Numeric separators cannot be digits.");
        DecimalSeparator = decimalSeparator;
        ThousandsSeparator = thousandsSeparator;
        AllowLeadingSign = allowLeadingSign;
        AllowParentheses = allowParentheses;
        CurrencySymbols = Array.AsReadOnly((currencySymbols ?? throw new ArgumentNullException(nameof(currencySymbols)))
            .Where(value => !string.IsNullOrEmpty(value)).Distinct(StringComparer.Ordinal).OrderByDescending(value => value.Length).ToArray());
        AllowPercent = allowPercent;
    }

    public char DecimalSeparator { get; }
    public char ThousandsSeparator { get; }
    public bool AllowLeadingSign { get; }
    public bool AllowParentheses { get; }
    public IReadOnlyList<string> CurrencySymbols { get; }
    public bool AllowPercent { get; }

    public static TextNumberConversionOptions InvariantFinancial { get; } =
        new TextNumberConversionOptions('.', ',', true, true, new[] { "$" }, true);
}

public sealed class DataCleaningCommand
{
    private readonly CommandDescriptor _descriptor;

    public DataCleaningCommand(CommandDescriptor descriptor) =>
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

    public FormulaBlockPlan PlanTrimOuter(FormulaBlockSnapshot snapshot) =>
        PlanText(snapshot, "Trim outer Unicode whitespace", TrimOuter);

    public FormulaBlockPlan PlanCollapseWhitespace(FormulaBlockSnapshot snapshot) =>
        PlanText(snapshot, "Collapse Unicode whitespace", CollapseWhitespace);

    public FormulaBlockPlan PlanRemoveNonprinting(FormulaBlockSnapshot snapshot, bool preserveTabsAndNewlines) =>
        PlanText(snapshot, preserveTabsAndNewlines ? "Remove nonprinting controls; preserve tabs/newlines" : "Remove nonprinting controls",
            value => RemoveNonprinting(value, preserveTabsAndNewlines));

    public FormulaBlockPlan PlanDisplayConversion(FormulaBlockSnapshot snapshot, DisplayValueConversion conversion)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            if (!TryConvertDisplay(current, conversion, out var proposed)) { skipped++; return current; }
            if (current.Equals(proposed)) { skipped++; return current; }
            changed++;
            AddSample(samples, row, column, current, proposed!);
            return proposed!;
        });
        return Build(snapshot, after, changed, skipped, samples, requiresPreview: true,
            $"{DisplayName(conversion)}: {changed:N0} matching constants changed; {skipped:N0} formulas/nonmatches skipped.",
            new[] { Pair("conversion", conversion.ToString().ToLowerInvariant()) });
    }

    public FormulaBlockPlan PlanTextToNumber(FormulaBlockSnapshot snapshot, TextNumberConversionOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return PlanTyped(snapshot, "Invariant text to number", current =>
        {
            if (current.Kind != FormulaCellKind.Text || !TryParseCompleteNumber(current.InvariantValue, options, out var number)) return null;
            return FormulaCellValue.Number(number);
        }, new[]
        {
            Pair("decimal_separator", options.DecimalSeparator.ToString()),
            Pair("thousands_separator", options.ThousandsSeparator.ToString()),
            Pair("leading_sign", options.AllowLeadingSign ? "allow" : "refuse"),
            Pair("parentheses", options.AllowParentheses ? "allow_negative" : "refuse"),
            Pair("currency_symbols", string.Join("|", options.CurrencySymbols)),
            Pair("percent", options.AllowPercent ? "allow_divide_100" : "refuse"),
        });
    }

    public FormulaBlockPlan PlanNumberToText(FormulaBlockSnapshot snapshot, string invariantFormat)
    {
        if (string.IsNullOrWhiteSpace(invariantFormat)) throw new ArgumentException("An explicit invariant numeric format is required.", nameof(invariantFormat));
        return PlanTyped(snapshot, "Number to invariant text", current =>
        {
            if (current.Kind != FormulaCellKind.Number) return null;
            string output;
            try { output = current.AsNumber().ToString(invariantFormat, CultureInfo.InvariantCulture); }
            catch (FormatException exception) { throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "The explicit number-to-text format is invalid: " + exception.Message); }
            return FormulaCellValue.Text(output);
        }, new[] { Pair("invariant_format", invariantFormat) });
    }

    public FormulaBlockPlan PlanNormalizeDateText(FormulaBlockSnapshot snapshot, IEnumerable<string> inputPatterns, string outputPattern)
    {
        var patterns = (inputPatterns ?? throw new ArgumentNullException(nameof(inputPatterns))).ToArray();
        if (patterns.Length == 0 || patterns.Any(string.IsNullOrWhiteSpace)) throw new ArgumentException("At least one explicit date input pattern is required.", nameof(inputPatterns));
        if (string.IsNullOrWhiteSpace(outputPattern)) throw new ArgumentException("An explicit date output pattern is required.", nameof(outputPattern));
        return PlanTyped(snapshot, "Normalize explicit date text", current =>
        {
            if (current.Kind != FormulaCellKind.Text) return null;
            if (!DateTime.TryParseExact(current.InvariantValue, patterns, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var date)) return null;
            string output;
            try { output = date.ToString(outputPattern, CultureInfo.InvariantCulture); }
            catch (FormatException exception) { throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "The explicit date output pattern is invalid: " + exception.Message); }
            return FormulaCellValue.Text(output);
        }, new[]
        {
            Pair("input_patterns", string.Join("|", patterns)),
            Pair("output_pattern", outputPattern),
            Pair("calendar", "gregorian"),
            Pair("timezone", "none_date_only"),
        });
    }

    public static string TrimOuter(string value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        var start = 0;
        while (start < value.Length && IsApprovedWhitespace(value[start])) start++;
        var end = value.Length - 1;
        while (end >= start && IsApprovedWhitespace(value[end])) end--;
        return value.Substring(start, end - start + 1);
    }

    public static string CollapseWhitespace(string value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        var trimmed = TrimOuter(value);
        var builder = new StringBuilder(trimmed.Length);
        var inWhitespace = false;
        foreach (var character in trimmed)
        {
            if (IsApprovedWhitespace(character))
            {
                inWhitespace = true;
                continue;
            }
            if (inWhitespace && builder.Length > 0) builder.Append(' ');
            builder.Append(character);
            inWhitespace = false;
        }
        return builder.ToString();
    }

    public static string RemoveNonprinting(string value, bool preserveTabsAndNewlines)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            var preservedLineCharacter = preserveTabsAndNewlines && (character == '\t' || character == '\n' || character == '\r');
            var disallowed = (character >= '\u0000' && character <= '\u001f') ||
                             (character >= '\u007f' && character <= '\u009f');
            if (!disallowed || preservedLineCharacter) builder.Append(character);
        }
        return builder.ToString();
    }

    public static bool IsApprovedWhitespace(char value) =>
        (value >= '\u0009' && value <= '\u000d') || value == '\u0020' || value == '\u0085' ||
        value == '\u00a0' || value == '\u1680' || (value >= '\u2000' && value <= '\u200a') ||
        value == '\u2028' || value == '\u2029' || value == '\u202f' || value == '\u205f' || value == '\u3000';

    public static bool TryParseCompleteNumber(string source, TextNumberConversionOptions options, out double value)
    {
        value = 0;
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (source.Length == 0 || source.Any(IsApprovedWhitespace)) return false;
        var text = source;
        var negative = false;
        if (text.Length >= 2 && text[0] == '(' && text[text.Length - 1] == ')')
        {
            if (!options.AllowParentheses) return false;
            negative = true;
            text = text.Substring(1, text.Length - 2);
        }
        else if (text.Length > 0 && (text[0] == '+' || text[0] == '-'))
        {
            if (!options.AllowLeadingSign) return false;
            negative = text[0] == '-';
            text = text.Substring(1);
        }
        foreach (var symbol in options.CurrencySymbols)
        {
            if (!text.StartsWith(symbol, StringComparison.Ordinal)) continue;
            text = text.Substring(symbol.Length);
            break;
        }
        var percent = text.EndsWith("%", StringComparison.Ordinal);
        if (percent)
        {
            if (!options.AllowPercent) return false;
            text = text.Substring(0, text.Length - 1);
        }
        if (text.Length == 0 || text.IndexOf('+') >= 0 || text.IndexOf('-') >= 0 || text.IndexOf('(') >= 0 || text.IndexOf(')') >= 0 || text.IndexOf('%') >= 0)
            return false;
        var decimalIndex = text.IndexOf(options.DecimalSeparator);
        if (decimalIndex != text.LastIndexOf(options.DecimalSeparator)) return false;
        var integerPart = decimalIndex < 0 ? text : text.Substring(0, decimalIndex);
        var fractionalPart = decimalIndex < 0 ? string.Empty : text.Substring(decimalIndex + 1);
        if (integerPart.Length == 0 || (decimalIndex >= 0 && fractionalPart.Length == 0) ||
            fractionalPart.Any(character => character < '0' || character > '9')) return false;
        var groups = integerPart.Split(options.ThousandsSeparator);
        if (groups.Any(group => group.Length == 0 || group.Any(character => character < '0' || character > '9'))) return false;
        if (groups.Length > 1)
        {
            if (groups[0].Length < 1 || groups[0].Length > 3 || groups.Skip(1).Any(group => group.Length != 3)) return false;
        }
        var canonical = string.Concat(groups) + (decimalIndex < 0 ? string.Empty : "." + fractionalPart);
        if (!double.TryParse(canonical, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var parsed) ||
            double.IsNaN(parsed) || double.IsInfinity(parsed)) return false;
        value = negative ? -parsed : parsed;
        if (percent) value /= 100d;
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private FormulaBlockPlan PlanText(FormulaBlockSnapshot snapshot, string label, Func<string, string> transform)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var sourceKinds = new HashSet<FormulaCellKind>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            sourceKinds.Add(current.Kind);
            if (current.Kind != FormulaCellKind.Text) { skipped++; return current; }
            var output = transform(current.InvariantValue);
            if (string.Equals(output, current.InvariantValue, StringComparison.Ordinal)) { skipped++; return current; }
            var proposed = FormulaCellValue.Text(output);
            changed++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        var mixed = sourceKinds.Count > 1;
        return Build(snapshot, after, changed, skipped, samples, mixed || changed > FormulaBlockCommand.DefaultImmediatePreviewLimit,
            $"{label}: {changed:N0} text constants changed; {skipped:N0} formulas/nontext/already-normalized cells skipped.",
            new[] { Pair("source_policy", "text_constants_only") });
    }

    private FormulaBlockPlan PlanTyped(FormulaBlockSnapshot snapshot, string label,
        Func<FormulaCellValue, FormulaCellValue?> transform, IEnumerable<KeyValuePair<string, string>> arguments)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            if (current.IsFormula) { skipped++; return current; }
            var proposed = transform(current);
            if (proposed is null || proposed.Equals(current)) { skipped++; return current; }
            changed++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        return Build(snapshot, after, changed, skipped, samples, requiresPreview: true,
            $"{label}: {changed:N0} exact matches changed; {skipped:N0} formulas/nonmatches/already-normalized cells skipped.",
            arguments);
    }

    private FormulaBlockPlan Build(FormulaBlockSnapshot snapshot, FormulaCellBlock after, int changed, int skipped,
        IEnumerable<string> samples, bool requiresPreview, string summary, IEnumerable<KeyValuePair<string, string>> arguments)
    {
        var fullArguments = arguments.Concat(new[]
        {
            Pair("before_sha256", snapshot.Contents.Fingerprint),
            Pair("after_sha256", after.Fingerprint),
            Pair("changed", changed.ToString(CultureInfo.InvariantCulture)),
            Pair("skipped", skipped.ToString(CultureInfo.InvariantCulture)),
        });
        var plan = new CommandPlan(_descriptor.Id, _descriptor.Impact, snapshot.Selection.Context,
            new[] { "value" }, changed, summary, snapshot.Contents.Fingerprint, _descriptor.ContractVersion,
            requiresPreview, fullArguments);
        return new FormulaBlockPlan(plan, snapshot, after, changed, skipped, samples);
    }

    private static bool TryConvertDisplay(FormulaCellValue current, DisplayValueConversion conversion, out FormulaCellValue? output)
    {
        output = null;
        if (current.IsFormula) return false;
        switch (conversion)
        {
            case DisplayValueConversion.BlankToZero:
                if (!current.IsBlank) return false; output = FormulaCellValue.Number(0); return true;
            case DisplayValueConversion.ZeroToBlank:
                if (current.Kind != FormulaCellKind.Number || current.AsNumber() != 0d) return false; output = FormulaCellValue.Blank(); return true;
            case DisplayValueConversion.BlankToNaText:
                if (!current.IsBlank) return false; output = FormulaCellValue.Text("N/A"); return true;
            case DisplayValueConversion.BlankToNmText:
                if (!current.IsBlank) return false; output = FormulaCellValue.Text("NM"); return true;
            case DisplayValueConversion.BlankToDashText:
                if (!current.IsBlank) return false; output = FormulaCellValue.Text("-"); return true;
            case DisplayValueConversion.NaTextToBlank:
                if (current.Kind != FormulaCellKind.Text || current.InvariantValue != "N/A") return false; output = FormulaCellValue.Blank(); return true;
            case DisplayValueConversion.NmTextToBlank:
                if (current.Kind != FormulaCellKind.Text || current.InvariantValue != "NM") return false; output = FormulaCellValue.Blank(); return true;
            case DisplayValueConversion.DashTextToBlank:
                if (current.Kind != FormulaCellKind.Text || current.InvariantValue != "-") return false; output = FormulaCellValue.Blank(); return true;
            default: return false;
        }
    }

    private static string DisplayName(DisplayValueConversion conversion)
    {
        switch (conversion)
        {
            case DisplayValueConversion.BlankToZero: return "Blank to numeric zero";
            case DisplayValueConversion.ZeroToBlank: return "Numeric zero to blank";
            case DisplayValueConversion.BlankToNaText: return "Blank to exact N/A text";
            case DisplayValueConversion.BlankToNmText: return "Blank to exact NM text";
            case DisplayValueConversion.BlankToDashText: return "Blank to exact dash text";
            case DisplayValueConversion.NaTextToBlank: return "Exact N/A text to blank";
            case DisplayValueConversion.NmTextToBlank: return "Exact NM text to blank";
            case DisplayValueConversion.DashTextToBlank: return "Exact dash text to blank";
            default: return conversion.ToString();
        }
    }

    private static void RequireSafe(FormulaBlockSnapshot snapshot)
    {
        var safety = snapshot.Selection.Safety;
        if (safety.AreaCount != 1 || safety.HasMergedCells) throw Refuse(RefusalCodes.SelectionUnsupported, "Data cleaning requires one unmerged rectangle.");
        if (safety.WorksheetProtected || safety.WorkbookReadOnly) throw Refuse(RefusalCodes.ProtectedTarget, "The target is protected or read-only.");
        if (!safety.DynamicArraySpillCheckSupported || safety.HasLegacyArray || safety.HasDynamicArraySpill)
            throw Refuse(RefusalCodes.ArrayOrSpillUnsafe, "The target intersects an unqualified array or spill state.");
    }

    private static KeyValuePair<string, string> Pair(string key, string value) => new KeyValuePair<string, string>(key, value);
    private static void AddSample(ICollection<string> samples, int row, int column, FormulaCellValue before, FormulaCellValue after)
    {
        if (samples.Count < 10) samples.Add($"R{row + 1}C{column + 1}: {before.Kind} {Escape(before.InvariantValue)} -> {after.Kind} {Escape(after.InvariantValue)}");
    }
    private static string Escape(string value) => value.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    private static CommandRefusedException Refuse(string code, string message) =>
        new CommandRefusedException(code, message, "Use one bounded, editable, unmerged range outside arrays and spills.");
}
