using System;
using System.Collections.Generic;

namespace ExcelAccel.Application.Commands;

/// <summary>
/// The keyboard route each command is reachable by, derived from the ribbon
/// definition itself.
///
/// This exists so a descriptor cannot drift from the ribbon. Every catalog reads
/// its route from here rather than composing one from a KeyTip fragment, which
/// is what previously let the two disagree without anything noticing.
///
/// Regenerate this file whenever the ribbon layout changes.
/// </summary>
public static class RibbonRoutes
{
    private static readonly IReadOnlyDictionary<string, string> Routes = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        { "assumption", "Alt, X, A, EY, A" },
        { "audit.dependents.direct", "Alt, X, A, AD, D" },
        { "audit.dependents.indirect", "Alt, X, A, AD, I" },
        { "audit.dependents.workbook", "Alt, X, A, AD, W" },
        { "audit.formula.inspect", "Alt, X, A, AF" },
        { "audit.precedents.direct", "Alt, X, A, AP, D" },
        { "audit.precedents.indirect", "Alt, X, A, AP, I" },
        { "bindings.cheat_sheet.export", "Alt, X, A, S, B" },
        { "clean.convert.date_normalize", "Alt, X, A, KC, DN" },
        { "clean.convert.number_to_text", "Alt, X, A, KC, NT" },
        { "clean.convert.text_to_number", "Alt, X, A, KC, TN" },
        { "clean.display.blank_to_dash_text", "Alt, X, A, KC, BD" },
        { "clean.display.blank_to_na_text", "Alt, X, A, KC, BN" },
        { "clean.display.blank_to_nm_text", "Alt, X, A, KC, BM" },
        { "clean.display.blank_to_zero", "Alt, X, A, KC, BZ" },
        { "clean.display.dash_text_to_blank", "Alt, X, A, KC, DB" },
        { "clean.display.na_text_to_blank", "Alt, X, A, KC, NB" },
        { "clean.display.nm_text_to_blank", "Alt, X, A, KC, MB" },
        { "clean.display.zero_to_blank", "Alt, X, A, KC, ZB" },
        { "clean.text.collapse_whitespace", "Alt, X, A, KC, CW" },
        { "clean.text.remove_nonprinting", "Alt, X, A, KC, RN" },
        { "clean.text.trim_outer", "Alt, X, A, KC, TO" },
        { "command.search.open", "Alt, X, A, Q" },
        { "date_header", "Alt, X, A, EY, DH" },
        { "favorite.add", "Alt, X, A, Q, then Ctrl+D" },
        { "favorite.invoke", "Alt, X, A, Q, then Enter on a favorite" },
        { "favorite.remove", "Alt, X, A, Q, then Ctrl+Shift+D" },
        { "fill.date_sequence", "Alt, X, A, RF, D" },
        { "fill.formula_from_above", "Alt, X, A, RF, F" },
        { "fill.numeric_sequence", "Alt, X, A, RF, N" },
        { "fill.value_from_above", "Alt, X, A, RF, V" },
        { "format.alignment.horizontal.cycle", "Alt, X, A, EA, H" },
        { "format.alignment.vertical.cycle", "Alt, X, A, EA, V" },
        { "format.border.remove", "Alt, X, A, EB, R" },
        { "format.border.sum_bar.apply", "Alt, X, A, EB, S" },
        { "format.center_across.apply", "Alt, X, A, EC" },
        { "format.column_width.cycle", "Alt, X, A, ES, C" },
        { "format.columns.autofit", "Alt, X, A, ES, W" },
        { "format.fill_color.cycle", "Alt, X, A, L" },
        { "format.font_color.cycle", "Alt, X, A, F" },
        { "format.font_size.cycle", "Alt, X, A, EF, S" },
        { "format.indent.decrease", "Alt, X, A, EA, D" },
        { "format.indent.increase", "Alt, X, A, EA, I" },
        { "format.number.boolean", "Alt, X, A, NB" },
        { "format.number.currency", "Alt, X, A, C" },
        { "format.number.date", "Alt, X, A, D" },
        { "format.number.decimals.decrease", "Alt, X, A, ND" },
        { "format.number.decimals.increase", "Alt, X, A, NI" },
        { "format.number.general", "Alt, X, A, G" },
        { "format.number.multiple", "Alt, X, A, NM" },
        { "format.number.percentage", "Alt, X, A, P" },
        { "format.row_height.cycle", "Alt, X, A, ES, R" },
        { "format.rows.autofit", "Alt, X, A, ES, A" },
        { "format.underline.cycle", "Alt, X, A, EF, U" },
        { "formula", "Alt, X, A, EY, F" },
        { "formula.copy.down", "Alt, X, A, RC, D" },
        { "formula.copy.right", "Alt, X, A, RC, R" },
        { "formula.iferror.toggle", "Alt, X, A, RE" },
        { "formula.sign.reverse", "Alt, X, A, RV" },
        { "formula.source.capture", "Alt, X, A, RP, C" },
        { "formula.spacing.columns", "Alt, X, A, RS, C" },
        { "formula.spacing.rows", "Alt, X, A, RS, R" },
        { "formula.transpose", "Alt, X, A, RP, R" },
        { "formula.units.from_millions", "Alt, X, A, RU, N" },
        { "formula.units.from_thousands", "Alt, X, A, RU, F" },
        { "formula.units.to_millions", "Alt, X, A, RU, M" },
        { "formula.units.to_thousands", "Alt, X, A, RU, T" },
        { "inspect.selection.summary", "Alt, X, A, I" },
        { "linked_formula", "Alt, X, A, EY, K" },
        { "major_header", "Alt, X, A, EY, MH" },
        { "minor_header", "Alt, X, A, EY, MI" },
        { "model_check.finding.ignore_local", "Alt, X, A, AM, I" },
        { "model_check.finding.unignore_local", "Alt, X, A, AM, M" },
        { "model_check.rescan", "Alt, X, A, AM, R" },
        { "model_check.results.export", "Alt, X, A, AM, E" },
        { "model_check.run.selection", "Alt, X, A, AM, S" },
        { "model_check.run.workbook", "Alt, X, A, AM, B" },
        { "model_check.run.worksheet", "Alt, X, A, AM, W" },
        { "navigate.bookmark.add_session", "Alt, X, A, VN, M" },
        { "navigate.bookmark.clear_session", "Alt, X, A, VN, C" },
        { "navigate.bookmark.next_session", "Alt, X, A, VN, J" },
        { "navigate.bookmark.previous_session", "Alt, X, A, VN, K" },
        { "navigate.cell.a1", "Alt, X, A, VN, A" },
        { "navigate.history.back", "Alt, X, A, VN, B" },
        { "navigate.history.forward", "Alt, X, A, VN, O" },
        { "navigate.region.edge.down", "Alt, X, A, VN, D" },
        { "navigate.region.edge.left", "Alt, X, A, VN, E" },
        { "navigate.region.edge.right", "Alt, X, A, VN, R" },
        { "navigate.region.edge.up", "Alt, X, A, VN, U" },
        { "navigate.sheet.next", "Alt, X, A, VN, N" },
        { "navigate.sheet.previous", "Alt, X, A, VN, P" },
        { "navigate.used.first", "Alt, X, A, VN, F" },
        { "navigate.used.last", "Alt, X, A, VN, L" },
        { "output", "Alt, X, A, EY, O" },
        { "paste.formats_only", "Alt, X, A, RP, T" },
        { "paste.formulas_only", "Alt, X, A, RP, F" },
        { "paste.values_only", "Alt, X, A, RP, V" },
        { "profile.export", "Alt, X, A, S, E" },
        { "profile.import.apply", "Alt, X, A, S, I" },
        { "profile.import.preview", "Alt, X, A, S, I (preview)" },
        { "selection.select.blanks", "Alt, X, A, KS, B" },
        { "selection.select.constants", "Alt, X, A, KS, C" },
        { "selection.select.external_formulas", "Alt, X, A, KS, X" },
        { "selection.select.formulas", "Alt, X, A, KS, F" },
        { "selection.select.numeric_hardcodes", "Alt, X, A, KS, N" },
        { "style.apply", "Alt, X, A, EY, L" },
        { "style.apply_builtin", "Alt, X, A, EY" },
        { "style.capture", "Alt, X, A, EY, L, then Capture New (Ctrl+N)" },
        { "style.delete_local", "Alt, X, A, EY, L, then Delete Local (Del)" },
        { "support.diagnostics.export", "Alt, X, A, S, D" },
        { "total", "Alt, X, A, EY, T" },
        { "undo.last_excelaccel_property", "Alt, X, A, U" },
        { "view.gridlines.toggle", "Alt, X, A, VG" },
        { "view.panes.freeze", "Alt, X, A, VF" },
        { "view.panes.unfreeze", "Alt, X, A, VU" },
        { "view.zoom.set", "Alt, X, A, VZ" },
        { "warning", "Alt, X, A, EY, W" },
    };

    public static IReadOnlyDictionary<string, string> All => Routes;

    /// <summary>
    /// The route for a command. A command with no ribbon entry falls back to a
    /// Command Search route, which is always true: every registered command is
    /// reachable there.
    /// </summary>
    public static string For(string commandId) =>
        Routes.TryGetValue(commandId, out var route) ? route : "Command Search: Alt, X, A, Q";

    public static bool Has(string commandId) => Routes.ContainsKey(commandId);
}
