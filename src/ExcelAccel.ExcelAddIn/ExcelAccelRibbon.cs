using ExcelDna.Integration.CustomUI;
using ExcelAccel.Application.Commands;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

public sealed class ExcelAccelRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) =>
        @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
            <ribbon>
              <tabs>
                <tab id='ExcelAccel.Tab' label='ExcelAccel' keytip='XA'>
                  <group id='ExcelAccel.Group.Diagnostics' label='Safe Tools'>
                    <button id='ExcelAccel.CommandSearch'
                            label='Search Commands'
                            keytip='Q'
                            size='large'
                            imageMso='FindDialog'
                            screentip='Search ExcelAccel commands'
                            supertip='Searches local command metadata and shows current availability without scanning or changing the workbook.'
                            onAction='OnOpenCommandSearch'/>
                    <button id='ExcelAccel.InspectSelection'
                            label='Inspect Selection'
                            keytip='I'
                            size='large'
                            imageMso='ReviewShowAllMarkup'
                            screentip='Read selection metadata'
                            supertip='Reads selection identity, size, formula state, and number format. It does not change the workbook.'
                            onAction='OnInspectSelection'/>
                    <button id='ExcelAccel.ApplyCurrency'
                            label='Currency Format'
                            keytip='C'
                            size='large'
                            imageMso='AccountingFormat'
                            screentip='Apply currency number format'
                            supertip='Changes only the NumberFormat property of the current selection.'
                            onAction='OnApplyCurrencyFormat'/>
                    <menu id='ExcelAccel.Formatting' label='Formatting' keytip='F' imageMso='FormatCellsDialog'>
                      <button id='ExcelAccel.FontColor' label='Cycle Font Color' keytip='FC' tag='format.font_color.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.FillColor' label='Cycle Fill Color' keytip='FI' tag='format.fill_color.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.General' label='General' keytip='NG' tag='format.number.general' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Percentage' label='Percentage' keytip='NP' tag='format.number.percentage' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Multiple' label='Multiple' keytip='NM' tag='format.number.multiple' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Date' label='Date' keytip='NT' tag='format.number.date' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Boolean' label='Boolean' keytip='NB' tag='format.number.boolean' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.MoreDecimals' label='Increase Decimals' keytip='NI' tag='format.number.decimals.increase' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.LessDecimals' label='Decrease Decimals' keytip='ND' tag='format.number.decimals.decrease' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.CenterAcross' label='Center Across Selection' keytip='CA' tag='format.center_across.apply' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.HorizontalAlignment' label='Cycle Horizontal Alignment' keytip='AH' tag='format.alignment.horizontal.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.VerticalAlignment' label='Cycle Vertical Alignment' keytip='AV' tag='format.alignment.vertical.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.IndentIncrease' label='Increase Indent' keytip='II' tag='format.indent.increase' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.IndentDecrease' label='Decrease Indent' keytip='ID' tag='format.indent.decrease' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Underline' label='Cycle Underline' keytip='FU' tag='format.underline.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.FontSize' label='Cycle Font Size' keytip='FS' tag='format.font_size.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.RowHeight' label='Cycle Row Height' keytip='RH' tag='format.row_height.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.ColumnWidth' label='Cycle Column Width' keytip='CW' tag='format.column_width.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.SumBar' label='Apply Sum Bar' keytip='BS' tag='format.border.sum_bar.apply' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.RemoveBorders' label='Remove Borders' keytip='BR' tag='format.border.remove' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.AutoFitRows' label='AutoFit Rows' keytip='AR' tag='format.rows.autofit' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.AutoFitColumns' label='AutoFit Columns' keytip='AC' tag='format.columns.autofit' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Gridlines' label='Toggle Gridlines' keytip='VG' tag='view.gridlines.toggle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Zoom' label='Set Zoom 100%' keytip='VZ' tag='view.zoom.set' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Freeze' label='Freeze Panes (Preview Required)' keytip='VF' tag='view.panes.freeze' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Unfreeze' label='Unfreeze Panes' keytip='VU' tag='view.panes.unfreeze' onAction='OnFormattingCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Profiles' label='Profiles' keytip='P' imageMso='FileSaveAs'>
                      <button id='ExcelAccel.ProfileExport' label='Export Profile...' keytip='E' onAction='OnExportProfile'/>
                      <button id='ExcelAccel.ProfileImport' label='Import Profile...' keytip='I' onAction='OnImportProfile'/>
                      <button id='ExcelAccel.BindingExport' label='Export Shortcut Cheat Sheet...' keytip='B' onAction='OnExportBindingCheatSheet'/>
                    </menu>
                    <menu id='ExcelAccel.Styles' label='Styles' keytip='Y' imageMso='CellStylesGallery'>
                      <button id='ExcelAccel.StyleLibrary' label='Style Library...' keytip='L' onAction='OnOpenStyleLibrary'/>
                      <button id='ExcelAccel.StyleMajorHeader' label='Major Header' keytip='MH' tag='major_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleMinorHeader' label='Minor Header' keytip='MI' tag='minor_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleDateHeader' label='Date Header' keytip='DH' tag='date_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleAssumption' label='Assumption' keytip='A' tag='assumption' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleFormula' label='Formula' keytip='F' tag='formula' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleLinkedFormula' label='Linked Formula' keytip='K' tag='linked_formula' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleOutput' label='Output' keytip='O' tag='output' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleWarning' label='Warning' keytip='W' tag='warning' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleTotal' label='Total' keytip='T' tag='total' onAction='OnApplyBuiltInStyle'/>
                    </menu>
                    <menu id='ExcelAccel.Formulas' label='Formulas' keytip='M' imageMso='FunctionWizard'>
                      <button id='ExcelAccel.FormulaSourceCapture' label='Capture Formula Source' keytip='SC' tag='formula.source.capture' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaTranspose' label='Transpose Captured Source Here...' keytip='TP' tag='formula.transpose' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.PasteFormulasOnly' label='Paste Formulas Only' keytip='PF' tag='paste.formulas_only' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.PasteValuesOnly' label='Paste Values Only...' keytip='PV' tag='paste.values_only' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaCopyDown' label='Smart Copy Down' keytip='CD' tag='formula.copy.down' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaCopyRight' label='Smart Copy Right' keytip='CR' tag='formula.copy.right' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaSpacingRows' label='Space Formulas by Rows...' keytip='SR' tag='formula.spacing.rows' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaSpacingColumns' label='Space Formulas by Columns...' keytip='SL' tag='formula.spacing.columns' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillFormulaAbove' label='Fill Formula from Above' keytip='FA' tag='fill.formula_from_above' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillValueAbove' label='Fill Value from Above...' keytip='VA' tag='fill.value_from_above' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillNumericSequence' label='Fill Numeric Sequence...' keytip='NS' tag='fill.numeric_sequence' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillDateSequence' label='Fill Date Sequence...' keytip='DS' tag='fill.date_sequence' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaIfError' label='Toggle IFERROR' keytip='IE' tag='formula.iferror.toggle' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaReverseSign' label='Reverse Sign' keytip='RS' tag='formula.sign.reverse' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaToThousands' label='To Thousands (÷1,000)' keytip='UT' tag='formula.units.to_thousands' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaFromThousands' label='From Thousands (×1,000)' keytip='UF' tag='formula.units.from_thousands' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaToMillions' label='To Millions (÷1,000,000)' keytip='UM' tag='formula.units.to_millions' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaFromMillions' label='From Millions (×1,000,000)' keytip='UN' tag='formula.units.from_millions' onAction='OnFormulaCommand'/>
                    </menu>
                    <menu id='ExcelAccel.DataCleaning' label='Data Cleaning' keytip='D' imageMso='DataValidation'>
                      <button id='ExcelAccel.TrimOuter' label='Trim Outer Whitespace' keytip='TO' tag='clean.text.trim_outer' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.CollapseWhitespace' label='Collapse Whitespace' keytip='CW' tag='clean.text.collapse_whitespace' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.RemoveNonprinting' label='Remove Nonprinting' keytip='RN' tag='clean.text.remove_nonprinting' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.TextToNumber' label='Invariant Text to Number...' keytip='TN' tag='clean.convert.text_to_number' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.NumberToText' label='Number to Invariant Text...' keytip='NT' tag='clean.convert.number_to_text' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.NormalizeDateText' label='Normalize Date Text...' keytip='DN' tag='clean.convert.date_normalize' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.BlankToZero' label='Blanks to Zero...' keytip='BZ' tag='clean.display.blank_to_zero' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.ZeroToBlank' label='Zeros to Blank...' keytip='ZB' tag='clean.display.zero_to_blank' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.BlankToNA' label='Blanks to N/A...' keytip='BN' tag='clean.display.blank_to_na_text' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.BlankToNM' label='Blanks to NM...' keytip='BM' tag='clean.display.blank_to_nm_text' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.BlankToDash' label='Blanks to Dash...' keytip='BD' tag='clean.display.blank_to_dash_text' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.NAToBlank' label='N/A to Blanks...' keytip='NB' tag='clean.display.na_text_to_blank' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.NMToBlank' label='NM to Blanks...' keytip='MB' tag='clean.display.nm_text_to_blank' onAction='OnDataCleaningCommand'/>
                      <button id='ExcelAccel.DashToBlank' label='Dashes to Blanks...' keytip='DB' tag='clean.display.dash_text_to_blank' onAction='OnDataCleaningCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Selection' label='Select' keytip='L' imageMso='SelectCurrentRegion'>
                      <button id='ExcelAccel.SelectFormulas' label='Formulas' keytip='FO' tag='selection.select.formulas' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectConstants' label='Constants' keytip='CO' tag='selection.select.constants' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectBlanks' label='True Blanks' keytip='BL' tag='selection.select.blanks' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectNumericHardcodes' label='Numeric Hardcodes' keytip='NH' tag='selection.select.numeric_hardcodes' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectExternalFormulas' label='External Formulas' keytip='EX' tag='selection.select.external_formulas' onAction='OnSelectionCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Navigation' label='Navigate' keytip='V' imageMso='GoTo'>
                      <button id='ExcelAccel.PreviousSheet' label='Previous Sheet' keytip='P' tag='navigate.sheet.previous' onAction='OnNavigate'/>
                      <button id='ExcelAccel.NextSheet' label='Next Sheet' keytip='N' tag='navigate.sheet.next' onAction='OnNavigate'/>
                      <button id='ExcelAccel.GoA1' label='Go to A1' keytip='A' tag='navigate.cell.a1' onAction='OnNavigate'/>
                      <button id='ExcelAccel.UsedFirst' label='First Used Cell' keytip='F' tag='navigate.used.first' onAction='OnNavigate'/>
                      <button id='ExcelAccel.UsedLast' label='Last Used Cell' keytip='L' tag='navigate.used.last' onAction='OnNavigate'/>
                      <button id='ExcelAccel.EdgeUp' label='Region Edge Up' keytip='U' tag='navigate.region.edge.up' onAction='OnNavigate'/>
                      <button id='ExcelAccel.EdgeDown' label='Region Edge Down' keytip='D' tag='navigate.region.edge.down' onAction='OnNavigate'/>
                      <button id='ExcelAccel.EdgeLeft' label='Region Edge Left' keytip='E' tag='navigate.region.edge.left' onAction='OnNavigate'/>
                      <button id='ExcelAccel.EdgeRight' label='Region Edge Right' keytip='R' tag='navigate.region.edge.right' onAction='OnNavigate'/>
                      <button id='ExcelAccel.NavBack' label='Back' keytip='B' tag='navigate.history.back' onAction='OnNavigate'/>
                      <button id='ExcelAccel.NavForward' label='Forward' keytip='O' tag='navigate.history.forward' onAction='OnNavigate'/>
                      <button id='ExcelAccel.BookmarkAdd' label='Add Bookmark' keytip='M' tag='navigate.bookmark.add_session' onAction='OnNavigate'/>
                      <button id='ExcelAccel.BookmarkNext' label='Next Bookmark' keytip='J' tag='navigate.bookmark.next_session' onAction='OnNavigate'/>
                      <button id='ExcelAccel.BookmarkPrevious' label='Previous Bookmark' keytip='K' tag='navigate.bookmark.previous_session' onAction='OnNavigate'/>
                      <button id='ExcelAccel.BookmarkClear' label='Clear Bookmarks' keytip='C' tag='navigate.bookmark.clear_session' onAction='OnNavigate'/>
                    </menu>
                    <button id='ExcelAccel.UndoLast' label='Undo ExcelAccel' keytip='U' imageMso='Undo' onAction='OnUndoLast'/>
                    <button id='ExcelAccel.ExportDiagnostics' label='Export Diagnostics' keytip='S' imageMso='FileSaveAs' onAction='OnExportDiagnostics'/>
                  </group>
                </tab>
              </tabs>
            </ribbon>
          </customUI>";

    public void OnInspectSelection(IRibbonControl control)
    {
        CallbackBoundary.Run(InspectSelectionCommand.Id, () =>
            CommandDispatcher.InspectSelection());
    }

    public void OnOpenCommandSearch(IRibbonControl control)
    {
        CallbackBoundary.Run("command.search.open", CommandSearchRuntime.Open, showResult: false);
    }

    public void OnApplyCurrencyFormat(IRibbonControl control)
    {
        CallbackBoundary.Run(ApplyCurrencyFormatCommand.Id, () =>
            CommandDispatcher.ApplyCurrencyFormat());
    }

    public void OnFormattingCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.ApplyProfileFormatting(commandId));
    }

    public void OnNavigate(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.Navigate(commandId));
    }

    public void OnFormulaCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.ApplyFormulaCommand(commandId));
    }

    public void OnDataCleaningCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.ApplyDataCleaningCommand(commandId));
    }

    public void OnSelectionCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.ApplySelectionCommand(commandId));
    }

    public void OnOpenStyleLibrary(IRibbonControl control)
    {
        CallbackBoundary.Run("style.apply", StyleLibraryRuntime.Open, showResult: false);
    }

    public void OnApplyBuiltInStyle(IRibbonControl control)
    {
        CallbackBoundary.Run("style.apply_builtin", () => CommandDispatcher.ApplyStyle(control.Tag, requireBuiltIn: true));
    }

    public void OnExportProfile(IRibbonControl control)
    {
        CallbackBoundary.Run("profile.export", CommandDispatcher.ExportProfile);
    }

    public void OnImportProfile(IRibbonControl control)
    {
        CallbackBoundary.Run("profile.import.apply", () => CommandDispatcher.ImportProfile(apply: true));
    }

    public void OnExportBindingCheatSheet(IRibbonControl control)
    {
        CallbackBoundary.Run("bindings.cheat_sheet.export", CommandDispatcher.ExportBindingCheatSheet);
    }

    public void OnUndoLast(IRibbonControl control)
    {
        CallbackBoundary.Run(ExcelAccel.Application.Undo.UndoLastCommand.Id, () => CommandDispatcher.UndoLastProperty());
    }

    public void OnExportDiagnostics(IRibbonControl control)
    {
        CallbackBoundary.Run("support.diagnostics.export", CommandDispatcher.ExportDiagnostics);
    }
}
