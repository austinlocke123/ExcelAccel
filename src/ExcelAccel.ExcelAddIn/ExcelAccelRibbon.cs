using ExcelDna.Integration.CustomUI;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

public sealed class ExcelAccelRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) =>
        @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
            <ribbon>
              <tabs>
                <tab id='ExcelAccel.Tab' label='ExcelAccel' keytip='XA'>

                  <group id='ExcelAccel.Group.Find' label='Find'>
                    <button id='ExcelAccel.CommandSearch' label='Search Commands' keytip='Q' size='large' imageMso='FindDialog'
                            screentip='Search ExcelAccel commands'
                            supertip='Searches local command metadata and shows current availability without scanning or changing the workbook.'
                            onAction='OnOpenCommandSearch'/>
                    <button id='ExcelAccel.InspectSelection' label='Inspect Selection' keytip='I' imageMso='ReviewShowAllMarkup'
                            screentip='Read selection metadata'
                            supertip='Reads selection identity, size, formula state, and number format. It does not change the workbook.'
                            onAction='OnInspectSelection'/>
                    <button id='ExcelAccel.UndoLast' label='Undo ExcelAccel' keytip='U' imageMso='Undo'
                            screentip='Reverse the last ExcelAccel property change'
                            onAction='OnUndoLast'/>
                  </group>

                  <group id='ExcelAccel.Group.NumberFormat' label='Number Format'>
                    <button id='ExcelAccel.ApplyCurrency' label='Currency' keytip='C' imageMso='AccountingFormat' onAction='OnApplyCurrencyFormat'/>
                    <button id='ExcelAccel.General' label='General' keytip='G' tag='format.number.general' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Percentage' label='Percentage' keytip='P' tag='format.number.percentage' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Date' label='Date' keytip='D' tag='format.number.date' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Multiple' label='Multiple' keytip='NM' tag='format.number.multiple' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Boolean' label='Boolean' keytip='NB' tag='format.number.boolean' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.MoreDecimals' label='More Decimals' keytip='NI' tag='format.number.decimals.increase' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.LessDecimals' label='Fewer Decimals' keytip='ND' tag='format.number.decimals.decrease' onAction='OnFormattingCommand'/>
                  </group>

                  <group id='ExcelAccel.Group.CellFormat' label='Cell Format'>
                    <button id='ExcelAccel.FontColor' label='Font Color' keytip='F' tag='format.font_color.cycle' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.FillColor' label='Fill Color' keytip='L' tag='format.fill_color.cycle' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.CenterAcross' label='Center Across' keytip='EC' tag='format.center_across.apply' onAction='OnFormattingCommand'/>
                    <menu id='ExcelAccel.Align' label='Align' keytip='EA' imageMso='AlignCenter'>
                      <button id='ExcelAccel.HorizontalAlignment' label='Cycle Horizontal' keytip='H' tag='format.alignment.horizontal.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.VerticalAlignment' label='Cycle Vertical' keytip='V' tag='format.alignment.vertical.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.IndentIncrease' label='Increase Indent' keytip='I' tag='format.indent.increase' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.IndentDecrease' label='Decrease Indent' keytip='D' tag='format.indent.decrease' onAction='OnFormattingCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Font' label='Font' keytip='EF' imageMso='FontDialog'>
                      <button id='ExcelAccel.FontSize' label='Cycle Font Size' keytip='S' tag='format.font_size.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.Underline' label='Cycle Underline' keytip='U' tag='format.underline.cycle' onAction='OnFormattingCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Borders' label='Borders' keytip='EB' imageMso='BordersAll'>
                      <button id='ExcelAccel.SumBar' label='Apply Sum Bar' keytip='S' tag='format.border.sum_bar.apply' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.RemoveBorders' label='Remove Borders' keytip='R' tag='format.border.remove' onAction='OnFormattingCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Size' label='Size' keytip='ES' imageMso='CellsMenu'>
                      <button id='ExcelAccel.RowHeight' label='Cycle Row Height' keytip='R' tag='format.row_height.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.ColumnWidth' label='Cycle Column Width' keytip='C' tag='format.column_width.cycle' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.AutoFitRows' label='AutoFit Rows' keytip='A' tag='format.rows.autofit' onAction='OnFormattingCommand'/>
                      <button id='ExcelAccel.AutoFitColumns' label='AutoFit Columns' keytip='W' tag='format.columns.autofit' onAction='OnFormattingCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Styles' label='Styles' keytip='EY' imageMso='CellStylesGallery'>
                      <button id='ExcelAccel.StyleMajorHeader' label='Major Header' keytip='MH' tag='major_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleMinorHeader' label='Minor Header' keytip='MI' tag='minor_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleDateHeader' label='Date Header' keytip='DH' tag='date_header' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleAssumption' label='Assumption' keytip='A' tag='assumption' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleFormula' label='Formula' keytip='F' tag='formula' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleLinkedFormula' label='Linked Formula' keytip='K' tag='linked_formula' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleOutput' label='Output' keytip='O' tag='output' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleWarning' label='Warning' keytip='W' tag='warning' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleTotal' label='Total' keytip='T' tag='total' onAction='OnApplyBuiltInStyle'/>
                      <button id='ExcelAccel.StyleLibrary' label='Style Library...' keytip='L' onAction='OnOpenStyleLibrary'/>
                    </menu>
                  </group>

                  <group id='ExcelAccel.Group.QuickFormulas' label='Quick Formulas'>
                    <button id='ExcelAccel.FormulaIfError' label='IFERROR' keytip='RE' tag='formula.iferror.toggle' onAction='OnFormulaCommand'/>
                    <button id='ExcelAccel.FormulaReverseSign' label='Reverse Sign' keytip='RV' tag='formula.sign.reverse' onAction='OnFormulaCommand'/>
                    <menu id='ExcelAccel.Units' label='Units' keytip='RU' imageMso='NumberFormatGallery'>
                      <button id='ExcelAccel.FormulaToThousands' label='To Thousands (divide by 1,000)' keytip='T' tag='formula.units.to_thousands' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaFromThousands' label='From Thousands (times 1,000)' keytip='F' tag='formula.units.from_thousands' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaToMillions' label='To Millions (divide by 1,000,000)' keytip='M' tag='formula.units.to_millions' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaFromMillions' label='From Millions (times 1,000,000)' keytip='N' tag='formula.units.from_millions' onAction='OnFormulaCommand'/>
                    </menu>
                    <menu id='ExcelAccel.SmartCopy' label='Smart Copy' keytip='RC' imageMso='Copy'>
                      <button id='ExcelAccel.FormulaCopyDown' label='Smart Copy Down' keytip='D' tag='formula.copy.down' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaCopyRight' label='Smart Copy Right' keytip='R' tag='formula.copy.right' onAction='OnFormulaCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Spacing' label='Spacing' keytip='RS' imageMso='DistributeVertically'>
                      <button id='ExcelAccel.FormulaSpacingRows' label='Space Formulas by Rows...' keytip='R' tag='formula.spacing.rows' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaSpacingColumns' label='Space Formulas by Columns...' keytip='C' tag='formula.spacing.columns' onAction='OnFormulaCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Fill' label='Fill' keytip='RF' imageMso='FillDown'>
                      <button id='ExcelAccel.FillFormulaAbove' label='Fill Formula from Above' keytip='F' tag='fill.formula_from_above' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillValueAbove' label='Fill Value from Above...' keytip='V' tag='fill.value_from_above' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillNumericSequence' label='Fill Numeric Sequence...' keytip='N' tag='fill.numeric_sequence' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FillDateSequence' label='Fill Date Sequence...' keytip='D' tag='fill.date_sequence' onAction='OnFormulaCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Paste' label='Paste' keytip='RP' imageMso='PasteSpecialDialog'>
                      <button id='ExcelAccel.FormulaSourceCapture' label='Capture Formula Source' keytip='C' tag='formula.source.capture' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.PasteFormulasOnly' label='Paste Formulas Only' keytip='F' tag='paste.formulas_only' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.PasteValuesOnly' label='Paste Values Only...' keytip='V' tag='paste.values_only' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.PasteFormatsOnly' label='Paste Formats Only...' keytip='T' tag='paste.formats_only' onAction='OnFormulaCommand'/>
                      <button id='ExcelAccel.FormulaTranspose' label='Transpose Captured Source Here...' keytip='R' tag='formula.transpose' onAction='OnFormulaCommand'/>
                    </menu>
                  </group>

                  <group id='ExcelAccel.Group.CleanSelect' label='Clean and Select'>
                    <menu id='ExcelAccel.DataCleaning' label='Clean' keytip='KC' imageMso='DataValidation'>
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
                    <menu id='ExcelAccel.Selection' label='Select' keytip='KS' imageMso='SelectCurrentRegion'>
                      <button id='ExcelAccel.SelectFormulas' label='Formulas' keytip='F' tag='selection.select.formulas' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectConstants' label='Constants' keytip='C' tag='selection.select.constants' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectBlanks' label='True Blanks' keytip='B' tag='selection.select.blanks' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectNumericHardcodes' label='Numeric Hardcodes' keytip='N' tag='selection.select.numeric_hardcodes' onAction='OnSelectionCommand'/>
                      <button id='ExcelAccel.SelectExternalFormulas' label='External Formulas' keytip='X' tag='selection.select.external_formulas' onAction='OnSelectionCommand'/>
                    </menu>
                  </group>

                  <group id='ExcelAccel.Group.Audit' label='Audit and Review'>
                    <menu id='ExcelAccel.Precedents' label='Precedents' keytip='AP' imageMso='TracePrecedents'>
                      <button id='ExcelAccel.DirectPrecedents' label='Direct Precedents' keytip='D' tag='audit.precedents.direct' onAction='OnAuditCommand'/>
                      <button id='ExcelAccel.IndirectPrecedents' label='Indirect Precedents' keytip='I' tag='audit.precedents.indirect' onAction='OnAuditCommand'/>
                    </menu>
                    <menu id='ExcelAccel.Dependents' label='Dependents' keytip='AD' imageMso='TraceDependents'>
                      <button id='ExcelAccel.DirectDependents' label='Direct Dependents' keytip='D' tag='audit.dependents.direct' onAction='OnAuditCommand'/>
                      <button id='ExcelAccel.IndirectDependents' label='Indirect Dependents' keytip='I' tag='audit.dependents.indirect' onAction='OnAuditCommand'/>
                      <button id='ExcelAccel.WorkbookDependents' label='Across Workbook' keytip='W' tag='audit.dependents.workbook' onAction='OnAuditCommand'/>
                    </menu>
                    <button id='ExcelAccel.InspectFormula' label='Inspect Formula' keytip='AF' imageMso='FunctionWizard'
                            screentip='Show the structure of the selected formula'
                            supertip='Opens a read-only tree of functions, operators, constants, and references with exact source spans. Nothing is evaluated, scored, or explained.'
                            onAction='OnAuditCommand'
                            tag='audit.formula.inspect'/>
                    <menu id='ExcelAccel.ModelCheck' label='Model Check' keytip='AM' imageMso='ReviewShowMarkupMenu'>
                      <button id='ExcelAccel.ModelCheckSelection' label='Check Selection' keytip='S' tag='model_check.run.selection' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckWorksheet' label='Check Worksheet' keytip='W' tag='model_check.run.worksheet' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckWorkbook' label='Check Workbook' keytip='B' tag='model_check.run.workbook' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckRescan' label='Rescan' keytip='R' tag='model_check.rescan' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckIgnore' label='Ignore Finding...' keytip='I' tag='model_check.finding.ignore_local' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckUnignore' label='Manage Ignores...' keytip='M' tag='model_check.finding.unignore_local' onAction='OnModelCheckCommand'/>
                      <button id='ExcelAccel.ModelCheckExport' label='Export Findings...' keytip='E' tag='model_check.results.export' onAction='OnModelCheckCommand'/>
                    </menu>
                  </group>

                  <group id='ExcelAccel.Group.NavigateView' label='Navigate and View'>
                    <menu id='ExcelAccel.Navigation' label='Navigate' keytip='VN' imageMso='GoTo'>
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
                    <button id='ExcelAccel.Freeze' label='Freeze Panes' keytip='VF' tag='view.panes.freeze' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Unfreeze' label='Unfreeze Panes' keytip='VU' tag='view.panes.unfreeze' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Gridlines' label='Gridlines' keytip='VG' tag='view.gridlines.toggle' onAction='OnFormattingCommand'/>
                    <button id='ExcelAccel.Zoom' label='Zoom 100%' keytip='VZ' tag='view.zoom.set' onAction='OnFormattingCommand'/>
                  </group>

                  <group id='ExcelAccel.Group.Settings' label='Settings'>
                    <menu id='ExcelAccel.Settings' label='Settings' keytip='S' imageMso='FileSaveAs'>
                      <button id='ExcelAccel.ProfileExport' label='Export Profile...' keytip='E' onAction='OnExportProfile'/>
                      <button id='ExcelAccel.ProfileImport' label='Import Profile...' keytip='I' onAction='OnImportProfile'/>
                      <button id='ExcelAccel.BindingExport' label='Export Shortcut Cheat Sheet...' keytip='B' onAction='OnExportBindingCheatSheet'/>
                      <button id='ExcelAccel.ExportDiagnostics' label='Export Diagnostics...' keytip='D' onAction='OnExportDiagnostics'/>
                    </menu>
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
        CallbackBoundary.Run("command.search.open", CommandSearchRuntime.Open);
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

    public void OnModelCheckCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.InvokeRegistered(commandId, null, InvocationSource.Ribbon),
            // The scan commands always present their own result view, including
            // refusals. Rescan can refuse before any view exists, so it is not
            // suppressed.
            showResult: commandId != ModelCheckCommandCatalog.RunSelectionId &&
                commandId != ModelCheckCommandCatalog.RunWorksheetId &&
                commandId != ModelCheckCommandCatalog.RunWorkbookId);
    }

    public void OnAuditCommand(IRibbonControl control)
    {
        var commandId = control.Tag;
        CallbackBoundary.Run(commandId, () => CommandDispatcher.InvokeRegistered(commandId, null, InvocationSource.Ribbon), showResult: false);
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
        CallbackBoundary.Run("style.apply", StyleLibraryRuntime.Open);
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
