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

    public void OnApplyCurrencyFormat(IRibbonControl control)
    {
        CallbackBoundary.Run(ApplyCurrencyFormatCommand.Id, () =>
            CommandDispatcher.ApplyCurrencyFormat());
    }
}
