[CmdletBinding()]
param(
    [string]$AddInPath = "",
    [ValidateRange(10, 120)]
    [int]$TimeoutSeconds = 40,
    [switch]$Worker
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Worker {
    param([string]$ResolvedAddInPath)

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelAccelNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $cell = $null
    $secondCell = $null
    $navigationCell = $null
    $auditSource = $null
    $auditFormula = $null
    $auditName = $null
    $auditViewWorkbook = $null
    $auditViewSheet = $null
    $auditViewSource = $null
    $auditViewFormula = $null
    $depSource = $null
    $depDirect = $null
    $depRange = $null
    $depIndirect = $null
    $chainA = $null
    $chainB = $null
    $chainC = $null
    $font = $null
    $interior = $null
    $multiArea = $null
    $mergedRange = $null
    $formulaRange = $null
    $formulaSource = $null
    $formulaDestination1 = $null
    $formulaDestination2 = $null
    $transposeSource = $null
    $transposeDestination = $null
    $transposeSource1 = $null
    $transposeSource2 = $null
    $transposeSource3 = $null
    $transposeSource4 = $null
    $transposeDestination1 = $null
    $transposeDestination2 = $null
    $transposeDestination3 = $null
    $transposeDestination4 = $null
    $dataRange = $null
    $dataText = $null
    $dataFormula = $null
    $dataZero = $null
    $dataTextZero = $null
    $dataBlank = $null
    $selectionRange = $null
    $selectionNumber1 = $null
    $selectionNumber2 = $null
    $selectionNumber3 = $null
    $selectionNumber4 = $null
    $selectionFormula1 = $null
    $selectionFormula2 = $null
    $selectionText = $null
    $typedRange = $null
    $typedNumber = $null
    $typedDate = $null
    $typedExistingNumber = $null
    $typedInvalid = $null
    $typedFormula = $null
    $pasteValueSource = $null
    $pasteValueDestination = $null
    $pasteValueSourceFormula = $null
    $pasteValueSourceText = $null
    $pasteValueDestination1 = $null
    $pasteValueDestination2 = $null
    $pasteValueDestination3 = $null
    $pasteValueDestination4 = $null
    $aboveSource = $null
    $aboveDestination = $null
    $aboveSource1 = $null
    $aboveSource2 = $null
    $aboveDestination1 = $null
    $aboveDestination2 = $null
    $aboveDestination3 = $null
    $aboveDestination4 = $null
    $aboveInput1 = $null
    $aboveInput2 = $null
    $aboveAbsolute = $null
    $sequenceRange = $null
    $sequence1 = $null
    $sequence2 = $null
    $sequence3 = $null
    $sequence4 = $null
    $formatSource = $null
    $formatDestination = $null
    $formatSourceFont = $null
    $formatSourceInterior = $null
    $formatDestinationFont = $null
    $formatDestinationInterior = $null
    $formatDestination1 = $null
    $formatDestination2 = $null
    $formatDestination3 = $null
    $formatDestination4 = $null
    $quitReturned = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        [uint32]$excelProcessId = 0
        [void][ExcelAccelNativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        [Console]::Out.Flush()

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $registered = $excel.RegisterXLL($ResolvedAddInPath)
        [Console]::WriteLine("registered=$registered")
        [Console]::Out.Flush()
        if (-not $registered) {
            throw 'Excel returned false from RegisterXLL.'
        }

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $cell = $worksheet.Range('A1')
        $cell.Formula = '=EXCELACCEL.VERSION()'
        $excel.Calculate()
        $version = [string]$cell.Value2
        [Console]::WriteLine("version=$version")
        [Console]::Out.Flush()
        if ([string]::IsNullOrWhiteSpace($version) -or $version.StartsWith('#')) {
            throw "The health function returned '$version'."
        }

        [void]$excel.Run('ExcelAccel.Smoke.OpenAndCloseCommandSearch')
        [Console]::WriteLine('command_search_ui=opened_and_closed')
        [Console]::Out.Flush()
        [void]$excel.Run('ExcelAccel.Smoke.OpenAndCloseStyleLibrary')
        [Console]::WriteLine('style_library_ui=opened_and_closed')
        [Console]::Out.Flush()

        $cell.Value2 = 1234.5
        $cell.NumberFormat = 'General'
        $valueBefore = $cell.Value2
        $formulaBefore = $cell.Formula
        [void]$cell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ApplyCurrencyFormat')
        $formatAfter = [string]$cell.NumberFormat
        $valueAfter = $cell.Value2
        $formulaAfter = $cell.Formula
        $formatPreservedContent = ($valueBefore -eq $valueAfter) -and ($formulaBefore -eq $formulaAfter)
        [Console]::WriteLine("currency_format=$formatAfter")
        [Console]::WriteLine("content_preserved=$formatPreservedContent")
        [Console]::Out.Flush()
        if ($formatAfter -ne '$#,##0.00;($#,##0.00);-' -or -not $formatPreservedContent) {
            throw 'The formatting command did not produce the exact property-scoped result.'
        }

        $font = $cell.Font
        $font.Color = 0x563412
        $valueBefore = $cell.Value2
        [void]$cell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ApplyFontColorCycle')
        $fontColorAfter = [int]$font.Color
        $fontColorPreservedContent = ($valueBefore -eq $cell.Value2)
        [Console]::WriteLine("font_color_after=$fontColorAfter")
        [Console]::WriteLine("font_color_content_preserved=$fontColorPreservedContent")
        [Console]::Out.Flush()
        if ($fontColorAfter -ne 0 -or -not $fontColorPreservedContent) {
            throw 'The profile font-color cycle did not change only the declared property.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $fontColorAfterUndo = [int]$font.Color
        [Console]::WriteLine("font_color_after_undo=$fontColorAfterUndo")
        [Console]::Out.Flush()
        if ($fontColorAfterUndo -ne 0x563412) {
            throw 'Optimistic session undo did not restore the exact prior font color.'
        }

        $interior = $cell.Interior
        $font.Bold = $false
        $font.Size = 11
        $font.Color = 0x332211
        $interior.Color = 0x112233
        $styleValueBefore = $cell.Value2
        [void]$cell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ApplyMajorHeaderStyle')
        $styleApplied = [bool]$font.Bold -and ([double]$font.Size -eq 14) -and ([int]$font.Color -eq 0xFFFFFF) -and ([int]$interior.Color -eq 0x784E1F)
        $styleContentPreserved = $styleValueBefore -eq $cell.Value2
        [Console]::WriteLine("style_applied_exactly=$styleApplied")
        [Console]::WriteLine("style_content_preserved=$styleContentPreserved")
        [Console]::Out.Flush()
        if (-not $styleApplied -or -not $styleContentPreserved) {
            throw 'The built-in style did not change exactly its declared formatting properties.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $styleUndoExact = (-not [bool]$font.Bold) -and ([double]$font.Size -eq 11) -and ([int]$font.Color -eq 0x332211) -and ([int]$interior.Color -eq 0x112233)
        [Console]::WriteLine("style_batch_undo_exact=$styleUndoExact")
        [Console]::Out.Flush()
        if (-not $styleUndoExact) {
            throw 'One session undo did not restore every built-in style property.'
        }

        $phase1bWatch = [Diagnostics.Stopwatch]::StartNew()
        $formulaRange = $worksheet.Range('A10:A12')
        $formulaSource = $worksheet.Range('A10')
        $formulaDestination1 = $worksheet.Range('A11')
        $formulaDestination2 = $worksheet.Range('A12')
        $formulaRange.ClearContents()
        $formulaSource.Formula = '=B10+$C$1'
        [void]$formulaRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.FormulaCopyDown')
        [Console]::WriteLine("formula_source=$([string]$formulaSource.Formula)")
        [Console]::WriteLine("formula_destination_1=$([string]$formulaDestination1.Formula)")
        [Console]::WriteLine("formula_destination_2=$([string]$formulaDestination2.Formula)")
        $formulaCopyExact =
            ([string]$formulaSource.Formula -eq '=B10+$C$1') -and
            ([string]$formulaDestination1.Formula -eq '=B11+$C$1') -and
            ([string]$formulaDestination2.Formula -eq '=B12+$C$1')
        [Console]::WriteLine("formula_copy_down_exact=$formulaCopyExact")
        [Console]::Out.Flush()
        if (-not $formulaCopyExact) {
            throw 'Transactional formula copy-down did not produce the exact translated formulas.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $formulaUndoExact =
            ([string]$formulaSource.Formula -eq '=B10+$C$1') -and
            ($null -eq $formulaDestination1.Value2) -and
            ($null -eq $formulaDestination2.Value2)
        [Console]::WriteLine("formula_block_undo_exact=$formulaUndoExact")
        [Console]::Out.Flush()
        if (-not $formulaUndoExact) {
            throw 'Optimistic formula block undo did not restore the exact prior matrix.'
        }

        $transposeSource = $worksheet.Range('B20:C21')
        $transposeDestination = $worksheet.Range('E20:F21')
        $transposeSource1 = $worksheet.Range('B20')
        $transposeSource2 = $worksheet.Range('C20')
        $transposeSource3 = $worksheet.Range('B21')
        $transposeSource4 = $worksheet.Range('C21')
        $transposeDestination1 = $worksheet.Range('E20')
        $transposeDestination2 = $worksheet.Range('F20')
        $transposeDestination3 = $worksheet.Range('E21')
        $transposeDestination4 = $worksheet.Range('F21')
        $transposeSource.ClearContents()
        $transposeDestination.ClearContents()
        $transposeSource1.Formula = '=C20'
        $transposeSource2.Value2 = 7
        $transposeSource3.Value2 = 'x'
        $transposeSource4.Formula = '=$D21'
        [void]$transposeDestination.Select()
        [void]$excel.Run('ExcelAccel.Smoke.FormulaTranspose')
        $transposeExact =
            ([string]$transposeDestination1.Formula -eq '=E21') -and
            ([string]$transposeDestination2.Value2 -eq 'x') -and
            ([double]$transposeDestination3.Value2 -eq 7) -and
            ([string]$transposeDestination4.Formula -eq '=F$4')
        $transposeSelectionPreserved = ([string]$excel.Selection.Address($false, $false) -eq 'E20:F21')
        [Console]::WriteLine("formula_transpose_exact=$transposeExact")
        [Console]::WriteLine("formula_transpose_selection_preserved=$transposeSelectionPreserved")
        [Console]::Out.Flush()
        if (-not $transposeExact -or -not $transposeSelectionPreserved) {
            throw 'Off-selection transpose did not produce the exact result while preserving selection.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $transposeUndoExact =
            ($null -eq $transposeDestination1.Value2) -and
            ($null -eq $transposeDestination2.Value2) -and
            ($null -eq $transposeDestination3.Value2) -and
            ($null -eq $transposeDestination4.Value2)
        [Console]::WriteLine("formula_transpose_undo_exact=$transposeUndoExact")
        [Console]::Out.Flush()
        if (-not $transposeUndoExact) {
            throw 'Transpose undo did not restore the complete destination matrix.'
        }

        $dataRange = $worksheet.Range('A30:E30')
        $dataText = $worksheet.Range('A30')
        $dataFormula = $worksheet.Range('B30')
        $dataZero = $worksheet.Range('C30')
        $dataTextZero = $worksheet.Range('D30')
        $dataBlank = $worksheet.Range('E30')
        $dataOriginalText = " `tclean me" + [char]0x00A0
        $dataText.Value2 = $dataOriginalText
        $dataFormula.Formula = '=" keep formula whitespace "'
        $dataZero.Value2 = 0
        $dataTextZero.NumberFormat = '@'
        $dataTextZero.Value2 = '0'
        $dataBlank.ClearContents()
        [void]$dataRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.DataCleaning')
        [Console]::WriteLine("data_text=$([string]$dataText.Value2)")
        [Console]::WriteLine("data_formula=$([string]$dataFormula.Formula)")
        [Console]::WriteLine("data_zero_is_null=$($null -eq $dataZero.Value2)")
        [Console]::WriteLine("data_zero_value=$([string]$dataZero.Value2)")
        if ($null -ne $dataZero.Value2) { [Console]::WriteLine("data_zero_type=$($dataZero.Value2.GetType().FullName)") }
        [Console]::WriteLine("data_text_zero=$([string]$dataTextZero.Value2)")
        [Console]::WriteLine("data_text_zero_type=$($dataTextZero.Value2.GetType().FullName)")
        [Console]::WriteLine("data_blank_is_null=$($null -eq $dataBlank.Value2)")
        $dataCleaningExact =
            ([string]$dataText.Value2 -eq 'clean me') -and
            ([string]$dataFormula.Formula -eq '=" keep formula whitespace "') -and
            ($null -eq $dataZero.Value2) -and
            ([string]$dataTextZero.Value2 -eq '0') -and
            ($null -eq $dataBlank.Value2)
        [Console]::WriteLine("data_cleaning_exact=$dataCleaningExact")
        [Console]::Out.Flush()
        if (-not $dataCleaningExact) {
            throw 'Transactional data cleaning changed a formula/nonmatch or missed an exact target.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $dataZeroUndoExact = ([double]$dataZero.Value2 -eq 0) -and ([string]$dataText.Value2 -eq 'clean me')
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $dataTrimUndoExact = ([string]$dataText.Value2 -eq $dataOriginalText)
        [Console]::WriteLine("data_cleaning_two_receipt_undo_exact=$($dataZeroUndoExact -and $dataTrimUndoExact)")
        [Console]::Out.Flush()
        if (-not $dataZeroUndoExact -or -not $dataTrimUndoExact) {
            throw 'Data-cleaning receipts did not restore exact values in reverse order.'
        }

        $selectionRange = $worksheet.Range('A40:D42')
        $selectionRange.ClearContents()
        $selectionNumber1 = $worksheet.Range('A40')
        $selectionNumber2 = $worksheet.Range('C40')
        $selectionNumber3 = $worksheet.Range('B41')
        $selectionNumber4 = $worksheet.Range('D42')
        $selectionFormula1 = $worksheet.Range('B40')
        $selectionFormula2 = $worksheet.Range('C41')
        $selectionText = $worksheet.Range('D41')
        $selectionNumber1.Value2 = 10
        $selectionNumber2.Value2 = 20
        $selectionNumber3.Value2 = 30
        $selectionNumber4.Value2 = 40
        $selectionFormula1.Formula = '=A40*2'
        $selectionFormula2.Formula = "='[Book One.xlsx]Model'!A1"
        $selectionText.NumberFormat = '@'
        $selectionText.Value2 = '30'
        [void]$selectionRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.SelectNumericHardcodes')
        $selectionAddress = [string]$excel.Selection.Address($false, $false)
        $selectionExact =
            ($selectionAddress -eq 'A40,C40,B41,D42') -and
            ([double]$selectionNumber1.Value2 -eq 10) -and
            ([double]$selectionNumber2.Value2 -eq 20) -and
            ([double]$selectionNumber3.Value2 -eq 30) -and
            ([double]$selectionNumber4.Value2 -eq 40) -and
            ([string]$selectionFormula1.Formula -eq '=A40*2') -and
            ([string]$selectionFormula2.Formula -eq "='[Book One.xlsx]Model'!A1") -and
            ([string]$selectionText.Value2 -eq '30')
        [Console]::WriteLine("numeric_hardcode_selection=$selectionAddress")
        [Console]::WriteLine("selection_content_preserved=$selectionExact")
        [Console]::Out.Flush()
        if (-not $selectionExact) {
            throw 'Numeric-hardcode selection was not exact or changed workbook content.'
        }

        $typedRange = $worksheet.Range('A50:E50')
        $typedNumber = $worksheet.Range('A50')
        $typedDate = $worksheet.Range('B50')
        $typedExistingNumber = $worksheet.Range('C50')
        $typedInvalid = $worksheet.Range('D50')
        $typedFormula = $worksheet.Range('E50')
        $typedNumber.NumberFormat = '@'
        $typedDate.NumberFormat = '@'
        $typedInvalid.NumberFormat = '@'
        $typedNumber.Value2 = '$1,234.50'
        $typedDate.Value2 = '2026/08/19'
        $typedExistingNumber.Value2 = 12.5
        $typedInvalid.Value2 = '12,34'
        $typedFormula.Formula = '=1+1'
        [void]$typedRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.TypedDataConversions')
        $typedExact =
            ([string]$typedNumber.Value2 -eq '1234.5') -and
            ([string]$typedDate.Value2 -eq '2026-08-19') -and
            ([string]$typedExistingNumber.Value2 -eq '12.5') -and
            ([string]$typedInvalid.Value2 -eq '12,34') -and
            ([string]$typedFormula.Formula -eq '=1+1') -and
            ([string]$typedNumber.Value2.GetType().FullName -eq 'System.String') -and
            ([string]$typedExistingNumber.Value2.GetType().FullName -eq 'System.String')
        [Console]::WriteLine("typed_conversions_exact=$typedExact")
        [Console]::Out.Flush()
        if (-not $typedExact) {
            throw 'Typed number/date conversions were not exact or changed a formula/nonmatch.'
        }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $typedUndoExact =
            ([string]$typedNumber.Value2 -eq '$1,234.50') -and
            ([string]$typedDate.Value2 -eq '2026/08/19') -and
            ([double]$typedExistingNumber.Value2 -eq 12.5) -and
            ([string]$typedInvalid.Value2 -eq '12,34') -and
            ([string]$typedFormula.Formula -eq '=1+1')
        [Console]::WriteLine("typed_conversions_undo_exact=$typedUndoExact")
        [Console]::Out.Flush()
        if (-not $typedUndoExact) {
            throw 'Typed conversion receipts did not restore the exact prior matrix.'
        }

        $pasteValueSource = $worksheet.Range('A60:B60')
        $pasteValueDestination = $worksheet.Range('D60:E61')
        $pasteValueSourceFormula = $worksheet.Range('A60')
        $pasteValueSourceText = $worksheet.Range('B60')
        $pasteValueDestination1 = $worksheet.Range('D60')
        $pasteValueDestination2 = $worksheet.Range('E60')
        $pasteValueDestination3 = $worksheet.Range('D61')
        $pasteValueDestination4 = $worksheet.Range('E61')
        $pasteValueSourceFormula.Formula = '=10+5'
        $pasteValueSourceText.Value2 = 'source'
        $pasteValueDestination1.Formula = '=99'
        $pasteValueDestination2.Value2 = 'old'
        $pasteValueDestination3.Value2 = 8
        $pasteValueDestination4.ClearContents()
        $excel.Calculate()
        [void]$pasteValueDestination.Select()
        [void]$excel.Run('ExcelAccel.Smoke.PasteValues')
        $pasteValuesExact =
            ([double]$pasteValueDestination1.Value2 -eq 15) -and
            ([string]$pasteValueDestination2.Value2 -eq 'source') -and
            ([double]$pasteValueDestination3.Value2 -eq 15) -and
            ([string]$pasteValueDestination4.Value2 -eq 'source') -and
            ([string]$pasteValueSourceFormula.Formula -eq '=10+5') -and
            ([string]$pasteValueSourceText.Value2 -eq 'source')
        [Console]::WriteLine("paste_values_underlying_exact=$pasteValuesExact")
        [Console]::Out.Flush()
        if (-not $pasteValuesExact) { throw 'Values-only paste did not use exact underlying values.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $pasteValuesUndoExact =
            ([string]$pasteValueDestination1.Formula -eq '=99') -and
            ([string]$pasteValueDestination2.Value2 -eq 'old') -and
            ([double]$pasteValueDestination3.Value2 -eq 8) -and
            ($null -eq $pasteValueDestination4.Value2)
        [Console]::WriteLine("paste_values_undo_exact=$pasteValuesUndoExact")
        [Console]::Out.Flush()
        if (-not $pasteValuesUndoExact) { throw 'Values-only paste undo did not restore destination formulas and values.' }

        $aboveSource = $worksheet.Range('A70:B70')
        $aboveDestination = $worksheet.Range('A71:B72')
        $aboveSource1 = $worksheet.Range('A70')
        $aboveSource2 = $worksheet.Range('B70')
        $aboveDestination1 = $worksheet.Range('A71')
        $aboveDestination2 = $worksheet.Range('B71')
        $aboveDestination3 = $worksheet.Range('A72')
        $aboveDestination4 = $worksheet.Range('B72')
        $aboveInput1 = $worksheet.Range('C70')
        $aboveInput2 = $worksheet.Range('D70')
        $aboveAbsolute = $worksheet.Range('D1')
        $aboveInput1.Value2 = 5
        $aboveInput2.Value2 = 6
        $aboveAbsolute.Value2 = 10
        $aboveSource1.Formula = '=C70+$D$1'
        $aboveSource2.Formula = '=D70'
        $aboveDestination.ClearContents()
        $excel.Calculate()
        [void]$aboveDestination.Select()
        [void]$excel.Run('ExcelAccel.Smoke.FormulaFromAbove')
        $formulaAboveExact =
            ([string]$aboveDestination1.Formula -eq '=C71+$D$1') -and
            ([string]$aboveDestination2.Formula -eq '=D71') -and
            ([string]$aboveDestination3.Formula -eq '=C72+$D$1') -and
            ([string]$aboveDestination4.Formula -eq '=D72')
        [Console]::WriteLine("formula_from_above_exact=$formulaAboveExact")
        [Console]::Out.Flush()
        if (-not $formulaAboveExact) { throw 'Formula-from-above translation was not exact.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        [void]$aboveDestination.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ValueFromAbove')
        $valueAboveExact =
            ([double]$aboveDestination1.Value2 -eq 15) -and
            ([double]$aboveDestination2.Value2 -eq 6) -and
            ([double]$aboveDestination3.Value2 -eq 15) -and
            ([double]$aboveDestination4.Value2 -eq 6) -and
            ([string]$aboveSource1.Formula -eq '=C70+$D$1') -and
            ([string]$aboveSource2.Formula -eq '=D70')
        [Console]::WriteLine("value_from_above_exact=$valueAboveExact")
        [Console]::Out.Flush()
        if (-not $valueAboveExact) { throw 'Value-from-above did not use exact underlying formula values.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')

        $sequenceRange = $worksheet.Range('G70:H71')
        $sequence1 = $worksheet.Range('G70')
        $sequence2 = $worksheet.Range('H70')
        $sequence3 = $worksheet.Range('G71')
        $sequence4 = $worksheet.Range('H71')
        $sequenceRange.ClearContents()
        [void]$sequenceRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.NumericSequence')
        $numericSequenceExact =
            ([double]$sequence1.Value2 -eq 1) -and ([double]$sequence2.Value2 -eq 3) -and
            ([double]$sequence3.Value2 -eq 5) -and ([double]$sequence4.Value2 -eq 7)
        [Console]::WriteLine("numeric_sequence_exact=$numericSequenceExact")
        [Console]::Out.Flush()
        if (-not $numericSequenceExact) { throw 'Numeric sequence did not follow explicit row-major direction.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        [void]$sequenceRange.Select()
        [void]$excel.Run('ExcelAccel.Smoke.DateSequence')
        $dateStartSerial = [DateTime]::ParseExact('2026-08-19', 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).ToOADate()
        $dateSequenceExact =
            ([double]$sequence1.Value2 -eq $dateStartSerial) -and
            ([double]$sequence2.Value2 -eq ($dateStartSerial + 7)) -and
            ([double]$sequence3.Value2 -eq ($dateStartSerial + 14)) -and
            ([double]$sequence4.Value2 -eq ($dateStartSerial + 21))
        [Console]::WriteLine("date_sequence_exact=$dateSequenceExact")
        [Console]::Out.Flush()
        if (-not $dateSequenceExact) { throw 'Date sequence did not follow the workbook 1900 date system and explicit direction.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')

        $formatSource = $worksheet.Range('A80')
        $formatDestination = $worksheet.Range('D80:E81')
        $formatSourceFont = $formatSource.Font
        $formatSourceInterior = $formatSource.Interior
        $formatDestinationFont = $formatDestination.Font
        $formatDestinationInterior = $formatDestination.Interior
        $formatDestination1 = $worksheet.Range('D80')
        $formatDestination2 = $worksheet.Range('E80')
        $formatDestination3 = $worksheet.Range('D81')
        $formatDestination4 = $worksheet.Range('E81')
        $formatSource.Value2 = 'format source value'
        $formatSource.NumberFormat = '0.00'
        $formatSourceFont.Name = 'Arial'
        $formatSourceFont.Size = 12
        $formatSourceFont.Bold = $true
        $formatSourceFont.Italic = $true
        $formatSourceFont.Underline = 2
        $formatSourceFont.Color = 0x0000FF
        $formatSourceInterior.Color = 0x0000FF
        $formatSource.HorizontalAlignment = -4152
        $formatSource.VerticalAlignment = -4108
        $formatSource.IndentLevel = 1
        $formatDestination1.Formula = '=1+1'
        $formatDestination2.Value2 = 'keep'
        $formatDestination3.Value2 = 8
        $formatDestination4.ClearContents()
        $formatDestination.NumberFormat = 'General'
        $formatDestinationFont.Name = 'Aptos'
        $formatDestinationFont.Size = 11
        $formatDestinationFont.Bold = $false
        $formatDestinationFont.Italic = $false
        $formatDestinationFont.Underline = -4142
        $formatDestinationFont.Color = 0x008000
        $formatDestinationInterior.Color = 0xFF0000
        $formatDestination.HorizontalAlignment = -4131
        $formatDestination.VerticalAlignment = -4160
        $formatDestination.IndentLevel = 0
        [void]$formatDestination.Select()
        [void]$excel.Run('ExcelAccel.Smoke.PasteFormats')
        $formatPasteExact =
            ([string]$formatDestination.NumberFormat -eq '0.00') -and
            ([string]$formatDestinationFont.Name -eq 'Arial') -and
            ([double]$formatDestinationFont.Size -eq 12) -and
            ([bool]$formatDestinationFont.Bold) -and ([bool]$formatDestinationFont.Italic) -and
            ([int]$formatDestinationFont.Underline -eq 2) -and
            ([int]$formatDestination.HorizontalAlignment -eq -4152) -and
            ([int]$formatDestination.VerticalAlignment -eq -4108) -and
            ([int]$formatDestination.IndentLevel -eq 1) -and
            ([int]$formatDestinationFont.Color -eq 0x008000) -and
            ([int]$formatDestinationInterior.Color -eq 0xFF0000) -and
            ([string]$formatDestination1.Formula -eq '=1+1') -and
            ([string]$formatDestination2.Value2 -eq 'keep') -and
            ([double]$formatDestination3.Value2 -eq 8) -and
            ($null -eq $formatDestination4.Value2)
        [Console]::WriteLine("formats_only_exact=$formatPasteExact")
        [Console]::Out.Flush()
        if (-not $formatPasteExact) { throw 'Formats-only paste changed excluded properties/content or missed an approved property.' }
        [void]$excel.Run('ExcelAccel.Smoke.UndoLastProperty')
        $formatPasteUndoExact =
            ([string]$formatDestination.NumberFormat -eq 'General') -and
            ([string]$formatDestinationFont.Name -eq 'Aptos') -and
            ([double]$formatDestinationFont.Size -eq 11) -and
            (-not [bool]$formatDestinationFont.Bold) -and (-not [bool]$formatDestinationFont.Italic) -and
            ([int]$formatDestinationFont.Underline -eq -4142) -and
            ([int]$formatDestination.HorizontalAlignment -eq -4131) -and
            ([int]$formatDestination.VerticalAlignment -eq -4160) -and
            ([int]$formatDestination.IndentLevel -eq 0) -and
            ([int]$formatDestinationFont.Color -eq 0x008000) -and
            ([int]$formatDestinationInterior.Color -eq 0xFF0000)
        [Console]::WriteLine("formats_only_undo_exact=$formatPasteUndoExact")
        [Console]::Out.Flush()
        if (-not $formatPasteUndoExact) { throw 'Formats-only paste undo did not restore the exact approved property matrix.' }
        $phase1bWatch.Stop()
        $phase1bFeatureSuiteMs = [int64][Math]::Ceiling($phase1bWatch.Elapsed.TotalMilliseconds)
        [Console]::WriteLine("phase1b_feature_suite_ms=$phase1bFeatureSuiteMs")
        [Console]::Out.Flush()
        if ($phase1bFeatureSuiteMs -gt 5000) { throw "The bounded Phase 1B smoke suite exceeded its 5,000 ms qualification budget ($phase1bFeatureSuiteMs ms)." }

        $navigationCell = $worksheet.Range('D5')
        [void]$navigationCell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.NavigateA1')
        $navigationAddress = [string]$excel.Selection.Address($false, $false)
        [Console]::WriteLine("navigation_address=$navigationAddress")
        [Console]::Out.Flush()
        if ($navigationAddress -ne 'A1') {
            throw 'Read-only A1 navigation did not select the exact target.'
        }

        $auditSource = $worksheet.Range('A90')
        $auditFormula = $worksheet.Range('B90')
        $auditSource.Value2 = 42
        $auditFormula.Formula = '=A90'
        [void]$auditFormula.Select()
        $auditResult = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedents')
        $auditSelectionPreserved = ([string]$excel.Selection.Address($false, $false) -eq 'B90')
        $auditContentPreserved = ([double]$auditSource.Value2 -eq 42) -and ([string]$auditFormula.Formula -eq '=A90')
        [Console]::WriteLine("direct_precedents=$auditResult")
        [Console]::WriteLine("direct_precedents_selection_preserved=$auditSelectionPreserved")
        [Console]::WriteLine("direct_precedents_content_preserved=$auditContentPreserved")
        [Console]::Out.Flush()
        if ($auditResult -ne 'Complete|1|0|0|Value' -or -not $auditSelectionPreserved -or -not $auditContentPreserved) {
            throw 'Direct-precedent capture did not return the exact read-only result.'
        }
        $auditName = $workbook.Names.Add('AuditRate', "=$([string]$worksheet.Name)!`$A`$90")
        $auditFormula.Formula = '=AuditRate'
        [void]$auditFormula.Select()
        $auditNameResult = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedents')
        [Console]::WriteLine("direct_precedents_name=$auditNameResult")
        [Console]::Out.Flush()
        if ($auditNameResult -ne 'Partial|1|0|0|Value') {
            throw 'Direct-precedent capture did not resolve the supported workbook name exactly.'
        }

        $auditFormula.Formula = '=A90'
        [void]$auditFormula.Select()
        $auditViewResult = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedentsView')
        $auditViewSelectionPreserved = ([string]$excel.Selection.Address($false, $false) -eq 'B90')
        $auditViewContentPreserved = ([double]$auditSource.Value2 -eq 42) -and ([string]$auditFormula.Formula -eq '=A90')
        [Console]::WriteLine("direct_precedents_view=$auditViewResult")
        [Console]::WriteLine("direct_precedents_view_selection_preserved=$auditViewSelectionPreserved")
        [Console]::WriteLine("direct_precedents_view_content_preserved=$auditViewContentPreserved")
        [Console]::Out.Flush()
        if ($auditViewResult -ne 'open|success' -or -not $auditViewSelectionPreserved -or -not $auditViewContentPreserved) {
            throw 'The read-only direct-precedent view did not open without touching the workbook.'
        }

        $auditViewRetained = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedentsViewRevalidate')
        [Console]::WriteLine("direct_precedents_view_open_workbook=$auditViewRetained")
        [Console]::Out.Flush()
        if ($auditViewRetained -ne 'retained|open') {
            throw 'The direct-precedent view discarded a result whose source workbook is still open.'
        }

        $auditViewWorkbook = $excel.Workbooks.Add()
        $auditViewSheet = $auditViewWorkbook.Worksheets.Item(1)
        $auditViewSource = $auditViewSheet.Range('A1')
        $auditViewFormula = $auditViewSheet.Range('B1')
        $auditViewSource.Value2 = 7
        $auditViewFormula.Formula = '=A1'
        [void]$auditViewFormula.Select()
        $auditViewSecondResult = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedentsView')
        [Console]::WriteLine("direct_precedents_view_second_workbook=$auditViewSecondResult")
        [Console]::Out.Flush()
        if ($auditViewSecondResult -ne 'open|success') {
            throw 'The direct-precedent view did not present the second workbook result.'
        }
        Release-ComObject $chainC
        Release-ComObject $chainB
        Release-ComObject $chainA
        Release-ComObject $depIndirect
        Release-ComObject $depRange
        Release-ComObject $depDirect
        Release-ComObject $depSource
        Release-ComObject $auditViewFormula
        $auditViewFormula = $null
        Release-ComObject $auditViewSource
        $auditViewSource = $null
        Release-ComObject $auditViewSheet
        $auditViewSheet = $null
        $auditViewWorkbook.Close($false)
        Release-ComObject $auditViewWorkbook
        $auditViewWorkbook = $null
        $auditViewDiscarded = [string]$excel.Run('ExcelAccel.Smoke.DirectPrecedentsViewRevalidate')
        [Console]::WriteLine("direct_precedents_view_closed_workbook=$auditViewDiscarded")
        [Console]::Out.Flush()
        if ($auditViewDiscarded -ne 'discarded|closed') {
            throw 'The direct-precedent view survived the close of its source workbook.'
        }

        [void]$workbook.Activate()
        [void]$auditFormula.Select()
        [void]$excel.Run('ExcelAccel.Smoke.DirectPrecedentsView')
        $auditViewExplicitClose = [string]$excel.Run('ExcelAccel.Smoke.CloseDirectPrecedentsView')
        [Console]::WriteLine("direct_precedents_view_explicit_close=$auditViewExplicitClose")
        [Console]::Out.Flush()
        if ($auditViewExplicitClose -ne 'closed') {
            throw 'The direct-precedent view did not release on the explicit close path.'
        }

        $depSource = $worksheet.Range('A200')
        $depDirect = $worksheet.Range('B200')
        $depRange = $worksheet.Range('C200')
        $depIndirect = $worksheet.Range('D200')
        $depSource.Value2 = 42
        $depDirect.Formula = '=A200'
        $depRange.Formula = '=SUM(A200:A210)'
        $depIndirect.Formula = '=B200'
        [void]$depSource.Select()
        $dependentResult = [string]$excel.Run('ExcelAccel.Smoke.DirectDependents')
        $dependentSelectionPreserved = ([string]$excel.Selection.Address($false, $false) -eq 'A200')
        $dependentContentPreserved =
            ([double]$depSource.Value2 -eq 42) -and
            ([string]$depDirect.Formula -eq '=A200') -and
            ([string]$depRange.Formula -eq '=SUM(A200:A210)') -and
            ([string]$depIndirect.Formula -eq '=B200')
        [Console]::WriteLine("direct_dependents=$dependentResult")
        [Console]::WriteLine("direct_dependents_selection_preserved=$dependentSelectionPreserved")
        [Console]::WriteLine("direct_dependents_content_preserved=$dependentContentPreserved")
        [Console]::Out.Flush()
        $dependentParts = $dependentResult.Split('|')
        if ($dependentParts[0] -ne 'Complete') {
            throw "The bounded worksheet dependent scan did not claim completeness: $dependentResult"
        }
        if ($dependentParts[3] -ne '0') {
            throw "The dependent scan reported an unexpected coverage gap: $dependentResult"
        }
        if ($dependentParts[1] -ne 'B200,C200') {
            throw "The dependent scan did not return the exact direct dependents: $dependentResult"
        }
        if ([int]$dependentParts[2] -le 0) {
            throw 'The dependent scan reported no scanned formulas.'
        }
        if ($dependentParts[4] -ne 'Completed') {
            throw "The dependent scan did not finish in the Completed progress phase: $dependentResult"
        }
        if (-not $dependentSelectionPreserved -or -not $dependentContentPreserved) {
            throw 'The dependent scan changed the selection or workbook contents.'
        }

        [void]$depSource.Select()
        $dependentCancelled = [string]$excel.Run('ExcelAccel.Smoke.DirectDependentsCancelled')
        [Console]::WriteLine("direct_dependents_cancelled=$dependentCancelled")
        [Console]::Out.Flush()
        if ($dependentCancelled -ne 'Refused|AUDIT_SCAN_CANCELLED|0') {
            throw 'A cancelled dependent scan did not fail closed.'
        }

        [void]$depSource.Select()
        $dependentViewResult = [string]$excel.Run('ExcelAccel.Smoke.DirectDependentsView')
        $dependentViewSelectionPreserved = ([string]$excel.Selection.Address($false, $false) -eq 'A200')
        [Console]::WriteLine("direct_dependents_view=$dependentViewResult")
        [Console]::WriteLine("direct_dependents_view_selection_preserved=$dependentViewSelectionPreserved")
        [Console]::Out.Flush()
        if ($dependentViewResult -ne 'open|success' -or -not $dependentViewSelectionPreserved) {
            throw 'The registered dependent-view route did not open a read-only result.'
        }
        $dependentViewClose = [string]$excel.Run('ExcelAccel.Smoke.CloseDirectDependentsView')
        [Console]::WriteLine("direct_dependents_view_explicit_close=$dependentViewClose")
        [Console]::Out.Flush()
        if ($dependentViewClose -ne 'closed') {
            throw 'The dependent view did not release on the explicit close path.'
        }

        $chainA = $worksheet.Range('A210')
        $chainB = $worksheet.Range('B210')
        $chainC = $worksheet.Range('C210')
        $chainA.Value2 = 5
        $chainB.Formula = '=A210'
        $chainC.Formula = '=B210'
        [void]$chainC.Select()
        $indirectPrecedents = [string]$excel.Run('ExcelAccel.Smoke.IndirectTrace', 'precedents')
        [Console]::WriteLine("indirect_precedents_view=$indirectPrecedents")
        [Console]::Out.Flush()
        if ($indirectPrecedents -ne 'open|success') {
            throw "The registered indirect-precedent route did not open a read-only result: $indirectPrecedents"
        }

        [void]$chainA.Select()
        $indirectDependents = [string]$excel.Run('ExcelAccel.Smoke.IndirectTrace', 'dependents')
        [Console]::WriteLine("indirect_dependents_view=$indirectDependents")
        [Console]::Out.Flush()
        if ($indirectDependents -ne 'open|success') {
            throw "The registered indirect-dependent route did not open a read-only result: $indirectDependents"
        }

        $traceNavigate = [string]$excel.Run('ExcelAccel.Smoke.TraceNavigate', [string]$worksheet.Name, 'C210')
        [Console]::WriteLine("trace_navigate=$traceNavigate")
        [Console]::Out.Flush()
        if ($traceNavigate -ne 'C210|recorded') {
            throw "Trace navigation did not select the target and record return history: $traceNavigate"
        }
        $chainContentPreserved =
            ([double]$chainA.Value2 -eq 5) -and
            ([string]$chainB.Formula -eq '=A210') -and
            ([string]$chainC.Formula -eq '=B210')
        [Console]::WriteLine("indirect_trace_content_preserved=$chainContentPreserved")
        [Console]::Out.Flush()
        if (-not $chainContentPreserved) { throw 'The indirect trace changed workbook contents.' }

        $indirectClose = [string]$excel.Run('ExcelAccel.Smoke.CloseIndirectTrace')
        [Console]::WriteLine("indirect_trace_explicit_close=$indirectClose")
        [Console]::Out.Flush()
        if ($indirectClose -ne 'closed') {
            throw 'The indirect trace view did not release on the explicit close path.'
        }

        $excel.ScreenUpdating = $true
        $excel.EnableEvents = $true
        [void]$excel.Run('ExcelAccel.Smoke.ThrowInsideStateGuard')
        $stateRestored = [bool]$excel.ScreenUpdating -and [bool]$excel.EnableEvents
        [Console]::WriteLine("state_restored_after_fault=$stateRestored")
        [Console]::Out.Flush()
        if (-not $stateRestored) {
            throw 'Excel application state was not restored after the injected mutation failure.'
        }

        $cell.NumberFormat = 'General'
        [void]$cell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ApplyCurrencyFormatAfterInterveningChange')
        $stalePropertyRefused = [string]$cell.NumberFormat -eq '0.00'
        [Console]::WriteLine("stale_property_refused=$stalePropertyRefused")
        [Console]::Out.Flush()
        if (-not $stalePropertyRefused) {
            throw 'The formatting command applied a stale plan after its planned property changed.'
        }

        $cell.NumberFormat = 'General'
        $worksheet.Protect()
        try {
            [void]$cell.Select()
            [void]$excel.Run('ExcelAccel.Smoke.ApplyCurrencyFormat')
            $protectedTargetRefused = ([string]$cell.NumberFormat -eq 'General')
        }
        finally {
            $worksheet.Unprotect()
        }
        [Console]::WriteLine("protected_target_refused=$protectedTargetRefused")
        [Console]::Out.Flush()
        if (-not $protectedTargetRefused) {
            throw 'The formatting command mutated a protected worksheet target.'
        }

        $secondCell = $worksheet.Range('C1')
        $multiArea = $worksheet.Range('A1,C1')
        $cell.NumberFormat = 'General'
        $secondCell.NumberFormat = 'General'
        [void]$multiArea.Select()
        [void]$excel.Run('ExcelAccel.Smoke.ApplyCurrencyFormat')
        $multiAreaRefused =
            ([string]$cell.NumberFormat -eq 'General') -and
            ([string]$secondCell.NumberFormat -eq 'General')
        [Console]::WriteLine("multi_area_refused=$multiAreaRefused")
        [Console]::Out.Flush()
        if (-not $multiAreaRefused) {
            throw 'The formatting command mutated a multi-area selection.'
        }

        $mergedRange = $worksheet.Range('A2:B2')
        $mergedRange.Merge()
        try {
            $mergedRange.NumberFormat = 'General'
            [void]$mergedRange.Select()
            [void]$excel.Run('ExcelAccel.Smoke.ApplyCurrencyFormat')
            $mergedTargetRefused = ([string]$mergedRange.NumberFormat -eq 'General')
        }
        finally {
            $mergedRange.UnMerge()
        }
        [Console]::WriteLine("merged_target_refused=$mergedTargetRefused")
        [Console]::Out.Flush()
        if (-not $mergedTargetRefused) {
            throw 'The formatting command mutated a merged-cell selection.'
        }

        $excelProcess = Get-Process -Id ([int]$excelProcessId)
        $excelProcess.Refresh()
        [Console]::WriteLine("working_set_bytes=$($excelProcess.WorkingSet64)")
        [Console]::WriteLine("private_memory_bytes=$($excelProcess.PrivateMemorySize64)")
        [Console]::WriteLine("handle_count=$($excelProcess.HandleCount)")
        [Console]::Out.Flush()

        $workbook.Close($false)
        $workbook = $null
        [Console]::WriteLine('workbook_closed=true')
        [Console]::Out.Flush()
        $excel.Quit()
        $quitReturned = $true
        [Console]::WriteLine('quit_returned=true')
        [Console]::Out.Flush()
    }
    finally {
        try {
            if ($null -ne $workbook) {
                $workbook.Close($false)
            }
        }
        catch {
        }
        try {
            if ($null -ne $excel -and -not $quitReturned) {
                $excel.Quit()
            }
        }
        catch {
        }
        Release-ComObject $mergedRange
        Release-ComObject $auditFormula
        Release-ComObject $auditSource
        Release-ComObject $auditViewFormula
        Release-ComObject $auditViewSource
        Release-ComObject $auditViewSheet
        Release-ComObject $auditViewWorkbook
        Release-ComObject $auditName
        Release-ComObject $formulaDestination2
        Release-ComObject $formulaDestination1
        Release-ComObject $formulaSource
        Release-ComObject $formulaRange
        Release-ComObject $transposeDestination4
        Release-ComObject $transposeDestination3
        Release-ComObject $transposeDestination2
        Release-ComObject $transposeDestination1
        Release-ComObject $transposeSource4
        Release-ComObject $transposeSource3
        Release-ComObject $transposeSource2
        Release-ComObject $transposeSource1
        Release-ComObject $transposeDestination
        Release-ComObject $transposeSource
        Release-ComObject $dataBlank
        Release-ComObject $dataTextZero
        Release-ComObject $dataZero
        Release-ComObject $dataFormula
        Release-ComObject $dataText
        Release-ComObject $dataRange
        Release-ComObject $selectionText
        Release-ComObject $selectionFormula2
        Release-ComObject $selectionFormula1
        Release-ComObject $selectionNumber4
        Release-ComObject $selectionNumber3
        Release-ComObject $selectionNumber2
        Release-ComObject $selectionNumber1
        Release-ComObject $selectionRange
        Release-ComObject $typedFormula
        Release-ComObject $typedInvalid
        Release-ComObject $typedExistingNumber
        Release-ComObject $typedDate
        Release-ComObject $typedNumber
        Release-ComObject $typedRange
        Release-ComObject $sequence4
        Release-ComObject $sequence3
        Release-ComObject $sequence2
        Release-ComObject $sequence1
        Release-ComObject $sequenceRange
        Release-ComObject $formatDestination4
        Release-ComObject $formatDestination3
        Release-ComObject $formatDestination2
        Release-ComObject $formatDestination1
        Release-ComObject $formatDestinationInterior
        Release-ComObject $formatDestinationFont
        Release-ComObject $formatSourceInterior
        Release-ComObject $formatSourceFont
        Release-ComObject $formatDestination
        Release-ComObject $formatSource
        Release-ComObject $aboveAbsolute
        Release-ComObject $aboveInput2
        Release-ComObject $aboveInput1
        Release-ComObject $aboveDestination4
        Release-ComObject $aboveDestination3
        Release-ComObject $aboveDestination2
        Release-ComObject $aboveDestination1
        Release-ComObject $aboveSource2
        Release-ComObject $aboveSource1
        Release-ComObject $aboveDestination
        Release-ComObject $aboveSource
        Release-ComObject $pasteValueDestination4
        Release-ComObject $pasteValueDestination3
        Release-ComObject $pasteValueDestination2
        Release-ComObject $pasteValueDestination1
        Release-ComObject $pasteValueSourceText
        Release-ComObject $pasteValueSourceFormula
        Release-ComObject $pasteValueDestination
        Release-ComObject $pasteValueSource
        Release-ComObject $multiArea
        Release-ComObject $secondCell
        Release-ComObject $navigationCell
        Release-ComObject $interior
        Release-ComObject $font
        Release-ComObject $cell
        Release-ComObject $worksheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($Worker) {
    Invoke-Worker -ResolvedAddInPath $AddInPath
    exit 0
}

if ([string]::IsNullOrWhiteSpace($AddInPath)) {
    $AddInPath = Join-Path $PSScriptRoot '..\src\ExcelAccel.ExcelAddIn\bin\Debug\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
}

$resolvedPath = (Resolve-Path -LiteralPath $AddInPath).Path
$runId = [Guid]::NewGuid().ToString('N')
$outputPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-smoke-$runId.out"
$errorPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-smoke-$runId.err"
$workerProcess = $null

try {
    $arguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$PSCommandPath`"",
        '-Worker',
        '-AddInPath', "`"$resolvedPath`""
    )

    $workerProcess = Start-Process powershell.exe `
        -ArgumentList $arguments `
        -WindowStyle Hidden `
        -RedirectStandardOutput $outputPath `
        -RedirectStandardError $errorPath `
        -PassThru

    $completed = $workerProcess.WaitForExit($TimeoutSeconds * 1000)
    $output = if (Test-Path -LiteralPath $outputPath) { Get-Content -LiteralPath $outputPath -Raw } else { '' }
    $errors = if (Test-Path -LiteralPath $errorPath) { Get-Content -LiteralPath $errorPath -Raw } else { '' }

    if (-not $completed) {
        Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
    }
    else {
        $workerProcess.Refresh()
    }

    $excelProcessId = [regex]::Match($output, '(?m)^excel_pid=(\d+)$').Groups[1].Value
    if ($excelProcessId) {
        $excelProcess = Get-Process -Id ([int]$excelProcessId) -ErrorAction SilentlyContinue
        if ($excelProcess) {
            $exitedAfterGracePeriod = $excelProcess.WaitForExit(5000)
            if (-not $exitedAfterGracePeriod) {
                Stop-Process -Id $excelProcess.Id -Force
                throw "Excel PID $excelProcessId did not exit cleanly within the smoke-test window. Worker output:`n$output"
            }
        }
    }

    if (-not $completed) {
        throw "Smoke-test worker timed out after $TimeoutSeconds seconds. Worker output:`n$output`nWorker errors:`n$errors"
    }

    if (-not [string]::IsNullOrWhiteSpace($errors)) {
        throw "Smoke-test worker reported an error. Worker output:`n$output`nWorker errors:`n$errors"
    }

    $requiredEvidence = @(
        'registered=True',
        'version=',
        'currency_format=$#,##0.00;($#,##0.00);-',
        'content_preserved=True',
        'font_color_after=0',
        'command_search_ui=opened_and_closed',
        'style_library_ui=opened_and_closed',
        'style_applied_exactly=True',
        'style_content_preserved=True',
        'style_batch_undo_exact=True',
        'font_color_content_preserved=True',
        'font_color_after_undo=5649426',
        'navigation_address=A1',
        'direct_precedents=Complete|1|0|0|Value',
        'direct_precedents_selection_preserved=True',
        'direct_precedents_content_preserved=True',
        'direct_precedents_name=Partial|1|0|0|Value',
        'direct_precedents_view=open|success',
        'direct_precedents_view_selection_preserved=True',
        'direct_precedents_view_content_preserved=True',
        'direct_precedents_view_open_workbook=retained|open',
        'direct_precedents_view_second_workbook=open|success',
        'direct_precedents_view_closed_workbook=discarded|closed',
        'direct_precedents_view_explicit_close=closed',
        'direct_dependents=Complete|B200,C200|',
        'direct_dependents_selection_preserved=True',
        'direct_dependents_content_preserved=True',
        'direct_dependents_cancelled=Refused|AUDIT_SCAN_CANCELLED|0',
        'direct_dependents_view=open|success',
        'direct_dependents_view_selection_preserved=True',
        'direct_dependents_view_explicit_close=closed',
        'indirect_precedents_view=open|success',
        'indirect_dependents_view=open|success',
        'trace_navigate=C210|recorded',
        'indirect_trace_content_preserved=True',
        'indirect_trace_explicit_close=closed',
        'numeric_hardcode_selection=A40,C40,B41,D42',
        'selection_content_preserved=True',
        'typed_conversions_exact=True',
        'typed_conversions_undo_exact=True',
        'paste_values_underlying_exact=True',
        'paste_values_undo_exact=True',
        'formula_from_above_exact=True',
        'value_from_above_exact=True',
        'numeric_sequence_exact=True',
        'date_sequence_exact=True',
        'formats_only_exact=True',
        'formats_only_undo_exact=True',
        'phase1b_feature_suite_ms=',
        'state_restored_after_fault=True',
        'stale_property_refused=True',
        'protected_target_refused=True',
        'multi_area_refused=True',
        'merged_target_refused=True',
        'working_set_bytes=',
        'private_memory_bytes=',
        'handle_count=',
        'workbook_closed=true',
        'quit_returned=true'
    )
    foreach ($evidenceLine in $requiredEvidence) {
        if ($output -notmatch "(?m)^$([regex]::Escape($evidenceLine))") {
            throw "Smoke-test evidence is incomplete; missing '$evidenceLine'. Worker output:`n$output"
        }
    }

    [pscustomobject]@{
        AddIn = $resolvedPath
        Passed = $true
        Evidence = $output.Trim()
    }
}
finally {
    foreach ($temporaryPath in @($outputPath, $errorPath)) {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}
