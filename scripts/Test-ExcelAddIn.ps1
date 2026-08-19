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
    $font = $null
    $interior = $null
    $multiArea = $null
    $mergedRange = $null
    $formulaRange = $null
    $formulaSource = $null
    $formulaDestination1 = $null
    $formulaDestination2 = $null
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

        $navigationCell = $worksheet.Range('D5')
        [void]$navigationCell.Select()
        [void]$excel.Run('ExcelAccel.Smoke.NavigateA1')
        $navigationAddress = [string]$excel.Selection.Address($false, $false)
        [Console]::WriteLine("navigation_address=$navigationAddress")
        [Console]::Out.Flush()
        if ($navigationAddress -ne 'A1') {
            throw 'Read-only A1 navigation did not select the exact target.'
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
        Release-ComObject $formulaDestination2
        Release-ComObject $formulaDestination1
        Release-ComObject $formulaSource
        Release-ComObject $formulaRange
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
