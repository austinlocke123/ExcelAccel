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
    $multiArea = $null
    $mergedRange = $null
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
        Release-ComObject $multiArea
        Release-ComObject $secondCell
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
        'state_restored_after_fault=True',
        'stale_property_refused=True',
        'protected_target_refused=True',
        'multi_area_refused=True',
        'merged_target_refused=True',
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
