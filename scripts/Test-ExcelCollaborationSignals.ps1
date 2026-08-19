[CmdletBinding()]
param(
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
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelAccelCollaborationNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

    $excel = $null
    $workbooks = $null
    $workbook = $null
    $worksheets = $null
    $worksheet = $null
    $cell = $null
    $quitReturned = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        [uint32]$excelProcessId = 0
        [void][ExcelAccelCollaborationNativeMethods]::GetWindowThreadProcessId(
            [IntPtr]$excel.Hwnd,
            [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        [Console]::WriteLine("excel_version=$([string]$excel.Version)")
        [Console]::WriteLine("excel_build=$([string]$excel.Build)")
        [Console]::Out.Flush()

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Add()
        $worksheets = $workbook.Worksheets
        $worksheet = $worksheets.Item(1)
        $cell = $worksheet.Range('A1')

        $autoSaveSupported = $false
        $autoSaveBefore = $null
        try {
            $autoSaveBefore = [bool]$workbook.AutoSaveOn
            $autoSaveSupported = $true
        }
        catch {
        }

        $legacySharedSupported = $false
        $legacyShared = $null
        try {
            $legacyShared = [bool]$workbook.MultiUserEditing
            $legacySharedSupported = $true
        }
        catch {
        }

        $pathIsEmpty = [string]::IsNullOrEmpty([string]$workbook.Path)
        $cell.NumberFormat = 'General'
        $plannedFormat = [string]$cell.NumberFormat
        $cell.NumberFormat = '0.00'
        $currentFormat = [string]$cell.NumberFormat
        $interveningPropertyChangeDetected = $plannedFormat -ne $currentFormat

        $autoSaveUnchanged = $true
        if ($autoSaveSupported) {
            $autoSaveAfter = [bool]$workbook.AutoSaveOn
            $autoSaveUnchanged = $autoSaveBefore -eq $autoSaveAfter
        }

        [Console]::WriteLine("autosave_read_supported=$autoSaveSupported")
        [Console]::WriteLine("autosave_value=$autoSaveBefore")
        [Console]::WriteLine("autosave_unchanged=$autoSaveUnchanged")
        [Console]::WriteLine("legacy_shared_read_supported=$legacySharedSupported")
        [Console]::WriteLine("legacy_shared_value=$legacyShared")
        [Console]::WriteLine("unsaved_path=$pathIsEmpty")
        [Console]::WriteLine("intervening_property_change_detected=$interveningPropertyChangeDetected")
        [Console]::Out.Flush()

        if (-not $autoSaveSupported -or -not $autoSaveUnchanged) {
            throw 'The read-only AutoSave probe was unavailable or changed AutoSave state.'
        }

        if (-not $legacySharedSupported -or $legacyShared) {
            throw 'The temporary workbook did not expose the expected non-shared legacy state.'
        }

        if (-not $pathIsEmpty -or -not $interveningPropertyChangeDetected) {
            throw 'The temporary-workbook signal fixture did not match its expected preconditions.'
        }

        $workbook.Close($false)
        [Console]::WriteLine('workbook_closed=true')
        [Console]::Out.Flush()

        Release-ComObject $cell
        $cell = $null
        Release-ComObject $worksheet
        $worksheet = $null
        Release-ComObject $worksheets
        $worksheets = $null
        Release-ComObject $workbook
        $workbook = $null
        Release-ComObject $workbooks
        $workbooks = $null

        $excel.Quit()
        $quitReturned = $true
        [Console]::WriteLine('quit_returned=true')
        [Console]::Out.Flush()
        Release-ComObject $excel
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
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
        Release-ComObject $cell
        Release-ComObject $worksheet
        Release-ComObject $worksheets
        Release-ComObject $workbook
        Release-ComObject $workbooks
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($Worker) {
    Invoke-Worker
    exit 0
}

$runId = [Guid]::NewGuid().ToString('N')
$outputPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-collaboration-$runId.out"
$errorPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-collaboration-$runId.err"
$workerProcess = $null

try {
    $arguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$PSCommandPath`"",
        '-Worker'
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
                throw "Excel PID $excelProcessId did not exit cleanly. Worker output:`n$output"
            }
        }
    }

    if (-not $completed) {
        throw "Collaboration-signal worker timed out after $TimeoutSeconds seconds. Worker output:`n$output`nWorker errors:`n$errors"
    }

    if (-not [string]::IsNullOrWhiteSpace($errors)) {
        throw "Collaboration-signal worker reported an error. Worker output:`n$output`nWorker errors:`n$errors"
    }

    $requiredEvidence = @(
        'excel_version=',
        'excel_build=',
        'autosave_read_supported=True',
        'autosave_unchanged=True',
        'legacy_shared_read_supported=True',
        'legacy_shared_value=False',
        'unsaved_path=True',
        'intervening_property_change_detected=True',
        'workbook_closed=true',
        'quit_returned=true'
    )
    foreach ($evidenceLine in $requiredEvidence) {
        if ($output -notmatch "(?m)^$([regex]::Escape($evidenceLine))") {
            throw "Collaboration-signal evidence is incomplete; missing '$evidenceLine'. Worker output:`n$output"
        }
    }

    [pscustomobject]@{
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
