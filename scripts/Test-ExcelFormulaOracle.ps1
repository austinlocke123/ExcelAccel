[CmdletBinding()]
param(
    [string]$CorpusPath = "",
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
    param([string]$ResolvedCorpusPath)

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelAccelFormulaOracleNativeMethods
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
        $corpus = Get-Content -LiteralPath $ResolvedCorpusPath -Raw | ConvertFrom-Json
        $oracleCases = @($corpus | Where-Object { $_.excelOracle -eq $true })
        if ($oracleCases.Count -eq 0) {
            throw 'The formula corpus contains no real-Excel oracle cases.'
        }

        $excel = New-Object -ComObject Excel.Application
        [uint32]$excelProcessId = 0
        [void][ExcelAccelFormulaOracleNativeMethods]::GetWindowThreadProcessId(
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
        $cell = $worksheet.Range('Z100')

        $passed = 0
        foreach ($testCase in $oracleCases) {
            $cell.ClearContents()
            $formula = [string]$testCase.formula
            if ([string]$testCase.notation -eq 'R1C1') {
                $cell.FormulaR1C1 = $formula
                $roundTrip = [string]$cell.FormulaR1C1
            }
            else {
                $cell.Formula = $formula
                $roundTrip = [string]$cell.Formula
            }

            if ($roundTrip -ne $formula) {
                throw "Formula oracle case '$($testCase.id)' changed '$formula' to '$roundTrip'."
            }

            $passed++
        }

        [Console]::WriteLine("oracle_cases=$passed")
        [Console]::WriteLine('oracle_round_trip=True')
        [Console]::Out.Flush()

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
    Invoke-Worker -ResolvedCorpusPath $CorpusPath
    exit 0
}

if ([string]::IsNullOrWhiteSpace($CorpusPath)) {
    $CorpusPath = Join-Path $PSScriptRoot '..\tests\ExcelAccel.Core.Tests\Fixtures\formula-v1-corpus.json'
}

$resolvedPath = (Resolve-Path -LiteralPath $CorpusPath).Path
$runId = [Guid]::NewGuid().ToString('N')
$outputPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-formula-oracle-$runId.out"
$errorPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-formula-oracle-$runId.err"
$workerProcess = $null

try {
    $arguments = @(
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$PSCommandPath`"",
        '-Worker',
        '-CorpusPath', "`"$resolvedPath`""
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
        throw "Formula-oracle worker timed out after $TimeoutSeconds seconds. Worker output:`n$output`nWorker errors:`n$errors"
    }

    if (-not [string]::IsNullOrWhiteSpace($errors)) {
        throw "Formula-oracle worker reported an error. Worker output:`n$output`nWorker errors:`n$errors"
    }

    $requiredEvidence = @(
        'oracle_cases=',
        'excel_version=',
        'excel_build=',
        'oracle_round_trip=True',
        'workbook_closed=true',
        'quit_returned=true'
    )
    foreach ($evidenceLine in $requiredEvidence) {
        if ($output -notmatch "(?m)^$([regex]::Escape($evidenceLine))") {
            throw "Formula-oracle evidence is incomplete; missing '$evidenceLine'. Worker output:`n$output"
        }
    }

    [pscustomobject]@{
        Corpus = $resolvedPath
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
