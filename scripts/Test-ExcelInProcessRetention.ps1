<#
.SYNOPSIS
    Measures whether one long-lived Excel session retains resources as ExcelAccel
    operations are repeated inside it.

.DESCRIPTION
    The reliability soak launches a fresh Excel process per iteration, so it
    proves process cleanup and cross-session stability and says so in its own
    limitations. What it cannot show, at any iteration count, is whether a single
    session grows while the user works in it — the process ends before drift
    could accumulate.

    This harness does the opposite: one Excel process, one workbook, and the
    Phase 2 read-only operations repeated many times, sampling working set,
    private memory, and handle count throughout. Drift is measured between the
    first and last samples, after a warm-up that lets one-time allocations settle.

    Every operation is read-only and the workbook is never saved.
#>
[CmdletBinding()]
param(
    [ValidateRange(4, 200)]
    [int]$Cycles = 30,

    [ValidateRange(1, 50)]
    [int]$WarmupCycles = 5,

    [ValidateRange(1, 60)]
    [double]$MaximumDriftPercent = 10,

    [string]$AddInPath,

    [ValidateRange(60, 7200)]
    [int]$TimeoutSeconds = 1800,

    [switch]$Worker
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Resolve-AddInPath {
    param([string]$Candidate)
    if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
        if (-not (Test-Path -LiteralPath $Candidate)) { throw "The add-in was not found at '$Candidate'." }
        return (Resolve-Path -LiteralPath $Candidate).Path
    }

    $default = Join-Path $PSScriptRoot '..\src\ExcelAccel.ExcelAddIn\bin\Debug\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
    if (-not (Test-Path -LiteralPath $default)) {
        throw "Build the Debug packed add-in first; it was not found at '$default'."
    }

    return (Resolve-Path -LiteralPath $default).Path
}

if ($Worker) {
    $resolvedPath = Resolve-AddInPath -Candidate $AddInPath
    $excel = $null
    $workbook = $null
    $worksheet = $null
    try {
        Add-Type -Namespace ExcelAccelRetention -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int processId);
'@

        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = 0
        [void][ExcelAccelRetention.NativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $registered = $excel.RegisterXLL($resolvedPath)
        [Console]::WriteLine("registered=$registered")
        if (-not $registered) { throw 'The add-in did not register.' }

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'Retention'

        # A modest corpus: large enough that the operations do real work, small
        # enough that hundreds of cycles finish in a sensible time.
        $rows = 400
        $inputs = New-Object 'object[,]' $rows, 1
        for ($row = 1; $row -le $rows; $row++) { $inputs[($row - 1), 0] = $row }
        $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($rows, 1)).Value2 = $inputs

        $formulas = New-Object 'object[,]' $rows, 3
        for ($row = 1; $row -le $rows; $row++) {
            for ($column = 1; $column -le 3; $column++) {
                $formulas[($row - 1), ($column - 1)] = "=`$A$row*$column"
            }
        }
        $worksheet.Range($worksheet.Cells.Item(1, 2), $worksheet.Cells.Item($rows, 4)).Formula = $formulas
        $excel.Calculate()

        $process = Get-Process -Id $excelProcessId
        $samples = New-Object System.Collections.Generic.List[object]

        for ($cycle = 1; $cycle -le ($WarmupCycles + $Cycles); $cycle++) {
            [void]$worksheet.Range('B10').Select()
            [void]$excel.Run('ExcelAccel.Perf.DirectPrecedents')
            [void]$worksheet.Range('A10').Select()
            [void]$excel.Run('ExcelAccel.Perf.DependentScan')
            [void]$worksheet.Range('B10').Select()
            [void]$excel.Run('ExcelAccel.Perf.IndirectPrecedents')
            [void]$worksheet.Range('A10').Select()
            [void]$excel.Run('ExcelAccel.Perf.ModelCheckWorksheet')

            if ($cycle -le $WarmupCycles) { continue }

            # Collect inside Excel's CLR, where the managed add-in lives. A
            # PowerShell GC only collects this controller process and says
            # nothing about allocations retained by ExcelAccel.
            $gcStatus = [string]$excel.Run('ExcelAccel.Perf.CollectGarbage')
            if ($gcStatus -ne 'collected') {
                throw "The in-process collection hook returned '$gcStatus'."
            }
            $process.Refresh()
            $samples.Add([pscustomobject]@{
                cycle                = $cycle - $WarmupCycles
                working_set_bytes    = $process.WorkingSet64
                private_memory_bytes = $process.PrivateMemorySize64
                handle_count         = $process.HandleCount
            })
        }

        # Compare the first and last thirds rather than single samples, so one
        # noisy reading cannot decide the result either way.
        $window = [Math]::Max(1, [int][Math]::Floor($samples.Count / 3))
        $firstSlice = $samples[0..($window - 1)]
        $lastSlice = $samples[($samples.Count - $window)..($samples.Count - 1)]

        foreach ($field in @('working_set_bytes', 'private_memory_bytes', 'handle_count')) {
            $firstAverage = ($firstSlice | Measure-Object -Property $field -Average).Average
            $lastAverage = ($lastSlice | Measure-Object -Property $field -Average).Average
            $drift = if ($firstAverage -gt 0) { (($lastAverage - $firstAverage) / $firstAverage) * 100 } else { 0 }
            [Console]::WriteLine("drift $field first=$([Math]::Round($firstAverage,0)) last=$([Math]::Round($lastAverage,0)) percent=$([Math]::Round($drift,2))")
            if ($drift -gt $MaximumDriftPercent) {
                throw "$field grew $([Math]::Round($drift,2))% across one session, above the $MaximumDriftPercent% ceiling."
            }
        }

        [Console]::WriteLine("cycles_measured=$($samples.Count)")
        [Console]::WriteLine('gc_in_excel=True')

        $contentPreserved = ([string]$worksheet.Range('B10').Formula -eq '=$A10*1')
        [Console]::WriteLine("content_preserved=$contentPreserved")
        if (-not $contentPreserved) { throw 'A read-only operation changed the corpus.' }

        $outputDirectory = Join-Path $PSScriptRoot '..\.tools\reliability'
        [void][System.IO.Directory]::CreateDirectory($outputDirectory)
        $report = [pscustomobject]@{
            schema_version = 1
            work_package   = 'WP-R-06'
            generated_utc  = (Get-Date).ToUniversalTime().ToString('o')
            add_in         = $resolvedPath
            cycles         = $samples.Count
            warmup_cycles  = $WarmupCycles
            samples        = $samples
            limitations    = @(
                'One Excel process throughout, which is what makes in-process retention visible; process cleanup is covered by the reliability soak instead.',
                'The corpus is deliberately small so many cycles fit in a sensible run; it measures retention, not throughput.',
                'Drift compares the first and last thirds of the samples after a warm-up, so a single noisy reading cannot decide the outcome.'
            )
        }
        $reportPath = Join-Path $outputDirectory 'in-process-retention-latest.json'
        $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
        [Console]::WriteLine("report=$reportPath")

        $workbook.Close($false)
        [Console]::WriteLine('workbook_closed=true')
        Release-ComObject $worksheet
        $worksheet = $null
        Release-ComObject $workbook
        $workbook = $null
        $excel.Quit()
        [Console]::WriteLine('quit_returned=true')
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
        Release-ComObject $worksheet
        Release-ComObject $workbook
        if ($null -ne $excel) { try { $excel.Quit() } catch { } }
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return
}

# Parent: an existing Excel would share the process this harness measures.
if (@(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'In-process retention measurement requires Excel to be closed before it starts; no existing Excel process will be terminated.'
}

$resolvedPath = Resolve-AddInPath -Candidate $AddInPath
$outputPath = [System.IO.Path]::GetTempFileName()
$errorPath = [System.IO.Path]::GetTempFileName()

try {
    $arguments = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-File', "`"$PSCommandPath`"",
        '-Worker',
        '-AddInPath', "`"$resolvedPath`"",
        '-Cycles', $Cycles,
        '-WarmupCycles', $WarmupCycles,
        '-MaximumDriftPercent', $MaximumDriftPercent
    )

    $workerProcess = Start-Process powershell.exe `
        -ArgumentList $arguments `
        -WindowStyle Hidden `
        -RedirectStandardOutput $outputPath `
        -RedirectStandardError $errorPath `
        -PassThru

    $completed = $workerProcess.WaitForExit($TimeoutSeconds * 1000)
    if ($completed) {
        # The parameterless wait completes redirected-stream handling and makes
        # ExitCode reliable after the timed wait reports that the process ended.
        $workerProcess.WaitForExit()
    }
    $output = if (Test-Path -LiteralPath $outputPath) { [string](Get-Content -LiteralPath $outputPath -Raw) } else { '' }
    $errors = if (Test-Path -LiteralPath $errorPath) { [string](Get-Content -LiteralPath $errorPath -Raw) } else { '' }
    if ($null -eq $output) { $output = '' }
    if ($null -eq $errors) { $errors = '' }

    if (-not $completed) {
        Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
        throw "The retention worker did not finish within $TimeoutSeconds seconds. Output:`n$output`n$errors"
    }

    # Target the reported PID only; never terminate Excel by name.
    $excelProcessId = [regex]::Match($output, '(?m)^excel_pid=(\d+)').Groups[1].Value
    if ($excelProcessId) {
        $excelProcess = Get-Process -Id ([int]$excelProcessId) -ErrorAction SilentlyContinue
        if ($excelProcess) {
            if (-not $excelProcess.WaitForExit(5000)) {
                Stop-Process -Id $excelProcess.Id -Force
                throw "Excel PID $excelProcessId did not exit cleanly. Output:`n$output"
            }
        }
    }

    # Some Windows PowerShell hosts leave ExitCode unset on a Start-Process
    # wrapper even after both waits. A known non-zero code is a failure; when it
    # is unavailable, the required worker evidence below remains the fail-closed
    # completion check.
    if (($null -ne $workerProcess.ExitCode -and $workerProcess.ExitCode -ne 0) -or $errors.Trim().Length -gt 0) {
        throw "The retention worker reported an error. Output:`n$output`nErrors:`n$errors"
    }

    foreach ($expected in @('registered=True', 'gc_in_excel=True', 'content_preserved=True', 'workbook_closed=true', 'quit_returned=true')) {
        if ($output -notmatch [regex]::Escape($expected)) {
            throw "The retention run did not report '$expected'. Output:`n$output"
        }
    }

    Write-Output $output.Trim()
}
finally {
    Remove-Item -LiteralPath $outputPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $errorPath -ErrorAction SilentlyContinue
}
