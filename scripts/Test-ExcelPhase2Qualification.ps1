[CmdletBinding()]
param(
    [string]$AddInPath = "",
    [string]$CorpusPath = "",
    [ValidateSet('Quick', 'Qualification')]
    [string]$Profile = 'Quick',
    [ValidateRange(60, 900)]
    [int]$TimeoutSeconds = 420,
    [ValidateRange(1, 20)]
    [int]$Iterations = 1,
    [switch]$Worker
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Get-Percentile {
    param([double[]]$Samples, [double]$Probability)

    $sorted = @($Samples | Sort-Object)
    if ($sorted.Count -eq 0) { return 0.0 }
    if ($sorted.Count -eq 1) { return [double]$sorted[0] }
    $rank = $Probability * ($sorted.Count - 1)
    $lower = [Math]::Floor($rank)
    $upper = [Math]::Ceiling($rank)
    if ($lower -eq $upper) { return [double]$sorted[[int]$rank] }
    return [double]$sorted[[int]$lower] + (($rank - $lower) * ([double]$sorted[[int]$upper] - [double]$sorted[[int]$lower]))
}

function Invoke-Worker {
    param([string]$ResolvedAddInPath, [string]$ResolvedCorpusPath, [string]$ResolvedProfile)

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelAccelPhase2NativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

    $corpus = Get-Content -LiteralPath $ResolvedCorpusPath -Raw | ConvertFrom-Json
    $profileSettings = $corpus.profiles.($ResolvedProfile.ToLowerInvariant())
    $warmups = [int]$profileSettings.warmup_iterations
    $measured = [int]$profileSettings.measured_iterations

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $quitReturned = $false
    $needle = 'EXCELACCELSEED' + ([Guid]::NewGuid().ToString('N').Substring(0, 12).ToUpperInvariant())

    try {
        $excel = New-Object -ComObject Excel.Application
        [uint32]$excelProcessId = 0
        [void][ExcelAccelPhase2NativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $registered = [bool]$excel.RegisterXLL($ResolvedAddInPath)
        [Console]::WriteLine("registered=$registered")
        [Console]::Out.Flush()
        if (-not $registered) { throw "The add-in at '$ResolvedAddInPath' did not register." }

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        # The worksheet name carries the marker, so a leak through sheet identity
        # would be caught as well as one through a formula or value.
        $worksheet.Name = 'C' + $needle

        $rows = [int]$corpus.corpus.formula_rows
        $columns = [int]$corpus.corpus.formula_columns

        # Build the corpus with array assignment: cell-by-cell COM writes would
        # dominate the run time and tell us nothing about the add-in.
        # Computed indices are parenthesised because PowerShell would otherwise
        # read "$row - 1, 0" as subtracting the array (1, 0).
        $inputs = New-Object 'object[,]' $rows, 1
        for ($row = 1; $row -le $rows; $row++) { $inputs[($row - 1), 0] = $row }
        $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item($rows, 1)).Value2 = $inputs

        $formulas = New-Object 'object[,]' $rows, $columns
        for ($row = 1; $row -le $rows; $row++) {
            for ($column = 1; $column -le $columns; $column++) {
                $formulas[($row - 1), ($column - 1)] = "=`$A$row*$column"
            }
        }

        # Seeded rule violations at known rows.
        $formulas[9, 0] = '=$A10*99'
        $formulas[499, 0] = '=$A500*97'
        $formulas[999, 0] = '=$A1000*93'
        $formulas[1499, 0] = '=$A1500*91'
        $formulas[19, 1] = '=1/0'
        $formulas[29, 1] = '=#REF!+1'
        $formulas[39, 2] = "='[OtherBook.xlsx]Data'!A1"
        $formulas[49, 2] = "='[OtherBook.xlsx]Data'!A2"
        $formulas[1998, 4] = '=' + $needle

        $worksheet.Range($worksheet.Cells.Item(1, 2), $worksheet.Cells.Item($rows, 1 + $columns)).Formula = $formulas

        # The marker also appears as a defined name and as a literal cell value.
        [void]$workbook.Names.Add($needle, "='$($worksheet.Name)'!`$A`$1")
        $worksheet.Range('H1').Value2 = $needle

        $deepColumn = 1 + $columns + 2
        $worksheet.Cells.Item(1, $deepColumn).Formula = '=$A$1*2'
        for ($row = 2; $row -le 40; $row++) {
            $worksheet.Cells.Item($row, $deepColumn).Formula = "=$([char](64 + $deepColumn))$($row - 1)+1"
        }

        $excel.Calculate()
        [Console]::WriteLine("corpus_rows=$rows")
        [Console]::WriteLine("corpus_columns=$columns")
        [Console]::Out.Flush()

        $results = @{}
        foreach ($workload in $corpus.workloads) {
            $samples = New-Object System.Collections.Generic.List[double]
            $detail = ''
            for ($iteration = 1; $iteration -le ($warmups + $measured); $iteration++) {
                switch ($workload.operation) {
                    'direct_precedents' {
                        [void]$worksheet.Cells.Item(40, $deepColumn).Select()
                        $raw = [string]$excel.Run('ExcelAccel.Perf.DirectPrecedents')
                    }
                    'dependent_scan' {
                        [void]$worksheet.Range('A100').Select()
                        $raw = [string]$excel.Run('ExcelAccel.Perf.DependentScan')
                    }
                    'indirect_precedents' {
                        [void]$worksheet.Cells.Item(40, $deepColumn).Select()
                        $raw = [string]$excel.Run('ExcelAccel.Perf.IndirectPrecedents')
                    }
                    'model_check_worksheet' {
                        [void]$worksheet.Range('A100').Select()
                        $raw = [string]$excel.Run('ExcelAccel.Perf.ModelCheckWorksheet')
                    }
                    default { throw "Unknown workload operation '$($workload.operation)'." }
                }

                $parts = $raw.Split('|')
                if ($iteration -gt $warmups) { $samples.Add([double]$parts[0]) }
                $detail = $parts[1]
            }

            $p95 = Get-Percentile -Samples $samples.ToArray() -Probability 0.95
            $results[$workload.id] = $p95
            [Console]::WriteLine("workload=$($workload.id) p95_ms=$([Math]::Round($p95, 1)) budget_ms=$($workload.provisional_p95_ms) detail=$detail")
            [Console]::Out.Flush()
            if ($p95 -gt [double]$workload.provisional_p95_ms) {
                throw "Workload '$($workload.id)' P95 $([Math]::Round($p95,1)) ms exceeded its $($workload.provisional_p95_ms) ms budget."
            }
        }

        # Cancellation at corpus scale must fail closed and return promptly.
        [void]$worksheet.Range('A100').Select()
        $cancelRaw = [string]$excel.Run('ExcelAccel.Perf.ModelCheckCancelled')
        $cancelParts = $cancelRaw.Split('|')
        $cancelMs = [double]$cancelParts[0]
        [Console]::WriteLine("cancellation_ms=$cancelMs detail=$($cancelParts[1])")
        [Console]::Out.Flush()
        if ($cancelParts[1] -notlike 'Refused*') { throw "A cancelled scan did not refuse: $cancelRaw" }
        if ($cancelParts[1] -notlike "*$($corpus.cancellation.expected_refusal_code)*") {
            throw "A cancelled scan did not carry the expected refusal code: $cancelRaw"
        }
        if ($cancelMs -gt [double]$corpus.cancellation.maximum_ms) {
            throw "Cancellation took $cancelMs ms, above its $($corpus.cancellation.maximum_ms) ms ceiling."
        }

        # Privacy: after every Phase 2 operation, the marker must not survive
        # into the exported diagnostics.
        $privacyRaw = [string]$excel.Run('ExcelAccel.Perf.PrivacyProbe', $needle)
        [Console]::WriteLine("privacy=$privacyRaw")
        [Console]::Out.Flush()
        if (-not $privacyRaw.StartsWith('clean|')) {
            throw "The seeded marker survived into the exported diagnostics: $privacyRaw"
        }

        # The corpus is read-only throughout; nothing is saved.
        $contentPreserved = ([string]$worksheet.Range('B10').Formula -eq '=$A10*99')
        [Console]::WriteLine("content_preserved=$contentPreserved")
        [Console]::Out.Flush()
        if (-not $contentPreserved) { throw 'A Phase 2 operation changed the corpus.' }

        $process = Get-Process -Id ([int]$excelProcessId) -ErrorAction SilentlyContinue
        if ($null -ne $process) {
            $process.Refresh()
            [Console]::WriteLine("working_set_bytes=$($process.WorkingSet64)")
            [Console]::WriteLine("handle_count=$($process.HandleCount)")
            [Console]::Out.Flush()
        }

        $workbook.Close($false)
        $workbook = $null
        [Console]::WriteLine('workbook_closed=true')
        $excel.Quit()
        $quitReturned = $true
        [Console]::WriteLine('quit_returned=true')
        [Console]::Out.Flush()
    }
    finally {
        try { if ($null -ne $workbook) { $workbook.Close($false) } } catch { }
        try { if ($null -ne $excel -and -not $quitReturned) { $excel.Quit() } } catch { }
        Release-ComObject $worksheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($Worker) {
    Invoke-Worker -ResolvedAddInPath $AddInPath -ResolvedCorpusPath $CorpusPath -ResolvedProfile $Profile
    exit 0
}

if ([string]::IsNullOrWhiteSpace($AddInPath)) {
    $AddInPath = Join-Path $PSScriptRoot '..\src\ExcelAccel.ExcelAddIn\bin\Debug\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
}

if ([string]::IsNullOrWhiteSpace($CorpusPath)) {
    $CorpusPath = Join-Path $PSScriptRoot '..\benchmarks\phase2-corpus-v1.json'
}

$resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath).Path
$resolvedCorpus = (Resolve-Path -LiteralPath $CorpusPath).Path
$iterationResults = New-Object System.Collections.Generic.List[object]

for ($iteration = 1; $iteration -le $Iterations; $iteration++) {
    $runId = [Guid]::NewGuid().ToString('N')
    $outputPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-phase2-$runId.out"
    $errorPath = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-phase2-$runId.err"
    $workerProcess = $null

    try {
        $arguments = @(
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', "`"$PSCommandPath`"", '-Worker',
            '-AddInPath', "`"$resolvedAddIn`"",
            '-CorpusPath', "`"$resolvedCorpus`"",
            '-Profile', $Profile
        )

        $workerProcess = Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden `
            -RedirectStandardOutput $outputPath -RedirectStandardError $errorPath -PassThru
        $completed = $workerProcess.WaitForExit($TimeoutSeconds * 1000)
        $output = if (Test-Path -LiteralPath $outputPath) { Get-Content -LiteralPath $outputPath -Raw } else { '' }
        $errors = if (Test-Path -LiteralPath $errorPath) { Get-Content -LiteralPath $errorPath -Raw } else { '' }

        if (-not $completed) {
            Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
            throw "Iteration $iteration timed out after $TimeoutSeconds seconds. Output:`n$output"
        }

        $excelProcessId = [regex]::Match($output, '(?m)^excel_pid=(\d+)$').Groups[1].Value
        if ($excelProcessId) {
            $excelProcess = Get-Process -Id ([int]$excelProcessId) -ErrorAction SilentlyContinue
            if ($excelProcess) {
                if (-not $excelProcess.WaitForExit(5000)) {
                    Stop-Process -Id $excelProcess.Id -Force
                    throw "Iteration $iteration left Excel PID $excelProcessId running. Output:`n$output"
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($errors)) {
            throw "Iteration $iteration reported an error. Output:`n$output`nErrors:`n$errors"
        }

        foreach ($required in @('registered=True', 'privacy=clean|', 'content_preserved=True', 'workbook_closed=true', 'quit_returned=true')) {
            if ($output -notmatch "(?m)^$([regex]::Escape($required))") {
                throw "Iteration $iteration evidence is incomplete; missing '$required'. Output:`n$output"
            }
        }

        $workloads = @{}
        foreach ($match in [regex]::Matches($output, '(?m)^workload=(\S+) p95_ms=(\S+) budget_ms=(\S+)')) {
            $workloads[$match.Groups[1].Value] = [double]$match.Groups[2].Value
        }

        $iterationResults.Add([pscustomobject]@{
            Iteration = $iteration
            Workloads = $workloads
            CancellationMs = [double]([regex]::Match($output, '(?m)^cancellation_ms=(\S+)').Groups[1].Value)
            HandleCount = [int]([regex]::Match($output, '(?m)^handle_count=(\d+)').Groups[1].Value)
            WorkingSetBytes = [int64]([regex]::Match($output, '(?m)^working_set_bytes=(\d+)').Groups[1].Value)
            Evidence = $output.Trim()
        })

        Write-Host "Iteration $iteration passed."
    }
    finally {
        Remove-Item -LiteralPath $outputPath -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $errorPath -ErrorAction SilentlyContinue
    }
}

$summary = [ordered]@{}
foreach ($id in ($iterationResults[0].Workloads.Keys | Sort-Object)) {
    $samples = @($iterationResults | ForEach-Object { $_.Workloads[$id] })
    $summary[$id] = [Math]::Round((Get-Percentile -Samples $samples -Probability 0.95), 1)
}

[pscustomobject]@{
    AddIn = $resolvedAddIn
    Corpus = $resolvedCorpus
    Profile = $Profile
    Iterations = $Iterations
    Passed = $true
    WorkloadP95Ms = $summary
    CancellationP95Ms = [Math]::Round((Get-Percentile -Samples @($iterationResults | ForEach-Object { $_.CancellationMs }) -Probability 0.95), 1)
    HandleCountP95 = [Math]::Round((Get-Percentile -Samples @($iterationResults | ForEach-Object { [double]$_.HandleCount }) -Probability 0.95), 0)
    WorkingSetP95Bytes = [Math]::Round((Get-Percentile -Samples @($iterationResults | ForEach-Object { [double]$_.WorkingSetBytes }) -Probability 0.95), 0)
    Evidence = $iterationResults[$iterationResults.Count - 1].Evidence
}
