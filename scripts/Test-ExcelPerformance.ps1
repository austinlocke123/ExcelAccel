[CmdletBinding()]
param(
    [string]$AddInPath = "",
    [string]$CorpusPath = "",
    [string]$OutputPath = "",
    [ValidateSet('Quick', 'Qualification')]
    [string]$Profile = 'Quick',
    [ValidateRange(30, 600)]
    [int]$TimeoutSeconds = 180,
    [switch]$Worker,
    [ValidateSet('startup', 'block_read', 'property_write', 'workbook_read')]
    [string]$Operation = 'startup',
    [int]$Rows = 0,
    [int]$Columns = 0,
    [int]$Sheets = 0,
    [int]$WarmupIterations = 0,
    [int]$MeasuredIterations = 0
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Worker {
    param(
        [string]$ResolvedAddInPath,
        [string]$RequestedOperation,
        [int]$RequestedRows,
        [int]$RequestedColumns,
        [int]$RequestedSheets,
        [int]$RequestedWarmups,
        [int]$RequestedMeasurements
    )

    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

public static class ExcelAccelPerformanceNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}

public sealed class ExcelAccelHeartbeatMonitor : IDisposable
{
    private readonly IntPtr _window;
    private readonly Thread _thread;
    private volatile bool _stop;
    private long _maximumMilliseconds;
    private int _samples;
    private int _timeouts;

    public ExcelAccelHeartbeatMonitor(IntPtr window)
    {
        _window = window;
        _thread = new Thread(Run) { IsBackground = true, Name = "ExcelAccel performance heartbeat" };
        _thread.Start();
    }

    public long MaximumMilliseconds { get { return Interlocked.Read(ref _maximumMilliseconds); } }
    public int Samples { get { return Volatile.Read(ref _samples); } }
    public int Timeouts { get { return Volatile.Read(ref _timeouts); } }

    public void Dispose()
    {
        _stop = true;
        if (!_thread.Join(1000)) throw new TimeoutException("Heartbeat monitor did not stop within one second.");
    }

    private void Run()
    {
        while (!_stop)
        {
            var watch = Stopwatch.StartNew();
            IntPtr result;
            var succeeded = SendMessageTimeout(_window, 0, IntPtr.Zero, IntPtr.Zero, 2, 250, out result) != IntPtr.Zero;
            watch.Stop();
            Interlocked.Increment(ref _samples);
            if (!succeeded) Interlocked.Increment(ref _timeouts);
            var elapsed = watch.ElapsedMilliseconds;
            long prior;
            while (elapsed > (prior = Interlocked.Read(ref _maximumMilliseconds)) &&
                   Interlocked.CompareExchange(ref _maximumMilliseconds, elapsed, prior) != prior) { }
            Thread.Sleep(25);
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam,
        uint flags, uint timeout, out IntPtr result);
}
'@

    $excel = $null
    $workbook = $null
    $ranges = [System.Collections.Generic.List[object]]::new()
    $worksheets = [System.Collections.Generic.List[object]]::new()
    $quitReturned = $false
    $result = $null
    $heartbeat = $null
    $heartbeatSamples = 0
    $heartbeatTimeouts = 0
    $maximumHeartbeatMs = 0

    try {
        $launchWatch = [Diagnostics.Stopwatch]::StartNew()
        $excel = New-Object -ComObject Excel.Application
        $launchWatch.Stop()

        [uint32]$excelProcessId = 0
        [void][ExcelAccelPerformanceNativeMethods]::GetWindowThreadProcessId(
            [IntPtr]$excel.Hwnd,
            [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        [Console]::Out.Flush()

        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $registrationWatch = [Diagnostics.Stopwatch]::StartNew()
        $registered = $excel.RegisterXLL($ResolvedAddInPath)
        $registrationWatch.Stop()
        if (-not $registered) {
            throw 'Excel returned false from RegisterXLL.'
        }

        $excelVersion = [string]$excel.Version
        $excelBuild = try { [string]$excel.Build } catch { 'unknown' }
        $excelBitness = if ([Environment]::Is64BitProcess) { 64 } else { 32 }
        $startupTotalMs = $launchWatch.Elapsed.TotalMilliseconds + $registrationWatch.Elapsed.TotalMilliseconds

        if ($RequestedOperation -ne 'startup') {
            if ($RequestedRows -lt 1 -or $RequestedColumns -lt 1 -or $RequestedSheets -lt 1) {
                throw 'Non-startup workloads require positive rows, columns, and sheets.'
            }
            if ($RequestedMeasurements -lt 2 -or $RequestedWarmups -lt 0) {
                throw 'Measured iterations must be at least two and warmups cannot be negative.'
            }

            $workbook = $excel.Workbooks.Add()
            while ([int]$workbook.Worksheets.Count -lt $RequestedSheets) {
                $addedWorksheet = $workbook.Worksheets.Add()
                Release-ComObject $addedWorksheet
            }

            $values = [object[,]]::new($RequestedRows, $RequestedColumns)
            $seedValue = 1
            for ($rowIndex = 0; $rowIndex -lt $RequestedRows; $rowIndex++) {
                for ($columnIndex = 0; $columnIndex -lt $RequestedColumns; $columnIndex++) {
                    $values[$rowIndex, $columnIndex] = [double]$seedValue
                    $seedValue++
                }
            }

            for ($sheetIndex = 1; $sheetIndex -le $RequestedSheets; $sheetIndex++) {
                $worksheet = $workbook.Worksheets.Item($sheetIndex)
                $worksheets.Add($worksheet)
                $cells = $null
                $firstCell = $null
                $lastCell = $null
                try {
                    $cells = $worksheet.Cells
                    $firstCell = $cells.Item(1, 1)
                    $lastCell = $cells.Item($RequestedRows, $RequestedColumns)
                    $range = $worksheet.Range($firstCell, $lastCell)
                    $range.Value2 = $values
                    $ranges.Add($range)
                }
                finally {
                    Release-ComObject $lastCell
                    Release-ComObject $firstCell
                    Release-ComObject $cells
                }
            }

            $excelProcess = Get-Process -Id ([int]$excelProcessId)
            $excelProcess.Refresh()
            $workingSetBefore = [int64]$excelProcess.WorkingSet64
            $samples = [System.Collections.Generic.List[double]]::new()
            $maximumCallMs = 0.0
            $totalIterations = $RequestedWarmups + $RequestedMeasurements
            $heartbeat = [ExcelAccelHeartbeatMonitor]::new([IntPtr]$excel.Hwnd)

            for ($iteration = 0; $iteration -lt $totalIterations; $iteration++) {
                $iterationWatch = [Diagnostics.Stopwatch]::StartNew()

                if ($RequestedOperation -eq 'block_read') {
                    $callWatch = [Diagnostics.Stopwatch]::StartNew()
                    $snapshot = $ranges[0].Value2
                    $callWatch.Stop()
                    $maximumCallMs = [Math]::Max($maximumCallMs, $callWatch.Elapsed.TotalMilliseconds)
                    [GC]::KeepAlive($snapshot)
                }
                elseif ($RequestedOperation -eq 'property_write') {
                    $formatCode = if (($iteration % 2) -eq 0) { '0.00' } else { 'General' }
                    $callWatch = [Diagnostics.Stopwatch]::StartNew()
                    $ranges[0].NumberFormat = $formatCode
                    $callWatch.Stop()
                    $maximumCallMs = [Math]::Max($maximumCallMs, $callWatch.Elapsed.TotalMilliseconds)
                }
                elseif ($RequestedOperation -eq 'workbook_read') {
                    foreach ($range in $ranges) {
                        $callWatch = [Diagnostics.Stopwatch]::StartNew()
                        $snapshot = $range.Value2
                        $callWatch.Stop()
                        $maximumCallMs = [Math]::Max($maximumCallMs, $callWatch.Elapsed.TotalMilliseconds)
                        [GC]::KeepAlive($snapshot)
                    }
                }
                else {
                    throw "Unsupported worker operation '$RequestedOperation'."
                }

                $iterationWatch.Stop()
                if ($iteration -ge $RequestedWarmups) {
                    $samples.Add([Math]::Round($iterationWatch.Elapsed.TotalMilliseconds, 4))
                }
            }

            $heartbeat.Dispose()
            $heartbeatSamples = $heartbeat.Samples
            $heartbeatTimeouts = $heartbeat.Timeouts
            $maximumHeartbeatMs = $heartbeat.MaximumMilliseconds
            $heartbeat = $null

            $excelProcess.Refresh()
            $workingSetAfter = [int64]$excelProcess.WorkingSet64
        }
        else {
            $samples = [System.Collections.Generic.List[double]]::new()
            $maximumCallMs = 0.0
            $workingSetBefore = 0
            $workingSetAfter = 0
        }

        $result = [ordered]@{
            operation = $RequestedOperation
            excel_version = $excelVersion
            excel_build = $excelBuild
            excel_bitness = $excelBitness
            excel_pid = [int]$excelProcessId
            launch_ms = [Math]::Round($launchWatch.Elapsed.TotalMilliseconds, 4)
            registration_ms = [Math]::Round($registrationWatch.Elapsed.TotalMilliseconds, 4)
            startup_total_ms = [Math]::Round($startupTotalMs, 4)
            samples_ms = @($samples)
            maximum_excel_call_ms = [Math]::Round($maximumCallMs, 4)
            working_set_before_bytes = $workingSetBefore
            working_set_after_bytes = $workingSetAfter
            working_set_delta_bytes = $workingSetAfter - $workingSetBefore
            heartbeat_samples = $heartbeatSamples
            heartbeat_timeouts = $heartbeatTimeouts
            maximum_heartbeat_ms = $maximumHeartbeatMs
        }

        $snapshot = $null
        $range = $null
        $worksheet = $null
        for ($index = $ranges.Count - 1; $index -ge 0; $index--) {
            Release-ComObject $ranges[$index]
        }
        $ranges.Clear()
        for ($index = $worksheets.Count - 1; $index -ge 0; $index--) {
            Release-ComObject $worksheets[$index]
        }
        $worksheets.Clear()

        if ($null -ne $workbook) {
            $workbook.Close($false)
            Release-ComObject $workbook
            $workbook = $null
        }
        $excel.Quit()
        $quitReturned = $true
        Release-ComObject $excel
        $excel = $null
    }
    finally {
        if ($null -ne $heartbeat) {
            try { $heartbeat.Dispose() } catch { }
        }
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
        for ($index = $ranges.Count - 1; $index -ge 0; $index--) {
            Release-ComObject $ranges[$index]
        }
        for ($index = $worksheets.Count - 1; $index -ge 0; $index--) {
            Release-ComObject $worksheets[$index]
        }
        Release-ComObject $workbook
        Release-ComObject $excel
    }

    if ($null -eq $result -or -not $quitReturned) {
        throw 'The benchmark worker did not complete cleanly.'
    }

    $resultJson = $result | ConvertTo-Json -Compress -Depth 5
    $resultBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($resultJson))
    [Console]::WriteLine("result_base64=$resultBase64")
    [Console]::Out.Flush()
}

if ($Worker) {
    Invoke-Worker `
        -ResolvedAddInPath $AddInPath `
        -RequestedOperation $Operation `
        -RequestedRows $Rows `
        -RequestedColumns $Columns `
        -RequestedSheets $Sheets `
        -RequestedWarmups $WarmupIterations `
        -RequestedMeasurements $MeasuredIterations
    exit 0
}

function Get-DistributionSummary {
    param([double[]]$Samples)

    if ($null -eq $Samples -or $Samples.Count -lt 1) {
        throw 'At least one sample is required.'
    }

    $sorted = @($Samples | Sort-Object)
    $mean = ($sorted | Measure-Object -Average).Average
    $p95Index = [Math]::Max(0, [Math]::Ceiling(0.95 * $sorted.Count) - 1)
    $medianIndex = [Math]::Max(0, [Math]::Ceiling(0.50 * $sorted.Count) - 1)
    $deviationTotal = 0.0
    foreach ($sample in $sorted) {
        $deviationTotal += [Math]::Pow($sample - $mean, 2)
    }
    $standardDeviation = if ($sorted.Count -gt 1) {
        [Math]::Sqrt($deviationTotal / ($sorted.Count - 1))
    }
    else {
        0.0
    }

    [ordered]@{
        count = $sorted.Count
        minimum_ms = [Math]::Round($sorted[0], 4)
        maximum_ms = [Math]::Round($sorted[$sorted.Count - 1], 4)
        mean_ms = [Math]::Round($mean, 4)
        median_ms = [Math]::Round($sorted[$medianIndex], 4)
        p95_ms = [Math]::Round($sorted[$p95Index], 4)
        sample_standard_deviation_ms = [Math]::Round($standardDeviation, 4)
        coefficient_of_variation = if ($mean -eq 0) { 0 } else { [Math]::Round($standardDeviation / $mean, 6) }
    }
}

function Invoke-IsolatedWorker {
    param(
        [string]$ResolvedAddInPath,
        [string]$RequestedOperation,
        [int]$RequestedRows,
        [int]$RequestedColumns,
        [int]$RequestedSheets,
        [int]$RequestedWarmups,
        [int]$RequestedMeasurements
    )

    $workerId = [Guid]::NewGuid().ToString('N')
    $outputFile = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-performance-$workerId.out"
    $errorFile = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-performance-$workerId.err"
    $workerProcess = $null

    try {
        $arguments = @(
            '-NoProfile',
            '-NonInteractive',
            '-ExecutionPolicy', 'Bypass',
            '-File', "`"$PSCommandPath`"",
            '-Worker',
            '-AddInPath', "`"$ResolvedAddInPath`"",
            '-Operation', $RequestedOperation,
            '-Rows', $RequestedRows,
            '-Columns', $RequestedColumns,
            '-Sheets', $RequestedSheets,
            '-WarmupIterations', $RequestedWarmups,
            '-MeasuredIterations', $RequestedMeasurements
        )

        $workerProcess = Start-Process powershell.exe `
            -ArgumentList $arguments `
            -WindowStyle Hidden `
            -RedirectStandardOutput $outputFile `
            -RedirectStandardError $errorFile `
            -PassThru

        $completed = $workerProcess.WaitForExit($TimeoutSeconds * 1000)
        $output = if (Test-Path -LiteralPath $outputFile) { Get-Content -LiteralPath $outputFile -Raw } else { '' }
        $errors = if (Test-Path -LiteralPath $errorFile) { Get-Content -LiteralPath $errorFile -Raw } else { '' }

        if (-not $completed) {
            Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue
        }
        else {
            # The parameterless wait flushes redirected streams and makes the
            # final exit code reliable after the timed wait succeeds.
            $workerProcess.WaitForExit()
            $workerProcess.Refresh()
        }

        $excelProcessId = [regex]::Match($output, '(?m)^excel_pid=(\d+)').Groups[1].Value
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
            throw "Performance worker timed out after $TimeoutSeconds seconds. Output:`n$output`nErrors:`n$errors"
        }
        $exitCode = $workerProcess.ExitCode
        if (($null -ne $exitCode -and $exitCode -ne 0) -or -not ([string]::IsNullOrWhiteSpace($errors))) {
            throw "Performance worker failed with exit code $exitCode. Output:`n$output`nErrors:`n$errors"
        }

        $resultMarker = 'result_base64='
        $resultIndex = $output.IndexOf($resultMarker, [StringComparison]::Ordinal)
        if ($resultIndex -lt 0) {
            throw "Performance worker returned no encoded result. Output:`n$output"
        }

        # Windows PowerShell may visually wrap redirected console output. Base64
        # permits removing all whitespace without changing the payload.
        $encodedResult = $output.Substring($resultIndex + $resultMarker.Length) -replace '\s', ''
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encodedResult))
        return $json | ConvertFrom-Json
    }
    finally {
        foreach ($temporaryFile in @($outputFile, $errorFile)) {
            if (Test-Path -LiteralPath $temporaryFile) {
                Remove-Item -LiteralPath $temporaryFile -Force
            }
        }
    }
}

$existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($existingExcel.Count -ne 0) {
    throw 'Close all Excel processes before running the isolated performance harness.'
}

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($AddInPath)) {
    $AddInPath = Join-Path $repositoryRoot 'src\ExcelAccel.ExcelAddIn\bin\Release\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
}
if ([string]::IsNullOrWhiteSpace($CorpusPath)) {
    $CorpusPath = Join-Path $repositoryRoot 'benchmarks\performance-corpus-v1.json'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repositoryRoot '.tools\performance\wp-p0-07-latest.json'
}

$resolvedAddInPath = (Resolve-Path -LiteralPath $AddInPath).Path
$resolvedCorpusPath = (Resolve-Path -LiteralPath $CorpusPath).Path
$corpus = Get-Content -LiteralPath $resolvedCorpusPath -Raw | ConvertFrom-Json
$profileKey = $Profile.ToLowerInvariant()
$selectedProfile = $corpus.profiles.$profileKey
if ($null -eq $selectedProfile) {
    throw "The corpus does not define profile '$Profile'."
}

$startupResults = [System.Collections.Generic.List[object]]::new()
for ($sampleIndex = 0; $sampleIndex -lt [int]$selectedProfile.startup_samples; $sampleIndex++) {
    [Console]::WriteLine("progress=startup $($sampleIndex + 1)/$([int]$selectedProfile.startup_samples)")
    [Console]::Out.Flush()
    $startupResults.Add((Invoke-IsolatedWorker `
        -ResolvedAddInPath $resolvedAddInPath `
        -RequestedOperation 'startup' `
        -RequestedRows 0 `
        -RequestedColumns 0 `
        -RequestedSheets 0 `
        -RequestedWarmups 0 `
        -RequestedMeasurements 0))
}

$workloadResults = [System.Collections.Generic.List[object]]::new()
foreach ($workload in $corpus.workloads) {
    [Console]::WriteLine("progress=workload $([string]$workload.id)")
    [Console]::Out.Flush()
    $workerResult = Invoke-IsolatedWorker `
        -ResolvedAddInPath $resolvedAddInPath `
        -RequestedOperation ([string]$workload.operation) `
        -RequestedRows ([int]$workload.rows) `
        -RequestedColumns ([int]$workload.columns) `
        -RequestedSheets ([int]$workload.sheets) `
        -RequestedWarmups ([int]$selectedProfile.warmup_iterations) `
        -RequestedMeasurements ([int]$selectedProfile.measured_iterations)

    $workloadResults.Add([ordered]@{
        id = [string]$workload.id
        governing_requirement = [string]$workload.governing_requirement
        provisional_p95_ms = [double]$workload.provisional_p95_ms
        rows = [int]$workload.rows
        columns = [int]$workload.columns
        sheets = [int]$workload.sheets
        samples_ms = @($workerResult.samples_ms)
        distribution = Get-DistributionSummary -Samples @($workerResult.samples_ms)
        maximum_excel_call_ms = [double]$workerResult.maximum_excel_call_ms
        working_set_delta_bytes = [int64]$workerResult.working_set_delta_bytes
        heartbeat_samples = [int]$workerResult.heartbeat_samples
        heartbeat_timeouts = [int]$workerResult.heartbeat_timeouts
        maximum_heartbeat_ms = [int64]$workerResult.maximum_heartbeat_ms
    })
}

foreach ($workloadResult in $workloadResults) {
    if ($workloadResult.heartbeat_samples -lt 1 -or $workloadResult.heartbeat_timeouts -ne 0 -or $workloadResult.maximum_heartbeat_ms -gt 500) {
        throw "UI heartbeat gate failed for '$($workloadResult.id)': samples=$($workloadResult.heartbeat_samples), timeouts=$($workloadResult.heartbeat_timeouts), max_ms=$($workloadResult.maximum_heartbeat_ms)."
    }
    if ($workloadResult.distribution.p95_ms -gt $workloadResult.provisional_p95_ms) {
        throw "Performance gate failed for '$($workloadResult.id)': p95_ms=$($workloadResult.distribution.p95_ms), budget_ms=$($workloadResult.provisional_p95_ms)."
    }
}

$startupSamples = @($startupResults | ForEach-Object { [double]$_.startup_total_ms })
$coldSample = @($startupSamples[0])
$warmSamples = if ($startupSamples.Count -gt 1) { @($startupSamples[1..($startupSamples.Count - 1)]) } else { @() }
$firstResult = $startupResults[0]
$cpuName = $env:PROCESSOR_IDENTIFIER
$physicalMemory = 0
try {
    $computerSystem = Get-CimInstance Win32_ComputerSystem
    $physicalMemory = [int64]$computerSystem.TotalPhysicalMemory
}
catch {
    $physicalMemory = 0
}

$commit = 'unknown'
try {
    $commit = (& git -c "safe.directory=$repositoryRoot" -C $repositoryRoot rev-parse HEAD 2>$null).Trim()
}
catch {
    $commit = 'unknown'
}

$report = [ordered]@{
    schema_version = 1
    work_package = 'WP-P0-07'
    generated_utc = [DateTime]::UtcNow.ToString('o')
    commit = $commit
    fixture = [ordered]@{
        id = [string]$corpus.fixture_id
        schema_version = [int]$corpus.schema_version
        profile = $profileKey
    }
    machine = [ordered]@{
        windows_version = [Environment]::OSVersion.VersionString
        process_architecture = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
        cpu = $cpuName
        logical_processors = [Environment]::ProcessorCount
        physical_memory_bytes = $physicalMemory
        powershell_version = $PSVersionTable.PSVersion.ToString()
        dotnet_runtime = [Environment]::Version.ToString()
        excel_version = [string]$firstResult.excel_version
        excel_build = [string]$firstResult.excel_build
        excel_bitness = [int]$firstResult.excel_bitness
    }
    startup = [ordered]@{
        classification = 'first sample cold; subsequent fresh-process samples warm'
        cold_samples_ms = $coldSample
        cold_distribution = Get-DistributionSummary -Samples $coldSample
        warm_samples_ms = $warmSamples
        warm_distribution = if ($warmSamples.Count -gt 0) { Get-DistributionSummary -Samples $warmSamples } else { $null }
    }
    workloads = @($workloadResults)
    limitations = @(
        'Quick profile validates the harness only and cannot freeze budgets.',
        'Win32 WM_NULL heartbeat is an independent process-window responsiveness probe; it is not an end-user input-latency trace.',
        'Working-set delta from one process is not a retained-memory or leak qualification.'
    )
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
}
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding UTF8

$remainingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($remainingExcel.Count -ne 0) {
    throw 'The performance harness left an Excel process running.'
}

[pscustomobject]@{
    Passed = $true
    Profile = $Profile
    Fixture = [string]$corpus.fixture_id
    Output = (Resolve-Path -LiteralPath $OutputPath).Path
    ColdStartupMs = $coldSample[0]
    WarmStartupP95Ms = if ($warmSamples.Count -gt 0) { (Get-DistributionSummary -Samples $warmSamples).p95_ms } else { $null }
    Workloads = $workloadResults.Count
}
