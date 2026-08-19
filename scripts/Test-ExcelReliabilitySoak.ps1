[CmdletBinding()]
param(
    [string]$AddInPath = "",
    [ValidateRange(3, 100)]
    [int]$Iterations = 10,
    [ValidateRange(10, 120)]
    [int]$TimeoutSeconds = 40,
    [string]$OutputPath = ""
)

$ErrorActionPreference = 'Stop'

function Get-Percentile {
    param(
        [double[]]$Samples,
        [ValidateRange(0, 100)]
        [double]$Percentile
    )

    if ($null -eq $Samples -or $Samples.Count -eq 0) {
        throw 'At least one sample is required.'
    }

    $sorted = @($Samples | Sort-Object)
    $index = [Math]::Ceiling(($Percentile / 100.0) * $sorted.Count) - 1
    return [double]$sorted[[Math]::Max(0, $index)]
}

function Get-RequiredMetric {
    param(
        [string]$Evidence,
        [string]$Name
    )

    $match = [regex]::Match($Evidence, "(?m)^$([regex]::Escape($Name))=(\d+)\r?$")
    if (-not $match.Success) {
        throw "Smoke evidence did not contain metric '$Name'."
    }

    return [int64]$match.Groups[1].Value
}

function Assert-AddInUnlocked {
    param([string]$Path)

    $stream = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Wait-ForNoExcel {
    param([ValidateRange(1, 30)][int]$TimeoutSeconds)

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $remaining = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            return
        }

        Start-Sleep -Milliseconds 200
    }
    while ([DateTime]::UtcNow -lt $deadline)

    $processIds = @($remaining | ForEach-Object { $_.Id }) -join ', '
    throw "Excel did not exit naturally within $TimeoutSeconds seconds. Remaining PID(s): $processIds."
}

if ([string]::IsNullOrWhiteSpace($AddInPath)) {
    $AddInPath = Join-Path $PSScriptRoot '..\src\ExcelAccel.ExcelAddIn\bin\Debug\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
}

$resolvedAddInPath = (Resolve-Path -LiteralPath $AddInPath).Path
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repositoryRoot '.tools\reliability\phase0-soak-latest.json'
}

$existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($existingExcel.Count -ne 0) {
    throw 'Reliability soak requires Excel to be closed before it starts; no existing Excel process will be terminated.'
}

$samples = [System.Collections.Generic.List[object]]::new()
$soakWatch = [Diagnostics.Stopwatch]::StartNew()

for ($iteration = 1; $iteration -le $Iterations; $iteration++) {
    $iterationWatch = [Diagnostics.Stopwatch]::StartNew()
    $result = & (Join-Path $PSScriptRoot 'Test-ExcelAddIn.ps1') `
        -AddInPath $resolvedAddInPath `
        -TimeoutSeconds $TimeoutSeconds
    $iterationWatch.Stop()

    if (-not $result.Passed) {
        throw "Reliability soak iteration $iteration did not pass."
    }

    Wait-ForNoExcel -TimeoutSeconds 10

    Assert-AddInUnlocked -Path $resolvedAddInPath
    $samples.Add([ordered]@{
        iteration = $iteration
        duration_ms = [Math]::Round($iterationWatch.Elapsed.TotalMilliseconds, 4)
        working_set_bytes = Get-RequiredMetric -Evidence $result.Evidence -Name 'working_set_bytes'
        private_memory_bytes = Get-RequiredMetric -Evidence $result.Evidence -Name 'private_memory_bytes'
        handle_count = Get-RequiredMetric -Evidence $result.Evidence -Name 'handle_count'
        clean_process_exit = $true
        add_in_unlocked = $true
    })
}

$soakWatch.Stop()
$durationSamples = [double[]]@($samples | ForEach-Object { $_.duration_ms })
$workingSetSamples = [double[]]@($samples | ForEach-Object { $_.working_set_bytes })
$privateMemorySamples = [double[]]@($samples | ForEach-Object { $_.private_memory_bytes })
$handleSamples = [double[]]@($samples | ForEach-Object { $_.handle_count })

$report = [ordered]@{
    schema_version = 1
    work_package = 'PHASE-0-CLOSURE'
    generated_utc = [DateTime]::UtcNow.ToString('o')
    add_in = $resolvedAddInPath
    iterations = $Iterations
    total_duration_ms = [Math]::Round($soakWatch.Elapsed.TotalMilliseconds, 4)
    summary = [ordered]@{
        duration_p95_ms = Get-Percentile -Samples $durationSamples -Percentile 95
        working_set_p95_bytes = Get-Percentile -Samples $workingSetSamples -Percentile 95
        private_memory_p95_bytes = Get-Percentile -Samples $privateMemorySamples -Percentile 95
        handle_count_p95 = Get-Percentile -Samples $handleSamples -Percentile 95
        handle_count_range = [double](($handleSamples | Measure-Object -Maximum).Maximum - ($handleSamples | Measure-Object -Minimum).Minimum)
    }
    samples = @($samples)
    limitations = @(
        'Each iteration uses a fresh Excel process, so this proves process cleanup and cross-session stability rather than in-process long-duration retention.',
        'Resource values are observational distributions; Phase 1 budgets require explicit approval.'
    )
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
}
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8

[pscustomobject]@{
    Passed = $true
    Iterations = $Iterations
    Output = (Resolve-Path -LiteralPath $OutputPath).Path
    DurationP95Ms = $report.summary.duration_p95_ms
    WorkingSetP95Bytes = $report.summary.working_set_p95_bytes
    PrivateMemoryP95Bytes = $report.summary.private_memory_p95_bytes
    HandleCountP95 = $report.summary.handle_count_p95
}
