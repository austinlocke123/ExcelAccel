[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$PackageDirectory,
    [ValidateRange(1, 512)]
    [int]$MaximumArtifactMegabytes = 100,
    [switch]$RequireValidSignature,
    [switch]$LoadInExcel,
    [ValidateRange(10, 120)]
    [int]$TimeoutSeconds = 40,
    [switch]$Worker,
    [string]$ArtifactPath = ''
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-LoadWorker {
    param([string]$ResolvedArtifactPath)

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelAccelPackageNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $cell = $null
    $quitReturned = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        [uint32]$excelProcessId = 0
        [void][ExcelAccelPackageNativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
        [Console]::WriteLine("excel_pid=$excelProcessId")
        [Console]::Out.Flush()

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $registered = $excel.RegisterXLL($ResolvedArtifactPath)
        if (-not $registered) {
            throw 'Excel returned false from RegisterXLL.'
        }

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $cell = $worksheet.Range('A1')
        $cell.Formula = '=EXCELACCEL.VERSION()'
        $excel.Calculate()
        $version = [string]$cell.Value2
        if ([string]::IsNullOrWhiteSpace($version) -or $version.StartsWith('#')) {
            throw "The packaged health function returned '$version'."
        }
        [Console]::WriteLine("registered=$registered")
        [Console]::WriteLine("version=$version")
        [Console]::Out.Flush()

        $workbook.Close($false)
        $workbook = $null
        $excel.Quit()
        $quitReturned = $true
        [Console]::WriteLine('workbook_closed=true')
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
        Release-ComObject $cell
        Release-ComObject $worksheet
        Release-ComObject $workbook
        Release-ComObject $excel
    }
}

if ($Worker) {
    Invoke-LoadWorker -ResolvedArtifactPath $ArtifactPath
    exit 0
}

$resolvedPackage = (Resolve-Path -LiteralPath $PackageDirectory).Path
$manifestPath = Join-Path $resolvedPackage 'package-manifest.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ([int]$manifest.schema_version -ne 1 -or [string]$manifest.product -ne 'ExcelAccel') {
    throw 'The package manifest schema or product is unsupported.'
}

$relativePath = [string]$manifest.artifact.relative_path
if ([string]::IsNullOrWhiteSpace($relativePath) -or [IO.Path]::IsPathRooted($relativePath)) {
    throw 'The package artifact path must be relative.'
}
$components = $relativePath.Replace([IO.Path]::AltDirectorySeparatorChar, [IO.Path]::DirectorySeparatorChar).Split([IO.Path]::DirectorySeparatorChar)
if (@($components | Where-Object { [string]::IsNullOrWhiteSpace($_) -or $_ -eq '.' -or $_ -eq '..' }).Count -ne 0) {
    throw 'The package artifact path contains an unsafe component.'
}

$resolvedArtifact = [IO.Path]::GetFullPath((Join-Path $resolvedPackage $relativePath))
$packagePrefix = $resolvedPackage.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $resolvedArtifact.StartsWith($packagePrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The package artifact resolves outside the package directory.'
}

$artifact = Get-Item -LiteralPath $resolvedArtifact
$maximumBytes = [int64]$MaximumArtifactMegabytes * 1024 * 1024
if ($artifact.Length -gt $maximumBytes -or [int64]$manifest.artifact.length -gt $maximumBytes) {
    throw 'The package artifact exceeds the configured size limit.'
}
if ($artifact.Length -ne [int64]$manifest.artifact.length) {
    throw 'The package artifact length does not match the manifest.'
}
$actualHash = (Get-FileHash -LiteralPath $resolvedArtifact -Algorithm SHA256).Hash.ToUpperInvariant()
if ($actualHash -ne ([string]$manifest.artifact.sha256).ToUpperInvariant()) {
    throw 'The package artifact SHA-256 does not match the manifest.'
}

$signature = Get-AuthenticodeSignature -LiteralPath $resolvedArtifact
if ($RequireValidSignature -and $signature.Status -ne [Management.Automation.SignatureStatus]::Valid) {
    throw "The package signature status is '$($signature.Status)', not Valid."
}
if ($null -ne $manifest.authenticode.signer_thumbprint) {
    if ($null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -ne [string]$manifest.authenticode.signer_thumbprint) {
        throw 'The embedded signer certificate does not match the manifest signer thumbprint.'
    }
}

$zoneStreamPresent = $false
try {
    $zoneStreamPresent = $null -ne (Get-Item -LiteralPath $resolvedArtifact -Stream Zone.Identifier -ErrorAction Stop)
}
catch {
    $zoneStreamPresent = $false
}

$loadEvidence = $null
if ($LoadInExcel) {
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
        throw 'Close all Excel processes before running the isolated package-load test.'
    }

    $runId = [Guid]::NewGuid().ToString('N')
    $outputFile = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-package-$runId.out"
    $errorFile = Join-Path ([IO.Path]::GetTempPath()) "excelaccel-package-$runId.err"
    try {
        $arguments = @(
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', "`"$PSCommandPath`"",
            '-Worker',
            '-PackageDirectory', "`"$resolvedPackage`"",
            '-ArtifactPath', "`"$resolvedArtifact`""
        )
        $process = Start-Process powershell.exe `
            -ArgumentList $arguments `
            -WindowStyle Hidden `
            -RedirectStandardOutput $outputFile `
            -RedirectStandardError $errorFile `
            -PassThru
        $completed = $process.WaitForExit($TimeoutSeconds * 1000)
        $output = if (Test-Path -LiteralPath $outputFile) { Get-Content -LiteralPath $outputFile -Raw } else { '' }
        $errors = if (Test-Path -LiteralPath $errorFile) { Get-Content -LiteralPath $errorFile -Raw } else { '' }
        if (-not $completed) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }

        $excelProcessId = [regex]::Match($output, '(?m)^excel_pid=(\d+)').Groups[1].Value
        if ($excelProcessId) {
            $excelProcess = Get-Process -Id ([int]$excelProcessId) -ErrorAction SilentlyContinue
            if ($excelProcess -and -not $excelProcess.WaitForExit(5000)) {
                Stop-Process -Id $excelProcess.Id -Force
                throw "Excel PID $excelProcessId did not exit cleanly."
            }
        }

        if (-not $completed) {
            throw "The package-load worker timed out. Output:`n$output`nErrors:`n$errors"
        }
        if (-not [string]::IsNullOrWhiteSpace($errors)) {
            throw "The package-load worker failed. Output:`n$output`nErrors:`n$errors"
        }
        foreach ($requiredLine in @('registered=True', 'version=', 'workbook_closed=true', 'quit_returned=true')) {
            if ($output -notmatch "(?m)^$([regex]::Escape($requiredLine))") {
                throw "Package-load evidence is missing '$requiredLine'. Output:`n$output"
            }
        }
        $loadEvidence = $output.Trim()
    }
    finally {
        foreach ($temporaryFile in @($outputFile, $errorFile)) {
            if (Test-Path -LiteralPath $temporaryFile) {
                Remove-Item -LiteralPath $temporaryFile -Force
            }
        }
    }
}

[pscustomobject]@{
    Passed = $true
    PackageVersion = [string]$manifest.package_version
    Artifact = $resolvedArtifact
    Length = [int64]$artifact.Length
    Sha256 = $actualHash
    SignatureStatus = $signature.Status.ToString()
    SignerThumbprint = if ($null -ne $signature.SignerCertificate) { $signature.SignerCertificate.Thumbprint } else { $null }
    MarkOfTheWeb = $zoneStreamPresent
    ExcelLoadEvidence = $loadEvidence
}
