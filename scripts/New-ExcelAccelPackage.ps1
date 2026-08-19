[CmdletBinding()]
param(
    [ValidatePattern('^[0-9A-Za-z][0-9A-Za-z.-]{0,63}$')]
    [string]$Version = '0.0.0-phase0',
    [string]$InputXll = '',
    [string]$OutputRoot = '',
    [string]$CertificateThumbprint = '',
    [string]$SignToolPath = '',
    [string]$TimestampUrl = '',
    [switch]$RequireValidSignature
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path

if ([string]::IsNullOrWhiteSpace($InputXll)) {
    $InputXll = Join-Path $repositoryRoot 'src\ExcelAccel.ExcelAddIn\bin\Release\net48\publish\ExcelAccel.ExcelAddIn-AddIn64-packed.xll'
}
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $repositoryRoot '.tools\packages'
}

$resolvedInput = (Resolve-Path -LiteralPath $InputXll).Path
$resolvedOutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$packageDirectory = [IO.Path]::GetFullPath((Join-Path $resolvedOutputRoot "ExcelAccel-$Version-x64"))
$outputPrefix = $resolvedOutputRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $packageDirectory.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The package directory resolves outside the requested output root.'
}
if (Test-Path -LiteralPath $packageDirectory) {
    throw "Package output already exists and will not be overwritten: $packageDirectory"
}

[void](New-Item -ItemType Directory -Path $packageDirectory -Force)
$artifactName = 'ExcelAccel-AddIn64-packed.xll'
$artifactPath = Join-Path $packageDirectory $artifactName
Copy-Item -LiteralPath $resolvedInput -Destination $artifactPath

if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    if ([string]::IsNullOrWhiteSpace($SignToolPath)) {
        $candidate = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin' `
            -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object FullName -Match '\\x64\\signtool\.exe$' |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($null -eq $candidate) {
            throw 'SignTool was not found. Install a Windows SDK or provide -SignToolPath.'
        }
        $SignToolPath = $candidate.FullName
    }

    $signArguments = @(
        'sign',
        '/sha1', $CertificateThumbprint.Replace(' ', ''),
        '/s', 'My',
        '/fd', 'SHA256'
    )
    if (-not [string]::IsNullOrWhiteSpace($TimestampUrl)) {
        $signArguments += @('/tr', $TimestampUrl, '/td', 'SHA256')
    }
    $signArguments += @('/v', $artifactPath)

    & $SignToolPath @signArguments
    if ($LASTEXITCODE -ne 0) {
        throw "SignTool failed with exit code $LASTEXITCODE."
    }
}

$signature = Get-AuthenticodeSignature -LiteralPath $artifactPath
if ($RequireValidSignature -and $signature.Status -ne [Management.Automation.SignatureStatus]::Valid) {
    throw "The package artifact signature status is '$($signature.Status)', not Valid."
}

$artifact = Get-Item -LiteralPath $artifactPath
$hash = Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256
$manifest = [ordered]@{
    schema_version = 1
    product = 'ExcelAccel'
    package_version = $Version
    architecture = 'x64'
    runtime = 'net48'
    created_utc = [DateTime]::UtcNow.ToString('o')
    artifact = [ordered]@{
        relative_path = $artifactName
        length = [int64]$artifact.Length
        sha256 = $hash.Hash.ToUpperInvariant()
    }
    authenticode = [ordered]@{
        status = $signature.Status.ToString()
        status_message = [string]$signature.StatusMessage
        signer_subject = if ($null -ne $signature.SignerCertificate) { $signature.SignerCertificate.Subject } else { $null }
        signer_thumbprint = if ($null -ne $signature.SignerCertificate) { $signature.SignerCertificate.Thumbprint } else { $null }
        timestamp_subject = if ($null -ne $signature.TimeStamperCertificate) { $signature.TimeStamperCertificate.Subject } else { $null }
    }
}

$manifestPath = Join-Path $packageDirectory 'package-manifest.json'
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

[pscustomobject]@{
    PackageDirectory = $packageDirectory
    Manifest = $manifestPath
    Artifact = $artifactPath
    Sha256 = $hash.Hash.ToUpperInvariant()
    SignatureStatus = $signature.Status.ToString()
    SignerThumbprint = if ($null -ne $signature.SignerCertificate) { $signature.SignerCertificate.Thumbprint } else { $null }
}
