[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$VersionAPackage,
    [Parameter(Mandatory)]
    [string]$VersionBPackage,
    [string]$InstallRoot = '',
    [switch]$AllowUntrustedPrototype,
    [switch]$KeepSandbox
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot '.tools\install-sandbox'))
if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $allowedRoot 'ExcelAccel'
}

$resolvedInstallRoot = [IO.Path]::GetFullPath($InstallRoot)
$allowedPrefix = $allowedRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $resolvedInstallRoot.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The local lifecycle harness only permits install roots under .tools/install-sandbox/.'
}
if (Test-Path -LiteralPath $resolvedInstallRoot) {
    throw "The install sandbox already exists and will not be overwritten: $resolvedInstallRoot"
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close all Excel processes before running the package lifecycle harness.'
}

$verifyScript = Join-Path $PSScriptRoot 'Test-ExcelAccelPackage.ps1'
$packageA = (Resolve-Path -LiteralPath $VersionAPackage).Path
$packageB = (Resolve-Path -LiteralPath $VersionBPackage).Path
$requireValid = -not $AllowUntrustedPrototype
$createdRoot = $false
$evidence = [System.Collections.Generic.List[string]]::new()

function Test-Package {
    param([string]$Path, [switch]$Load)

    $parameters = @{
        PackageDirectory = $Path
        RequireValidSignature = $requireValid
        LoadInExcel = $Load
    }
    & $verifyScript @parameters
}

function Get-PackageVersion {
    param([string]$Path)

    $manifest = Get-Content -LiteralPath (Join-Path $Path 'package-manifest.json') -Raw | ConvertFrom-Json
    $version = [string]$manifest.package_version
    if ($version -notmatch '^[0-9A-Za-z][0-9A-Za-z.-]{0,63}$') {
        throw "Package version '$version' is unsafe for a version directory."
    }
    return $version
}

function Set-CurrentVersion {
    param([string]$Version)

    $pointer = [ordered]@{
        schema_version = 1
        product = 'ExcelAccel'
        active_version = $Version
        updated_utc = [DateTime]::UtcNow.ToString('o')
    }
    $temporaryPointer = Join-Path $resolvedInstallRoot 'current.json.tmp'
    $currentPointer = Join-Path $resolvedInstallRoot 'current.json'
    $pointer | ConvertTo-Json | Set-Content -LiteralPath $temporaryPointer -Encoding UTF8
    Move-Item -LiteralPath $temporaryPointer -Destination $currentPointer -Force
}

function Remove-OwnedPath {
    param([string]$Path, [switch]$Recurse)

    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse:$Recurse -Force
            return
        }
        catch [UnauthorizedAccessException], [IO.IOException] {
            if ($attempt -eq 10) {
                throw
            }
            Start-Sleep -Milliseconds 200
        }
    }
}

function Wait-ForExcelExit {
    for ($attempt = 1; $attempt -le 50; $attempt++) {
        if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0) {
            return
        }
        Start-Sleep -Milliseconds 200
    }
    throw 'The lifecycle rehearsal still observed Excel after the bounded shutdown window.'
}

try {
    $verificationA = Test-Package -Path $packageA
    $verificationB = Test-Package -Path $packageB
    $versionA = Get-PackageVersion -Path $packageA
    $versionB = Get-PackageVersion -Path $packageB
    if ($versionA -eq $versionB) {
        throw 'The lifecycle rehearsal requires two distinct package versions.'
    }

    [void](New-Item -ItemType Directory -Path (Join-Path $resolvedInstallRoot 'versions') -Force)
    $createdRoot = $true
    $marker = [ordered]@{
        schema_version = 1
        owner = 'ExcelAccel WP-P0-08 local lifecycle harness'
        created_utc = [DateTime]::UtcNow.ToString('o')
    }
    $marker | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $resolvedInstallRoot 'sandbox-owner.json') -Encoding UTF8

    $installedA = Join-Path $resolvedInstallRoot "versions\$versionA"
    Copy-Item -LiteralPath $packageA -Destination $installedA -Recurse
    [void](Test-Package -Path $installedA -Load)
    Set-CurrentVersion -Version $versionA
    $evidence.Add("installed_and_loaded=$versionA")

    $installedB = Join-Path $resolvedInstallRoot "versions\$versionB"
    Copy-Item -LiteralPath $packageB -Destination $installedB -Recurse
    [void](Test-Package -Path $installedB -Load)
    Set-CurrentVersion -Version $versionB
    $evidence.Add("upgraded_and_loaded=$versionB")

    Set-CurrentVersion -Version $versionA
    [void](Test-Package -Path $installedA -Load)
    $evidence.Add("rolled_back_and_loaded=$versionA")

    $currentPointer = Join-Path $resolvedInstallRoot 'current.json'
    $disabledPointer = Join-Path $resolvedInstallRoot 'disabled.json'
    Move-Item -LiteralPath $currentPointer -Destination $disabledPointer
    if (Test-Path -LiteralPath $currentPointer) {
        throw 'Disable failed because the active-version pointer still exists.'
    }
    $evidence.Add('disabled=true')

    Remove-OwnedPath -Path $disabledPointer
    Remove-OwnedPath -Path $installedB -Recurse
    Remove-OwnedPath -Path $installedA -Recurse
    Remove-OwnedPath -Path (Join-Path $resolvedInstallRoot 'versions')
    Remove-OwnedPath -Path (Join-Path $resolvedInstallRoot 'sandbox-owner.json')
    Remove-OwnedPath -Path $resolvedInstallRoot
    $createdRoot = $false
    $evidence.Add('uninstalled=true')

    Wait-ForExcelExit

    [pscustomobject]@{
        Passed = $true
        ProductionSignatureRequired = $requireValid
        VersionA = $versionA
        VersionB = $versionB
        SandboxRemoved = -not (Test-Path -LiteralPath $resolvedInstallRoot)
        Evidence = @($evidence)
    }
}
finally {
    if ($createdRoot -and -not $KeepSandbox) {
        $ownerMarker = Join-Path $resolvedInstallRoot 'sandbox-owner.json'
        if (Test-Path -LiteralPath $ownerMarker) {
            Remove-OwnedPath -Path $resolvedInstallRoot -Recurse
        }
    }
}
