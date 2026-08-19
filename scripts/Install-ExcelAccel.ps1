[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Install', 'Upgrade', 'Disable', 'Enable', 'Rollback', 'Uninstall')]
    [string]$Action = 'Install',
    [string]$PackageDirectory = '',
    [string]$InstallRoot = '',
    [switch]$AllowUntrustedPrototype
)

$ErrorActionPreference = 'Stop'
$productRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'ExcelAccel'))
if ([string]::IsNullOrWhiteSpace($InstallRoot)) { $InstallRoot = Join-Path $productRoot 'app' }
$installPath = [IO.Path]::GetFullPath($InstallRoot)
$allowedPrefix = $productRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not $installPath.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The per-user installer only owns a child directory of LocalAppData\ExcelAccel.'
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close every Excel process before install, update, rollback, disable, enable, or uninstall.'
}

$statePath = Join-Path $installPath 'install-state.json'
$ownerPath = Join-Path $installPath 'installer-owner.json'
$optionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$ownerToken = 'ExcelAccel.PerUserInstaller.v1'

function Read-State {
    if (-not (Test-Path -LiteralPath $statePath)) { return $null }
    $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    if ($state.schema_version -ne 1 -or $state.owner -ne $ownerToken) { throw 'The install state is not owned by this installer.' }
    if ([string]$state.registry_value_name -notmatch '^OPEN(?:[1-9][0-9]?)?$') { throw 'The owned Excel OPEN value name is invalid.' }
    if ([string]$state.active_version -notmatch '^[0-9A-Za-z][0-9A-Za-z.-]{0,63}$') { throw 'The active version in install state is invalid.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$state.previous_version) -and [string]$state.previous_version -notmatch '^[0-9A-Za-z][0-9A-Za-z.-]{0,63}$') { throw 'The prior version in install state is invalid.' }
    $expectedXll = Join-Path $installPath "versions\$($state.active_version)\ExcelAccel-AddIn64-packed.xll"
    $expectedData = '/R "' + $expectedXll + '"'
    if ([string]$state.registry_value_data -ne $expectedData) { throw 'The registry value recorded in install state is outside the owned active version.' }
    return $state
}

function Write-State {
    param([object]$State)
    [void](New-Item -ItemType Directory -Path $installPath -Force)
    $temporary = Join-Path $installPath ('.install-state.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $temporary -Encoding UTF8
    if (Test-Path -LiteralPath $statePath) {
        [IO.File]::Replace($temporary, $statePath, $null, $true)
    }
    else { [IO.File]::Move($temporary, $statePath) }
}

function Get-OwnedRegistryValue {
    param([object]$State)
    if (-not (Test-Path -LiteralPath $optionsKey)) { return $null }
    $properties = Get-ItemProperty -LiteralPath $optionsKey
    return $properties.($State.registry_value_name)
}

function Assert-OwnedRegistryValue {
    param([object]$State, [switch]$AllowMissing)
    $actual = Get-OwnedRegistryValue -State $State
    if ($null -eq $actual -and $AllowMissing) { return }
    if ([string]$actual -ne [string]$State.registry_value_data) {
        throw "Excel startup value '$($State.registry_value_name)' is missing or no longer owned by ExcelAccel. No registry change was made."
    }
}

function Find-FreeOpenValueName {
    [void](New-Item -Path $optionsKey -Force)
    $properties = Get-ItemProperty -LiteralPath $optionsKey
    foreach ($index in 0..99) {
        $name = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }
        if ($null -eq $properties.$name) { return $name }
    }
    throw 'No free Excel OPEN startup value is available within the bounded search.'
}

function Set-OwnedRegistration {
    param([object]$State, [string]$XllPath)
    $valueData = '/R "' + $XllPath + '"'
    $existing = Get-OwnedRegistryValue -State $State
    if ($null -ne $existing -and [string]$existing -ne [string]$State.registry_value_data) {
        throw 'The owned registry slot changed externally; refusing to overwrite it.'
    }
    New-ItemProperty -LiteralPath $optionsKey -Name $State.registry_value_name -PropertyType String -Value $valueData -Force | Out-Null
    $State.registry_value_data = $valueData
    $State.disabled = $false
}

function Remove-OwnedRegistration {
    param([object]$State)
    Assert-OwnedRegistryValue -State $State -AllowMissing
    if ($null -ne (Get-OwnedRegistryValue -State $State)) {
        Remove-ItemProperty -LiteralPath $optionsKey -Name $State.registry_value_name
    }
    $State.disabled = $true
}

if ($Action -in @('Install', 'Upgrade')) {
    if ([string]::IsNullOrWhiteSpace($PackageDirectory)) { throw 'PackageDirectory is required for install or upgrade.' }
    $package = (Resolve-Path -LiteralPath $PackageDirectory).Path
    $verify = Join-Path $PSScriptRoot 'Test-ExcelAccelPackage.ps1'
    $verification = & $verify -PackageDirectory $package -RequireValidSignature:(-not $AllowUntrustedPrototype) -LoadInExcel
    $manifest = Get-Content -LiteralPath (Join-Path $package 'package-manifest.json') -Raw | ConvertFrom-Json
    $version = [string]$manifest.package_version
    if ($version -notmatch '^[0-9A-Za-z][0-9A-Za-z.-]{0,63}$') { throw 'The package version is unsafe.' }
    $versionPath = [IO.Path]::GetFullPath((Join-Path $installPath "versions\$version"))
    if (Test-Path -LiteralPath $versionPath) { throw 'The requested version is already installed; no overwrite is allowed.' }
    $existingState = Read-State
    if ($Action -eq 'Install' -and $null -ne $existingState) { throw 'An installation already exists; use Upgrade.' }
    if ($Action -eq 'Upgrade' -and $null -eq $existingState) { throw 'No installation exists; use Install.' }
    if ($PSCmdlet.ShouldProcess($installPath, "$Action ExcelAccel $version")) {
        [void](New-Item -ItemType Directory -Path (Join-Path $installPath 'versions') -Force)
        if (-not (Test-Path -LiteralPath $ownerPath)) {
            @{ schema_version = 1; owner = $ownerToken } | ConvertTo-Json | Set-Content -LiteralPath $ownerPath -Encoding UTF8
        }
        Copy-Item -LiteralPath $package -Destination $versionPath -Recurse
        $xll = Join-Path $versionPath 'ExcelAccel-AddIn64-packed.xll'
        $state = $existingState
        if ($null -eq $state) {
            $state = [pscustomobject]@{ schema_version = 1; owner = $ownerToken; active_version = $version; previous_version = $null; registry_value_name = (Find-FreeOpenValueName); registry_value_data = ''; disabled = $true }
        }
        else {
            Assert-OwnedRegistryValue -State $state -AllowMissing:$state.disabled
            $state.previous_version = $state.active_version
            $state.active_version = $version
        }
        Set-OwnedRegistration -State $state -XllPath $xll
        Write-State -State $state
        [pscustomobject]@{ Action = $Action; Version = $version; RegisteredValue = $state.registry_value_name; Signature = $verification.SignatureStatus }
    }
    return
}

$state = Read-State
if ($null -eq $state) { throw 'No owned ExcelAccel installation state exists.' }
if ($Action -eq 'Disable') {
    if ($PSCmdlet.ShouldProcess($optionsKey, 'Disable ExcelAccel startup registration')) { Remove-OwnedRegistration -State $state; Write-State -State $state }
}
elseif ($Action -eq 'Enable') {
    if (-not $state.disabled) { throw 'ExcelAccel is already enabled.' }
    Assert-OwnedRegistryValue -State $state -AllowMissing
    $xll = Join-Path $installPath "versions\$($state.active_version)\ExcelAccel-AddIn64-packed.xll"
    if (-not (Test-Path -LiteralPath $xll)) { throw 'The active version artifact is missing.' }
    if ($PSCmdlet.ShouldProcess($optionsKey, 'Enable ExcelAccel startup registration')) { Set-OwnedRegistration -State $state -XllPath $xll; Write-State -State $state }
}
elseif ($Action -eq 'Rollback') {
    if ([string]::IsNullOrWhiteSpace([string]$state.previous_version)) { throw 'No prior version is recorded.' }
    Assert-OwnedRegistryValue -State $state -AllowMissing:$state.disabled
    $prior = $state.previous_version
    $xll = Join-Path $installPath "versions\$prior\ExcelAccel-AddIn64-packed.xll"
    if (-not (Test-Path -LiteralPath $xll)) { throw 'The prior version artifact is missing.' }
    if ($PSCmdlet.ShouldProcess($optionsKey, "Roll back ExcelAccel to $prior")) {
        $state.previous_version = $state.active_version; $state.active_version = $prior
        Set-OwnedRegistration -State $state -XllPath $xll; Write-State -State $state
    }
}
elseif ($Action -eq 'Uninstall') {
    if (-not (Test-Path -LiteralPath $ownerPath)) { throw 'The installer owner marker is missing; refusing recursive removal.' }
    $owner = Get-Content -LiteralPath $ownerPath -Raw | ConvertFrom-Json
    if ($owner.owner -ne $ownerToken) { throw 'The installer owner marker is invalid.' }
    Assert-OwnedRegistryValue -State $state -AllowMissing:$state.disabled
    if ($PSCmdlet.ShouldProcess($installPath, 'Unregister and remove the owned ExcelAccel per-user installation')) {
        if (-not $state.disabled) { Remove-OwnedRegistration -State $state }
        Remove-Item -LiteralPath $installPath -Recurse -Force
    }
}
