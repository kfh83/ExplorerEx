#requires -Version 5.1
#requires -RunAsAdministrator
<#
Installs the ReleaseDLL bundle beside this script. Run from 64-bit PowerShell.
Disable the old Windhawk ExplorerEx loader before restarting Explorer.
The proxy is Windows\dxgi.dll; System32\dxgi.dll is never replaced.
Uninstall with Uninstall-ExplorerEx.ps1.
No running shell is stopped by this script.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) {
    throw 'Run this script from 64-bit PowerShell.'
}
$destination = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'ExplorerEx'
$proxy = Join-Path ([Environment]::GetFolderPath('Windows')) 'dxgi.dll'

function Assert-OurProxy {
    if (Test-Path -LiteralPath $proxy) {
        $info = [Diagnostics.FileVersionInfo]::GetVersionInfo($proxy)
        if ($info.ProductName -ne 'ExplorerEx DXGI Loader') {
            throw "Refusing to replace/remove another DXGI proxy: $proxy. Remove the conflicting loader first."
        }
    }
}

Assert-OurProxy

$required = @('dxgi.dll', 'ExplorerEx.dll', 'ExplorerEx.Shell32.dll', 'ExplorerEx.Loader.exe', 'THIRD-PARTY-NOTICES.txt')
foreach ($name in $required) {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name) -PathType Leaf)) {
        throw "Missing $name. Run the script from a complete ReleaseDLL output/artifact directory."
    }
}
if ([Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $PSScriptRoot 'dxgi.dll')).ProductName -ne 'ExplorerEx DXGI Loader') {
    throw 'The supplied dxgi.dll is not the ExplorerEx proxy.'
}
if (-not $PSCmdlet.ShouldProcess("$destination and $proxy", 'Prepare symbols and install standalone ExplorerEx')) {
    return
}

$stage = Join-Path ([IO.Path]::GetTempPath()) ('ExplorerEx-install-' + [guid]::NewGuid().ToString('N'))
$backup = Join-Path $stage 'backup'
$changed = [Collections.Generic.List[object]]::new()
$rollbackFailed = $false
try {
    New-Item -ItemType Directory -Path $stage, $backup | Out-Null
    & (Join-Path $PSScriptRoot 'ExplorerEx.Loader.exe') prepare $stage
    if ($LASTEXITCODE -ne 0) {
        throw "Symbol preparation failed ($LASTEXITCODE); the installed shell was not changed."
    }
    $files = @()
    foreach ($name in $required | Where-Object { $_ -ne 'dxgi.dll' }) {
        $files += [pscustomobject]@{ Source = Join-Path $PSScriptRoot $name; Target = Join-Path $destination $name }
    }
    $files += [pscustomobject]@{ Source = Join-Path $stage 'explorer-hook.bin'; Target = Join-Path $destination 'explorer-hook.bin' }
    foreach ($folder in Get-ChildItem -LiteralPath $PSScriptRoot -Directory) {
        foreach ($mui in Get-ChildItem -LiteralPath $folder.FullName -Filter '*.mui' -File) {
            $files += [pscustomobject]@{ Source = $mui.FullName; Target = Join-Path (Join-Path $destination $folder.Name) $mui.Name }
        }
    }
    # Publish the proxy last, after the payload and verified symbol cache exist.
    $files += [pscustomobject]@{ Source = Join-Path $PSScriptRoot 'dxgi.dll'; Target = $proxy }
    foreach ($file in $files) {
        if ($file.Target -eq $proxy) { Assert-OurProxy }
        $saved = $null
        if (Test-Path -LiteralPath $file.Target) {
            $saved = Join-Path $backup ([guid]::NewGuid().ToString('N'))
            Copy-Item -LiteralPath $file.Target -Destination $saved
        }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $file.Target) | Out-Null
        $changed.Add([pscustomobject]@{ Target = $file.Target; Backup = $saved })
        Copy-Item -LiteralPath $file.Source -Destination $file.Target -Force
    }
    Write-Host "Installed $proxy and $destination."
    Write-Host 'Disable any Windhawk ExplorerEx loader, then restart Explorer. This script did not stop or restart it.'
    Write-Host 'After a Windows Explorer update, rerun this installer to refresh the exact-image symbol cache.'
} catch {
    $failure = $_
    for ($i = $changed.Count - 1; $i -ge 0; --$i) {
        $file = $changed[$i]
        try {
            if ($file.Backup) { Copy-Item -LiteralPath $file.Backup -Destination $file.Target -Force }
            elseif (Test-Path -LiteralPath $file.Target) { Remove-Item -LiteralPath $file.Target }
        } catch {
            $rollbackFailed = $true
            Write-Warning "Could not restore $($file.Target). Backup retained at $backup."
        }
    }
    throw $failure
} finally {
    if (-not $rollbackFailed) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
}
