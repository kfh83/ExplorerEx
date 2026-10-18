#requires -Version 5.1
#requires -RunAsAdministrator
<#
Removes the standalone ExplorerEx loader: Windows\dxgi.dll (only if it is ours)
and Program Files\ExplorerEx. System32\dxgi.dll is never touched.
Files still loaded by a running Explorer are scheduled for deletion at reboot.
No running shell is stopped by this script; restart Explorer afterwards.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) {
    throw 'Run this script from 64-bit PowerShell.'
}
$destination = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'ExplorerEx'
$proxy = Join-Path ([Environment]::GetFolderPath('Windows')) 'dxgi.dll'

Add-Type -Namespace ExplorerEx -Name Native -MemberDefinition @'
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool MoveFileExW(string existing, string replacement, int flags);
'@
$pendingReboot = $false

function Remove-OrSchedule([string]$path) {
    try {
        Remove-Item -LiteralPath $path -Force
    } catch [UnauthorizedAccessException], [IO.IOException] {
        # MOVEFILE_DELAY_UNTIL_REBOOT: in-use DLLs cannot be deleted while loaded.
        if (-not [ExplorerEx.Native]::MoveFileExW($path, $null, 4)) {
            throw "Could not remove or schedule removal of ${path}: $($_.Exception.Message)"
        }
        $script:pendingReboot = $true
        Write-Warning "$path is in use; it will be deleted at the next reboot."
    }
}

if (Test-Path -LiteralPath $proxy) {
    if ([Diagnostics.FileVersionInfo]::GetVersionInfo($proxy).ProductName -ne 'ExplorerEx DXGI Loader') {
        throw "Refusing to remove $proxy; it belongs to another loader."
    }
    if ($PSCmdlet.ShouldProcess($proxy, 'Remove ExplorerEx DXGI proxy')) {
        Remove-OrSchedule $proxy
    }
}

if ((Test-Path -LiteralPath $destination) -and $PSCmdlet.ShouldProcess($destination, 'Remove ExplorerEx payload')) {
    # Deepest paths first so directories are empty before removal.
    $items = Get-ChildItem -LiteralPath $destination -Recurse -Force |
        Sort-Object { $_.FullName.Length } -Descending
    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            if (-not (Get-ChildItem -LiteralPath $item.FullName -Force)) { Remove-Item -LiteralPath $item.FullName -Force }
            elseif ([ExplorerEx.Native]::MoveFileExW($item.FullName, $null, 4)) { $pendingReboot = $true }
        } else {
            Remove-OrSchedule $item.FullName
        }
    }
    if (Get-ChildItem -LiteralPath $destination -Force) {
        # Leftover files are scheduled; schedule the directory after them.
        [void][ExplorerEx.Native]::MoveFileExW($destination, $null, 4)
        $pendingReboot = $true
    } else {
        Remove-Item -LiteralPath $destination -Force
    }
}

if ($pendingReboot) {
    Write-Host 'ExplorerEx uninstalled; some files are in use and will be removed at the next reboot.'
} else {
    Write-Host 'ExplorerEx uninstalled.'
}
Write-Host 'Restart Explorer to return to the native shell.'
