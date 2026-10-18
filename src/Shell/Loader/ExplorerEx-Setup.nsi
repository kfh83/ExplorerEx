; ExplorerEx NSIS Installer
; Requires NSIS 3.x Unicode
; Assembles from build\bin\ReleaseDLL\x64 output

!include "LogicLib.nsh"
!include "x64.nsh"
!include "FileFunc.nsh"
!include "WinMessages.nsh"
!include "Sections.nsh"

;--------------------------------
; General

Unicode true
Name "ExplorerEx"
!ifdef OUTDIR
    OutFile "${OUTDIR}\ExplorerEx-Setup.exe"
!else
    OutFile "ExplorerEx-Setup.exe"
!endif
InstallDir "$PROGRAMFILES64\ExplorerEx"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
BrandingText "ExplorerEx"
LicenseData "${BUILDDIR}\THIRD-PARTY-NOTICES.txt"
ShowInstDetails show
ShowUnInstDetails show


;--------------------------------
; Pages

Page license
Page components
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles



;--------------------------------
; Macros

!macro StopExplorer
    ; Prevent Winlogon from auto-restarting explorer while we work
    DetailPrint "Disabling AutoRestartShell..."
    ReadRegDWORD $2 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "AutoRestartShell"
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "AutoRestartShell" 0
    DetailPrint "Stopping Explorer..."
    FindWindow $0 "Shell_TrayWnd"
    ${If} $0 != 0
        System::Call 'user32::GetWindowThreadProcessId(p $0, *i .r1) i'
        System::Call 'user32::PostThreadMessageW(i r1, i ${WM_QUIT}, p 0, p 0)'
        Sleep 2000
    ${EndIf}
    nsExec::ExecToLog 'taskkill /F /IM explorer.exe'
    Sleep 1000
!macroend

!macro StartExplorer
    ; Restore AutoRestartShell before launching
    DetailPrint "Restoring AutoRestartShell..."
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "AutoRestartShell" $2
    DetailPrint "Starting Explorer..."
    Exec '"$WINDIR\explorer.exe"'
!macroend

!macro TakeOwnership path
    ; Take ownership from TrustedInstaller and grant Administrators full control
    nsExec::ExecToLog 'takeown /F "${path}" /A'
    nsExec::ExecToLog 'icacls "${path}" /grant Administrators:F'
!macroend

;--------------------------------
; Installer

Section "ExplorerEx Core" SecCore
    SectionIn RO

    ; Stop Explorer before replacing files
    !insertmacro StopExplorer

    ; Install payload to Program Files\ExplorerEx
    SetOutPath "$INSTDIR"
    File "${BUILDDIR}\ExplorerEx.dll"
    File "${BUILDDIR}\ExplorerEx.Shell32.dll"
    File "${BUILDDIR}\ExplorerEx.Loader.exe"
    File "${BUILDDIR}\THIRD-PARTY-NOTICES.txt"
    File "${BUILDDIR}\Uninstall-ExplorerEx.ps1"
    File "${BUILDDIR}\Install-ExplorerEx.ps1"
    File "${SRCDIR}\ExplorerEx.reg"

    ; English is always included
    SetOutPath "$INSTDIR\en-US"
    File "${BUILDDIR}\en-US\ExplorerEx.dll.mui"

    ; Prepare symbol cache for exact native Explorer
    DetailPrint "Resolving native Explorer symbols..."
    nsExec::ExecToLog '"$INSTDIR\ExplorerEx.Loader.exe" prepare "$INSTDIR"'
    Pop $0
    ${If} $0 != 0
        MessageBox MB_OK|MB_ICONSTOP "Symbol preparation failed (exit $0).$\n$\nThe installer cannot hook this Explorer build.$\nExplorerEx will NOT be active after installation."
    ${EndIf}

    ; Import registry entries
    DetailPrint "Importing registry settings..."
    nsExec::ExecToLog 'regedit /s "$INSTDIR\ExplorerEx.reg"'

    ; Install DXGI proxy beside explorer.exe with TrustedInstaller handling.
    ; The genuine System32\dxgi.dll is never replaced.
    ; The proxy is embedded in the installer and extracted to a temp directory first.
    SetOutPath "$PLUGINSDIR"
    File "${BUILDDIR}\dxgi.dll"
    ${If} ${FileExists} "$WINDIR\dxgi.dll"
        !insertmacro TakeOwnership "$WINDIR\dxgi.dll"
        Delete "$WINDIR\dxgi.dll"
    ${EndIf}
    CopyFiles /SILENT "$PLUGINSDIR\dxgi.dll" "$WINDIR\dxgi.dll"
    nsExec::ExecToLog 'icacls "$WINDIR\dxgi.dll" /setowner Administrators'

    ; Write uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"

    ; Add/Remove Programs entry
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "DisplayName" "ExplorerEx"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "UninstallString" '"$INSTDIR\Uninstall.exe"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "Publisher" "ExplorerEx Contributors"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "NoRepair" 1
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
    IntFmt $0 "0x%08X" $0
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "EstimatedSize" $0

    ; Restart Explorer
    !insertmacro StartExplorer
SectionEnd

;--------------------------------
; Translation sections

SectionGroup "Language Packs" SecLangs

Section "Español (es-ES)" SecES
    SetOutPath "$INSTDIR\es-ES"
    File "${BUILDDIR}\es-ES\ExplorerEx.dll.mui"
SectionEnd

Section "日本語 (ja-JP)" SecJA
    SetOutPath "$INSTDIR\ja-JP"
    File "${BUILDDIR}\ja-JP\ExplorerEx.dll.mui"
SectionEnd

Section "한국어 (ko-KR)" SecKO
    SetOutPath "$INSTDIR\ko-KR"
    File "${BUILDDIR}\ko-KR\ExplorerEx.dll.mui"
SectionEnd

Section "Polski (pl-PL)" SecPL
    SetOutPath "$INSTDIR\pl-PL"
    File "${BUILDDIR}\pl-PL\ExplorerEx.dll.mui"
SectionEnd

Section "Русский (ru-RU)" SecRU
    SetOutPath "$INSTDIR\ru-RU"
    File "${BUILDDIR}\ru-RU\ExplorerEx.dll.mui"
SectionEnd

SectionGroupEnd


;--------------------------------
; Uninstaller

Section "Uninstall"
    ; Stop Explorer before removing files
    !insertmacro StopExplorer

    ; Remove DXGI proxy with TrustedInstaller handling
    ${If} ${FileExists} "$WINDIR\dxgi.dll"
        !insertmacro TakeOwnership "$WINDIR\dxgi.dll"
        Delete "$WINDIR\dxgi.dll"
    ${EndIf}

    ; Remove payload directory
    RMDir /r "$INSTDIR\en-US"
    RMDir /r "$INSTDIR\es-ES"
    RMDir /r "$INSTDIR\ja-JP"
    RMDir /r "$INSTDIR\ko-KR"
    RMDir /r "$INSTDIR\pl-PL"
    RMDir /r "$INSTDIR\ru-RU"
    RMDir /r "$INSTDIR\symbols"
    Delete "$INSTDIR\ExplorerEx.dll"
    Delete "$INSTDIR\ExplorerEx.Shell32.dll"
    Delete "$INSTDIR\ExplorerEx.Loader.exe"
    Delete "$INSTDIR\ExplorerEx.reg"
    Delete "$INSTDIR\explorer-hook.bin"
    Delete "$INSTDIR\THIRD-PARTY-NOTICES.txt"
    Delete "$INSTDIR\Install-ExplorerEx.ps1"
    Delete "$INSTDIR\Uninstall-ExplorerEx.ps1"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"

    ; Remove Add/Remove Programs entry
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx"

    ; Restart Explorer
    !insertmacro StartExplorer
SectionEnd

