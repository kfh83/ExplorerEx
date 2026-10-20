; ExplorerEx NSIS Installer
; Requires NSIS 3.x Unicode
; Assembles from build\bin\ReleaseDLL\x64 output

; Must precede all code and data (including the Functions below).
Unicode true

!include "LogicLib.nsh"
!include "x64.nsh"
!include "FileFunc.nsh"
!include "WinMessages.nsh"
!include "Sections.nsh"

;--------------------------------
; General

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

; Saved Winlogon AutoRestartShell (a dedicated variable: $0-$9 are scratch).
Var AutoRestartShell

; This installer is a 32-bit NSIS program on 64-bit Windows. Without these two
; settings its HKLM writes land in WOW6432Node and `regedit`/`taskkill` resolve
; to 32-bit SysWOW64 copies, so Winlogon, the .reg import and the uninstall
; entry would all target the wrong registry view.
Function .onInit
    ${IfNot} ${RunningX64}
        MessageBox MB_OK|MB_ICONSTOP "ExplorerEx requires 64-bit Windows."
        Abort
    ${EndIf}
    SetRegView 64
    ${DisableX64FSRedirection}
FunctionEnd

Function un.onInit
    SetRegView 64
    ${DisableX64FSRedirection}
FunctionEnd


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
    ClearErrors
    ReadRegDWORD $AutoRestartShell HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "AutoRestartShell"
    ${If} ${Errors}
        StrCpy $AutoRestartShell 1 ; value absent: Windows default is enabled
    ${EndIf}
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
    WriteRegDWORD HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "AutoRestartShell" $AutoRestartShell
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
    File "${SRCDIR}\ExplorerEx.reg"

    ; English is always included
    SetOutPath "$INSTDIR\en-US"
    File "${BUILDDIR}\en-US\ExplorerEx.dll.mui"

    ; Prepare exact-image caches for native Explorer and optional menu hooks.
    DetailPrint "Resolving native shell and menu symbols..."
    nsExec::ExecToLog '"$INSTDIR\ExplorerEx.Loader.exe" prepare "$INSTDIR"'
    Pop $0
    ${If} $0 == 8
        ; Exit 8: ExplorerEx is ready; one or more optional menu hooks are unavailable.
        DetailPrint "Warning: some multicolumn/immersive menu hooks could not be prepared; continuing with available hooks."
    ${ElseIf} $0 != 0
        MessageBox MB_OK|MB_ICONSTOP "Symbol preparation failed (exit $0).$\n$\nThe installer cannot hook this Explorer build.$\nExplorerEx will NOT be active after installation."
    ${EndIf}

    DetailPrint "Importing registry settings..."
    nsExec::ExecToLog '"$WINDIR\regedit.exe" /s "$INSTDIR\ExplorerEx.reg"'
    Pop $0
    DetailPrint "regedit exit code: $0"
    ClearErrors
    ReadRegStr $1 HKLM "SOFTWARE\Classes\CLSID\{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}\InProcServer32" ""
    ${If} ${Errors}
        nsExec::ExecToLog '"$WINDIR\System32\reg.exe" import "$INSTDIR\ExplorerEx.reg" /reg:64'
        Pop $0
        DetailPrint "reg import exit code: $0"
        ClearErrors
        ReadRegStr $1 HKLM "SOFTWARE\Classes\CLSID\{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}\InProcServer32" ""
        ${If} ${Errors}
        ${EndIf}
    ${EndIf}

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

    ; Earlier installer builds wrote this key to the 32-bit view; migrate it.
    SetRegView 32
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx"
    SetRegView 64
    ; Add/Remove Programs entry
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "DisplayName" "ExplorerEx"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "UninstallString" '"$INSTDIR\Uninstall.exe"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "Publisher" "ExplorerEx Contributors"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx" "NoRepair" 1
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $3
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
    Delete "$INSTDIR\shell32-hook.bin"
    Delete "$INSTDIR\immersive-menu-hook.bin"
    Delete "$INSTDIR\THIRD-PARTY-NOTICES.txt"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"

    ; Remove Add/Remove Programs entry
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ExplorerEx"

    ; Restart Explorer
    !insertmacro StartExplorer
SectionEnd

