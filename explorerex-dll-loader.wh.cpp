// ==WindhawkMod==
// @id              explorerex-dll-loader
// @name            ExplorerEx DLL Loader
// @description     Loads the DLL build of ExplorerEx
// @version         1.0.0
// @author          aubymori
// @github          https://github.com/aubymori
// @include         explorer.exe
// @architecture    x86-64
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# ExplorerEx DLL Loader
This mod allows you to use the DLL build of ExplorerEx, which allows you to use
UWP/Immersive apps.

# IMPORTANT: READ!
Windhawk needs to hook into `winlogon.exe` to successfully capture Explorer starting. Please
navigate to Windhawk's Settings, Advanced settings, More advanced settings, and make sure that
`winlogon.exe` is in the Process inclusion list.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- path: "ExplorerEx.dll"
  $name: DLL path
  $description: Path to the ExplorerEx DLL.
- console: true
  $name: Enable console
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

void *CreateDesktopAndTray_orig = nullptr;

void Wh_ModSettingsChanged(void)
{
    if (Wh_GetIntSetting(L"console"))
        AllocConsole();
    else
        FreeConsole();
}

BOOL Wh_ModInit(void)
{
    Wh_Log(L"init");
    Wh_ModSettingsChanged();
    WindhawkUtils::StringSetting path = WindhawkUtils::StringSetting::make(L"path");

    HMODULE hEpTaskbar = LoadLibraryW(path.get());
    if (!hEpTaskbar)
    {
        Wh_Log(L"Failed to load ExplorerEx DLL");
        return FALSE;
    }

    void *CreateDesktopAndTray = (void *)GetProcAddress(hEpTaskbar, "CreateDesktopAndTray");
    if (!CreateDesktopAndTray)
    {
        Wh_Log(L"Failed to load CreateDesktopAndTray from ExplorerEx DLL");
        return FALSE;
    }

    const WindhawkUtils::SYMBOL_HOOK hook = {
        {
            L"void * __cdecl CreateDesktopAndTray(void)"
        },
        &CreateDesktopAndTray_orig,
        CreateDesktopAndTray,
        false
    };

    if (!WindhawkUtils::HookSymbols(
        GetModuleHandleW(NULL),
        &hook, 1
    ))
    {
        Wh_Log(L"Failed to hook CreateDesktopAndTray in explorer.exe");
        return FALSE;
    }

    return TRUE;
}

void Wh_ModUninit(void)
{
    FreeConsole();
}