// ==WindhawkMod==
// @id              fix-multicolumn-menus-theme-background-infinite-expansion
// @name            Fix Multicolumn Shell Menus Infinite Expansion with Theme Background Margins
// @version         1.0
// @author          Isabella Lulamoon (kawapure)
// @github          https://github.com/kawapure
// @twitter         https://twitter.com/kawaipure
// @homepage        https://kawapure.github.io/
// @include         explorer.exe
// @compilerOptions -lcomdlg32
// @license         MIT
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# Fix Multicolumn Shell Menus Infinite Expansion with Theme Background Margins

As the title suggests, this mod fixes shell menus expanding infinitely when the theme background has margins
and the menu is multicolumn. This is a bug that was introduced sometime after such themes would not be used
in Windows anymore (my guesses without doing any research are either Windows Vista or Windows 10 version 1703).

This mod is useful for the start menu in ExplorerEx with the Luna theme.
*/
// ==/WindhawkModReadme==

#include <windhawk_utils.h>

void (__cdecl *CMenuToolbarBase__GetSize)(void *pThis, SIZE *pSize) = nullptr;
void (__cdecl *CMenuSFToolbar__GetSize_orig)(void *pThis, SIZE *pSize) = nullptr;
void __cdecl CMenuSFToolbar__GetSize_hook(void *pThis, SIZE *pSize)
{
    CMenuToolbarBase__GetSize(pThis, pSize);
}

// shell32.dll
WindhawkUtils::SYMBOL_HOOK c_rghookShell32[] = {
    {
        {
            L"public: virtual void __cdecl CMenuSFToolbar::GetSize(struct tagSIZE *)",
        },
        &CMenuSFToolbar__GetSize_orig,
        CMenuSFToolbar__GetSize_hook,
    },
    {
        {
            L"public: virtual void __cdecl CMenuToolbarBase::GetSize(struct tagSIZE *)",
        },
        &CMenuToolbarBase__GetSize,
    },
};

// The mod is being initialized, load settings, hook functions, and do other
// initialization stuff if required.
BOOL Wh_ModInit()
{
    Wh_Log(L"Init");

    HMODULE hmShell32 = LoadLibraryW(L"shell32");

    if (!WindhawkUtils::HookSymbols(hmShell32, c_rghookShell32, ARRAYSIZE(c_rghookShell32)))
    {
        Wh_Log(L"Failed to hook symbols in shell32.");
    }

    return TRUE;
}

// The mod is being unloaded, free all allocated resources.
void Wh_ModUninit()
{
    Wh_Log(L"Uninit");
}