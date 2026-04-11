#define UXCTRL_VERSION 0x0100

#include "w4warn.h"
/*
 *   Level 4 warnings to be turned on.
 *   Do not disable any more level 4 warnings.
 */
#pragma warning(disable:4189)    // local variable is initialized but not referenced
#pragma warning(disable:4245)    // conversion from 'const int' to 'UINT', signed/unsign
#pragma warning(disable:4701)    // local variable 'pszPic' may be used without having been initiali
#pragma warning(disable:4706)    // assignment within conditional expression
#pragma warning(disable:4328)    // indirection alignment of formal parameter 1(4) is greater than the actual argument alignment (1)

#define _BROWSEUI_          // See HACKS OF DEATH in sfthost.cpp
#include "pch.h"
#include "desktop2.h"
#include "shfusion.h"

EXTERN_C HWND v_hwndTray;
EXTERN_C HWND v_hwndStartPane;

#define REGSTR_PATH_STARTFAVS       TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartPage")
#define REGSTR_VAL_STARTFAVS        TEXT("Favorites")
#define REGSTR_VAL_STARTFAVCHANGES  TEXT("FavoritesChanges")

#define REGSTR_VAL_PROGLIST         TEXT("ProgramsCache")

// When we want to get a tick count for the starting time of some interval
// and ensure that it is not zero (because we use zero to mean "not started").
// If we actually get zero back, then change it to -1.  Don't change it
// to 1, or somebody who does GetTickCount() - dwStart will get a time of
// 49 days.

__inline DWORD NonzeroGetTickCount()
{
    DWORD dw = GetTickCount();
    return dw ? dw : -1;
}

#include "resource.h"
