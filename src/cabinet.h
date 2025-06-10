#ifndef _CABINET_H
#define _CABINET_H

#include "w4warn.h"
/*
 *   Level 4 warnings to be turned on.
 *   Do not disable any more level 4 warnings.
 */
#pragma warning(disable:4127)    // conditional expression is constant
#pragma warning(disable:4189)    // 'fIoctlSuccess' : local variable is initialized but not referenced
#pragma warning(disable:4201)    // nonstandard extension used : nameless struct/union
#pragma warning(disable:4245)    // conversion signed/unsigned mismatch
#pragma warning(disable:4509)    // nonstandard extension used: 'GetUserAssist' uses SEH and 'debug_crit' has destructor
#pragma warning(disable:4701)    // local variable 'hfontOld' may be used without having been initialized
#pragma warning(disable:4706)    // assignment within conditional expression
#pragma warning(disable:4328)    // indirection alignment of formal parameter 1(4) is greater than the actual argument alignment (1)


#define _WINMM_ // for DECLSPEC_IMPORT

#define OEMRESOURCE

#define OVERRIDE_SHLWAPI_PATH_FUNCTIONS     // see comment in shsemip.h

#ifdef WINNT
#include <nt.h>         // Some of the NT specific code calls Rtl functions
#include <ntrtl.h>      // which requires all of these header files...
#include <nturtl.h>
#endif


// if you include atlstuff.h, you don't get windowsx.h.  so we define needed functions here

#ifdef UNICODE
#define CP_WINNATURAL   CP_WINUNICODE
#else
#define CP_WINNATURAL   CP_WINANSI
#endif

#define DISALLOW_Assert

#include "pch.h"
#include "debug.h"          // our version of Assert etc.
#include "port32.h"
#include "dbt.h"
#include "trayp.h"
#include "ieguidp.h"
#include "shdguid.h"
#include "shundoc.h"

//
// Trace/dump/break flags specific to explorer.
//   (Standard flags defined in shellp.h)
//

// Trace flags
#define TF_DDE              0x00000100      // DDE traces
#define TF_TARGETFRAME      0x00000200      // Target frame
#define TF_TRAYDOCK         0x00000400      // Tray dock
#define TF_TRAY             0x00000800      // Tray 

// "Olde names"
#define DM_DDETRACE         TF_DDE
#define DM_TARGETFRAME      TF_TARGETFRAME
#define DM_TRAYDOCK         TF_TRAYDOCK

// Function trace flags
#define FTF_DDE             0x00000001      // DDE functions
#define FTF_TARGETFRAME     0x00000002      // Target frame methods

// Dump flags
#define DF_DDE              0x00000001      // DDE package
#define DF_DELAYLOADDLL     0x00000002      // Delay load

#ifdef __cplusplus
extern "C" {            /* Assume C declarations for C++ */
#endif  /* __cplusplus */

//---------------------------------------------------------------------------
// Globals
extern HINSTANCE hinstCabinet;  // Instance handle of the app.

extern HWND v_hwndDesktop;

extern HKEY g_hkeyExplorer;

//
// Is Mirroring APIs enabled (BiDi Memphis and NT5 only)
//
extern BOOL g_bMirroredOS;

// Global System metrics.  the desktop wnd proc will be responsible
// for watching wininichanges and keeping these up to date.

extern int g_fCleanBoot;
extern BOOL g_fFakeShutdown;
extern int g_fDragFullWindows;
extern int g_cxEdge;
extern int g_cyEdge;
extern int g_cxPaddedBorder;
extern int g_cySize;
extern int g_cyTabSpace;
extern int g_cxTabSpace;
extern int g_cxBorder;
extern int g_cyBorder;
extern int g_cxPrimaryDisplay;
extern int g_cyPrimaryDisplay;
extern int g_cxDlgFrame;
extern int g_cyDlgFrame;
extern int g_cxFrame;
extern int g_cyFrame;
extern int g_cxMinimized;
extern int g_cxVScroll;
extern int g_cyHScroll;
extern BOOL g_fNoDesktop;
extern UINT g_uDoubleClick;


extern HWND v_hwndTray;
extern HWND v_hwndStartPane;
extern BOOL g_fDesktopRaised;

extern const WCHAR c_wzTaskbarTheme[];

// the order of these is IMPORTANT for move-tracking and profile stuff
// also for the STUCK_HORIZONTAL macro
#define STICK_FIRST     ABE_LEFT
#define STICK_LEFT      ABE_LEFT
#define STICK_TOP       ABE_TOP
#define STICK_RIGHT     ABE_RIGHT
#define STICK_BOTTOM    ABE_BOTTOM
#define STICK_LAST      ABE_BOTTOM
#define STICK_MAX       ABE_MAX
#define STUCK_HORIZONTAL(x)     (x & 0x1)

#if STUCK_HORIZONTAL(STICK_LEFT) || STUCK_HORIZONTAL(STICK_RIGHT) || \
   !STUCK_HORIZONTAL(STICK_TOP)  || !STUCK_HORIZONTAL(STICK_BOTTOM)
#error Invalid STICK_* constants
#endif

#define InRange(id, idFirst, idLast)      ((UINT)((id)-(idFirst)) <= (UINT)((idLast)-(idFirst)))
#define IsInRange                   InRange

#define IsValidSTUCKPLACE(stick) IsInRange(stick, STICK_FIRST, STICK_LAST)

// initcab.cpp
HKEY GetSessionKey(REGSAM samDesired);
void RunStartupApps();
void WriteCleanShutdown(DWORD dwValue);


//
// Debug helper functions
//

void InvokeURLDebugDlg(HWND hwnd);

void Cabinet_InitGlobalMetrics(WPARAM, LPTSTR);


#define REGSTR_PATH_ADVANCED        REGSTR_PATH_EXPLORER TEXT("\\Advanced")
#define REGSTR_PATH_SMADVANCED      REGSTR_PATH_EXPLORER TEXT("\\StartMenu")
#define REGSTR_PATH_RUN_POLICY      REGSTR_PATH_POLICIES TEXT("\\Explorer\\Run")
#define REGSTR_EXPLORER_ADVANCED    REGSTR_PATH_ADVANCED
#define REGSTR_POLICIES_EXPLORER    TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer")


#ifdef __cplusplus
};       /* End of extern "C" { */


#endif // __cplusplus

#define PERF_ENABLESETMARK
#ifdef PERF_ENABLESETMARK
void DoSetMark(LPCSTR pszMark, ULONG cbSz);
#define PERFSETMARK(text)   DoSetMark(text, sizeof(text))
#else
#define PERFSETMARK(text)
#endif  // PERF_ENABLESETMARK


#endif  // _CABINET_H
