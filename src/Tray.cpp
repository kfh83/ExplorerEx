#include "pch.h"
#include "cocreateinstancehook.h"
#include "shguidp.h"
#include "cabinet.h"
#include "dlg.h"
#include "trayclok.h"
#include "util.h"
#include "tray.h"

#if defined(FE_IME)
#include <immp.h>
#endif

#include "bandsite.h"
#include "startmnu.h"
#include "uemapp.h"

#define NO_NOTIFYSUBCLASSWNDPROC
#include "cwndproc.h"
#include "desktop2.h"
#include "mixer.h"
#include "port32.h"
#include "shundoc.h"
#include "vssym32.h"
#include "startids.h"
#include "debug.h"
#include <immersive/ImmersiveInit.h>

#import "C:\Windows\System32\HelpPaneProxy.dll" named_guids no_namespace exclude("GUID") exclude("IUnknown")

#define DM_FOCUS        0           // focus
#define DM_SHUTDOWN     TF_TRAY     // shutdown
#define DM_UEMTRACE     TF_TRAY     // timer service, other UEM stuff
#define DM_MISC         0           // miscellany

const GUID CLSID_MSUTBDeskBand = { 0x540d8a8b, 0x1c3f, 0x4e32, 0x81, 0x32, 0x53, 0x0f, 0x6a, 0x50, 0x20, 0x90 };

// From Desktop2\proglist.cpp
HRESULT AddMenuItemsCacheTask(IShellTaskScheduler* pSystemScheduler, BOOL fKeepCacheWhenFinished);

// import the WIN31 Compatibility HACKs from the shell32.dll
//STDAPI_(void) CheckWinIniForAssocs(void);
//
//// hooks to Shell32.dll
//STDAPI CheckDiskSpace();
//STDAPI CheckStagingArea();

// startmnu.cpp
void HandleFirstTime();

HWND v_hwndDesktop = NULL;
HWND v_hwndTray = NULL;
HWND v_hwndStartPane = NULL;

BOOL g_fDesktopRaised = FALSE;
BOOL g_fInSizeMove = FALSE;

UINT _uMsgEnableUserTrackedBalloonTips = 0;
UINT _uMsgShowOnlyQuickLaunchDeskBand = 0;

void ClearRecentDocumentsAndMRUStuff(BOOL fBroadcastChange);
void DoTaskBarProperties(HWND hwnd, DWORD dwFlags);

void ClassFactory_Start();
void ClassFactory_Stop();

void SetupMergedFolderKeys(LPCTSTR clsid);
void ComServer_Stop(LPCTSTR clsid);

//
// Settings UI entry point types.
//
typedef void (WINAPI* PTRAYPROPSHEETCALLBACK)(DWORD nStartPage);
typedef void (WINAPI* PSETTINGSUIENTRY)(PTRAYPROPSHEETCALLBACK);

// Shell perf automation
extern DWORD g_dwShellStartTime;
extern DWORD g_dwShellStopTime;
extern DWORD g_dwStopWatchMode;

CTray c_tray;

// from explorer\desktop2
STDAPI DesktopV2_Create(
    IMenuPopup** ppmp, IMenuBand** ppmb, void** ppvStartPane, IUnknown** ppunk, HWND hwnd);
STDAPI DesktopV2_Build(void* pvStartPane);

// dyna-res change for multi-config hot/warm-doc
void HandleDisplayChange(int x, int y, BOOL fCritical);
DWORD GetMinDisplayRes(void);

// timer IDs
#define IDT_AUTOHIDE            2
#define IDT_AUTOUNHIDE          3
#ifdef DELAYWININICHANGE
#define IDT_DELAYWININICHANGE   5
#endif
#define IDT_DESKTOP             6
#define IDT_PROGRAMS            IDM_PROGRAMS
#define IDT_RECENT              IDM_RECENT
#define IDT_REBUILDMENU         7
#define IDT_HANDLEDELAYBOOTSTUFF 8
#define IDT_REVERTPROGRAMS      9
#define IDT_REVERTRECENT        10
#define IDT_REVERTFAVORITES     11

#define IDT_STARTMENU           12

#define IDT_ENDUNHIDEONTRAYNOTIFY 13

#define IDT_SERVICE0            14
#define IDT_SERVICE1            15
#define IDT_SERVICELAST         IDT_SERVICE1
#define IDT_SAVESETTINGS        17
#define IDT_ENABLEUNDO          18
#define IDT_STARTUPFAILED       19
#define IDT_CHECKDISKSPACE      21
#define IDT_STARTBUTTONBALLOON  22
#define IDT_CHANGENOTIFY        23
#define IDT_COFREEUNUSED        24

#define FADEINDELAY             100
#define BALLOONTIPDELAY         10000 // default balloon time copied from traynot.cpp


// INSTRUMENTATION WARNING: If you change anything here, make sure to update instrument.c
// we need to start at 500 because we're now sharing the hotkey handler
// with shortcuts..  they use an index array so they need to be 0 based
// NOTE, this constant is also in desktop.cpp, so that we can forward hotkeys from the desktop for
// NOTE, app compatibility.
#define GHID_FIRST 500
enum
{
    GHID_RUN = GHID_FIRST,
    GHID_MINIMIZEALL,
    GHID_UNMINIMIZEALL,
    GHID_HELP,
    GHID_EXPLORER,
    GHID_FINDFILES,
    GHID_FINDCOMPUTER,
    GHID_TASKTAB,
    GHID_TASKSHIFTTAB,
    GHID_SYSPROPERTIES,
    GHID_DESKTOP,
    GHID_TRAYNOTIFY,
    GHID_MAX
};

const DWORD GlobalKeylist[] =
{
    MAKELONG(TEXT('R'), MOD_WIN),
    MAKELONG(TEXT('M'), MOD_WIN),
    MAKELONG(TEXT('M'), MOD_SHIFT | MOD_WIN),
    MAKELONG(VK_F1,MOD_WIN),
    MAKELONG(TEXT('E'),MOD_WIN),
    MAKELONG(TEXT('F'),MOD_WIN),
    MAKELONG(TEXT('F'), MOD_CONTROL | MOD_WIN),
    MAKELONG(VK_TAB, MOD_WIN),
    MAKELONG(VK_TAB, MOD_WIN | MOD_SHIFT),
    MAKELONG(VK_PAUSE,MOD_WIN),
    MAKELONG(TEXT('D'),MOD_WIN),
    MAKELONG(TEXT('B'),MOD_WIN),
};

CTray::CTray()
    : _fCanSizeMove(TRUE)
    , _fIsLogoff(FALSE)
    , _fIsDesktopConnected(TRUE)
    , _stb(this)
    , _dwWatermarkPolicy(0)
{
}

void CTray::EnableTooltips(BOOL bEnable)
{
    SendMessage(_hwndTrayTips, TTM_ACTIVATE, bEnable, NULL);
}

void CTray::PurgeRebuildRequests()
{
    MSG msg;
    while (PeekMessage(&msg, _hwnd, SBM_REBUILDMENU, SBM_REBUILDMENU, PM_REMOVE | PM_NOYIELD))
    {
        // Keep sucking them out
    }
}

BOOL CTray::ShouldUseSmallIcons()
{
    return _fSMSmallIcons;  //_uModalMode possibly
}

VOID CTray::HandleFullScreenApp2(HWND hwnd)
{
    return HandleFullScreenApp(hwnd);
}

EXTERN_C BOOL WINAPI Tray_StartPanelEnabled()
{
    SHELLSTATE  ss = { 0 };
    SHGetSetSettings(&ss, SSF_STARTPANELON, FALSE);
    return ss.fStartPanelOn;
}

void CTray::TellTaskBandWeWantToOpenThisShit()
{
    static UINT WM_ShellHook = RegisterWindowMessage(TEXT("SHELLHOOK"));
    SendMessage(_hwndTasks, WM_ShellHook, 7, 0);
}

BOOL CTray::_CreateClockWindow()
{
    _hwndNotify = _trayNotify.TrayNotifyCreate(_hwnd, IDC_CLOCK, g_hinstCabinet);
    SendMessage(_hwndNotify, TNM_UPDATEVERTICAL, 0, !STUCK_HORIZONTAL(_uStuckPlace));

    return BOOLFROMPTR(_hwndNotify);
}

BOOL CTray::_InitTrayClass()
{
    WNDCLASS wc = { 0 };

    wc.lpszClassName = TEXT("Shell_TrayWnd");
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = s_WndProc;
    wc.hInstance = g_hinstCabinet;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.cbWndExtra = sizeof(LONG_PTR);

    return RegisterClass(&wc);
}


HFONT CTray::_CreateStartFont(HWND hwndTray)
{
    HFONT hfontStart = NULL;

    HTHEME hthemeStart = OpenThemeData(hwndTray, L"Button");
    if (hthemeStart)
    {
        LOGFONT lf;
        HDC hdc = GetDC(hwndTray);
        if (SUCCEEDED(GetThemeFont(hthemeStart, hdc, BP_PUSHBUTTON, PBS_NORMAL, TMT_FONT, &lf)))
        {
            hfontStart = CreateFontIndirect(&lf);
        }
        ReleaseDC(hwndTray, hdc);
        CloseThemeData(hthemeStart);
    }

    // Fallback to classic font if we can't get the theme font somehow
    if (!hfontStart)
    {
        NONCLIENTMETRICS ncm;
        ncm.cbSize = sizeof(ncm);
        if (SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, FALSE))
        {
            WORD wLang = GetUserDefaultLangID();

            // Select normal weight font for chinese language.
            if (PRIMARYLANGID(wLang) == LANG_CHINESE &&
                ((SUBLANGID(wLang) == SUBLANG_CHINESE_TRADITIONAL) ||
                    (SUBLANGID(wLang) == SUBLANG_CHINESE_SIMPLIFIED)))
                ncm.lfCaptionFont.lfWeight = FW_NORMAL;
            else
                ncm.lfCaptionFont.lfWeight = FW_BOLD;

            hfontStart = CreateFontIndirect(&ncm.lfCaptionFont);
        }
    }

    return hfontStart;
}

// Set the stuck monitor for the tray window
void CTray::_SetStuckMonitor()
{
    // use STICK_LEFT because most of the multi-monitors systems are set up
    // side by side. use DEFAULTTONULL because we don't want to get the wrong one
    // use the center point to call again in case we failed the first time.
    _hmonStuck = MonitorFromRect(&_arStuckRects[STICK_LEFT],
        MONITOR_DEFAULTTONULL);
    if (!_hmonStuck)
    {
        POINT pt;
        pt.x = (_arStuckRects[STICK_LEFT].left + _arStuckRects[STICK_LEFT].right) / 2;
        pt.y = (_arStuckRects[STICK_LEFT].top + _arStuckRects[STICK_LEFT].bottom) / 2;
        _hmonStuck = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    }

    _hmonOld = _hmonStuck;
}

DWORD _GetDefaultTVSDFlags()
{
    DWORD dwFlags = TVSD_TOPMOST;

    // if we are on a remote hydra session and if there is no previous saved value,
    // do not display the clock.
    if (SHGetMachineInfo(GMI_TSCLIENT))
    {
        dwFlags |= TVSD_HIDECLOCK;
    }
    return dwFlags;
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
void CTray::_GetSaveStateAndInitRects()
{
    TVSDCOMPAT tvsd;
    RECT rcDisplay;
    DWORD dwTrayFlags;
    UINT uStick;
    SIZE size;

    // first fill in the defaults
    SetRect(&rcDisplay, 0, 0, g_cxPrimaryDisplay, g_cyPrimaryDisplay);

    // size gets defaults
    size.cx = _stb._sizeStart.cx + 2 * (g_cxDlgFrame + g_cxBorder);
    size.cy = _stb._sizeStart.cy + 2 * (g_cyDlgFrame + g_cyBorder);

    // sStuckWidths gets minimum
    _sStuckWidths.cx = 2 * (g_cxDlgFrame + g_cxBorder);
    _sStuckWidths.cy = _stb._sizeStart.cy + 2 * (g_cyDlgFrame + g_cyBorder);

    _uStuckPlace = STICK_BOTTOM;
    dwTrayFlags = _GetDefaultTVSDFlags();

    _uAutoHide = 0;

    // now try to load saved vaules

    // BUG : 231077
    // Since Tasbar properties don't roam from NT5 to NT4, (NT4 -> NT5 yes)
    // Allow roaming from NT4 to NT5 only for the first time the User logs
    // on to NT5, so that future changes to NT5 are not lost when the user
    // logs on to NT4 after customizing the taskbar properties on NT5.

    DWORD cbData1 = sizeof(tvsd);
    DWORD cbData2 = sizeof(tvsd);
    if (Reg_GetStruct(g_hkeyExplorer, TEXT("StuckRectsXP2"), TEXT("Settings"),
        &tvsd, &cbData1)
        ||
        Reg_GetStruct(g_hkeyExplorer, TEXT("StuckRectsXP"), TEXT("Settings"),
            &tvsd, &cbData2))
    {
        if (IS_CURRENT_TVSD(tvsd.t) && IsValidSTUCKPLACE(tvsd.t.uStuckPlace))
        {
            _GetDisplayRectFromRect(&rcDisplay, &tvsd.t.rcLastStuck,
                MONITOR_DEFAULTTONEAREST);

            size = tvsd.t.sStuckWidths;
            _uStuckPlace = tvsd.t.uStuckPlace;

            dwTrayFlags = tvsd.t.dwFlags;
        }
        else if (MAYBE_WIN95_TVSD(tvsd.w95) &&
            IsValidSTUCKPLACE(tvsd.w95.uStuckPlace))
        {
            _uStuckPlace = tvsd.w95.uStuckPlace;
            dwTrayFlags = tvsd.w95.dwFlags;
            if (tvsd.w95.uAutoHide & AH_ON)
                dwTrayFlags |= TVSD_AUTOHIDE;

            switch (_uStuckPlace)
            {
                case STICK_LEFT:
                    size.cx = tvsd.w95.dxLeft;
                    break;

                case STICK_RIGHT:
                    size.cx = tvsd.w95.dxRight;
                    break;

                case STICK_BOTTOM:
                    size.cy = tvsd.w95.dyBottom;
                    break;

                case STICK_TOP:
                    size.cy = tvsd.w95.dyTop;
                    break;
            }
        }
    }


    ASSERT(IsValidSTUCKPLACE(_uStuckPlace));

    //
    // use the size only if it is not bogus
    //
    if (_sStuckWidths.cx < size.cx)
        _sStuckWidths.cx = size.cx;

    if (_sStuckWidths.cy < size.cy)
        _sStuckWidths.cy = size.cy;

    //
    // set the tray flags
    //
    _fAlwaysOnTop = BOOLIFY(dwTrayFlags & TVSD_TOPMOST);
    _fSMSmallIcons = BOOLIFY(dwTrayFlags & TVSD_SMSMALLICONS);
    _fHideClock = SHRestricted(REST_HIDECLOCK) || BOOLIFY(dwTrayFlags & TVSD_HIDECLOCK);
	_fNoThumbnails = SHWindowsPolicy(POLID_TaskbarNoThumbnail) || BOOLIFY(dwTrayFlags & 0x80);

    _fHideSCA[SCA_VOLUME] = BOOLIFY(dwTrayFlags & TVSD_HIDESCAVOLUME);
    _fHideSCA[SCA_NETWORK] = BOOLIFY(dwTrayFlags & TVSD_HIDESCANETWORK);
    _fHideSCA[SCA_POWER] = BOOLIFY(dwTrayFlags & TVSD_HIDESCAPOWER);

    _uAutoHide = (dwTrayFlags & TVSD_AUTOHIDE) ? AH_ON | AH_HIDING : 0;
    _RefreshSettings();

    //
    // initialize stuck rects
    //
    for (uStick = STICK_LEFT; uStick <= STICK_BOTTOM; uStick++)
        _MakeStuckRect(&_arStuckRects[uStick], &rcDisplay, _sStuckWidths, uStick);

    _UpdateVertical(_uStuckPlace);
    // Determine which monitor the tray is on using its stuck rectangles
    _SetStuckMonitor();
}

IBandSite* BandSite_CreateView();
HRESULT BandSite_SaveView(IUnknown* pbs);
LRESULT BandSite_OnMarshallBS(WPARAM wParam, LPARAM lParam);

void CTray::_SaveTrayStuff(void)
{
    TVSD tvsd;

    tvsd.dwSize = sizeof(tvsd);
    tvsd.lSignature = TVSDSIG_CURRENT;

    // position
    CopyRect(&tvsd.rcLastStuck, &_arStuckRects[_uStuckPlace]);
    tvsd.sStuckWidths = _sStuckWidths;
    tvsd.uStuckPlace = _uStuckPlace;

    tvsd.dwFlags = 0;
    if (_fAlwaysOnTop)      tvsd.dwFlags |= TVSD_TOPMOST;
    if (_fSMSmallIcons)     tvsd.dwFlags |= TVSD_SMSMALLICONS;
    if (_fHideClock && !SHRestricted(REST_HIDECLOCK))        tvsd.dwFlags |= TVSD_HIDECLOCK;
    if (_fNoThumbnails)     tvsd.dwFlags |= 0x80;
    if (_uAutoHide & AH_ON) tvsd.dwFlags |= TVSD_AUTOHIDE;

    // Save in Stuck rects.
    Reg_SetStruct(g_hkeyExplorer, TEXT("StuckRectsXP2"), TEXT("Settings"), &tvsd, sizeof(tvsd));

    BandSite_SaveView(_ptbs);

    return;
}

//  Helper function for CDesktopHost so clicking twice on the Start Button
//  treats the second click as a dismiss rather than a redisplay.
//
//  The crazy state machine goes like this:
//
//  SBSM_NORMAL - normal state, nothing exciting
//
//  When user opens Start Pane, we become
//
//  SBSM_SPACTIVE - start pane is active
//
//  If user clicks Start Button while SBSM_SPACTIVE, then we become
//
//      SBSM_EATING - eat mouse clicks
//
//      Until we receive a WM_MOUSEFIRST/WM_MOUSELAST message, and then
//      we return to SBSM_NORMAL.
//
//  If user dismisses Start Pane, we go straight to SBSM_NORMAL.
//
//
//      We eat the mouse clicks so that the click that the user made
//      to "unclick" the start button doesn't cause it to get pushed down
//      again (and cause the Start Menu to reopen).
//
#define SBSM_NORMAL         0
#define SBSM_SPACTIVE       1
#define SBSM_EATING         2

void Tray_SetStartPaneActive(BOOL fActive)
{
    if (fActive)
    {   // Start Pane appearing
        c_tray._uStartButtonState = SBSM_SPACTIVE;
    }
    else if (c_tray._uStartButtonState != SBSM_EATING)
    {   // Start Pane dismissing, not eating messages -> return to normal
        c_tray._uStartButtonState = SBSM_NORMAL;
    }
}

EXTERN_C const WCHAR c_wzTaskbarTheme[] = L"Taskbar";
EXTERN_C const WCHAR c_wzTaskbarVertTheme[] = L"TaskbarVert";

// create the toolbar with the three buttons and align windows

HWND CTray::_CreateStartButton()
{
    _uLogoffUser = RegisterWindowMessage(TEXT("Logoff User"));

    HWND hwndStart = _stb.CreateStartButton(_hwnd);
    BOOL fFlip3dAttribute = TRUE;
    DwmSetWindowAttribute(hwndStart, DWMWA_FLIP3D_POLICY, &fFlip3dAttribute, sizeof(fFlip3dAttribute));
    wprintf(L"Returning hwndStart %p\n", hwndStart);
    return hwndStart;
}

void CTray::_GetWindowSizes(UINT uStuckPlace, PRECT prcClient, PRECT prcView, PRECT prcNotify)
{
#ifdef DEAD_CODE
    prcView->top = 0;
    prcView->left = 0;
    prcView->bottom = prcClient->bottom;
    prcView->right = prcClient->right;

    if (STUCK_HORIZONTAL(uStuckPlace))
    {
        DWORD_PTR dwNotifySize = SendMessage(_hwndNotify, WM_CALCMINSIZE, prcClient->right / 2, prcClient->bottom);
        prcNotify->top = 0;
        prcNotify->left = prcClient->right - LOWORD(dwNotifySize);
        prcNotify->bottom = HIWORD(dwNotifySize);
        prcNotify->right = prcClient->right;

        prcView->left = _stb._sizeStart.cx + g_cxFrame + 1;
        prcView->right = prcNotify->left;
    }
    else
    {
        DWORD_PTR dwNotifySize = SendMessage(_hwndNotify, WM_CALCMINSIZE, prcClient->right, prcClient->bottom / 2);
        prcNotify->top = prcClient->bottom - HIWORD(dwNotifySize);
        prcNotify->left = 0;
        prcNotify->bottom = prcClient->bottom;
        prcNotify->right = LOWORD(dwNotifySize);

        prcView->top = _stb._sizeStart.cy + g_cyTabSpace;
        prcView->bottom = prcNotify->top;
    }
#endif
    if ((uStuckPlace & 1) != 0)
    {
        DWORD_PTR dwNotifySize = SendMessage(this->_hwndNotify, WM_CALCMINSIZE, prcClient->right / 2, prcClient->bottom);
        prcNotify->top = 0;
        prcNotify->left = prcClient->right - LOWORD(dwNotifySize);
        prcNotify->bottom = HIWORD(dwNotifySize);
        prcNotify->right = prcClient->right;
        prcView->top = 0;
        prcView->left = _stb._sizeStart.cx + (_hTheme == nullptr ? g_cxFrame + 1 : 0);
        prcView->right = prcNotify->left;
        prcView->bottom = prcClient->bottom;
    }
    else
    {
        DWORD_PTR dwNotifySize = SendMessage(_hwndNotify, WM_CALCMINSIZE, prcClient->right, prcClient->bottom / 2);
        prcNotify->left = 0;
        prcNotify->top = prcClient->bottom - HIWORD(dwNotifySize);
        prcNotify->bottom = prcClient->bottom;
        prcNotify->right = LOWORD(dwNotifySize);
        prcView->left = 0;
        prcView->top = g_cyTabSpace + this->_stb._sizeStart.cy;
        prcView->right = prcClient->right;
        prcView->bottom = prcNotify->top;
    }
}

void CTray::_RestoreWindowPos()
{
    WINDOWPLACEMENT wp;

    //first restore the stuck postitions
    _GetSaveStateAndInitRects();

    wp.length = sizeof(wp);
    wp.showCmd = SW_HIDE;

    _uMoveStuckPlace = (UINT)-1;
    _GetDockedRect(&wp.rcNormalPosition, FALSE);

    SendMessage(_hwndNotify, TNM_TRAYHIDE, 0, _fHideClock);
    SetWindowPlacement(_hwnd, &wp);
}

// Get the display (monitor) rectangle from the given arbitrary point
HMONITOR CTray::_GetDisplayRectFromPoint(LPRECT prcDisplay, POINT pt, UINT uFlags)
{
    RECT rcEmpty = { 0 };
    HMONITOR hmon = MonitorFromPoint(pt, uFlags);
    if (hmon && prcDisplay)
        GetMonitorRect(hmon, prcDisplay);
    else if (prcDisplay)
        *prcDisplay = rcEmpty;

    return hmon;
}

// Get the display (monitor) rectangle from the given arbitrary rectangle
HMONITOR CTray::_GetDisplayRectFromRect(LPRECT prcDisplay, LPCRECT prcIn, UINT uFlags)
{
    RECT rcEmpty = { 0 };
    HMONITOR hmon = MonitorFromRect(prcIn, uFlags);
    if (hmon && prcDisplay)
        GetMonitorRect(hmon, prcDisplay);
    else if (prcDisplay)
        *prcDisplay = rcEmpty;

    return hmon;
}

// Get the display (monitor) rectangle where the taskbar is currently on,
// if that monitor is invalid, get the nearest one.
void CTray::_GetStuckDisplayRect(UINT uStuckPlace, LPRECT prcDisplay)
{
    ASSERT(prcDisplay);
    BOOL fValid = GetMonitorRect(_hmonStuck, prcDisplay);

    if (!fValid)
        _GetDisplayRectFromRect(prcDisplay, &_arStuckRects[uStuckPlace], MONITOR_DEFAULTTONEAREST);
}

// Snap a StuckRect to the edge of a containing rectangle
// fClip determines whether to clip the rectangle if it's off the display or move it onto the screen
void CTray::_MakeStuckRect(LPRECT prcStick, LPCRECT prcBound, SIZE size, UINT uStick)
{
    CopyRect(prcStick, prcBound);

    if (!_hTheme)
    {
        InflateRect(prcStick, g_cxEdge, g_cyEdge);
    }

    if (size.cx < 0) size.cx *= -1;
    if (size.cy < 0) size.cy *= -1;

    switch (uStick)
    {
        case STICK_LEFT:   prcStick->right = (prcStick->left + size.cx); break;
        case STICK_TOP:    prcStick->bottom = (prcStick->top + size.cy); break;
        case STICK_RIGHT:  prcStick->left = (prcStick->right - size.cx); break;
        case STICK_BOTTOM: prcStick->top = (prcStick->bottom - size.cy); break;
    }
}

// the screen size has changed, so the docked rectangles need to be
// adjusted to the new screen.
void CTray::_ResizeStuckRects(RECT* arStuckRects)
{
    RECT rcDisplay;
    _GetStuckDisplayRect(_uStuckPlace, &rcDisplay);
    for (UINT uStick = STICK_LEFT; uStick <= STICK_BOTTOM; uStick++)
    {
        _MakeStuckRect(&arStuckRects[uStick], &rcDisplay, _sStuckWidths, uStick);
    }
}


//***   CTray::InvisibleUnhide -- temporary 'invisible' un-autohide
// DESCRIPTION
//  various tray resize routines need the tray to be un-autohide'd for
// stuff to be calculated correctly.  so we un-autohide it (invisibly...)
// here.  note the WM_SETREDRAW to prevent flicker (nt5:182340).
//  note that this is kind of a hack -- ideally the tray code would do
// stuff correctly even if hidden.
//
void CTray::InvisibleUnhide(BOOL fShowWindow)
{
    if (fShowWindow == FALSE)
    {
        if (_cHided++ == 0)
        {
            _ClipWindow(TRUE);
            SendMessageW(_hwnd, WM_SETREDRAW, FALSE, 0);
            ShowWindow(_hwnd, SW_HIDE);
            Unhide();
        }
    }
    else
    {
        ASSERT(_cHided > 0); // 1378
        if (_cHided-- == 1)
        {
            _Hide();
            ShowWindow(_hwnd, SW_SHOWNA);
            SendMessageW(_hwnd, WM_SETREDRAW, TRUE, 0);
        }
    }
}

// EXEX-VISTA: Arguments change. Very much so unmodified from XP.
void CTray::VerifySize(BOOL fWinIni, BOOL fRoundUp /* = FALSE */)
{
    RECT rc;
    BOOL fHiding;

    fHiding = (_uAutoHide & AH_HIDING);
    if (fHiding)
    {
        // force it visible so various calculations will happen relative
        // to unhidden size/position.
        //
        // fixes (e.g.) ie5:154536, where dropping a large-icon ISFBand
        // onto hidden tray didn't do size negotiation.
        //
        InvisibleUnhide(FALSE);
    }

    rc = _arStuckRects[_uStuckPlace];
    _HandleSizing(0, NULL, _uStuckPlace, FALSE); // EXEX-VISTA TODO: CHANGE WHEN ARGUMENTS CHANGE.

    if (!EqualRect(&rc, &_arStuckRects[_uStuckPlace]))
    {
        if (fWinIni)
        {
            // if we're changing size or position, we need to be unhidden
            Unhide();
            SizeWindows();
        }
        rc = _arStuckRects[_uStuckPlace];

        if (EVAL((_uAutoHide & (AH_ON | AH_HIDING)) != (AH_ON | AH_HIDING)))
        {
            _fSelfSizing = TRUE;
            SetWindowPos(_hwnd, NULL,
                rc.left, rc.top,
                RECTWIDTH(rc), RECTHEIGHT(rc),
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);

            _fSelfSizing = FALSE;
        }

        _StuckTrayChange();
    }

    if (fWinIni)
        SizeWindows();

    if (fHiding)
    {
        InvisibleUnhide(TRUE);
    }
}

HWND CTray::_GetClockWindow(void)
{
    return (HWND)SendMessage(_hwndNotify, TNM_GETCLOCK, 0, 0L);
}

void CTray::_CreateTrayTips()
{
    _hwndTrayTips = CreateWindowEx(WS_EX_TRANSPARENT, TOOLTIPS_CLASS, NULL,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT,
        CW_USEDEFAULT, CW_USEDEFAULT,
        _hwnd, NULL, g_hinstCabinet,
        NULL);

    if (_hwndTrayTips)
    {
        // taskbar windows are themed under Taskbar subapp name
        SendMessage(_hwndTrayTips, TTM_SETWINDOWTHEME, 0, (LPARAM)c_wzTaskbarTheme);

        //SetWindowZorder(_hwndTrayTips, HWND_TOPMOST);

        TOOLINFO ti;

        ti.cbSize = sizeof(ti);
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS | TTF_EXCLUDETOOLAREA;
        ti.hwnd = _hwnd;
        ti.uId = (UINT_PTR)_stb._hwndStart;
        ti.lpszText = (LPTSTR)MAKEINTRESOURCE(IDS_STARTBUTTONTIP);
        ti.hinst = g_hinstCabinet;
        SendMessage(_hwndTrayTips, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);

        // EXEX - VISTA TODO(allison) : Remove this because vista uses a custom clock tooltip.
        HWND hwndClock = _GetClockWindow();
        if (hwndClock)
        {
            ti.uFlags = TTF_EXCLUDETOOLAREA;
            ti.uId = (UINT_PTR)hwndClock;
            ti.lpszText = LPSTR_TEXTCALLBACK;
            ti.rect.left = ti.rect.top = ti.rect.bottom = ti.rect.right = 0;
            SendMessage(_hwndTrayTips, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
        }
    }
}

DEFINE_GUID(POLID_NoRestartOnTimer, 0x58AE0986, 0x0358, 0x4A44, 0x80, 0x89, 0x9C, 0x8C, 0x17, 0x7A, 0xA3, 0xB7);

#define SHCNE_STAGINGAREANOTIFICATIONS (SHCNE_CREATE | SHCNE_MKDIR | SHCNE_UPDATEDIR | SHCNE_UPDATEITEM)

// EXEX-VISTA: Slightly modified. Revalidate later.
LRESULT CTray::_CreateWindows()
{
#ifdef DEAD_CODE
    if (_CreateStartButton() && _CreateClockWindow())
    {
        // Initialise the theme.
        _OnThemeChanged();

        //
        //  We need to set the tray position, before creating
        // the view window, because it will call back our
        // GetWindowRect member functions.
        //
        _RestoreWindowPos();

        _CreateTrayTips();

        SendMessage(_hwndNotify, TNM_HIDECLOCK, 0, _fHideClock);

        _ptbs = BandSite_CreateView();
        if (_ptbs)
        {
            IUnknown_GetWindow(_ptbs, &_hwndRebar);
            _SetRebarTheme();
            SetWindowStyle(_hwndRebar, RBS_BANDBORDERS, FALSE);
            SetWindowStyle(_hwndRebar, 0x10, 1);

            // No need to check the disk space thing for non-privileged users, this reduces activity in the TS case
            // and only admins can properly free disk space anyways.
            if (IsUserAnAdmin() && !SHRestricted(REST_NOLOWDISKSPACECHECKS))
            {
                SetTimer(_hwnd, IDT_CHECKDISKSPACE, 60 * 1000, NULL);   // 60 seconds poll
            }

            if (!SHRestricted(REST_NOCDBURNING))
            {
                PIDLIST_ABSOLUTE pidlStaging;
                if (SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_CDBURN_AREA | CSIDL_FLAG_CREATE, NULL, 0, &pidlStaging)))
                {
                    SHChangeNotifyEntry fsne;
                    fsne.fRecursive = FALSE;
                    fsne.pidl = pidlStaging;
                    _uNotify = SHChangeNotifyRegister(_hwnd, SHCNRF_NewDelivery | SHCNRF_ShellLevel | SHCNRF_InterruptLevel,
                        SHCNE_STAGINGAREANOTIFICATIONS, TM_CHANGENOTIFY, 1, &fsne);

                    // start off by checking the first time.
                    _CheckStagingAreaOnTimer();

                    ILFree(pidlStaging);
                }
            }
            return 1;
        }
    }

    return -1;
#endif
    // Bail out if we fail to create either the start button or clock window.
    if (!_CreateStartButton() || !_CreateClockWindow())
        return -1;

    // Setup the theme for the taskbar.
    _OnThemeChanged();

    _RestoreWindowPos();
    _CreateTrayTips();

    SendMessageW(_hwndNotify, TNM_HIDECLOCK, 0, _fHideClock);

    _ptbs = BandSite_CreateView();
    if (!_ptbs)
        return -1;

    IUnknown_GetWindow(_ptbs, &_hwndRebar);
    _SetRebarTheme();

    SetWindowStyle(_hwndRebar, RBS_BANDBORDERS, 0);
    SetWindowStyle(_hwndRebar, 0x10u, 1);

    if (!SHRestricted(REST_NOLOWDISKSPACECHECKS))
        SetTimer(_hwnd, 0x15u, 0xEA60u, nullptr);

    LPITEMIDLIST pidlStaging;
    if (!SHRestricted(REST_NOCDBURNING) && SHGetFolderLocation(nullptr, 0x803B, nullptr, 0, &pidlStaging) >= 0 && ILRemoveLastID(pidlStaging))
    {
        SHChangeNotifyEntry fsne;
        fsne.pidl = pidlStaging;
        fsne.fRecursive = TRUE;
        _uNotify = SHChangeNotifyRegister(_hwnd, 32771, 0x800300A, 0x551u, 1, &fsne);
        _CheckStagingAreaOnTimer();
        ILFree(pidlStaging);
    }

    if (!SHWindowsPolicy(POLID_NoRestartOnTimer))
    {
        DWORD dwValue;
        DWORD cbData = sizeof(dwValue);
        if (!SHRegGetValueW(HKEY_LOCAL_MACHINE, SZ_EXPLORER_REGKEY, L"RestartTimer", 16, nullptr, &dwValue, &cbData) && dwValue)
        {
            SetTimer(_hwnd, 0x1Au, 3600000 * dwValue, nullptr);
        }
    }
    return 1;
}

LRESULT CTray::_InitStartButtonEtc()
{
    _stb.InitBackgroundBitmap();

    UpdateWindow(_hwnd);

    _stb.BuildStartMenu();

    _RegisterDropTargets();

    if (_CheckAssociations())
        CheckWinIniForAssocs();

    SendNotifyMessage(HWND_BROADCAST,
        RegisterWindowMessage(TEXT("TaskbarCreated")), 0, 0);

    return 1;
}

void CTray::_AdjustMinimizedMetrics()
{
    MINIMIZEDMETRICS mm;

    mm.cbSize = sizeof(mm);
    SystemParametersInfo(SPI_GETMINIMIZEDMETRICS, sizeof(mm), &mm, FALSE);
    mm.iArrange |= ARW_HIDE;
    SystemParametersInfo(SPI_SETMINIMIZEDMETRICS, sizeof(mm), &mm, FALSE);
}

void CTray::_UpdateBandSiteStyle()
{
    if (_ptbs)
    {
        BANDSITEINFO bsi;
        bsi.dwMask = BSIM_STYLE;
        _ptbs->GetBandSiteInfo(&bsi);

        BOOL fCanMoveBands = _fCanSizeMove && !SHRestricted(REST_NOMOVINGBAND);

        DWORD dwStyleNew;
        if (fCanMoveBands)
        {
            dwStyleNew = (bsi.dwStyle & ~(BSIS_NOGRIPPER | BSIS_LOCKED)) | BSIS_AUTOGRIPPER
                | BSIS_PREFERNOLINEBREAK;
        }
        else
        {
            dwStyleNew = (bsi.dwStyle & ~BSIS_AUTOGRIPPER) | BSIS_NOGRIPPER | BSIS_LOCKED
                | BSIS_PREFERNOLINEBREAK;
        }

        // only bother with refresh if something's changed
        if (bsi.dwStyle ^ dwStyleNew)
        {
            bsi.dwStyle = dwStyleNew;
            _ptbs->SetBandSiteInfo(&bsi);
            IUnknown_Exec(_ptbs, &CGID_DeskBand, DBID_BANDINFOCHANGED, 0, NULL, NULL);
        }
    }
}

BOOL _IsSizeMoveRestricted()
{
    return SHRegGetBoolUSValue(REGSTR_POLICIES_EXPLORER, TEXT("LockTaskbar"), FALSE, FALSE);
}

BOOL _IsSizeMoveEnabled()
{
    BOOL fCanSizeMove;

    if (_IsSizeMoveRestricted())
    {
        fCanSizeMove = FALSE;
    }
    else
    {
        fCanSizeMove = SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("TaskbarSizeMove"), FALSE, TRUE);
    }

    return fCanSizeMove;
}

// EXEX-VISTA: Validated.
void CTray::_RefreshSettings()
{
    _fCanSizeMove = _IsSizeMoveEnabled();
}

LRESULT CTray::_OnCreateAsync()
{
    if (g_dwProfileCAP & 0x00000004)
    {
        StartCAP();
    }

    if (g_dwProfileCAP & 0x00000004)
    {
        StopCAP();
    }

    _hMainAccel = LoadAccelerators(g_hinstCabinet, MAKEINTRESOURCE(ACCEL_TRAY));

    _RegisterGlobalHotkeys();

    // We spin a thread that will process "Load=", "Run=", CU\Run, and CU\RunOnce
    RunStartupApps();

    // If there were any startup failures that occurred before we were
    // ready to handle them, re-raise the failure now that we're ready.
    if (_fEarlyStartupFailure)
        LogFailedStartupApp();

    // we run the tray thread that handles Ctrl-Esc with a high priority
    // class so that it can respond even on a stressed system.
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_ABOVE_NORMAL);

    return 1;
}

// d:\\longhorn\\shell\\explorer\\logo.cpp

int g_bmpWidth = 0;
int g_bmpHeight = 0;
HWND g_hwndWatermark = NULL;

BOOL WINAPI DrawLogoBitmap(HDC hdc, HBITMAP hbm)
{
    BOOL fRet = FALSE;
    HDC hdcMem = CreateCompatibleDC(hdc);
    if (hdcMem)
    {
        HBITMAP hbmMem = (HBITMAP)SelectObject(hdcMem, hbm);
        fRet = BitBlt(hdc, 0, 0, g_bmpWidth, g_bmpHeight, hdcMem, 0, 0, SRCCOPY);
        SelectObject(hdcMem, hbmMem);
        DeleteDC(hdcMem);
    }
    return fRet;
}

void WINAPI GetWatermarkOffset(POINT *pPt, RECT *pClientRect)
{
    ASSERT(pPt != NULL); // 419
    ASSERT(pClientRect != NULL); // 420
    pPt->x = IsBiDiLocalizedSystem() ? (pClientRect->left + 10) : (pClientRect->right - g_bmpWidth - 10);
    pPt->y = pClientRect->bottom - g_bmpHeight - 10;
}

void __stdcall RePositionWatermark(HWND hWnd, BOOL bRepaint)
{
    RECT rc = { 0 };
    SystemParametersInfo(SPI_GETWORKAREA, 0, &rc, 0);
    POINT pt = { 0, 0 };
    GetWatermarkOffset(&pt, &rc);
    MoveWindow(hWnd, pt.x, pt.y, g_bmpWidth, g_bmpHeight, bRepaint);
}

void WINAPI DeleteWatermark(HINSTANCE hInstance)
{
    if (g_hwndWatermark)
    {
        DestroyWindow(g_hwndWatermark);
        UnregisterClass(TEXT("CSWatermark"), hInstance);
        g_hwndWatermark = NULL;
    }
}

LRESULT CALLBACK WatermarkWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (uMsg == 5)
    {
        if (wParam == 1)
        {
            ShowWindow(hWnd, 10);
            return 0;
        }
        return DefWindowProc(hWnd, 5u, wParam, lParam);
    }
    if (uMsg == 15)
    {
        COLORREF crPixel = 0;

        PAINTSTRUCT ps = { 0 };
        HDC hdc = BeginPaint(hWnd, &ps);
        if (hdc)
        {
            HBITMAP h = LoadBitmap(g_hinstCabinet, (LPCWSTR)0xB8);
            if (h)
            {
                HDC CompatibleDC = CreateCompatibleDC(hdc);
                HDC v10 = CompatibleDC;
                if (CompatibleDC)
                {
                    HGDIOBJ v12 = SelectObject(CompatibleDC, h);
                    if (!v12 || (crPixel = GetPixel(v10, 0, 0), crPixel != -1))
                    {
                        int DeviceCaps = GetDeviceCaps(hdc, BITSPIXEL);
                        SetLayeredWindowAttributes(hWnd, crPixel, 0xA5u, 2 * (DeviceCaps >= 16) + 1);
                    }
                    SelectObject(v10, v12);
                    DeleteDC(v10);
                }
                DrawLogoBitmap(hdc, h);
                DeleteObject(h);
            }
            EndPaint(hWnd, &ps);
        }
    }
    else
    {
        if (uMsg != 16)
        {
            switch (uMsg)
            {
                case WM_SETTINGCHANGE:
                    if (wParam == 47)
                        RePositionWatermark(hWnd, TRUE);
                    break;
                case WM_WINDOWPOSCHANGING:
                    if (lParam)
                    {
                        WINDOWPOS* pwp = (WINDOWPOS*)lParam;
                        UINT flags = pwp->flags;
                        if ((flags & 0x80u) != 0)
                            pwp->flags = flags & 0xFFFFFF7F;
                        int v7 = pwp->flags;
                        if ((v7 & 4) != 0)
                        {
                            pwp->hwndInsertAfter = (HWND)-1;
                            pwp->flags = v7 & 0xFFFFFFFB;
                        }
                    }
                    break;
                case WM_DISPLAYCHANGE:
                    RePositionWatermark(hWnd, FALSE);
                    InvalidateRect(hWnd, NULL, FALSE);
                    break;
                default:
                    return DefWindowProc(hWnd, uMsg, wParam, lParam);
            }
        }
    }
    return 0;
}

HWND WINAPI CreateWatermark(HINSTANCE hInstance)
{
    COLORREF crKey = CLR_INVALID;
    HWND hwnd = NULL;
    HDC hDC = NULL;

    if (g_hwndWatermark)
        DeleteWatermark(hInstance);

    WNDCLASSEX wcex;
    ZeroMemory(&wcex, sizeof(wcex));

    wcex.hInstance = hInstance;
    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WatermarkWndProc;
    wcex.lpszClassName = TEXT("CSWatermark");

    if (!RegisterClassEx(&wcex))
        return NULL;

    HBITMAP hbmWatermark = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_STARTER_WATERMARK));
    if (!hbmWatermark)
        return NULL;

    HWND hwndDesktop = GetDesktopWindow();

    BITMAP bm = { 0 };
    if (GetObject(hbmWatermark, sizeof(BITMAP), &bm))
    {
        g_bmpWidth = bm.bmWidth;
        g_bmpHeight = bm.bmHeight;
        hDC = GetDC(hwndDesktop);
        if (hDC)
        {
            HDC hdcMem = CreateCompatibleDC(hDC);
            ReleaseDC(hwndDesktop, hDC);
            if (hdcMem)
            {
                HBITMAP hbmMem = (HBITMAP)SelectObject(hdcMem, hbmWatermark);
                if (hbmMem)
                {
                    crKey = GetPixel(hdcMem, 0, 0);
                }
                SelectObject(hdcMem, hbmMem);
                DeleteDC(hdcMem);

                RECT rcWorkArea;
                if (crKey != CLR_INVALID && SystemParametersInfo(SPI_GETWORKAREA, 0, &rcWorkArea, 0))
                {
                    POINT pPt{ 0, 0 };
                    GetWatermarkOffset(&pPt, &rcWorkArea);
                    hwnd = SHFusionCreateWindowEx(
                        WS_EX_TOPMOST | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED,
                        TEXT("CSWatermark"),
                        NULL,
                        WS_POPUP,
                        pPt.x, pPt.y,
                        g_bmpWidth, g_bmpHeight,
                        NULL,
                        NULL,
                        hInstance,
                        NULL);
                }
            }
        }
    }

    DeleteObject(hbmWatermark);
    if (hwnd)
    {
        SetLayeredWindowAttributes(hwnd, crKey, 0xA5, (GetDeviceCaps(hDC, BITSPIXEL) >= 16) ? (LWA_COLORKEY | LWA_ALPHA) : LWA_COLORKEY);
        ShowWindow(hwnd, SW_SHOW);
        UpdateWindow(hwnd);
        g_hwndWatermark = hwnd;
    }

    return hwnd;
}

#include <slpublic.h>
#pragma comment(lib, "slc.lib")

void CTray::_StarterWatermarkCreate(BOOL fCreate)
{
    DWORD dwWatermarkOverride;
    DWORD cbData = sizeof(dwWatermarkOverride); // @MOD: Add override for showing the watermark
    if ((SUCCEEDED(SLGetWindowsInformationDWORD(L"explorer-StarterEditionWatermark", &_dwWatermarkPolicy)) && _dwWatermarkPolicy == 1)
        || SHRegGetUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("ShowStarterWatermark"), NULL, &dwWatermarkOverride, &cbData, FALSE, 0, 0) == ERROR_SUCCESS && dwWatermarkOverride == 1)
    {
        if (fCreate == 1)
            CreateWatermark(g_hinstCabinet);
        else
            DeleteWatermark(g_hinstCabinet);
    }
    else
    {
        _dwWatermarkPolicy = 0;
    }
}

// EXEX-VISTA: Validated.
LRESULT CTray::_OnCreate(HWND hwnd)
{
    LRESULT lres = -1;
    v_hwndTray = hwnd;

    // EXEX-VISTA TODO: Uncommented when implemented CSystemMixer
#if 0
    _pSystemMixer = new (std::nothrow) CSystemMixer(_hwnd);
#endif

    SendMessage(_hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_INITIALIZE, 0), 0);

    _AdjustMinimizedMetrics();

    SetWindowStyle(_hwnd, WS_BORDER | WS_THICKFRAME, !_hTheme);

    // Force Refresh of frame
    SetWindowPos(_hwnd, NULL, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_FRAMECHANGED);

    if (_HotkeyCreate())
    {
        lres = _CreateWindows();
    }

    _StarterWatermarkCreate(TRUE);

    BOOL fFlip3dPolicy = TRUE;
    DwmSetWindowAttribute(_hwnd, DWMWA_FLIP3D_POLICY, &fFlip3dPolicy, sizeof(fFlip3dPolicy));

    return lres;
}

typedef struct tagFSEPDATA
{
    LPRECT prc;
    HMONITOR hmon;
    CTray* ptray;
}
FSEPDATA, * PFSEPDATA;

BOOL WINAPI CTray::FullScreenEnumProc(HMONITOR hmon, HDC hdc, LPRECT prc, LPARAM dwData)
{
    BOOL fFullScreen;   // Is there a rude app on this monitor?

    PFSEPDATA pd = (PFSEPDATA)dwData;
    if (pd->hmon == hmon)
    {
        fFullScreen = TRUE;
    }
    else if (pd->prc)
    {
        RECT rc, rcMon;
        GetMonitorRect(hmon, &rcMon);
        IntersectRect(&rc, &rcMon, pd->prc);
        fFullScreen = EqualRect(&rc, &rcMon);
    }
    else
    {
        fFullScreen = FALSE;
    }

    if (hmon == pd->ptray->_hmonStuck)
    {
        pd->ptray->_fStuckRudeApp = fFullScreen;
    }

    //
    // Tell all the appbars on the same display to get out of the way too
    //
    pd->ptray->_AppBarNotifyAll(hmon, ABN_FULLSCREENAPP, NULL, fFullScreen);

    return TRUE;
}

void CTray::HandleFullScreenApp(HWND hwnd)
{
#ifdef DEAD_CODE
    //
    // First check to see if something has actually changed
    //
    _hwndRude = hwnd;

    //
    // Enumerate all the monitors, see if the app is rude on each, adjust
    // app bars and _fStuckRudeApp as necessary.  (Some rude apps, such
    // as the NT Logon Screen Saver, span multiple monitors.)
    //
    FSEPDATA d = { 0 };
    RECT rc;
    if (hwnd && GetWindowRect(hwnd, &rc))
    {
        d.prc = &rc;
        d.hmon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL);
    }
    d.ptray = this;

    EnumDisplayMonitors(NULL, NULL, FullScreenEnumProc, (LPARAM)&d);

    //
    // Now that we've set _fStuckRudeApp, update the tray's z-order position
    //
    _ResetZorder(0);

    //
    // stop the clock so we don't eat cycles and keep tons of code paged in
    //
    SendMessage(_hwndNotify, TNM_TRAYHIDE, 0, _fStuckRudeApp);

    //
    // Finally, let traynot know about whether the tray is hiding
    //
    SendMessage(_hwndNotify, TNM_RUDEAPP, _fStuckRudeApp, 0);
#else
    int fProcessingDesktopRaise; // eax
    struct tagRECT rc; // [esp+Ch] [ebp-24h] BYREF
    int v6; // [esp+28h] [ebp-8h]
    int fStuckRudeApp; // [esp+2Ch] [ebp-4h]

    if (g_fDesktopRaised || _fProcessingDesktopRaise)
        hwnd = 0;
    fStuckRudeApp = _fStuckRudeApp;

	FSEPDATA d = { 0 };
    v6 = 0;
    fProcessingDesktopRaise = _fProcessingDesktopRaise;
    _hwndRude = hwnd;
    v6 = fProcessingDesktopRaise;
    if (hwnd && GetWindowRect(hwnd, &rc))
    {
        d.prc = &rc;
        d.hmon = MonitorFromWindow(hwnd, 0);
    }
    d.ptray = this;

    EnumDisplayMonitors(0, 0, CTray::FullScreenEnumProc, (LPARAM)&d);
    if (fStuckRudeApp != _fStuckRudeApp)
    {
        CTray::_ResetZorder(0);
        SendMessageW(_hwndNotify, 0x403u, 0, _fStuckRudeApp);
        SendMessageW(_hwndNotify, 0x407u, _fStuckRudeApp, 0);
    }
#endif
}

void CTray::StartButtonClicked()
{
    SendMessage(_hwnd, WM_COMMAND, IDC_START, NULL);
}

// to not break functionality
void Tray_OnStartMenuDismissed()
{
    // called only from classic start menu, thus
    // the start button works propely, only while
    // classic start button is enabled

    printf("tray_onstartmenu dismissing\n");
    c_tray.OnStartMenuDismissed();
}

void CTray::OnStartMenuDismissed()
{
    ForceStartButtonUp();
    _bMainMenuInit = 0;
    PostMessageW(v_hwndTray, 0x590, TRUE, 0);
}

// EXEX-VISTA: Reimplemented.
int CTray::GetStartButtonMinHeight()
{
    if (_hwndTasks)
    {
        return SendMessageW(_hwndTasks, TBC_BUTTONHEIGHT, 0, 0);
    }

    return 0;
}

UINT CTray::GetStartMenuStuckPlace()
{
    return _uStuckPlace;
}

void CTray::SetUnhideTimer(LONG x, LONG y)
{
    _SetUnhideTimer(x, y);
}

void CTray::OnStartButtonClosing()
{
    _DoExitWindows(v_hwndDesktop, FALSE, FALSE);
}

BOOL CTray::_IsTopmost()
{
    return BOOLIFY(GetWindowLong(_hwnd, GWL_EXSTYLE) & WS_EX_TOPMOST);
}

BOOL CTray::_IsPopupMenuVisible()
{
#ifdef DEAD_CODE
    HWND hwnd;
    return ((SUCCEEDED(IUnknown_GetWindow(_pmpStartMenu, &hwnd)) && IsWindowVisible(hwnd)) ||
        (SUCCEEDED(IUnknown_GetWindow(_pmpStartPane, &hwnd)) && IsWindowVisible(hwnd)) ||
        (SUCCEEDED(IUnknown_GetWindow(_pmpTasks, &hwnd)) && IsWindowVisible(hwnd)));
#else
    HWND hwnd;
    return (_stb.IsPopupMenuVisible()) || (SUCCEEDED(IUnknown_GetWindow(_pmpTasks, &hwnd)) && IsWindowVisible(hwnd));
#endif
}

BOOL CTray::_IsActive()
{
    //
    // We say the tray is "active" iff:
    //
    // (a) the foreground window is the tray or a window owned by the tray, or
    // (b) the start menu is showing
    //

    BOOL fActive = FALSE;
    HWND hwnd = GetForegroundWindow();

    if (hwnd != NULL &&
        (hwnd == _hwnd || (GetWindowOwner(hwnd) == _hwnd)))
    {
        fActive = TRUE;
    }
    else if (_IsPopupMenuVisible())
    {
        fActive = TRUE;
    }

    return fActive;
}

BOOL CALLBACK CTray::s_EnumTooltipWindowsProc(HWND hwnd, LPARAM lParam)
{
    WCHAR szClass[MAX_PATH];

    if (hwnd != v_hwndTray
        && GetParent(hwnd) == v_hwndTray
        && GetClassName(hwnd, szClass, ARRAYSIZE(szClass))
        && !StrCmpIC(szClass, TOOLTIPS_CLASS))
    {
        SetWindowPos(hwnd, (HWND)lParam, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }
    return 1;
}

static void LogZ(const wchar_t* tag, HWND h) {
    DWORD ex = (DWORD)GetWindowLongPtr(h, GWL_EXSTYLE);
    HWND parent = GetParent(h);
    HWND owner = GetWindow(h, GW_OWNER);
    RECT rc; GetWindowRect(h, &rc);
    wprintf(L"[Z] %s hwnd=%p ex=%08x topmost=%d parent=%p owner=%p rect=(%d,%d %d,%d)\n",
        tag, h, ex, !!(ex & WS_EX_TOPMOST), parent, owner, rc.left, rc.top, rc.right, rc.bottom);
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
void CTray::_ResetZorder(int a2)
{
    LPARAM v3; // esi

    if (g_fDesktopRaised || _fProcessingDesktopRaise || _fAlwaysOnTop && !_fStuckRudeApp)
    {
        v3 = -1;
    }
    else if (_IsActive())
    {
        v3 = 0;
    }
    else
    {
        v3 = _fStuckRudeApp != 0 ? 1 : -2;
    }
    wprintf(L"[Z] _ResetZorder v3=%lld active=%d alwaysOnTop=%d rude=%d desktopRaised=%d a2=%d\n",
        (long long)v3, _IsActive(), _fAlwaysOnTop, _fStuckRudeApp, g_fDesktopRaised, a2);

    LogZ(L"tray:before", _hwnd);
    LogZ(L"start:before", _stb._hwndStart);

    if (v3 != (_IsTopmost() != 0) - 2 || a2)
    {
        SetWindowPos(_stb._hwndStart, (HWND)v3, 0, 0, 0, 0, 0x213u);
        SHForceWindowZorder(_hwnd, (HWND)v3);
        EnumWindows(CTray::s_EnumTooltipWindowsProc, v3);
    }

    LogZ(L"tray:after", _hwnd);
    LogZ(L"start:after", _stb._hwndStart);
}

void CTray::_MessageLoop()
{
#ifdef DEAD_CODE
    for (;;)
    {
        MSG  msg;
        if (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT)
            {
                if (_hwnd && IsWindow(_hwnd))
                {
                    // Tell the tray to save everything off if we got here
                    // without it being destroyed.
                    SendMessage(_hwnd, WM_ENDSESSION, 1, 0);
                }
                return;  // break all the way out of the main loop
            }

            if (_pmbTasks)
            {
                HRESULT hr = _pmbTasks->IsMenuMessage(&msg);
                if (hr == E_FAIL)
                {
                    if (_hwndTasks)
                        SendMessage(_hwndTasks, TBC_FREEPOPUPMENUS, 0, 0);
                }
                else if (hr == S_OK)
                {
                    continue;
                }
            }

            // Note that this needs to come before _pmbStartMenu since
            // the start pane sometimes hosts the start menu and it needs
            // to handle the start menu messages in that case.
            if (_pmbStartPane &&
                _pmbStartPane->IsMenuMessage(&msg) == S_OK)
            {
                continue;
            }

            if (_pmbStartMenu &&
                _pmbStartMenu->IsMenuMessage(&msg) == S_OK)
            {
                continue;
            }

            if (_hMainAccel && TranslateAccelerator(_hwnd, _hMainAccel, &msg))
            {
                continue;
            }

            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        else
        {
            WaitMessage();
        }
    }
#else
    HRESULT v2; // eax
    HWND hwndTasks; // eax
    HACCEL hMainAccel; // eax
    MSG msg; // [esp+Ch] [ebp-1Ch] BYREF

    while (1)
    {
        while (!PeekMessage(&msg, 0, 0, 0, 1u))
        {
            //ResponseMonitor::IndicateIdle(&this->field_640);
            WaitMessage();
        }
        if (msg.message == 0x12)
            break;
        //ResponseMonitor::UpdateLatency(&this->field_640, &msg);
        //if (msg.message == 0x100)
        //{
        //    if (msg.wParam == 0xD)
        //        SHTracePerf(&ShellTraceId_CTray_MessageLoop_Return_Info);
        //}
        //else if (msg.message == WM_LBUTTONDOWN || msg.message == WM_LBUTTONUP)
        //{
        //    if (WPP_GLOBAL_Control != &WPP_GLOBAL_Control && (*((_BYTE*)WPP_GLOBAL_Control + 28) & 8) != 0)
        //        WPP_SF_d(*((_QWORD*)WPP_GLOBAL_Control + 2), 0xAu, &MessageGuid, msg.message);
        //    SHTracePerfLoop(&ShellTraceId_CTray_MessageLoop_LButtonAction_Info, msg.message);
        //}

        if (this->_pmbTasks)
        {
            v2 = this->_pmbTasks->IsMenuMessage(&msg);
            if (v2 == 0x80004005)
            {
                hwndTasks = this->_hwndTasks;
                if (hwndTasks)
                    SendMessageW(hwndTasks, 0x438u, 0, 0);
                goto LABEL_17;
            }
            if (v2)
                goto LABEL_17;
        }
        else
        {
        LABEL_17:
            if (_stb.IsMenuMessage(&msg))
            {
                hMainAccel = this->_hMainAccel;
                if (!hMainAccel || !TranslateAcceleratorW(this->_hwnd, hMainAccel, &msg))
                {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }
        }
    }
    if (this->_hwnd && IsWindow(this->_hwnd))
        SendMessageW(this->_hwnd, WM_ENDSESSION, 1u, 0);
#endif
}

BOOL CTray::Init()
{
    return SHCreateThread(
        MainThreadProc, this, CTF_THREAD_REF | CTF_COINIT | CTF_REF_COUNTED, SyncThreadProc) && (_hwnd != nullptr);
}

// EXEX-VISTA: Validated.
int CTray::_GetPart(UINT uStuckPlace)
{
    switch (uStuckPlace)
    {
        case STICK_BOTTOM: return TBP_BACKGROUNDBOTTOM;
        case STICK_LEFT: return TBP_BACKGROUNDLEFT;
        case STICK_TOP: return TBP_BACKGROUNDTOP;
        case STICK_RIGHT: return TBP_BACKGROUNDRIGHT;
    }

    return 0;
}

void CTray::_UpdateVertical(UINT uStuckPlace, BOOL fForce)
{
    static UINT _uOldStuckPlace = STICK_MAX + 1;

    if ((_uOldStuckPlace != uStuckPlace) || fForce)
    {
        _uOldStuckPlace = uStuckPlace;

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_uv tray is now %s"), STUCK_HORIZONTAL(uStuckPlace) ? TEXT("HORIZONTAL") : TEXT("VERTICAL"));

        if (_ptbs)
        {
            // This following function will cause a WINDOWPOSCHANGING which will call DoneMoving
            // DoneMoving will then go a screw up all of the window sizing
            _fIgnoreDoneMoving = TRUE;
            BandSite_SetMode(_ptbs, STUCK_HORIZONTAL(uStuckPlace) ? 0 : DBIF_VIEWMODE_VERTICAL);
            _SetBandSiteTheme();
            _fIgnoreDoneMoving = FALSE;
        }

        SendMessage(_hwndNotify, TNM_UPDATEVERTICAL, 0, !STUCK_HORIZONTAL(uStuckPlace));

        if (_hTheme)
        {
            HDC hdc = GetDC(_hwnd);
            GetThemeMetric(_hTheme, hdc, 0, 0, TMT_HEIGHT, &_iSizingBarHeight);
            ReleaseDC(_hwnd, hdc);
            _trayNotify.field_334 = _iSizingBarHeight;
        }
    }
}

void CTray::_InitBandsite()
{
    ASSERT(_hwnd);

    // we initilize the contents after all the infrastructure is created and sized properly
    // need to notify which side we're on.
    // nt5:211881: set mode *before* load, o.w. Update->RBAutoSize messed up
    _UpdateBandSiteStyle();

    BandSite_Load();

    BandSite_FixUpComposition(_ptbs);

    // now that the mode is set, we need to force an update because we
    // explicitly avoided the update during BandSite_Load
    _UpdateVertical(_uStuckPlace, TRUE);
    BandSite_Update(_ptbs);
    BandSite_UIActivateDBC(_ptbs, DBC_SHOW);

    BandSite_FindBand(_ptbs, CLSID_TaskBand, IID_PPV_ARG(IDeskBand, &_pdbTasks), NULL, NULL);
    IUnknown_GetWindow(_pdbTasks, &_hwndTasks);

    SendMessage(_hwndTasks, 0x43F, 0, !_fNoThumbnails);

    // Now that bandsite is ready, set the correct size
    VerifySize(FALSE, TRUE);
    _AccountAllBandsForTaskbarSizingBar();
}

void CTray::_KickStartAutohide()
{
    if (_uAutoHide & AH_ON)
    {
        // tray always starts out hidden on autohide
        _uAutoHide = AH_ON | AH_HIDING;

        // we and many apps rely upon us having calculated the size correctly
        Unhide();

        // register it
        if (!_AppBarSetAutoHideBar2(_hwnd, TRUE, _uStuckPlace))
        {
            // don't bother putting up UI in this case
            // if someone is there just silently convert to normal
            // (the shell is booting who would be there anyway?)
            _SetAutoHideState(FALSE);
        }
    }
}

void CTray::_InitNonzeroGlobals()
{
    // initalize globals that need to be non-zero

    if (GetSystemMetrics(SM_SLOWMACHINE))
    {
        _dtSlideHide = 0;       // dont slide the tray out
        _dtSlideShow = 0;
    }
    else
    {
        _dtSlideHide = 400;
        _dtSlideShow = 200;
    }

    _RefreshSettings();
}

void CTray::_CreateTrayWindow()
{
    _InitTrayClass();

    _uMsgEnableUserTrackedBalloonTips = RegisterWindowMessage(ENABLE_BALLOONTIP_MESSAGE);
    _uMsgShowOnlyQuickLaunchDeskBand = RegisterWindowMessage(TEXT("ShowOnlyQuickLaunchDeskBand"));

    _fNoToolbarsOnTaskbarPolicyEnabled = (SHRestricted(REST_NOTOOLBARSONTASKBAR) != 0);
    _fTaskbarLockAllPolicyEnabled = (SHWindowsPolicy(POLID_TaskbarLockAll) != 0);
    _fTaskbarNoRedockPolicyEnabled = (SHWindowsPolicy(POLID_TaskbarNoRedock) != 0);
    _fTaskbarNoResizePolicyEnabled = (SHWindowsPolicy(POLID_TaskbarNoResize) != 0);

    DWORD dwExStyle = WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW;
    // Don't fadein because layered windows suck
    // If you create a layered window on a non-active desktop then the window goes black
    dwExStyle |= IS_BIDI_LOCALIZED_SYSTEM() ? dwExStyleRTLMirrorWnd : 0L;

    SHFusionCreateWindowEx(
        dwExStyle, TEXT("Shell_TrayWnd"), nullptr, WS_CLIPCHILDREN | WS_POPUP, 0, 0, 0, 0, nullptr, nullptr,
        g_hinstCabinet, this);

    // Fix for DWM borders on classic theme
    if (!IsAppThemed())
    {
        DWMNCRENDERINGPOLICY ncrp = DWMNCRP_DISABLED;
        DwmSetWindowAttribute(_hwnd, DWMWA_NCRENDERING_POLICY, &ncrp, sizeof(DWMNCRENDERINGPOLICY));
    }
}

DWORD WINAPI CTray::SyncThreadProc(void* pv)
{
    CTray* ptray = (CTray*)pv;
    return ptray->_SyncThreadProc();
}

DWORD CTray::_SyncThreadProc()
{
    //if (g_dwStopWatchMode)
    //    StopWatch_StartTimed(SWID_STARTUP, TEXT("_SyncThreadProc"), SPMODE_SHELL | SPMODE_DEBUGOUT, GetPerfTime());

    if (g_dwProfileCAP & 0x00000002)
        StartCAP();


    InitializeCriticalSection(&_csHotkey);

    OleInitialize(NULL);    // matched in MainThreadProc()
    ClassFactory_Start();

    SetupMergedFolderKeys(L"{865e5e76-ad83-4dca-a109-50dc2113ce9c}");
    SetupMergedFolderKeys(L"{865e5e76-ad83-4dca-a109-50dc2113ce9d}");
    SetupMergedFolderKeys(L"{865e5e76-ad83-4dca-a109-50dc2113ce9e}");

    _InitNonzeroGlobals();
    _ssomgr.Init();

    //
    //  Watch the registry key that tells us which app is the default
    //  web browser, so we can track it in
    //  HKLM\Software\Clients\StartMenuInternet.  We have to track it
    //  ourselves because downlevel browsers won't know about it.
    //
    //  We need to do this only if we have write access to the key.
    //  (If we don't have write access, then we can't change it,
    //  so there's no point watching for it to change...)
    //
    //  Well, okay, even if we only have read access, we have to do
    //  it once in case it changed while we were logged off.
    //
    //  The order of these operations is important...
    //
    //      1.  Migrate browser settings.
    //      2.  Build default MFU.  (Depends on browser settings.)
    //      3.  Create tray window.  (Relies on value MFU.)
    //

    _hHTTPEvent = CreateEvent(NULL, FALSE, TRUE, NULL);
    if (_hHTTPEvent)
    {
        //  Make one migration pass immediately so HandleFirstTime
        //  sees good information.  This also kick-starts the
        //  registry change notification process if the current user
        //  has write permission.
        _MigrateOldBrowserSettings();

        if (RegisterWaitForSingleObject(&_hHTTPWait, _hHTTPEvent,
            _MigrateOldBrowserSettingsCB, this,
            INFINITE, WT_EXECUTEDEFAULT))
        {
            // Yay, everything is fine.
        }
    }

    // Build the default MFU if necessary
    HandleFirstTime();

    _CreateTrayWindow();

    if (_hwnd && _ptbs)
    {
        // EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later. Partially reversed from Vista.

        _ResetZorder(0); // obey the "always on top" flag
        _KickStartAutohide();

        _InitBandsite();

        _ClipWindow(TRUE);  // make sure we clip the taskbar to the current monitor before showing it

        // @MOD - Skipped telemetry SHTracePerf(&ShellTraceId_Explorer_InitStartButton_Start)

        _InitStartButtonEtc();

        // @MOD - Skipped telemetry SHTracePerf(&ShellTraceId_Explorer_InitStartButton_Stop)

        BandSite_HandleDelayInitStuff(_ptbs);

        // it looks really strange for the tray to pop up and rehide at logon
        // if we are autohide don't activate the tray when we show it
        // if we aren't autohide do what Win95 did (tray is active by default)
        ShowWindow(_hwnd, (_uAutoHide & AH_HIDING) ? SW_SHOWNA : SW_SHOW);

        UpdateWindow(_hwnd);
        // _StuckTrayChange(); // seems removed in Vista

        ShowWindow(_stb._hwndStart, SW_SHOW);

        _stb.InitTheme();
        _stb.UpdateStartButton(true);

        //_StarterWatermarkUpdate();
        _bBool679 = true;   // ExplorerEx-Vista

        //_SignalShellDesktopSwitchEvent();

        SetTimer(_hwnd, IDT_HANDLEDELAYBOOTSTUFF, 5 * 1000, NULL);

        // get the system background scheduler thread
        if (!Tray_StartPanelEnabled())
        {
            IShellTaskScheduler *pScheduler;
            if (SUCCEEDED(CoCreateInstance(CLSID_SharedTaskScheduler, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pScheduler))))
            {
                AddMenuItemsCacheTask(pScheduler, FALSE);
                pScheduler->Release();
            }
        }
    }

    if (g_dwProfileCAP & 0x00020000)
        StopCAP();

    //if (g_dwStopWatchMode)
    //    StopWatch_StopTimed(SWID_STARTUP, TEXT("_SyncThreadProc"), SPMODE_SHELL | SPMODE_DEBUGOUT, GetPerfTime());

    return FALSE;
}

// the rest of the thread proc that includes the message loop
DWORD WINAPI CTray::MainThreadProc(void* pv)
{
    CTray* ptray = (CTray*)pv;

    if (!ptray->_hwnd)
        return FALSE;

    ptray->_OnCreateAsync();

    PERFSETMARK("ExplorerStartMenuReady");

    ptray->_MessageLoop();

    ClassFactory_Stop();
    OleUninitialize();      // matched in _SyncThreadProc()

    return FALSE;
}

#define DM_IANELHK 0

#define HKIF_NULL               0
#define HKIF_CACHED             1
#define HKIF_FREEPIDLS          2

typedef struct
{
    LPITEMIDLIST pidlFolder;
    LPITEMIDLIST pidlItem;
    WORD wGHotkey;
    WORD wFlags;
} HOTKEYITEM;


UINT CTray::_HotkeyGetFreeItemIndex(void)
{
    int i, cItems;
    HOTKEYITEM* phki;

    ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

    cItems = DSA_GetItemCount(_hdsaHKI);
    for (i = 0; i < cItems; i++)
    {
        phki = (HOTKEYITEM*)DSA_GetItemPtr(_hdsaHKI, i);
        if (!phki->wGHotkey)
        {
            ASSERT(!phki->pidlFolder);
            ASSERT(!phki->pidlItem);
            break;
        }
    }
    return i;
}

// Weird, Global hotkeys use different flags for modifiers than window hotkeys
// (and hotkeys returned by the hotkey control)
WORD _MapHotkeyToGlobalHotkey(WORD wHotkey)
{
    UINT nMod = 0;

    // Map the modifiers.
    if (HIBYTE(wHotkey) & HOTKEYF_SHIFT)
        nMod |= MOD_SHIFT;
    if (HIBYTE(wHotkey) & HOTKEYF_CONTROL)
        nMod |= MOD_CONTROL;
    if (HIBYTE(wHotkey) & HOTKEYF_ALT)
        nMod |= MOD_ALT;
    UINT nVirtKey = LOBYTE(wHotkey);
    return (WORD)((nMod * 256) + nVirtKey);
}

// NB This takes a regular window hotkey not a global hotkey (it does
// the convertion for you).
int CTray::HotkeyAdd(WORD wHotkey, LPCITEMIDLIST pidlFolder, LPCITEMIDLIST pidlItem, BOOL fClone)
{
    if (wHotkey)
    {
        LPCITEMIDLIST pidl1, pidl2;

        HOTKEYITEM hki;

        EnterCriticalSection(&_csHotkey);

        int i = _HotkeyGetFreeItemIndex();

        ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

        // DebugMsg(DM_IANELHK, "c.hl_a: Hotkey %x with id %d.", wHotkey, i);

        if (fClone)
        {
            pidl1 = ILClone(pidlFolder);
            pidl2 = ILClone(pidlItem);
            hki.wFlags = HKIF_FREEPIDLS;
        }
        else
        {
            pidl1 = pidlFolder;
            pidl2 = pidlItem;
            hki.wFlags = HKIF_NULL;
        }

        hki.pidlFolder = (LPITEMIDLIST)pidl1;
        hki.pidlItem = (LPITEMIDLIST)pidl2;
        hki.wGHotkey = _MapHotkeyToGlobalHotkey(wHotkey);
        DSA_SetItem(_hdsaHKI, i, &hki);

        LeaveCriticalSection(&_csHotkey);
        return i;
    }

    return -1;
}

// NB Cached hotkeys have their own pidls that need to be free but
// regular hotkeys just keep a pointer to pidls used by the startmenu and
// so don't.
int CTray::_HotkeyAddCached(WORD wGHotkey, LPITEMIDLIST pidl)
{
    int i = -1;

    if (wGHotkey)
    {
        LPITEMIDLIST pidlItem = ILClone(ILFindLastID(pidl));

        ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

        if (pidlItem)
        {
            if (ILRemoveLastID(pidl))
            {
                HOTKEYITEM hki;

                EnterCriticalSection(&_csHotkey);

                i = _HotkeyGetFreeItemIndex();

                // DebugMsg(DM_IANELHK, "c.hl_ac: Hotkey %x with id %d.", wGHotkey, i);

                hki.pidlFolder = pidl;
                hki.pidlItem = pidlItem;
                hki.wGHotkey = wGHotkey;
                hki.wFlags = HKIF_CACHED | HKIF_FREEPIDLS;
                DSA_SetItem(_hdsaHKI, i, &hki);

                LeaveCriticalSection(&_csHotkey);
            }
        }
    }

    return i;
}

// NB Again, this takes window hotkey not a Global one.
// NB This doesn't delete cached hotkeys.
int CTray::_HotkeyRemove(WORD wHotkey)
{
    int iRet = -1;
    if (EVAL(_hdsaHKI))
    {
        int i, cItems;
        HOTKEYITEM* phki;
        WORD wGHotkey;

        ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

        // DebugMsg(DM_IANELHK, "c.hl_r: Remove hotkey for %x" , wHotkey);

        // Unmap the modifiers.
        wGHotkey = _MapHotkeyToGlobalHotkey(wHotkey);

        EnterCriticalSection(&_csHotkey);

        cItems = DSA_GetItemCount(_hdsaHKI);
        for (i = 0; i < cItems; i++)
        {
            phki = (HOTKEYITEM*)DSA_GetItemPtr(_hdsaHKI, i);
            if (phki && !(phki->wFlags & HKIF_CACHED) && (phki->wGHotkey == wGHotkey))
            {
                // DebugMsg(DM_IANELHK, "c.hl_r: Invalidating %d", i);
                if (phki->wFlags & HKIF_FREEPIDLS)
                {
                    if (phki->pidlFolder)
                        ILFree(phki->pidlFolder);
                    if (phki->pidlItem)
                        ILFree(phki->pidlItem);
                }
                phki->wGHotkey = 0;
                phki->pidlFolder = NULL;
                phki->pidlItem = NULL;
                phki->wFlags &= ~HKIF_FREEPIDLS;
                iRet = i;
                break;
            }
        }
        LeaveCriticalSection(&_csHotkey);
    }

    return iRet;
}

// NB This takes a global hotkey.
int CTray::_HotkeyRemoveCached(WORD wGHotkey)
{
    int iRet = -1;
    int i, cItems;
    HOTKEYITEM* phki;

    ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

    // DebugMsg(DM_IANELHK, "c.hl_rc: Remove hotkey for %x" , wGHotkey);

    EnterCriticalSection(&_csHotkey);

    cItems = DSA_GetItemCount(_hdsaHKI);
    for (i = 0; i < cItems; i++)
    {
        phki = (HOTKEYITEM*)DSA_GetItemPtr(_hdsaHKI, i);
        if (phki && (phki->wFlags & HKIF_CACHED) && (phki->wGHotkey == wGHotkey))
        {
            // DebugMsg(DM_IANELHK, "c.hl_r: Invalidating %d", i);
            if (phki->wFlags & HKIF_FREEPIDLS)
            {
                if (phki->pidlFolder)
                    ILFree(phki->pidlFolder);
                if (phki->pidlItem)
                    ILFree(phki->pidlItem);
            }
            phki->pidlFolder = NULL;
            phki->pidlItem = NULL;
            phki->wGHotkey = 0;
            phki->wFlags &= ~(HKIF_CACHED | HKIF_FREEPIDLS);
            iRet = i;
            break;
        }
    }
    LeaveCriticalSection(&_csHotkey);

    return iRet;
}

// NB Some (the ones not marked HKIF_FREEPIDLS) of the items in the list of hotkeys
// have pointers to idlists used by the filemenu so they are only valid for
// the lifetime of the filemenu.
BOOL CTray::_HotkeyCreate(void)
{
    if (!_hdsaHKI)
    {
        // DebugMsg(DM_TRACE, "c.hkl_c: Creating global hotkey list.");
        _hdsaHKI = DSA_Create(sizeof(HOTKEYITEM), 0);
    }

    if (_hdsaHKI)
        return TRUE;

    return FALSE;
}

void CTray::_DestroyStartMenu()
{
    _stb.DestroyStartMenu();

    ATOMICRELEASET(_pmpTasks, IMenuPopup);
    ATOMICRELEASET(_pmbTasks, IMenuBand);
}

void CTray::ForceStartButtonUp()
{
    _stb.ForceButtonUp();

    if (_hwndTasks)
        SendMessage(_hwndTasks, TBC_SETPREVFOCUS, 0, NULL);

    PostMessage(_hwnd, TM_STARTMENUDISMISSED, 0, 0);
}

/*------------------------------------------------------------------
** Respond to a button's pressing by bringing up the appropriate menu.
** Clean up the button depression when the menu is dismissed.
**------------------------------------------------------------------*/
void CTray::_ToolbarMenu()
{
    // EXEX-VISTA: Completely reversed from Vista.

    if (_hwndTasks)
        SendMessage(_hwndTasks, TBC_FREEPOPUPMENUS, 0, 0);

    // Close any Context Menus
    SendMessage(_hwnd, WM_CANCELMODE, 0, 0);

    SetActiveWindow(_hwnd);
    _bMainMenuInit = TRUE;

    SendMessage(_hwndTasks, /* undocumented new TBC_ */ 0x43F, 1, 0);

    _stb.DisplayStartMenu();
}


HRESULT CTray::_AppBarSetState(UINT uFlags)
{
    if (uFlags & ~(ABS_AUTOHIDE | ABS_ALWAYSONTOP))
    {
        return E_INVALIDARG;
    }
    else
    {
        _SetAutoHideState(uFlags & ABS_AUTOHIDE);
        _UpdateAlwaysOnTop(uFlags & ABS_ALWAYSONTOP);
        return S_OK;
    }
}

//
// can't use SubtractRect sometimes because of inclusion limitations
//
void CTray::_AppBarSubtractRect(PAPPBAR pab, LPRECT lprc)
{
    switch (pab->uEdge) {
        case ABE_TOP:
            if (pab->rc.bottom > lprc->top)
                lprc->top = pab->rc.bottom;
            break;

        case ABE_LEFT:
            if (pab->rc.right > lprc->left)
                lprc->left = pab->rc.right;
            break;

        case ABE_BOTTOM:
            if (pab->rc.top < lprc->bottom)
                lprc->bottom = pab->rc.top;
            break;

        case ABE_RIGHT:
            if (pab->rc.left < lprc->right)
                lprc->right = pab->rc.left;
            break;
    }
}

void CTray::_AppBarSubtractRects(HMONITOR hmon, LPRECT lprc)
{
    int i;

    if (!_hdpaAppBars)
        return;

    i = DPA_GetPtrCount(_hdpaAppBars);

    while (i--)
    {
        PAPPBAR pab = (PAPPBAR)DPA_GetPtr(_hdpaAppBars, i);

        //
        // autohide bars are not in our DPA or live on the edge
        // don't subtract the appbar if it's on a different display
        // don't subtract the appbar if we are in a locked desktop
        //
        // if (hmon == MonitorFromRect(&pab->rc, MONITOR_DEFAULTTONULL))
        if (hmon == MonitorFromRect(&pab->rc, MONITOR_DEFAULTTONULL) && !_fIsDesktopLocked)
            _AppBarSubtractRect(pab, lprc);
    }
}

#define RWA_NOCHANGE      0
#define RWA_CHANGED       1
#define RWA_BOTTOMMOSTTRAY 2

// (dli) This is a hack put in because bottommost tray is wierd, once
// it becomes a toolbar, this code should go away.
// In the bottommost tray case, even though the work area has not changed,
// we should notify the desktop.
int CTray::_RecomputeWorkArea(HWND hwndCause, HMONITOR hmon, LPRECT prcWork)
{
    int iRet = RWA_NOCHANGE;
    MONITORINFO mi;
    mi.cbSize = sizeof(mi);

    if (_fIsLogoff)
    {
        if (GetMonitorInfo(hmon, &mi))
        {
            *prcWork = mi.rcMonitor;
            iRet = RWA_CHANGED;
        }
        return iRet;
    }

    ASSERT(!_fIsLogoff);

    //
    // tell everybody that this window changed positions _on_this_monitor_
    // note that this notify happens even if we don't change the work area
    // since it may cause another app to change the work area...
    //
    PostMessage(_hwnd, TM_RELAYPOSCHANGED, (WPARAM)hwndCause, (LPARAM)hmon);

    //
    // get the current info for this monitor
    // we subtract down from the display rectangle to build the work area
    //
    if (GetMonitorInfo(hmon, &mi))
    {
        //
        // don't subtract the tray if it is autohide
        // don't subtract the tray if it is not always on top
        // don't subtract the tray if it's on a different display
        // don't subtract the tray if it is on a different desktop
        //
        if (!(_uAutoHide & AH_ON) && _fAlwaysOnTop &&
            (hmon == _hmonStuck) && !_fIsDesktopLocked)
        {
            SubtractRect(prcWork, &mi.rcMonitor,
                &_arStuckRects[_uStuckPlace]);
        }
        else
            *prcWork = mi.rcMonitor;

        //
        // now subtract off all the appbars on this display
        //
        _AppBarSubtractRects(hmon, prcWork);

        //
        // return whether we changed anything
        //
        if (!EqualRect(prcWork, &mi.rcWork))
            iRet = RWA_CHANGED;
        else if (!(_uAutoHide & AH_ON) && (!_fAlwaysOnTop) &&
            (!IsRectEmpty(&_arStuckRects[_uStuckPlace])))
            // NOTE: This is the bottommost case, it only applies for the tray.
            // this should be taken out when bottommost tray becomes toolbar
            iRet = RWA_BOTTOMMOSTTRAY;
    }
    else
    {
        iRet = RWA_NOCHANGE;
    }

    return iRet;
}

void RedrawDesktop(RECT* prcWork)
{
    // This rect point should always be valid (dli)
    //RIP(prcWork);

    if (v_hwndDesktop && g_fCleanBoot)
    {
        MapWindowRect(NULL, v_hwndDesktop, prcWork);

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac invalidating desktop rect {%d,%d,%d,%d}"), prcWork->left, prcWork->top, prcWork->right, prcWork->bottom);
        RedrawWindow(v_hwndDesktop, prcWork, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
    }
}

void CTray::_StuckAppChange(HWND hwndCause, LPCRECT prcOld, LPCRECT prcNew, BOOL bTray)
{
    RECT rcWork1, rcWork2;
    HMONITOR hmon1, hmon2 = 0;
    int iChange = 0;

    //
    // PERF FEATURE:
    // there are cases where we end up setting the work area multiple times
    // we need to keep a static array of displays that have changed and a
    //  reenter count so we can avoid pain of sending notifies to the whole
    //  planet...
    //
    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac from_AppBar %08X"), hwndCause);

    //
    // see if the work area changed on the display containing prcOld
    //
    if (prcOld)
    {
        if (bTray)
            hmon1 = _hmonOld;
        else
            hmon1 = MonitorFromRect(prcOld, MONITOR_DEFAULTTONEAREST);

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac old pos {%d,%d,%d,%d} on monitor %08X"), prcOld->left, prcOld->top, prcOld->right, prcOld->bottom, hmon1);

        if (hmon1)
        {
            int iret = _RecomputeWorkArea(hwndCause, hmon1, &rcWork1);
            if (iret == RWA_CHANGED)
                iChange = 1;
            if (iret == RWA_BOTTOMMOSTTRAY)
                iChange = 4;
        }
    }
    else
        hmon1 = NULL;

    //
    // see if the work area changed on the display containing prcNew
    //
    if (prcNew)
    {
        hmon2 = MonitorFromRect(prcNew, MONITOR_DEFAULTTONULL);

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac new pos {%d,%d,%d,%d} on monitor %08X"), prcNew->left, prcNew->top, prcNew->right, prcNew->bottom, hmon2);

        if (hmon2 && (hmon2 != hmon1))
        {
            int iret = _RecomputeWorkArea(hwndCause, hmon2, &rcWork2);
            if (iret == RWA_CHANGED)
                iChange |= 2;
            else if (iret == RWA_BOTTOMMOSTTRAY && (!iChange))
                iChange = 4;
        }
    }

    //
    // did the prcOld's display's work area change?
    //
    if (iChange & 1)
    {
        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac changing work area for monitor %08X"), hmon1);

        // only send SENDWININICHANGE if the desktop has been created (otherwise
        // we will hang the explorer because the main thread is currently blocked)
        SystemParametersInfo(SPI_SETWORKAREA, TRUE, &rcWork1,
            (iChange == 1 && v_hwndDesktop) ? SPIF_SENDWININICHANGE : 0);

        RedrawDesktop(&rcWork1);
    }

    //
    // did the prcOld's display's work area change?
    //
    if (iChange & 2)
    {
        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sac changing work area for monitor %08X"), hmon2);

        // only send SENDWININICHANGE if the desktop has been created (otherwise
        // we will hang the explorer because the main thread is currently blocked)
        SystemParametersInfo(SPI_SETWORKAREA, TRUE, &rcWork2,
            v_hwndDesktop ? SPIF_SENDWININICHANGE : 0);

        RedrawDesktop(&rcWork2);
    }

    // only send if the desktop has been created...
    // need to send if it's from the tray or any outside app that causes size change
    // from the tray because autohideness will affect desktop size even if it's not always on top
    if ((bTray || iChange == 4) && v_hwndDesktop)
        SendMessage(v_hwndDesktop, WM_SIZE, 0, 0);
}

void CTray::_StuckTrayChange()
{
    // We used to blow off the _StuckAppChange when the tray was in autohide
    // mode, since moving or resizing an autohid tray doesn't change the
    // work area.  Now we go ahead with the _StuckAppChange in this case
    // too.  The reason is that we can get into a state where the work area
    // size is incorrect, and we want the taskbar to always be self-repairing
    // in this case (so that resizing or moving the taskbar will correct the
    // work area size).

    //
    // pass a NULL window here since we don't want to hand out our window and
    // the tray doesn't get these anyway (nobody cares as long as its not them)
    //
    _StuckAppChange(NULL, &_rcOldTray,
        &_arStuckRects[_uStuckPlace], TRUE);

    //
    // save off the new tray position...
    //
    _rcOldTray = _arStuckRects[_uStuckPlace];
}

UINT CTray::_RecalcStuckPos(LPRECT prc)
{
    RECT rcDummy;
    POINT pt;

    if (!prc)
    {
        ////DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_rsp no rect supplied, using window rect"));

        prc = &rcDummy;
        GetWindowRect(_hwnd, prc);
    }

    // use the center of the original drag rect as a staring point
    pt.x = prc->left + RECTWIDTH(*prc) / 2;
    pt.y = prc->top + RECTHEIGHT(*prc) / 2;

    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_rsp rect is {%d, %d, %d, %d} point is {%d, %d}"), prc->left, prc->top, prc->right, prc->bottom, pt.x, pt.y);

    // reset this so the drag code won't give it preference
    _uMoveStuckPlace = (UINT)-1;

    // simulate a drag back to figure out where we originated from
    // you may be tempted to remove this.  before you do think about dragging
    // the tray across monitors and then hitting ESC...
    return _CalcDragPlace(pt);
}

/*------------------------------------------------------------------
** the position is changing in response to a move operation.
**
** if the docking status changed, we need to get a new size and
** maybe a new frame style.  change the WINDOWPOS to reflect
** these changes accordingly.
**------------------------------------------------------------------*/
void CTray::_DoneMoving(LPWINDOWPOS lpwp)
{
    RECT rc, * prc;

    if ((_uMoveStuckPlace == (UINT)-1) || (_fIgnoreDoneMoving))
        return;

    if (_fSysSizing)
        _fDeferedPosRectChange = TRUE;

    rc.left = lpwp->x;
    rc.top = lpwp->y;
    rc.right = lpwp->x + lpwp->cx;
    rc.bottom = lpwp->y + lpwp->cy;

    prc = &_arStuckRects[_uMoveStuckPlace];

    if (!EqualRect(prc, &rc))
    {
        _uMoveStuckPlace = _RecalcStuckPos(&rc);
        prc = &_arStuckRects[_uMoveStuckPlace];
    }

    // Get the new hmonitor
    _hmonStuck = MonitorFromRect(prc, MONITOR_DEFAULTTONEAREST);

    lpwp->x = prc->left;
    lpwp->y = prc->top;
    lpwp->cx = RECTWIDTH(*prc);
    lpwp->cy = RECTHEIGHT(*prc);

    lpwp->flags &= ~(SWP_NOMOVE | SWP_NOSIZE);

    // if we were autohiding, we need to update our appbar autohide rect
    if (_uAutoHide & AH_ON)
    {
        // unregister us from the old side
        _AppBarSetAutoHideBar2(_hwnd, FALSE, _uStuckPlace);
    }

    // All that work might've changed _uMoveStuckPlace (since there
    // was a lot of message traffic), so check one more time.
    // Somehow, NT Stress manages to get us in here with an invalid
    // uMoveStuckPlace.
    if (IsValidSTUCKPLACE(_uMoveStuckPlace))
    {
        // remember the new state
        _uStuckPlace = _uMoveStuckPlace;
    }
    _uMoveStuckPlace = (UINT)-1;
    _UpdateVertical(_uStuckPlace);

    _HandleSizing(0, prc, _uStuckPlace, FALSE);
    if ((_uAutoHide & AH_ON) &&
        !_AppBarSetAutoHideBar2(_hwnd, TRUE, _uStuckPlace))
    {
        _AutoHideCollision();
    }
}

UINT CTray::_CalcDragPlace(POINT pt)
{
    UINT uPlace = _uMoveStuckPlace;

    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_cdp starting point is {%d, %d}"), pt.x, pt.y);

    //
    // if the mouse is currently over the tray position leave it alone
    //
    if ((uPlace == (UINT)-1) || !PtInRect(&_arStuckRects[uPlace], pt))
    {
        HMONITOR hmonDrag;
        SIZE screen, error;
        UINT uHorzEdge, uVertEdge;
        RECT rcDisplay, * prcStick;

        //
        // which display is the mouse on?
        //
        hmonDrag = _GetDisplayRectFromPoint(&rcDisplay, pt,
            MONITOR_DEFAULTTOPRIMARY);

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_cdp monitor is %08X"), hmonDrag);

        //
        // re-origin at zero to make calculations simpler
        //
        screen.cx = RECTWIDTH(rcDisplay);
        screen.cy = RECTHEIGHT(rcDisplay);
        pt.x -= rcDisplay.left;
        pt.y -= rcDisplay.top;

        //
        // are we closer to the left or right side of this display?
        //
        if (pt.x < (screen.cx / 2))
        {
            uVertEdge = STICK_LEFT;
            error.cx = pt.x;
        }
        else
        {
            uVertEdge = STICK_RIGHT;
            error.cx = screen.cx - pt.x;
        }

        //
        // are we closer to the top or bottom side of this display?
        //
        if (pt.y < (screen.cy / 2))
        {
            uHorzEdge = STICK_TOP;
            error.cy = pt.y;
        }
        else
        {
            uHorzEdge = STICK_BOTTOM;
            error.cy = screen.cy - pt.y;
        }

        //
        // closer to a horizontal or vertical edge?
        //
        uPlace = ((error.cy * screen.cx) > (error.cx * screen.cy)) ?
            uVertEdge : uHorzEdge;

        // which StuckRect should we use?
        prcStick = &_arStuckRects[uPlace];

        //
        // need to recalc stuck rect for new monitor?
        //
        if ((hmonDrag != _GetDisplayRectFromRect(NULL, prcStick,
            MONITOR_DEFAULTTONULL)))
        {
            //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_cdp re-snapping rect for new display"));
            _MakeStuckRect(prcStick, &rcDisplay, _sStuckWidths, uPlace);
        }
    }

    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_cdp edge is %d, rect is {%d, %d, %d, %d}"), uPlace, _arStuckRects[uPlace].left, _arStuckRects[uPlace].top, _arStuckRects[uPlace].right, _arStuckRects[uPlace].bottom);
    ASSERT(IsValidSTUCKPLACE(uPlace));
    return uPlace;
}

void CTray::_HandleMoving(WPARAM wParam, LPRECT lprc)
{
    POINT ptCursor;
    GetCursorPos(&ptCursor);

    // If the cursor is not far from its starting point, then ignore it.
    // A very common problem is the user clicks near the corner of the clock,
    // twitches the mouse 5 pixels, and BLAM, their taskbar is now vertical
    // and they don't know what they did or how to get it back.

    if (g_fInSizeMove && PtInRect(&_rcSizeMoveIgnore, ptCursor))
    {
        // Ignore -- user is merely twitching
        _uMoveStuckPlace = _uStuckPlace;
    }
    else
    {
        _uMoveStuckPlace = _CalcDragPlace(ptCursor);
    }

    *lprc = _arStuckRects[_uMoveStuckPlace];
    _HandleSizing(wParam, lprc, _uMoveStuckPlace, FALSE);
}

// store the tray size when dragging is finished
void CTray::_SnapshotStuckRectSize(UINT uPlace)
{
    RECT rcDisplay, * prc = &_arStuckRects[uPlace];

    //
    // record the width of this stuck rect
    //
    if (STUCK_HORIZONTAL(uPlace))
        _sStuckWidths.cy = RECTHEIGHT(*prc);
    else
        _sStuckWidths.cx = RECTWIDTH(*prc);

    //
    // we only present a horizontal or vertical size to the end user
    // so update the StuckRect on the other side of the screen to match
    //
    _GetStuckDisplayRect(uPlace, &rcDisplay);

    uPlace += 2;
    uPlace %= 4;
    prc = &_arStuckRects[uPlace];

    _MakeStuckRect(prc, &rcDisplay, _sStuckWidths, uPlace);
}


// Size the icon area to fill as much of the tray window as it can.

// EXEX-VISTA: Fairly confidently reversed from Vista.
void CTray::SizeWindows()
{
    if (!_hwndRebar || !_hwnd || !_hwndNotify)
        return;

    BOOL fHiding = (_uAutoHide & AH_HIDING);
    if (fHiding)
    {
        InvisibleUnhide(FALSE);
    }

    // remember our current size
    _SnapshotStuckRectSize(_uStuckPlace);

    _stb.RecalcSize();

    RECT rcClient;
    GetClientRect(_hwnd, &rcClient);

    RECT rcView;
    RECT rcNotify;
    _GetWindowSizes(_uStuckPlace, &rcClient, &rcView, &rcNotify);

    // position the view
    SetWindowPos(
        _hwndRebar, nullptr, rcView.left, rcView.top, RECTWIDTH(rcView), RECTHEIGHT(rcView),
        SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);

    // And the clock
    SetWindowPos(
        _hwndNotify, nullptr, rcNotify.left, rcNotify.top, RECTWIDTH(rcNotify), RECTHEIGHT(rcNotify),
        SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);

    InvalidateRect(_hwnd, nullptr, TRUE);
    UpdateWindow(_hwnd);

    if (fHiding)
    {
        InvisibleUnhide(TRUE);
    }
}


void CTray::_HandleSize()
{
    //
    // if somehow we got minimized go ahead and un-minimize
    //
    if (((GetWindowLong(_hwnd, GWL_STYLE)) & WS_MINIMIZE))
    {
        ASSERT(FALSE);
        ShowWindow(_hwnd, SW_RESTORE);
    }

    //
    // if we are in the move/size loop and are visible then
    // re-snap the current stuck rect to the new window size
    //
    #ifdef DEBUG
    if (_fSysSizing && (_uAutoHide & AH_HIDING)) {
        TraceMsg(DM_TRACE, "fSysSize && hiding");
        ASSERT(0);
    }
    #endif
    if (_fSysSizing &&
        ((_uAutoHide & (AH_ON | AH_HIDING)) != (AH_ON | AH_HIDING)))
    {
        _uStuckPlace = _RecalcStuckPos(NULL);
        _UpdateVertical(_uStuckPlace);
    }

    //
    // if we are in fulldrag or we are not in the middle of a move/size then
    // we should resize all our child windows to reflect our new size
    //
    if (g_fDragFullWindows || !_fSysSizing)
        SizeWindows();

    //
    // if we are just plain resized and we are visible we may need re-dock
    //
    if (!_fSysSizing && !_fSelfSizing && IsWindowVisible(_hwnd))
    {
        if (_uAutoHide & AH_ON)
        {
            UINT uPlace = _uStuckPlace;
            HWND hwndOther = _AppBarGetAutoHideBar(uPlace);

            //
            // we sometimes defer checking for this until after a move
            // so as to avoid interrupting a full-window-drag in progress
            // if there is a different autohide window in our slot then whimper
            //
            if (hwndOther ?
                (hwndOther != _hwnd) :
                !_AppBarSetAutoHideBar2(_hwnd, TRUE, uPlace))
            {
                _AutoHideCollision();
            }
        }

        _StuckTrayChange();

        //
        // make sure we clip to tray to the current monitor (if necessary)
        //
        _ClipWindow(TRUE);
    }

    _stb.RepositionBalloon();
}

BOOL _IsSliverHeight(int cy)
{
    //
    // Is this height clearly bigger than the pure-border height that you
    // get when you resize the taskbar as small as it will go?
    //
    return (cy < (3 * (g_cyDlgFrame + g_cyBorder)));
}

SIZE *CTray::_GetStartButtonPadding(SIZE *pSize)
{
    pSize->cx = 2 * (GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXDLGFRAME));
    pSize->cy = 2 * (GetSystemMetrics(SM_CYDLGFRAME) + GetSystemMetrics(SM_CYBORDER));
    return pSize;
}

// EXEX-VISTA: Validated.
static int _GetRebarMinHeight(HWND hWnd)
{
    int iResult = 0;

    if (hWnd)
    {
        UINT cyMinChild = 0;
        DWORD numBands = SendMessageW(hWnd, RB_GETBANDCOUNT, 0, 0);

        for (int i = 0; i < numBands; i++)
        {
            REBARBANDINFO rbbi;
            rbbi.cbSize = sizeof(rbbi);
            rbbi.fMask = 33;
            SendMessageW(hWnd, RB_GETBANDINFOW, i, (LPARAM)&rbbi);
            if ((rbbi.fStyle & 8) == 0)
            {
                if ((rbbi.fStyle & 1) != 0)
                {
                    iResult += cyMinChild;
                    cyMinChild = 0;
                }
                if (cyMinChild < rbbi.cyMinChild)
                    cyMinChild = rbbi.cyMinChild;
            }
        }

        iResult += cyMinChild;
    }

    return iResult;
}

// EXEX-VISTA: Validated.
BOOL CTray::_HandleSizing(WPARAM code, LPRECT lprc, UINT uStuckPlace, BOOL fUpdateSize)
{
    BOOL fChangedSize = FALSE;

    if (!lprc || !_fCanSizeMove)
    {
        RECT rcTemp;
        rcTemp = _arStuckRects[uStuckPlace];
        lprc = &rcTemp;
    }

    if (!_fSelfSizing)
    {
        SIZE sizeStart;
        _GetStartButtonPadding(&sizeStart);

        if (STUCK_HORIZONTAL(uStuckPlace))
        {
            sizeStart.cy += _stb._sizeStart.cy;
        }
        else
        {
            sizeStart.cx += _stb._sizeStart.cx;
        }

        if (RECTHEIGHT(*lprc) < sizeStart.cy || RECTWIDTH(*lprc) < sizeStart.cx || _fTaskbarLockAllPolicyEnabled || _fTaskbarNoResizePolicyEnabled)
        {
            *lprc = _arStuckRects[uStuckPlace];
            return TRUE;
        }
    }

    BOOL fHiding = _uAutoHide & AH_HIDING;
    if (fHiding)
    {
        InvisibleUnhide(FALSE);
    }

    RECT rcDisplay;
    _GetDisplayRectFromRect(&rcDisplay, lprc, MONITOR_DEFAULTTONEAREST);

    if (code)
    {
        RECT rcMax;
        rcMax = rcDisplay;
        if (!_hTheme)
        {
            InflateRect(&rcMax, g_cxEdge, g_cyEdge);
        }
        switch (uStuckPlace)
        {
            case STICK_LEFT:
                lprc->left = rcMax.left;
                break;
            case STICK_TOP:
                lprc->top = rcMax.top;
                break;
            case STICK_RIGHT:
                lprc->right = rcMax.right;
                break;
            case STICK_BOTTOM:
                lprc->bottom = rcMax.bottom;
                lprc->top += rcMax.bottom - lprc->bottom;
                break;
        }
    }

    SIZE sNewWidths;
    sNewWidths.cx = min(RECTWIDTH(*lprc), RECTWIDTH(rcDisplay) / 2);
    sNewWidths.cy = min(RECTHEIGHT(*lprc), RECTHEIGHT(rcDisplay) / 2);

    _MakeStuckRect(lprc, &rcDisplay, sNewWidths, uStuckPlace);
    // DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hs starting rect is {%d, %d, %d, %d}"), lprc->left, lprc->top, lprc->right, lprc->bottom);

    // DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hs tray is being calculated for %s"), STUCK_HORIZONTAL(uStuckPlace) ? TEXT("HORIZONTAL") : TEXT("VERTICAL"));
    _UpdateVertical(uStuckPlace);

    IDeskBarClient* pdbc;
    if (_ptbs && /*dword38 &&*/ SUCCEEDED(_ptbs->QueryInterface(IID_PPV_ARGS(&pdbc))))
    {
        RECT rcClient;
        RECT rcOldClient;

        rcClient = *lprc;
        rcOldClient = _arStuckRects[uStuckPlace];
        if (!_hTheme)
        {
            InflateRect(&rcClient, -g_cxFrame, -g_cyFrame);
            InflateRect(&rcOldClient, -g_cxFrame, -g_cyFrame);
        }

        OffsetRect(&rcClient, -rcClient.left, -rcClient.top);
        OffsetRect(&rcOldClient, -rcOldClient.left, -rcOldClient.top);

        RECT rcView;
        RECT rcNotify;
        RECT rcOldView;

        // DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hs starting client rect is {%d, %d, %d, %d}"), rcClient.left, rcClient.top, rcClient.right, rcClient.bottom);
        _GetWindowSizes(uStuckPlace, &rcClient, &rcView, &rcNotify);
        _GetWindowSizes(uStuckPlace, &rcOldClient, &rcOldView, &rcNotify);

        OffsetRect(&rcView, -rcView.left, -rcView.top);
        OffsetRect(&rcOldView, -rcOldView.left, -rcOldView.top);

        if (!_fCanSizeMove || (RECTHEIGHT(rcView) && RECTWIDTH(rcView)))
        {
            int iRebarMinHeight = _GetRebarMinHeight(_hwndRebar);
            if (STUCK_HORIZONTAL(uStuckPlace))
            {
                int iMinButtonHeight = 1;
                if (rcNotify.bottom + SendMessage(_hwndTasks, TBC_BUTTONHEIGHT, 0, 0) >= 1)
                {
                    iMinButtonHeight = rcNotify.bottom + SendMessage(_hwndTasks, TBC_BUTTONHEIGHT, 0, 0);
                }

                int iMinStuckHeight;
                if (fUpdateSize)
                {
                    iMinStuckHeight = _arStuckHeights[uStuckPlace];
                }
                else
                {
                    iMinStuckHeight = (rcNotify.bottom + rcView.bottom) / iMinButtonHeight;
                    _arStuckHeights[uStuckPlace] = iMinStuckHeight;
                }
                rcView.bottom = max(iMinButtonHeight * iMinStuckHeight - rcNotify.bottom, iRebarMinHeight);
            }
            else if (rcView.right <= iRebarMinHeight)
            {
                rcView.right = iRebarMinHeight;
            }
        }

        if (STUCK_HORIZONTAL(uStuckPlace))
        {
            rcClient.top = rcView.top;
            rcClient.bottom = rcView.bottom;
        }
        else
        {
            rcClient.left = rcView.left;
            rcClient.right = rcView.right;
        }

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hs ending client rect is {%d, %d, %d, %d}"), rcClient.left, rcClient.top, rcClient.right, rcClient.bottom);
        if (!_hTheme)
        {
            InflateRect(&rcClient, g_cxFrame, g_cyFrame);
            InflateRect(&rcOldClient, g_cxFrame, g_cyFrame);
        }

        if (RECTHEIGHT(rcView) && RECTHEIGHT(rcOldView) && (RECTHEIGHT(rcClient) > (3 * RECTHEIGHT(rcOldClient))))
        {
            rcClient = rcOldClient;
        }

        if (STUCK_HORIZONTAL(uStuckPlace))
        {
            if (sNewWidths.cy != RECTHEIGHT(rcClient))
            {
                sNewWidths.cy = RECTHEIGHT(rcClient);
                fChangedSize = TRUE;
            }
        }
        else if (sNewWidths.cx != RECTWIDTH(rcClient))
        {
            sNewWidths.cx = RECTWIDTH(rcClient);
            fChangedSize = TRUE;
        }
        pdbc->Release();

        if (fChangedSize)
        {
            _MakeStuckRect(lprc, &rcDisplay, sNewWidths, uStuckPlace);
        }
    }

    // DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hs final rect is {%d, %d, %d, %d}"), lprc->left, lprc->top, lprc->right, lprc->bottom);
    _arStuckRects[uStuckPlace] = *lprc;

    if (fHiding)
    {
        InvisibleUnhide(TRUE);
    }

    _stb.RepositionBalloon();

    return fChangedSize;
}

/*-------------------------------------------------------------------
** the screen size changed, and we need to adjust some stuff, mostly
** globals.  if the tray was docked, it needs to be resized, too.
**
** TRICKINESS: the handling of WM_WINDOWPOSCHANGING is used to
** actually do all the real sizing work.  this saves a bit of
** extra code here.
**-------------------------------------------------------------------*/

BOOL WINAPI CTray::MonitorEnumProc(HMONITOR hMonitor, HDC hdc, LPRECT lprc, LPARAM lData)
{
    CTray* ptray = (CTray*)lData;

    RECT rcWork;
    int iRet = ptray->_RecomputeWorkArea(NULL, hMonitor, &rcWork);

    if (iRet == RWA_CHANGED)
    {
        // only send SENDWININICHANGE if the desktop has been created (otherwise
        // we will hang the explorer because the main thread is currently blocked)

        // PERF FEATURE: it will be nice to send WININICHANGE only once, but we can't
        // because each time the rcWork is different, and there is no way to do it all
        SystemParametersInfo(SPI_SETWORKAREA, TRUE, &rcWork, v_hwndDesktop ? SPIF_SENDWININICHANGE : 0);
        RedrawDesktop(&rcWork);
    }

    return TRUE;
}

void CTray::_RecomputeAllWorkareas()
{
    EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, (LPARAM)this);
}

void CTray::_ScreenSizeChange(HWND hwnd)
{
    // Set our new HMONITOR in case there is some change
    {
        MONITORINFO mi = { 0 };
        mi.cbSize = sizeof(mi);

        // Is our old HMONITOR still valid?
        // NOTE: This test is used to tell whether somethings happened to the
        // HMONITOR's or just the screen size changed
        if (!GetMonitorInfo(_hmonStuck, &mi))
        {
            // No, this means the HMONITORS changed, our monitor might have gone away
            _SetStuckMonitor();
            _fIsLogoff = FALSE;
            _RecomputeAllWorkareas();
        }
    }

    // screen size changed, so we need to adjust globals
    g_cxPrimaryDisplay = GetSystemMetrics(SM_CXSCREEN);
    g_cyPrimaryDisplay = GetSystemMetrics(SM_CYSCREEN);

    _ResizeStuckRects(_arStuckRects);

    if (hwnd)
    {
        //
        // set a bogus windowpos and actually repaint with the right
        // shape/size in handling the WINDOWPOSCHANGING message
        //
        SetWindowPos(hwnd, NULL, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    SizeWindows();

    RECT rc = _arStuckRects[_uStuckPlace];
    _HandleSizing(0, &rc, _uStuckPlace, FALSE);
    // In the multi-monitor case, we need to turn on clipping for the dynamic
    // monitor changes i.e. when the user add monitors or remove them from the
    // control panel.
    _ClipWindow(TRUE);
}

BOOL CTray::_UpdateAlwaysOnTop(BOOL fAlwaysOnTop)
{
    BOOL fChanged = ((_fAlwaysOnTop == 0) != (fAlwaysOnTop == 0));
    //
    // The user clicked on the AlwaysOnTop menu item, we should now toggle
    // the state and update the window accordingly...
    //
    _fAlwaysOnTop = fAlwaysOnTop;
    _ResetZorder(0);

    // Make sure the screen limits are update to the new state.
    _StuckTrayChange();
    return fChanged;
}



void CTray::_SaveTrayAndDesktop(void)
{
    _SaveTray();

    if (v_hwndDesktop)
        SendMessage(v_hwndDesktop, DTM_SAVESTATE, 0, 0);

    if (_hwndNotify)
        SendMessage(_hwndNotify, TNM_SAVESTATE, 0, 0);
}

void CTray::_SlideStep(HWND hwnd, const RECT* prcMonitor, const RECT* prcOld, const RECT* prcNew)
{
    SIZE sizeOld = { prcOld->right - prcOld->left, prcOld->bottom - prcOld->top };
    SIZE sizeNew = { prcNew->right - prcNew->left, prcNew->bottom - prcNew->top };
    BOOL fClipFirst = FALSE;
    UINT flags;

    DAD_ShowDragImage(FALSE);   // Make sure this is off - client function must turn back on!!!
    if (prcMonitor)
    {
        RECT rcClip, rcClipSafe, rcClipTest;

        _CalcClipCoords(&rcClip, prcMonitor, prcNew);

        rcClipTest = rcClip;

        OffsetRect(&rcClipTest, prcOld->left, prcOld->top);
        IntersectRect(&rcClipSafe, &rcClipTest, prcMonitor);

        fClipFirst = EqualRect(&rcClipTest, &rcClipSafe);

        if (fClipFirst)
            _ClipInternal(&rcClip);
    }

    flags = SWP_NOZORDER | SWP_NOACTIVATE;
    if ((sizeOld.cx == sizeNew.cx) && (sizeOld.cy == sizeNew.cy))
        flags |= SWP_NOSIZE;

    SetWindowPos(hwnd, NULL,
        prcNew->left, prcNew->top, sizeNew.cx, sizeNew.cy, flags);

    if (prcMonitor && !fClipFirst)
    {
        RECT rcClip;

        _CalcClipCoords(&rcClip, prcMonitor, prcNew);
        _ClipInternal(&rcClip);
    }
}

void CTray::_SlideWindow(HWND hwnd, RECT* prc, BOOL fShow)
{
    RECT rcLast;
    RECT rcMonitor;
    const RECT* prcMonitor;
    DWORD dt;
    BOOL fAnimate;

    if (!IsWindowVisible(hwnd))
    {
        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sw window is hidden, just moving"));
        MoveWindow(_hwnd, prc->left, prc->top, RECTWIDTH(*prc), RECTHEIGHT(*prc), FALSE);
        return;
    }

    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sw -----------------BEGIN-----------------"));

    if (GetSystemMetrics(SM_CMONITORS) > 1)
    {
        _GetStuckDisplayRect(_uStuckPlace, &rcMonitor);
        prcMonitor = &rcMonitor;
    }
    else
        prcMonitor = NULL;

    GetWindowRect(hwnd, &rcLast);

    dt = fShow ? _dtSlideShow : _dtSlideHide;

    // See if we can use animation effects.
    SystemParametersInfo(SPI_GETMENUANIMATION, 0, &fAnimate, 0);

    if (g_fDragFullWindows && fAnimate && (dt > 0))
    {
        RECT rcOld, rcNew, rcMove;
        int  dx, dy, priority;
        DWORD t, t2, t0;
        HANDLE me;

        rcOld = rcLast;
        rcNew = *prc;

        dx = ((rcNew.left + rcNew.right) - (rcOld.left + rcOld.right)) / 2;
        dy = ((rcNew.top + rcNew.bottom) - (rcOld.top + rcOld.bottom)) / 2;
        ASSERT(dx == 0 || dy == 0);

        me = GetCurrentThread();
        priority = GetThreadPriority(me);
        SetThreadPriority(me, THREAD_PRIORITY_HIGHEST);

        t2 = t0 = GetTickCount();

        rcMove = rcOld;
        while ((t = GetTickCount()) - t0 < dt)
        {
            int dtdiff;
            if (t != t2)
            {
                rcMove.right -= rcMove.left;
                rcMove.left = rcOld.left + (dx) * (t - t0) / dt;
                rcMove.right += rcMove.left;

                rcMove.bottom -= rcMove.top;
                rcMove.top = rcOld.top + (dy) * (t - t0) / dt;
                rcMove.bottom += rcMove.top;

                _SlideStep(hwnd, prcMonitor, &rcLast, &rcMove);

                if (fShow)
                    UpdateWindow(hwnd);

                rcLast = rcMove;
                t2 = t;
            }

            // don't draw frames faster than user can see, e.g. 20ms
            #define ONEFRAME 20
            dtdiff = GetTickCount();
            if ((dtdiff - t) < ONEFRAME)
                Sleep(ONEFRAME - (dtdiff - t));

            // try to give desktop a chance to update
            // only do it on hide because desktop doesn't need to paint on taskbar show
            if (!fShow)
            {
                DWORD_PTR lres;
                SendMessageTimeout(v_hwndDesktop, DTM_UPDATENOW, 0, 0, SMTO_ABORTIFHUNG, 50, &lres);
            }
        }

        SetThreadPriority(me, priority);
    }

    _SlideStep(hwnd, prcMonitor, &rcLast, prc);

    if (fShow)
        UpdateWindow(hwnd);

    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sw ------------------END------------------"));
}

void CTray::_UnhideNow()
{
    if (_uModalMode == MM_SHUTDOWN)
        return;

    _fSelfSizing = TRUE;
    DAD_ShowDragImage(FALSE);   // unlock the drag sink if we are dragging.
    _SlideWindow(_hwnd, &_arStuckRects[_uStuckPlace], _dtSlideShow);
    DAD_ShowDragImage(TRUE);    // restore the lock state.
    _fSelfSizing = FALSE;

    SendMessage(_hwndNotify, TNM_TRAYHIDE, 0, FALSE);
}

void CTray::_ComputeHiddenRect(LPRECT prc, UINT uStuck)
{
    int dwh;
    HMONITOR hMon;
    RECT rcMon;

    hMon = MonitorFromRect(prc, MONITOR_DEFAULTTONULL);
    if (!hMon)
        return;
    GetMonitorRect(hMon, &rcMon);

    if (STUCK_HORIZONTAL(uStuck))
        dwh = prc->bottom - prc->top;
    else
        dwh = prc->right - prc->left;

    switch (uStuck)
    {
        case STICK_LEFT:
            prc->right = rcMon.left + (g_cxFrame / 2);
            prc->left = prc->right - dwh;
            break;

        case STICK_RIGHT:
            prc->left = rcMon.right - (g_cxFrame / 2);
            prc->right = prc->left + dwh;
            break;

        case STICK_TOP:
            prc->bottom = rcMon.top + (g_cyFrame / 2);
            prc->top = prc->bottom - dwh;
            break;

        case STICK_BOTTOM:
            prc->top = rcMon.bottom - (g_cyFrame / 2);
            prc->bottom = prc->top + dwh;
            break;
    }
}

UINT CTray::_GetDockedRect(LPRECT prc, BOOL fMoving)
{
    UINT uPos;

    if (fMoving && (_uMoveStuckPlace != (UINT)-1))
        uPos = _uMoveStuckPlace;
    else
        uPos = _uStuckPlace;

    *prc = _arStuckRects[uPos];

    if ((_uAutoHide & (AH_ON | AH_HIDING)) == (AH_ON | AH_HIDING))
    {
        _ComputeHiddenRect(prc, uPos);
    }

    return uPos;
}

void CTray::_CalcClipCoords(RECT* prcClip, const RECT* prcMonitor, const RECT* prcNew)
{
    RECT rcMonitor;
    RECT rcWindow;

    if (!prcMonitor)
    {
        _GetStuckDisplayRect(_uStuckPlace, &rcMonitor);
        prcMonitor = &rcMonitor;
    }

    if (!prcNew)
    {
        GetWindowRect(_hwnd, &rcWindow);
        prcNew = &rcWindow;
    }

    IntersectRect(prcClip, prcMonitor, prcNew);
    OffsetRect(prcClip, -prcNew->left, -prcNew->top);
}

void CTray::_ClipInternal(const RECT* prcClip)
{
    HRGN hrgnClip;

    // don't worry about clipping if there's only one monitor
    if (GetSystemMetrics(SM_CMONITORS) <= 1)
        prcClip = NULL;

    if (prcClip)
    {
        _fMonitorClipped = TRUE;
        hrgnClip = CreateRectRgnIndirect(prcClip);
    }
    else
    {
        // SetWindowRgn is expensive, skip ones that are NOPs
        if (!_fMonitorClipped)
            return;

        _fMonitorClipped = FALSE;
        hrgnClip = NULL;
    }

    SetWindowRgn(_hwnd, hrgnClip, TRUE);
}

void CTray::_ClipWindow(BOOL fClipState)
{
    RECT rcClip;
    RECT* prcClip;

    if (_fSelfSizing || _fSysSizing)
    {
        //TraceMsg(TF_WARNING, "_ClipWindow: _fSelfSizing %x, _fSysSizing %x", _fSelfSizing, _fSysSizing);
        return;
    }

    if ((GetSystemMetrics(SM_CMONITORS) <= 1) || _hTheme)
        fClipState = FALSE;

    if (fClipState)
    {
        prcClip = &rcClip;
        _CalcClipCoords(prcClip, NULL, NULL);
    }
    else
        prcClip = NULL;

    _ClipInternal(prcClip);
}

void CTray::_Hide()
{
    RECT rcNew;

    // if we're in shutdown or if we're on boot up
    // don't hide
    if (_uModalMode == MM_SHUTDOWN)
    {
        //TraceMsg(TF_TRAY, "e.th: suppress hide (shutdown || Notify)");
        return;
    }

    KillTimer(_hwnd, IDT_AUTOHIDE);

    _fSelfSizing = TRUE;

    //
    // update the flags here to prevent race conditions
    //
    _uAutoHide = AH_ON | AH_HIDING;
    _GetDockedRect(&rcNew, FALSE);

    DAD_ShowDragImage(FALSE);   // unlock the drag sink if we are dragging.
    _SlideWindow(_hwnd, &rcNew, _dtSlideHide);
    DAD_ShowDragImage(FALSE);   // Another thread could have locked while we were gone
    DAD_ShowDragImage(TRUE);    // restore the lock state.

    SendMessage(_hwndNotify, TNM_TRAYHIDE, 0, TRUE);

    _fSelfSizing = FALSE;
}

void CTray::_AutoHideCollision()
{
    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_ahc COLLISION! (posting UI request)"));

    PostMessage(_hwnd, TM_WARNNOAUTOHIDE, ((_uAutoHide & AH_ON) != 0),
        0L);
}

LONG CTray::_SetAutoHideState(BOOL fAutoHide)
{
    //
    // make sure we have something to do
    //
    if ((fAutoHide != 0) == ((_uAutoHide & AH_ON) != 0))
    {
        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.sahs nothing to do"));
        return MAKELONG(FALSE, TRUE);
    }

    //
    // make sure we can do it
    //
    if (!_AppBarSetAutoHideBar2(_hwnd, fAutoHide, _uStuckPlace))
    {
        _AutoHideCollision();
        return MAKELONG(FALSE, FALSE);
    }

    //
    // do it
    //
    if (fAutoHide)
    {
        _uAutoHide = AH_ON;
        _RefreshSettings();
        _Hide();
        _AccountAllBandsForTaskbarSizingBar();
        #ifdef DEBUG
                // _Hide updates the flags for us (sanity)
        if (!(_uAutoHide & AH_ON))
        {
            //TraceMsg(DM_WARNING, "e.sahs: !AH_ON"); // ok to fail on boot/shutdown
        }
        #endif
    }
    else
    {
        _uAutoHide = 0;
        KillTimer(_hwnd, IDT_AUTOHIDE);
        _UnhideNow();
        _RefreshSettings();
    }

    //
    // brag about it
    //
    _StuckTrayChange();
    return MAKELONG(TRUE, TRUE);
}

void CTray::_HandleEnterMenuLoop()
{
    // kill the timer when we're in the menu loop so that we don't
    // pop done while browsing the menus.
    if (_uAutoHide & AH_ON)
    {
        KillTimer(_hwnd, IDT_AUTOHIDE);
    }
}

void CTray::_SetAutoHideTimer()
{
    if (_uAutoHide & AH_ON)
    {
        SetTimer(_hwnd, IDT_AUTOHIDE, 500, NULL);
    }
}

void CTray::_HandleExitMenuLoop()
{
    // when we leave the menu stuff, start checking again.
    _SetAutoHideTimer();
}

void CTray::Unhide()
{
    // handle autohide
    if ((_uAutoHide & AH_ON) &&
        (_uAutoHide & AH_HIDING))
    {
        _UnhideNow();
        _uAutoHide &= ~AH_HIDING;
        _SetAutoHideTimer();

        if (_fShouldResize)
        {
            ASSERT(0);
            ASSERT(!(_uAutoHide & AH_HIDING));
            SizeWindows();
            _fShouldResize = FALSE;
        }
    }
}

void CTray::_SetUnhideTimer(LONG x, LONG y)
{
    // handle autohide
    if ((_uAutoHide & AH_ON) &&
        (_uAutoHide & AH_HIDING))
    {
        LONG dx = x - _ptLastHittest.x;
        LONG dy = y - _ptLastHittest.y;
        LONG rr = dx * dx + dy * dy;
        LONG dd = GetSystemMetrics(SM_CXDOUBLECLK) * GetSystemMetrics(SM_CYDOUBLECLK);

        if (rr > dd)
        {
            SetTimer(_hwnd, IDT_AUTOUNHIDE, 50, NULL);
            _ptLastHittest.x = x;
            _ptLastHittest.y = y;
        }
    }
}

void CTray::_OnNewSystemSizes()
{
    //TraceMsg(TF_TRAY, "Handling win ini change.");
    _stb.StartButtonReset();
    VerifySize(TRUE);
}

//***   CheckWindowPositions -- flag which windows actually changed
// ENTRY/EXIT
//  _pPositions->hdsaWP[*]->fRestore modified
// NOTES
//  in order to correctly implement 'Undo Minimize-all(Cascade/Tile)',
// we need to tell which windows were changed by the 'Do' operation.
// (nt5:183421: we used to restore *every* top-level window).
int WINAPI CTray::CheckWndPosEnumProc(void* pItem, void* pData)
{
    HWNDANDPLACEMENT* pI2 = (HWNDANDPLACEMENT*)pItem;
    WINDOWPLACEMENT wp;

    wp.length = sizeof(wp);
    pI2->fRestore = TRUE;
    if (GetWindowPlacement(pI2->hwnd, &wp)) {
        if (memcmp(&pI2->wp, &wp, sizeof(wp)) == 0)
            pI2->fRestore = FALSE;
    }

    //TraceMsg(TF_TRAY, "cwp: (hwnd=0x%x) fRestore=%d", pI2->hwnd, pI2->fRestore);

    return 1;   // 1:continue enum
}

void CTray::CheckWindowPositions()
{
    ENTERCRITICAL;      // i think this is needed...
    if (_pPositions) {
        if (_pPositions->hdsaWP) {
            DSA_EnumCallback(_pPositions->hdsaWP, CheckWndPosEnumProc, NULL);
        }
    }
    LEAVECRITICAL;

    return;
}

BOOL BandSite_PermitAutoHide(IUnknown* punk)
{
    OLECMD cmd = { DBID_PERMITAUTOHIDE, 0 };
    if (SUCCEEDED(IUnknown_QueryStatus(punk, &IID_IDockingWindow, 1, &cmd, NULL)))
    {
        return !(cmd.cmdf & OLECMDF_SUPPORTED) || (cmd.cmdf & OLECMDF_ENABLED);
    }
    return TRUE;
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
void CTray::_HandleTimer(WPARAM wTimerID)
{
    switch (wTimerID)
    {
        case IDT_CHECKDISKSPACE:
            CheckDiskSpace();
            break;

        case IDT_CHANGENOTIFY:
            // did somebody send another notify since we last handled one?
            if (_fUseChangeNotifyTimer)
            {
                // yep.
                // all we do is check the staging area.
                CheckStagingArea();

                // kill it off the next time.
                _fUseChangeNotifyTimer = FALSE;
            }
            else
            {
                // nope.
                KillTimer(_hwnd, IDT_CHANGENOTIFY);
                _fChangeNotifyTimerRunning = FALSE;
            }
            break;

        case IDT_HANDLEDELAYBOOTSTUFF:
            KillTimer(_hwnd, IDT_HANDLEDELAYBOOTSTUFF);
            PostMessage(_hwnd, TM_HANDLEDELAYBOOTSTUFF, 0, 0);
            break;

        case IDT_STARTMENU:
            SetForegroundWindow(_hwnd);
            KillTimer(_hwnd, wTimerID);
            DAD_ShowDragImage(FALSE);       // unlock the drag sink if we are dragging.
            SendMessage(_stb._hwndStart, BM_SETSTATE, TRUE, 0);
            UpdateWindow(_stb._hwndStart);
            DAD_ShowDragImage(TRUE);        // restore the lock state.
            break;

        case IDT_SAVESETTINGS:
            KillTimer(_hwnd, IDT_SAVESETTINGS);
            _SaveTray();
            break;

        case IDT_ENABLEUNDO:
            KillTimer(_hwnd, IDT_ENABLEUNDO);
            CheckWindowPositions();
            _fUndoEnabled = TRUE;
            break;

        case IDT_AUTOHIDE:
            if (!_fSysSizing && (_uAutoHide & AH_ON))
            {
                POINT pt;
                RECT rc;

                // Don't hide if we're already hiding, a balloon tip is showing, or
                // (on NT5) if any apps are flashing.
                //
                if (!(_uAutoHide & AH_HIDING) && BandSite_PermitAutoHide(_ptbs) && !_fBalloonUp)
                {
                    // Get the cursor position.
                    GetCursorPos(&pt);

                    // Get the tray rect and inflate it a bit.
                    rc = _arStuckRects[_uStuckPlace];
                    InflateRect(&rc, g_cxEdge * 4, g_cyEdge * 4);

                    // Don't hide if cursor is within inflated tray rect.
                    if (!PtInRect(&rc, pt))
                    {
                        // Don't hide if the tray is active
                        if (!_IsActive() && _uStartButtonState != SBSM_SPACTIVE)
                        {
                            // Don't hide if the view has a system menu up.
                            if (!SendMessage(_hwndTasks, TBC_SYSMENUCOUNT, 0, 0L))
                            {
                                // Phew!  We made it.  Hide the tray.
                                _Hide();
                            }
                        }
                    }
                }
            }
            break;

        case IDT_AUTOUNHIDE:
            if (!_fSysSizing && (_uAutoHide & AH_ON))
            {
                POINT pt;
                RECT rc;

                KillTimer(_hwnd, wTimerID);
                _ptLastHittest.x = -0x0fff;
                _ptLastHittest.y = -0x0fff;
                GetWindowRect(_hwnd, &rc);
                if (_uAutoHide & AH_HIDING)
                {
                    GetCursorPos(&pt);
                    if (PtInRect(&rc, pt))
                        Unhide();
                }
            }
            break;

        // EXEX-VISTA: LAZILY MODIFIED. Redo later.
        case IDT_STARTBUTTONBALLOON:
            // _DestroyStartButtonBalloon();
            break;

        case IDT_COFREEUNUSED:
            CoFreeUnusedLibraries();
            KillTimer(_hwnd, IDT_COFREEUNUSED);
            break;

    }
}

void CTray::_CheckStagingAreaOnTimer()
{
    if (_fChangeNotifyTimerRunning)
    {
        // we're already running the timer, so force a check the next time it comes up
        _fUseChangeNotifyTimer = TRUE;
    }
    else
    {
        _fChangeNotifyTimerRunning = TRUE;

        // check once immediately
        CheckStagingArea();

        // check again in half a minute, but only if notifies have been happening in the meantime.
        SetTimer(_hwnd, IDT_CHANGENOTIFY, 30 * 1000, NULL);
    }
}

void CTray::_HandleChangeNotify(WPARAM wParam, LPARAM lParam)
{
    LPITEMIDLIST* ppidl;
    LONG lEvent;
    LPSHChangeNotificationLock pshcnl = SHChangeNotification_Lock((HANDLE)wParam, (DWORD)lParam, &ppidl, &lEvent);
    if (pshcnl)
    {
        if (lEvent & SHCNE_STAGINGAREANOTIFICATIONS)
        {
            // something has changed within the staging area.
            _CheckStagingAreaOnTimer();
        }
        SHChangeNotification_Unlock(pshcnl);
    }
}

BOOL _ExecItemByPidls(HWND hwnd, LPITEMIDLIST pidlFolder, LPITEMIDLIST pidlItem)
{
    BOOL fRes = FALSE;

    if (pidlFolder && pidlItem)
    {
        IShellFolder* psf = BindToFolder(pidlFolder);
        if (psf)
        {
            fRes = SUCCEEDED(SHInvokeDefaultCommand(hwnd, psf, pidlItem));
        }
        else
        {
            TCHAR szPath[MAX_PATH];
            SHGetPathFromIDList(pidlFolder, szPath);
            ShellMessageBox(g_hinstCabinet, hwnd, MAKEINTRESOURCE(IDS_CANTFINDSPECIALDIR),
                NULL, MB_ICONEXCLAMATION, szPath);
        }
    }
    return fRes;
}

void _DestroySavedWindowPositions(LPWINDOWPOSITIONS pPositions);

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
LRESULT CTray::_HandleDestroy()
{
    _fFromStart = TRUE;

    MINIMIZEDMETRICS mm;

    //TraceMsg(DM_SHUTDOWN, "_HD: enter");

    mm.cbSize = sizeof(mm);
    SystemParametersInfo(SPI_GETMINIMIZEDMETRICS, sizeof(mm), &mm, FALSE);
    mm.iArrange &= ~ARW_HIDE;
    SystemParametersInfo(SPI_SETMINIMIZEDMETRICS, sizeof(mm), &mm, FALSE);

    _RevokeDropTargets();
    _DestroyStartMenu();
    Mixer_Shutdown();

    // Tell the start menu to free all its cached darwin links
    SHRegisterDarwinLink(NULL, NULL, TRUE);

    _DestroySavedWindowPositions(_pPositions);
    _pPositions = NULL;

    if (_hTheme)
    {
        CloseThemeData(_hTheme);
        _hTheme = NULL;
    }

    _UnregisterGlobalHotkeys();

    if (_uNotify)
    {
        SHChangeNotifyDeregister(_uNotify);
        _uNotify = 0;
    }

    ATOMICRELEASE(_ptbs);
    ATOMICRELEASE(_pdbTasks);
    _hwndTasks = NULL;

    if (_hwndTrayTips)
    {
        DestroyWindow(_hwndTrayTips);
        _hwndTrayTips = NULL;
    }

    // REVIEW
    PostQuitMessage(0);

    if (_pSysTray)
    {
        _pSysTray->Exec(&CGID_ShellServiceObject, SSOCMDID_CLOSE, 0, nullptr, nullptr);
        IUnknown_SafeReleaseAndNullPtr(&_pSysTray);
    }

    if (_himlStartFlag)
    {
        ImageList_Destroy(_himlStartFlag);
    }

    // clean up service objects
    // _ssomgr.Destroy();

    if (_hShellReadyEvent)
    {
        ResetEvent(_hShellReadyEvent);
        CloseHandle(_hShellReadyEvent);
        _hShellReadyEvent = NULL;
    }

    if (_hShellDesktopSwitch)
    {
        ResetEvent(_hShellDesktopSwitch);
        CloseHandle(_hShellDesktopSwitch);
        _hShellDesktopSwitch = NULL;
    }

    if (_fHandledDelayBootStuff)
    {
        TBOOL(WinStationUnRegisterConsoleNotification(SERVERNAME_CURRENT, v_hwndTray));
    }

    DeleteCriticalSection(&_csHotkey);

    // The order in which we shut down the HTTP key monitoring is important.
    //
    // We must close the key before closing the event handle because
    // closing the key causes the event to be signalled and we don't
    // want ADVAPI32 to try to signal an event after we closed its handle...
    //
    // To avoid a spurious trigger when the event fires, we unregister
    // the wait before closing the key.
    //

    if (_hHTTPWait)
    {
        UnregisterWait(_hHTTPWait);
        _hHTTPWait = NULL;
    }

    if (_hkHTTP)
    {
        RegCloseKey(_hkHTTP);
        _hkHTTP = NULL;
    }

    if (_hHTTPEvent)
    {
        CloseHandle(_hHTTPEvent);
        _hHTTPEvent = NULL;
    }

    // End of order-sensitive operations ----------------------------------

    v_hwndTray = NULL;
    _stb._hwndStart = NULL;


    //TraceMsg(DM_SHUTDOWN, "_HD: leave");
    return 0;
}

void CTray::_SetFocus(HWND hwnd)
{
    IUnknown_UIActivateIO(_ptbs, FALSE, NULL);
    SetFocus(hwnd);
}

#define TRIEDTOOMANYTIMES 100

// EXEX-VISTA: SLIGHTLY MODIFIED. Confidently incomplete.
void CTray::_ActAsSwitcher()
{
#ifdef DEAD_CODE
    if (_uModalMode)
    {
        if (_uModalMode != MM_SHUTDOWN)
        {
            SwitchToThisWindow(GetLastActivePopup(_hwnd), TRUE);
        }
        MessageBeep(0);
    }
    else
    {
        HWND hwndForeground;
        HWND hwndActive;

        static int s_iRecurse = 0;
        s_iRecurse++;

        ASSERT(s_iRecurse < TRIEDTOOMANYTIMES);
        //TraceMsg(TF_TRAY, "s_iRecurse = %d", s_iRecurse);

        hwndForeground = GetForegroundWindow();
        hwndActive = GetActiveWindow();
        BOOL fIsTrayActive = (hwndForeground == _hwnd) && (hwndActive == _hwnd);
        if (v_hwndStartPane && hwndForeground == v_hwndStartPane && hwndActive == v_hwndStartPane)
        {
            fIsTrayActive = TRUE;
        }
        // only do the button once we're the foreground dude.
        if (fIsTrayActive)
        {
            // This code path causes the start button to do something because
            // of the keyboard. So reflect that with the focus rect.
            SendMessage(_stb._hwndStart, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR,
                UISF_HIDEFOCUS), 0);

            if (_stb.IsButtonPushed())
            {
                _stb.CloseStartMenu();
                ForceStartButtonUp();
            }
            else
            {
                // This pushes the start button and causes the start menu to popup.
                SendMessage(GetDlgItem(_hwnd, IDC_START), BM_SETSTATE, TRUE, 0);
            }
            s_iRecurse = 0;
        }
        else
        {
            // we don't want to loop endlessly trying to become
            // foreground.  With NT's new SetForegroundWindow rules, it would
            // be pointless to try and hopefully we won't need to anyhow.
            // Randomly, I picked a quarter as many times as the debug squirty would indicate
            // as the number of times to try on NT.
            // Hopefully that is enough on most machines.
            if (s_iRecurse > (TRIEDTOOMANYTIMES / 4))
            {
                s_iRecurse = 0;
                return;
            }
            // until then, try to come forward.
            HandleFullScreenApp(NULL);
            if (hwndForeground == v_hwndDesktop)
            {
                _SetFocus(_stb._hwndStart);
                if (GetFocus() != _stb._hwndStart)
                    return;
            }

            SwitchToThisWindow(_hwnd, TRUE);
            SetForegroundWindow(_hwnd);
            Sleep(20); // give some time for other async activation messages to get posted
            PostMessage(_hwnd, TM_ACTASTASKSW, 0, 0);
        }
    }
#else
    static int s_iRecurse;

    if (_uModalMode)
    {
        if (_uModalMode != MM_SHUTDOWN)
        {
            SwitchToThisWindow(GetLastActivePopup(_hwnd), TRUE);
        }
        MessageBeep(0);
    }
    else
    {
        ASSERT(s_iRecurse < TRIEDTOOMANYTIMES); // 5504
        //TraceMsg(TF_TRAY, "s_iRecurse = %d", s_iRecurse);


        HWND ForegroundWindow = GetForegroundWindow();
        HWND ActiveWindow = GetActiveWindow();
        HWND Ancestor = GetAncestor(_stb._hwndStart, GA_ROOTOWNER);
        HWND hwnd = Ancestor;
        BOOL v6 = ForegroundWindow == Ancestor && ActiveWindow == Ancestor;
        if (v_hwndStartPane && ForegroundWindow == v_hwndStartPane && ActiveWindow == v_hwndStartPane)
        {
            v6 = 1;
        }
        if (v6)
        {
            SendMessageW(this->_stb._hwndStart, 0x128u, 0x10002u, 0);
            if (_stb.IsButtonPushed())
            {
				_stb.CloseStartMenu();
                ForceStartButtonUp();
            }
            else
            {
                SendMessageW(this->_stb._hwndStart, 0xF3u, 1u, 0);
            }
            s_iRecurse = 0;
        }
        else if (s_iRecurse <= 25)
        {
            HandleFullScreenApp(NULL);
            if (ForegroundWindow != v_hwndDesktop
                || (_SetFocus(_stb._hwndStart), GetFocus() == this->_stb._hwndStart))
            {
                SwitchToThisWindow(hwnd, 1);
                SetForegroundWindow(hwnd);
                Sleep(20);
                PostMessage(_hwnd, 0x504u, 0, 0);
            }
        }
        else
        {
            s_iRecurse = 0;
        }
    }
#endif
}

// EXEX-VISTA: Slightly modified. Revalidate later.
void CTray::_OnWinIniChange(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    Cabinet_InitGlobalMetrics(wParam, (LPTSTR)lParam);

    // Reset the programs menu.
    // REVIEW IANEL - We should only need to listen to the SPI_SETNONCLIENT stuff
    // but deskcpl doesn't send one.
    if (wParam == SPI_SETNONCLIENTMETRICS || (!wParam && (!lParam || (lstrcmpi((LPTSTR)lParam, TEXT("WindowMetrics")) == 0))))
    {
#ifdef DEBUG
        if (wParam == SPI_SETNONCLIENTMETRICS)
        {
            //TraceMsg(TF_TRAY, "c.t_owic: Non-client metrics (probably) changed.");
        }
        else
        {
            //TraceMsg(TF_TRAY, "c.t_owic: Window metrics changed.");
        }
#endif

		_OnNewSystemSizes();
    }

    // Handle old extensions.
    if (!lParam || (lParam && (lstrcmpi((LPTSTR)lParam, TEXT("Extensions")) == 0)))
    {
        //TraceMsg(TF_TRAY, "t_owic: Extensions section change.");
        CheckWinIniForAssocs();
    }

    if (lParam && (0 == lstrcmpi((LPCTSTR)lParam, TEXT("TraySettings"))))
    {
        _RefreshStartMenu();
    }

    // Tell shell32 to refresh its cache
    SHSettingsChanged(wParam, lParam);
}

HWND CTray::_HotkeyInUse(WORD wHK)
{
    HWND hwnd;
    DWORD_PTR lrHKInUse = 0;
    int nMod;
    WORD wHKNew;
    #ifdef DEBUG
    TCHAR sz[MAX_PATH];
    #endif

        // Map the modifiers back.
    nMod = 0;
    if (HIBYTE(wHK) & MOD_SHIFT)
        nMod |= HOTKEYF_SHIFT;
    if (HIBYTE(wHK) & MOD_CONTROL)
        nMod |= HOTKEYF_CONTROL;
    if (HIBYTE(wHK) & MOD_ALT)
        nMod |= HOTKEYF_ALT;

    wHKNew = (WORD)((nMod * 256) + LOBYTE(wHK));

    //DebugMsg(DM_IANELHK, TEXT("c.hkl_hiu: Checking for %x"), wHKNew);
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD);
    while (hwnd)
    {
        SendMessageTimeout(hwnd, WM_GETHOTKEY, 0, 0, SMTO_ABORTIFHUNG | SMTO_BLOCK, 3000, &lrHKInUse);
        if (wHKNew == (WORD)lrHKInUse)
        {
            #ifdef DEBUG
            GetWindowText(hwnd, sz, ARRAYSIZE(sz));
            //DebugMsg(DM_IANELHK, TEXT("c.hkl_hiu: %s (%x) is using %x"), sz, hwnd, lrHKInUse);
            #endif
            return hwnd;
        }
        #ifdef DEBUG
        else if (lrHKInUse)
        {
            GetWindowText(hwnd, sz, ARRAYSIZE(sz));
            //DebugMsg(DM_IANELHK, TEXT("c.hkl_hiu: %s (%x) is using %x"), sz, hwnd, lrHKInUse);
        }
        #endif
        hwnd = GetWindow(hwnd, GW_HWNDNEXT);
    }
    return NULL;
}

#define DECLAREWAITCURSOR  HCURSOR hcursor_wait_cursor_save
#define SetWaitCursor()   hcursor_wait_cursor_save = SetCursor(LoadCursor(NULL, IDC_WAIT))
#define ResetWaitCursor() SetCursor(hcursor_wait_cursor_save)

void CTray::_HandleHotKey(int nID)
{
    //TraceMsg(TF_TRAY, "c.hkl_hh: Handling hotkey (%d).", nID);

    // Find it in the list.
    ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

    EnterCriticalSection(&_csHotkey);

    HOTKEYITEM* phki = (HOTKEYITEM*)DSA_GetItemPtr(_hdsaHKI, nID);
    if (phki && phki->wGHotkey)
    {
        //TraceMsg(TF_TRAY, "c.hkl_hh: Hotkey listed.");

        // Are global hotkeys enabled?
        if (!_fGlobalHotkeyDisable)
        {
            // Yep.
            HWND hwnd = _HotkeyInUse(phki->wGHotkey);
            // Make sure this hotkey isn't already in use by someone.
            if (hwnd)
            {
                //TraceMsg(TF_TRAY, "c.hkl_hh: Hotkey is already in use.");
                // Activate it.
                SwitchToThisWindow(GetLastActivePopup(hwnd), TRUE);
            }
            else
            {
                DECLAREWAITCURSOR;
                // Exec the item.
                SetWaitCursor();
                //TraceMsg(TF_TRAY, "c.hkl_hh: Hotkey is not in use, execing item.");
                ASSERT(phki->pidlFolder && phki->pidlItem);
                BOOL fRes = _ExecItemByPidls(_hwnd, phki->pidlFolder, phki->pidlItem);
                ResetWaitCursor();
                #ifdef DEBUG
                if (!fRes)
                {
                    //DebugMsg(DM_ERROR, TEXT("c.hkl_hh: Can't exec command ."));
                }
                #endif
            }
        }
        else
        {
            //DebugMsg(DM_ERROR, TEXT("c.hkl_hh: Global hotkeys have been disabled."));
        }
    }
    else
    {
        //DebugMsg(DM_ERROR, TEXT("c.hkl_hh: Hotkey not listed."));
    }
    LeaveCriticalSection(&_csHotkey);
}

LRESULT CTray::_UnregisterHotkey(HWND hwnd, int i)
{
    //TraceMsg(TF_TRAY, "c.t_uh: Unregistering hotkey (%d).", i);

    if (!UnregisterHotKey(hwnd, i))
    {
        //DebugMsg(DM_ERROR, TEXT("c.t_rh: Unable to unregister hotkey %d."), i);
    }
    return TRUE;
}

// Add hotkey to the shell's list of global hotkeys.
LRESULT CTray::_ShortcutRegisterHotkey(HWND hwnd, WORD wHotkey, ATOM atom)
{
    int i;
    LPITEMIDLIST pidl;
    TCHAR szPath[MAX_PATH];
    ASSERT(atom);

    if (GlobalGetAtomName(atom, szPath, MAX_PATH))
    {
        //TraceMsg(TF_TRAY, "c.t_srh: Hotkey %d for %s", wHotkey, szPath);

        pidl = ILCreateFromPath(szPath);
        if (pidl)
        {
            i = _HotkeyAddCached(_MapHotkeyToGlobalHotkey(wHotkey), pidl);
            if (i != -1)
            {
                _RegisterHotkey(_hwnd, i);
            }
        }
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

// Remove hotkey from shell's list.
LRESULT CTray::_ShortcutUnregisterHotkey(HWND hwnd, WORD wHotkey)
{
    // //DebugMsg(DM_TRACE, "c.t_suh: Hotkey %d", wHotkey);
    int i = _HotkeyRemove(wHotkey);
    if (i == -1)
        i = _HotkeyRemoveCached(_MapHotkeyToGlobalHotkey(wHotkey));

    if (i != -1)
        _UnregisterHotkey(hwnd, i);

    return TRUE;
}

LRESULT CTray::_RegisterHotkey(HWND hwnd, int i)
{
    HOTKEYITEM* phki;
    WORD wGHotkey = 0;

    ASSERT(IS_VALID_HANDLE(_hdsaHKI, DSA));

    //TraceMsg(TF_TRAY, "c.t_rh: Registering hotkey (%d).", i);

    EnterCriticalSection(&_csHotkey);

    phki = (HOTKEYITEM*)DSA_GetItemPtr(_hdsaHKI, i);
    ASSERT(phki);
    if (phki)
    {
        wGHotkey = phki->wGHotkey;
    }

    LeaveCriticalSection(&_csHotkey);

    if (wGHotkey)
    {
        // Is the hotkey available?
        if (RegisterHotKey(hwnd, i, HIBYTE(wGHotkey), LOBYTE(wGHotkey)))
        {
            // Yes.
            return TRUE;
        }
        else
        {
            // Delete any cached items that might be using this
            // hotkey.
            int iCached = _HotkeyRemoveCached(wGHotkey);
            ASSERT(iCached != i);
            if (iCached != -1)
            {
                // Free up the hotkey for us.
                _UnregisterHotkey(hwnd, iCached);
                // Yep, nuked the cached item. Try again.
                if (RegisterHotKey(hwnd, i, HIBYTE(wGHotkey), LOBYTE(wGHotkey)))
                {
                    return TRUE;
                }
            }
        }

        // Can't set hotkey for this item.
        //DebugMsg(DM_ERROR, TEXT("c.t_rh: Unable to register hotkey %d."), i);
        // Null out this item.
        phki->wGHotkey = 0;
        phki->pidlFolder = NULL;
        phki->pidlItem = NULL;
    }
    else
    {
        //DebugMsg(DM_ERROR, TEXT("c.t_rh: Hotkey item is invalid."));
    }
    return FALSE;
}


#define GetABDHWnd(pabd)   ((HWND)ULongToPtr((pabd)->dwWnd))

void CTray::_AppBarGetTaskBarPos(PTRAYAPPBARDATA ptabd)
{
    APPBARDATA3264* pabd;

    pabd = (APPBARDATA3264*)SHLockShared(UlongToPtr(ptabd->hSharedABD), ptabd->dwProcId);
    if (pabd)
    {
        pabd->rc = _arStuckRects[_uStuckPlace];
        pabd->uEdge = _uStuckPlace;     // compat: new to ie4
        SHUnlockShared(pabd);
    }
}

void CTray::_NukeAppBar(int i)
{
    LocalFree(DPA_GetPtr(_hdpaAppBars, i));
    DPA_DeletePtr(_hdpaAppBars, i);
}

void CTray::_AppBarRemove(PTRAYAPPBARDATA ptabd)
{
    int i;

    if (!_hdpaAppBars)
        return;

    i = DPA_GetPtrCount(_hdpaAppBars);

    while (i--)
    {
        PAPPBAR pab = (PAPPBAR)DPA_GetPtr(_hdpaAppBars, i);

        if (GetABDHWnd(&ptabd->abd) == pab->hwnd)
        {
            RECT rcNuke = pab->rc;

            _NukeAppBar(i);
            _StuckAppChange(GetABDHWnd(&ptabd->abd), &rcNuke, NULL, FALSE);
        }
    }
}

PAPPBAR CTray::_FindAppBar(HWND hwnd)
{
    if (_hdpaAppBars)
    {
        int i = DPA_GetPtrCount(_hdpaAppBars);

        while (i--)
        {
            PAPPBAR pab = (PAPPBAR)DPA_GetPtr(_hdpaAppBars, i);

            if (hwnd == pab->hwnd)
                return pab;
        }
    }

    return NULL;
}

void CTray::_AppBarNotifyAll(HMONITOR hmon, UINT uMsg, HWND hwndExclude, LPARAM lParam)
{
    if (!_hdpaAppBars)
        return;

    int i = DPA_GetPtrCount(_hdpaAppBars);

    while (i--)
    {
        PAPPBAR pab = (PAPPBAR)DPA_GetPtr(_hdpaAppBars, i);

        // We need to check pab here as an appbar can delete other
        // appbars on the callback.
        if (pab && (hwndExclude != pab->hwnd))
        {
            if (!IsWindow(pab->hwnd))
            {
                _NukeAppBar(i);
                continue;
            }

            //
            // if a monitor was specified only tell appbars on that display
            //
            if (hmon &&
                (hmon != MonitorFromWindow(pab->hwnd, MONITOR_DEFAULTTONULL)))
            {
                continue;
            }

            PostMessage(pab->hwnd, pab->uCallbackMessage, uMsg, lParam);
        }
    }
}

BOOL CTray::_AppBarNew(PTRAYAPPBARDATA ptabd)
{
    PAPPBAR pab;
    if (!_hdpaAppBars)
    {
        _hdpaAppBars = DPA_Create(4);
        if (!_hdpaAppBars)
            return FALSE;

    }
    else if (_FindAppBar(GetABDHWnd(&ptabd->abd)))
    {
        // already have this hwnd
        return FALSE;
    }

    pab = (PAPPBAR)LocalAlloc(LPTR, sizeof(APPBAR));
    if (!pab)
        return FALSE;

    pab->hwnd = GetABDHWnd(&ptabd->abd);
    pab->uCallbackMessage = ptabd->abd.uCallbackMessage;
    pab->uEdge = (UINT)-1;

    if (DPA_AppendPtr(_hdpaAppBars, pab) == -1)
    {
        // insertion failed
        LocalFree(pab);
        return FALSE;
    }

    return TRUE;
}


BOOL CTray::_AppBarOutsideOf(PAPPBAR pabReq, PAPPBAR pab)
{
    if (pabReq->uEdge == pab->uEdge)
    {
        switch (pab->uEdge)
        {
            case ABE_RIGHT:
                return (pab->rc.right >= pabReq->rc.right);

            case ABE_BOTTOM:
                return (pab->rc.bottom >= pabReq->rc.bottom);

            case ABE_TOP:
                return (pab->rc.top <= pabReq->rc.top);

            case ABE_LEFT:
                return (pab->rc.left <= pabReq->rc.left);
        }
    }
    return FALSE;
}

void CTray::_AppBarQueryPos(PTRAYAPPBARDATA ptabd)
{
    int i;
    PAPPBAR pabReq = _FindAppBar(GetABDHWnd(&ptabd->abd));

    if (pabReq)
    {
        APPBARDATA3264* pabd;

        pabd = (APPBARDATA3264*)SHLockShared(UlongToPtr(ptabd->hSharedABD), ptabd->dwProcId);
        if (pabd)
        {
            HMONITOR hmon;

            pabd->rc = ptabd->abd.rc;

            //
            // default to the primary display for this call because old appbars
            // sometimes pass a huge rect and let us pare it down.  if they do
            // something like that they don't support multiple displays anyway
            // so just put them on the primary display...
            //
            hmon = MonitorFromRect(&pabd->rc, MONITOR_DEFAULTTOPRIMARY);

            //
            // always subtract off the tray if it's on the same display
            //
            if (!_uAutoHide && (hmon == _hmonStuck))
            {
                APPBAR ab;

                ab.uEdge = _GetDockedRect(&ab.rc, FALSE);
                _AppBarSubtractRect(&ab, &pabd->rc);
            }

            i = DPA_GetPtrCount(_hdpaAppBars);

            while (i--)
            {
                PAPPBAR pab = (PAPPBAR)DPA_GetPtr(_hdpaAppBars, i);

                //
                // give top and bottom preference
                // ||
                // if we're not changing edges,
                //          subtract anything currently on the outside of us
                // ||
                // if we are changing sides,
                //          subtract off everything on the new side.
                //
                // of course ignore appbars which are not on the same display...
                //
                if ((((pabReq->hwnd != pab->hwnd) &&
                    STUCK_HORIZONTAL(pab->uEdge) &&
                    !STUCK_HORIZONTAL(ptabd->abd.uEdge)) ||
                    ((pabReq->hwnd != pab->hwnd) &&
                        (pabReq->uEdge == ptabd->abd.uEdge) &&
                        _AppBarOutsideOf(pabReq, pab)) ||
                    ((pabReq->hwnd != pab->hwnd) &&
                        (pabReq->uEdge != ptabd->abd.uEdge) &&
                        (pab->uEdge == ptabd->abd.uEdge))) &&
                    (hmon == MonitorFromRect(&pab->rc, MONITOR_DEFAULTTONULL)))
                {
                    _AppBarSubtractRect(pab, &pabd->rc);
                }
            }
            SHUnlockShared(pabd);
        }
    }
}

void CTray::_AppBarSetPos(PTRAYAPPBARDATA ptabd)
{
    PAPPBAR pab = _FindAppBar(GetABDHWnd(&ptabd->abd));

    if (pab)
    {
        RECT rcOld;
        APPBARDATA3264* pabd;
        BOOL fChanged = FALSE;

        _AppBarQueryPos(ptabd);

        pabd = (APPBARDATA3264*)SHLockShared(UlongToPtr(ptabd->hSharedABD), ptabd->dwProcId);
        if (pabd)
        {
            if (!EqualRect(&pab->rc, &pabd->rc)) {
                rcOld = pab->rc;
                pab->rc = pabd->rc;
                pab->uEdge = ptabd->abd.uEdge;
                fChanged = TRUE;
            }
            SHUnlockShared(pabd);
        }

        if (fChanged)
            _StuckAppChange(GetABDHWnd(&ptabd->abd), &rcOld, &pab->rc, FALSE);
    }
}

//
// FEATURE: need to get rid of this array-based implementation to allow autohide
// appbars on secondary display (or a/h tray on 2nd with a/h appbar on primary)
// change it to an _AppBarFindAutoHideBar that keeps flags on the appbardata...
//
HWND CTray::_AppBarGetAutoHideBar(UINT uEdge)
{
    if (uEdge >= ABE_MAX)
        return FALSE;
    else
    {
        HWND hwndAutoHide = _aHwndAutoHide[uEdge];
        if (!IsWindow(hwndAutoHide))
        {
            _aHwndAutoHide[uEdge] = NULL;
        }
        return _aHwndAutoHide[uEdge];
    }
}

BOOL CTray::_AppBarSetAutoHideBar2(HWND hwnd, BOOL fAutoHide, UINT uEdge)
{
    HWND hwndAutoHide = _aHwndAutoHide[uEdge];
    if (!IsWindow(hwndAutoHide))
    {
        _aHwndAutoHide[uEdge] = NULL;
    }

    if (fAutoHide)
    {
        // register
        if (!_aHwndAutoHide[uEdge])
        {
            _aHwndAutoHide[uEdge] = hwnd;
        }

        return _aHwndAutoHide[uEdge] == hwnd;
    }
    else
    {
        // unregister
        if (_aHwndAutoHide[uEdge] == hwnd)
        {
            _aHwndAutoHide[uEdge] = NULL;
        }

        return TRUE;
    }
}

BOOL CTray::_AppBarSetAutoHideBar(PTRAYAPPBARDATA ptabd)
{
    UINT uEdge = ptabd->abd.uEdge;
    if (uEdge >= ABE_MAX)
        return FALSE;
    else {
        return _AppBarSetAutoHideBar2(GetABDHWnd(&ptabd->abd), BOOLFROMPTR(ptabd->abd.lParam), uEdge);
    }
}

void CTray::_AppBarActivationChange2(HWND hwnd, UINT uEdge)
{
    //
    // FEATURE: make this multi-monitor cool
    //
    HWND hwndAutoHide = _AppBarGetAutoHideBar(uEdge);

    if (hwndAutoHide && (hwndAutoHide != hwnd))
    {
        //
        // the _AppBar got this notification inside a SendMessage from USER
        // and is now in a SendMessage to us.  don't try to do a SetWindowPos
        // right now...
        //
        PostMessage(_hwnd, TM_BRINGTOTOP, (WPARAM)hwndAutoHide, uEdge);
    }
}

void CTray::_AppBarActivationChange(PTRAYAPPBARDATA ptabd)
{
    PAPPBAR pab = _FindAppBar(GetABDHWnd(&ptabd->abd));
    if (pab)
    {
        // if this is an autohide bar and they're claiming to be on an edge not the same as their autohide edge,
        // we don't do any activation of other autohides
        for (UINT i = 0; i < ABE_MAX; i++)
        {
            if (_aHwndAutoHide[i] == GetABDHWnd(&ptabd->abd) &&
                i != pab->uEdge)
                return;
        }
        _AppBarActivationChange2(GetABDHWnd(&ptabd->abd), pab->uEdge);
    }
}

LRESULT CTray::_OnAppBarMessage(PCOPYDATASTRUCT pcds)
{
    PTRAYAPPBARDATA ptabd = (PTRAYAPPBARDATA)pcds->lpData;

    ASSERT(pcds->cbData == sizeof(TRAYAPPBARDATA));
    ASSERT(ptabd->abd.cbSize == sizeof(APPBARDATA3264));

    switch (ptabd->dwMessage) {
        case ABM_NEW:
            return _AppBarNew(ptabd);

        case ABM_REMOVE:
            _AppBarRemove(ptabd);
            break;

        case ABM_QUERYPOS:
            _AppBarQueryPos(ptabd);
            break;

        case ABM_SETPOS:
            _AppBarSetPos(ptabd);
            break;

        case ABM_GETSTATE:
            return _desktray.AppBarGetState();

        case ABM_SETSTATE:
            _AppBarSetState((UINT)ptabd->abd.lParam);
            break;

        case ABM_GETTASKBARPOS:
            _AppBarGetTaskBarPos(ptabd);
            break;

        case ABM_WINDOWPOSCHANGED:
        case ABM_ACTIVATE:
            _AppBarActivationChange(ptabd);
            break;

        case ABM_GETAUTOHIDEBAR:
            return (LRESULT)_AppBarGetAutoHideBar(ptabd->abd.uEdge);

        case ABM_SETAUTOHIDEBAR:
            return _AppBarSetAutoHideBar(ptabd);

        default:
            return FALSE;
    }

    return TRUE;

}

// EA486701-7F92-11cf-9E05-444553540000
const GUID CLSID_HIJACKINPROC = { 0xEA486701, 0x7F92, 0x11cf, 0x9E, 0x05, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00 };

HRESULT CTray::_LoadInProc(PCOPYDATASTRUCT pcds)
{
    ASSERT(pcds->cbData == sizeof(LOADINPROCDATA));

    PLOADINPROCDATA plipd = (PLOADINPROCDATA)pcds->lpData;

    // Hack to allow us to kill W95 shell extensions that do reall hacky things that
    // we can not support.    In this case Hijack pro
    if (IsEqualIID(plipd->clsid, CLSID_HIJACKINPROC))
    {
        return E_FAIL;
    }

    if (!_pSysTray)
    {
        return E_FAIL;
    }
    HRESULT hr = _pSysTray->Exec(&plipd->clsid, 2, plipd->dwFlags, nullptr, nullptr);
    return hr;
    //return _ssomgr.EnableObject(&plipd->clsid, plipd->dwFlags);
}

BOOL CTray::GlassEnabled()
{
    return _fGlassEnabled;
}

void CTray::EnableGlass(BOOL bEnable)
{
    if (_fGlassEnabled != bEnable)
    {
        _fGlassEnabled = bEnable;
        _RegisterForGlass();
        _OnThemeChanged();
    }
}

void CTray::_RegisterForGlass()
{
    DWM_BLURBEHIND blurBehind;
    blurBehind.dwFlags = DWM_BB_ENABLE | DWM_BB_BLURREGION | DWM_BB_TRANSITIONONMAXIMIZED;
    blurBehind.fEnable = _fGlassEnabled && IsCompositionActive();
    blurBehind.hRgnBlur = nullptr;
    blurBehind.fTransitionOnMaximized = TRUE;
    DwmEnableBlurBehindWindow(_hwnd, &blurBehind);
}

const WCHAR c_szTaskbarTheme[] = L"TaskBar";
const WCHAR c_szTaskbarCompositedTheme[] = L"TaskBarComposited";
const WCHAR c_szTaskbarThemeVert[] = L"TaskBarVert";
const WCHAR c_szTaskbarCompositedThemeVert[] = L"TaskBarVertComposited";

void CTray::_OpenTaskbarThemeData()
{
    if (IsCompositionActive() && _fGlassEnabled)
    {
        WCHAR szClassList[32];
        StringCchPrintfW(szClassList, ARRAYSIZE(szClassList), L"%s::%s", c_szTaskbarCompositedTheme, c_szTaskbarTheme);
        _hTheme = OpenThemeData(_hwnd, szClassList);
        wprintf(L"szClassList: %s\n", szClassList);
    }
    else
    {
        _hTheme = OpenThemeData(_hwnd, c_szTaskbarTheme);
        wprintf(L"szClassList: TaskBar\n");
    }
}

void CTray::_SetBandSiteTheme()
{
    LPCWSTR pszTheme;

    BOOL fComposited = IsCompositionActive() && _fGlassEnabled;

    if (STUCK_HORIZONTAL(_uStuckPlace))
    {
        pszTheme = fComposited ? c_szTaskbarCompositedTheme : c_szTaskbarTheme;
    }
    else
    {
        pszTheme = fComposited ? c_szTaskbarCompositedThemeVert : c_szTaskbarThemeVert;
    }
    BandSite_SetWindowTheme(_ptbs, (LPWSTR)pszTheme);
}

void CTray::_SetRebarTheme()
{
    if (_hTheme)
    {
        LPCWSTR pszTheme = (IsCompositionActive() && (_fGlassEnabled)) ? c_szTaskbarCompositedTheme : c_szTaskbarTheme;
        SendMessageW(_hwndRebar, RB_SETWINDOWTHEME, 0, (LPARAM)pszTheme);
    }
}

void CTray::_DrawBackupStartButton(const HDC hdc)
{
    if (_hTheme)
    {
        if (_stb._hTheme)
        {
            RECT rc;
            GetWindowRect(_stb._hwndStart, &rc);
            MapWindowPoints(0, _hwnd, (LPPOINT)&rc, 2u);
            InflateRect(&rc, RECTWIDTH(rc) / -20, RECTHEIGHT(rc) / -20);
            DrawThemeBackground(_stb._hTheme, hdc, 1, 1, &rc, 0);
        }
    }
    else
    {
        SendMessage(_stb._hwndStart, 0x318u, (WPARAM)hdc, 4);
    }
}

// @NOTE (allison): Cleanup/simplifiy
BOOL CTray::_ShouldSubclassForSizingBar()
{
    if (!_hTheme)
        return FALSE;

    BOOL bRet = TRUE;

    if (!_fCanSizeMove && (_uAutoHide & AH_ON) == 0)
        return FALSE;

    return bRet;
}

void CTray::_AccountAllBandsForTaskbarSizingBar()
{
    BOOL fShouldSubclass = _ShouldSubclassForSizingBar();

    if (GetWindowThreadProcessId(_hwnd, NULL) == GetCurrentThreadId())
    {
        if (_ptbs)
            BandSite_AccountAllBandsForTaskbarSizingBar(_ptbs, fShouldSubclass);
    }
    else
    {
        PostMessage(_hwnd, 0x5B6, fShouldSubclass, NULL);
    }
}

void CTray::_OnThemeChanged()
{
    if (_hTheme)
    {
        CloseThemeData(_hTheme);
        _hTheme = nullptr;
    }

    _OpenTaskbarThemeData();
    _SetBandSiteTheme();
    _SetRebarTheme();

    _iPaddedBorderWidth = 0;
    if (!_hTheme)
    {
        NONCLIENTMETRICSW ncm = { sizeof(ncm) };
        if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, 0, &ncm, FALSE))
            _iPaddedBorderWidth = ncm.iPaddedBorderWidth;
    }

    _UpdateVertical(_uStuckPlace, TRUE);
    SetWindowStyle(_hwnd, WS_THICKFRAME | WS_BORDER, !_hTheme);
    SetWindowPos(_hwnd, nullptr, 0, 0, 0, 0, SWP_DRAWFRAME | SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE);
    InvalidateRect(_hwnd, nullptr, TRUE);
    PostMessage(_hwnd, SBM_REBUILDMENU, 0, 0);
    _AccountAllBandsForTaskbarSizingBar();
}

// Allow the trays global hotkeys to be disabled for a while.
LRESULT CTray::_SetHotkeyEnable(HWND hwnd, BOOL fEnable)
{
    _fGlobalHotkeyDisable = !fEnable;
    return TRUE;
}

BOOL IsPosInHwnd(LPARAM lParam, HWND hwnd)
{
    RECT r1;
    POINT pt;

    pt.x = GET_X_LPARAM(lParam);
    pt.y = GET_Y_LPARAM(lParam);
    GetWindowRect(hwnd, &r1);
    return PtInRect(&r1, pt);
}

void CTray::_HandleWindowPosChanging(LPWINDOWPOS lpwp)
{
    //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hwpc"));

    if (_uMoveStuckPlace != (UINT)-1)
    {
        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hwpc handling pending move"));
        _DoneMoving(lpwp);
    }
    else if (_fSysSizing || !_fSelfSizing)
    {
        RECT rc;

        if (_fSysSizing)
        {
            GetWindowRect(_hwnd, &rc);
            if (!(lpwp->flags & SWP_NOMOVE))
            {
                rc.left = lpwp->x;
                rc.top = lpwp->y;
            }
            if (!(lpwp->flags & SWP_NOSIZE))
            {
                rc.right = rc.left + lpwp->cx;
                rc.bottom = rc.top + lpwp->cy;
            }

            //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hwpc sys sizing to rect {%d, %d, %d, %d}"), rc.left, rc.top, rc.right, rc.bottom);

            _uStuckPlace = _RecalcStuckPos(&rc);
            _UpdateVertical(_uStuckPlace);
        }

        _GetDockedRect(&rc, _fSysSizing);

        //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.t_hwpc using rect {%d, %d, %d, %d}"), rc.left, rc.top, rc.right, rc.bottom);

        lpwp->x = rc.left;
        lpwp->y = rc.top;
        lpwp->cx = RECTWIDTH(rc);
        lpwp->cy = RECTHEIGHT(rc);

        lpwp->flags &= ~(SWP_NOMOVE | SWP_NOSIZE);
    }

    lpwp->flags |= SWP_FRAMECHANGED;
}

void CTray::_HandlePowerStatus(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    BOOL fResetDisplay = FALSE;

    //
    // always reset the display when the machine wakes up from a
    // suspend.  NOTE: we don't need this for a standby suspend.
    //
    // a critical resume does not generate a WM_POWERBROADCAST
    // to windows for some reason, but it does generate an old
    // WM_POWER message.
    //
    switch (uMsg)
    {
        case WM_POWER:
            fResetDisplay = (wParam == PWR_CRITICALRESUME);
            break;

        case WM_POWERBROADCAST:
            switch (wParam)
            {
                case PBT_APMRESUMECRITICAL:
                    fResetDisplay = TRUE;
                    break;
            }
            break;
    }

    if (fResetDisplay)
        ChangeDisplaySettings(NULL, CDS_RESET);
}

//////////////////////////////////////////////////////

//
// This function checks whether we need to run the cleaner
// We will not run if user is guest, user has forced us not to, or if the requisite
// number of days have not yet elapsed
//
// We execute a great deal of code to decide whether to run or not that logically should be
// in fldrclnr.dll, but we execute it here so that we don't have to load fldrclnr.dll unless
// we absolutely have to, since we execute this path on every logon of explorer.exe
//

#define REGSTR_PATH_CLEANUPWIZ            REGSTR_PATH_EXPLORER TEXT("\\Desktop\\CleanupWiz")
#define REGSTR_OEM_PATH                   REGSTR_PATH_SETUP TEXT("\\OemStartMenuData")
#define REGSTR_VAL_TIME                   TEXT("Last used time")
#define REGSTR_VAL_DELTA_DAYS             TEXT("Days between clean up")
#define REGSTR_VAL_DONTRUN                TEXT("NoRun")

//////////////////////////////////////////////////////

//
// Try tacking a 1, 2, 3 or whatever on to a file or
// directory name until it is unique.  When on a file,
// stick it before the extension.
//
void _MakeBetterUniqueName(LPTSTR pszPathName, int cchPathName)
{
    TCHAR   szNewPath[MAX_PATH];
    int     i = 1;

    if (PathIsDirectory(pszPathName))
    {
        do
        {
            StringCchPrintf(szNewPath, ARRAYSIZE(szNewPath), TEXT("%s%d"), pszPathName, i++);
        } while (-1 != GetFileAttributes(szNewPath));

        StringCchCopy(pszPathName, cchPathName, szNewPath);
    }
    else
    {
        TCHAR   szExt[MAX_PATH];
        LPTSTR  pszExt;

        pszExt = PathFindExtension(pszPathName);
        if (pszExt)
        {
            StringCchCopy(szExt, ARRAYSIZE(szExt), pszExt);
            *pszExt = 0;

            do
            {
                StringCchPrintf(szNewPath, ARRAYSIZE(szNewPath), TEXT("%s%d%s"), pszPathName, i++, szExt);
            } while (-1 != GetFileAttributes(szNewPath));

            StringCchCopy(pszPathName, cchPathName, szNewPath);
        }
    }
}

BOOL_PTR WINAPI CTray::RogueProgramFileDlgProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
    TCHAR   szBuffer[MAX_PATH * 2];
    TCHAR   szBuffer2[MAX_PATH * 2];
    static TCHAR   szBetterPath[MAX_PATH];
    static TCHAR* pszPath = NULL;

    switch (iMsg)
    {
        case WM_INITDIALOG:
            pszPath = (TCHAR*)lParam;
            StringCchCopy(szBetterPath, ARRAYSIZE(szBetterPath), pszPath);
            _MakeBetterUniqueName(szBetterPath, ARRAYSIZE(szBetterPath));
            SendDlgItemMessage(hWnd, IDC_MSG, WM_GETTEXT, (WPARAM)(MAX_PATH * 2), (LPARAM)szBuffer);
            StringCchPrintf(szBuffer2, ARRAYSIZE(szBuffer2), szBuffer, pszPath, szBetterPath);
            SendDlgItemMessage(hWnd, IDC_MSG, WM_SETTEXT, (WPARAM)0, (LPARAM)szBuffer2);

            return TRUE;

        case WM_COMMAND:
            switch (LOWORD(wParam))
            {
                case IDC_RENAME:
                    //rename and fall through
                    if (pszPath)
                    {
                        MoveFile(pszPath, szBetterPath);
                    }
                    EndDialog(hWnd, IDC_RENAME);
                    return TRUE;

                case IDIGNORE:
                    EndDialog(hWnd, IDIGNORE);
                    return TRUE;
            }
            break;
    }
    return FALSE;
}

//
// Check to see if there are any files or folders that could interfere
// with the fact that Program Files has a space in it.
//
// An example would be a directory called: "C:\Program" or a file called"C:\Program.exe".
//
// This can prevent apps that dont quote strings in the registry or call CreateProcess with
// unquoted strings from working properly since CreateProcess wont know what the real exe is.
//
void CTray::_CheckForRogueProgramFile()
{
    TCHAR szProgramFilesPath[MAX_PATH];
    TCHAR szProgramFilesShortName[MAX_PATH];

    if (SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_ADMINS)
        && S_OK == SHGetFolderPath(NULL, CSIDL_PROGRAM_FILES, NULL, SHGFP_TYPE_CURRENT, szProgramFilesPath))
    {
        LPTSTR pszRoguePattern;
        size_t cchRoguePattern;

        pszRoguePattern = StrChr(szProgramFilesPath, TEXT(' '));
        cchRoguePattern = ARRAYSIZE(szProgramFilesPath) - (pszRoguePattern - szProgramFilesPath);

        if (pszRoguePattern)
        {
            HANDLE  hFind;
            WIN32_FIND_DATA wfd;

            // Remember short name for folder name comparison below
            *pszRoguePattern = TEXT('\0');
            StringCchCopy(szProgramFilesShortName, ARRAYSIZE(szProgramFilesShortName), szProgramFilesPath);

            // turn "C:\program files" into "C:\program.*"
            StringCchCopy(pszRoguePattern, cchRoguePattern, TEXT(".*"));
            pszRoguePattern = szProgramFilesPath;

            hFind = FindFirstFile(pszRoguePattern, &wfd);

            while (hFind != INVALID_HANDLE_VALUE)
            {
                int iRet = 0;
                TCHAR szRogueFileName[MAX_PATH];

                // we found a file (eg "c:\Program.txt")
                StringCchCopy(szRogueFileName, ARRAYSIZE(szRogueFileName), pszRoguePattern);
                PathRemoveFileSpec(szRogueFileName);
                StringCchCat(szRogueFileName, ARRAYSIZE(szRogueFileName), wfd.cFileName);

                // don't worry about folders unless they are called "Program"
                if (!((wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) &&
                    lstrcmpi(szProgramFilesShortName, szRogueFileName) != 0))
                {

                    iRet = SHMessageBoxCheckExW(GetDesktopWindow(),
                        g_hinstCabinet,
                        MAKEINTRESOURCE(DLG_PROGRAMFILECONFLICT),
                        RogueProgramFileDlgProc,
                        (void*)szRogueFileName,
                        IDIGNORE,
                        TEXT("RogueProgramName"));
                }

                if ((iRet == IDIGNORE) || !FindNextFile(hFind, &wfd))
                {
                    // user hit ignore or we are done, so don't keep going
                    break;
                }
            }

            if (hFind != INVALID_HANDLE_VALUE)
            {
                FindClose(hFind);
            }
        }
    }
}

#define NM_STARTWAIT            (NM_FIRST-9)
#define NM_ENDWAIT              (NM_FIRST-10)
#define NM_BTNCLK               (NM_FIRST-11)

void CTray::_OnWaitCursorNotify(LPNMHDR pnm)
{
    _iWaitCount += (pnm->code == NM_STARTWAIT ? 1 : -1);
    ASSERT(_iWaitCount >= 0);
    // Don't let it go negative or we'll never get rid of it.
    if (_iWaitCount < 0)
        _iWaitCount = 0;
    // what we really want is for user to simulate a mouse move/setcursor
    SetCursor(LoadCursor(NULL, _iWaitCount ? IDC_APPSTARTING : IDC_ARROW));
}

// EXEX-VISTA: Almost the same in Vista, but a new case using a COM object was added.
// Look into this later.
void CTray::_HandlePrivateCommand(LPARAM lParam)
{
    LPSTR psz = (LPSTR)lParam; // lParam always ansi.

    if (!lstrcmpiA(psz, "ToggleDesktop"))
    {
        _RaiseDesktop(!g_fDesktopRaised, TRUE);
    }
    else if (!lstrcmpiA(psz, "Explorer"))
    {
        // Fast way to bring up explorer window on root of
        // windows dir.
        SHELLEXECUTEINFO shei = { 0 };
        TCHAR szPath[MAX_PATH];

        if (GetWindowsDirectory(szPath, ARRAYSIZE(szPath)) != 0)
        {
            PathStripToRoot(szPath);

            shei.lpIDList = ILCreateFromPath(szPath);
            if (shei.lpIDList)
            {
                shei.cbSize = sizeof(shei);
                shei.fMask = SEE_MASK_IDLIST;
                shei.nShow = SW_SHOWNORMAL;
                shei.lpVerb = TEXT("explore");
                ShellExecuteEx(&shei);
                ILFree((LPITEMIDLIST)shei.lpIDList);
            }
        }
    }

    LocalFree(psz);
}

#define DTM_UIACTIVATEIO            (WM_USER + 88)
#define DTM_ONFOCUSCHANGEIS         (WM_USER + 89)

//***
//
void CTray::_OnFocusMsg(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
    BOOL fActivate = (BOOL)wParam;

    switch (uMsg)
    {
        case TM_UIACTIVATEIO:
        {
            #ifdef DEBUG
            {
                int dtb = (int)lParam;

                //TraceMsg(DM_FOCUS, "tiois: TM_UIActIO fAct=%d dtb=%d", fActivate, dtb);

                ASSERT(dtb == 1 || dtb == -1);
            }
            #endif

            if (fActivate)
            {
                // Since we are tabbing into the tray, turn the focus rect on.
                SendMessage(_hwnd, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR,
                    UISF_HIDEFOCUS), 0);

                SendMessage(v_hwndDesktop, DTM_ONFOCUSCHANGEIS, TRUE, (LPARAM)_hwnd);
                SetForegroundWindow(_hwnd);

                // fake an IUnknown_UIActivateIO(_ptbs, TRUE, &msg);
                if (GetAsyncKeyState(VK_SHIFT) < 0)
                {
                    _SetFocus(_hwndNotify);
                }
                else
                {
                    _SetFocus(_stb._hwndStart);
                }
            }
            else
            {
                Ldeact:
                IUnknown_UIActivateIO(_ptbs, FALSE, NULL);
                SetForegroundWindow(v_hwndDesktop);
            }

            break;
        }

        case TM_ONFOCUSCHANGEIS:
        {
            HWND hwnd = (HWND)lParam;

            //TraceMsg(DM_FOCUS, "tiois: TM_OnFocChgIS hwnd=%x fAct=%d", hwnd, fActivate);

            if (fActivate)
            {
                // someone else is activating, so we need to deactivate
                goto Ldeact;
            }

            break;
        }

        default:
            ASSERT(0);
            break;
    }

    return;
}

#define TSVC_NTIMER     (IDT_SERVICELAST - IDT_SERVICE0 + 1)

struct {
    #ifdef DEBUG
    UINT_PTR    idtWin;
    #endif
    TIMERPROC   pfnSvc;
} g_timerService[TSVC_NTIMER];

#define TSVC_IDToIndex(id)    ((id) - IDT_SERVICE0)
#define TSVC_IndexToID(i)     ((i) + IDT_SERVICE0)

// EXEX-VISTA: Validated. (I think... This code is very unreadable and decompiled form is honestly easier to read LOL)
int CTray::_OnTimerService(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    int i;
    UINT_PTR idt;
    TIMERPROC pfn;
    BOOL b;

    switch (uMsg) {
        case TM_SETTIMER:
            //TraceMsg(DM_UEMTRACE, "e.TM_SETTIMER: wP=0x%x lP=%x", wParam, lParam);
            ASSERT(IS_VALID_CODE_PTR(lParam, TIMERPROC));
            for (i = 0; i < TSVC_NTIMER; i++) {
                if (g_timerService[i].pfnSvc == 0) {
                    g_timerService[i].pfnSvc = (TIMERPROC)lParam;
                    idt = SetTimer(_hwnd, TSVC_IndexToID(i), (UINT)wParam, 0);
                    if (idt == 0) {
                        //TraceMsg(DM_UEMTRACE, "e.TM_SETTIMER: ST()=%d (!)", idt);
                        break;
                    }
                    ASSERT(idt == (UINT_PTR)TSVC_IndexToID(i));
                    DBEXEC(TRUE, (g_timerService[i].idtWin = idt));
                    //TraceMsg(DM_UEMTRACE, "e.TM_SETTIMER: ret=0x%x", TSVC_IndexToID(i));
                    return TSVC_IndexToID(i);   // idtWin
                }
            }
            //TraceMsg(DM_UEMTRACE, "e.TM_SETTIMER: ret=0 (!)");
            return 0;

        case TM_KILLTIMER:  // lP=idtWin
            //TraceMsg(DM_UEMTRACE, "e.TM_KILLTIMER: wP=0x%x lP=%x", wParam, lParam);
            if (EVAL(IDT_SERVICE0 <= lParam && lParam <= IDT_SERVICE0 + TSVC_NTIMER - 1)) {
                i = (int)TSVC_IDToIndex(lParam);
                if (g_timerService[i].pfnSvc) {
                    ASSERT(g_timerService[i].idtWin == (UINT)lParam);
                    b = KillTimer(_hwnd, lParam);
                    ASSERT(b);
                    g_timerService[i].pfnSvc = 0;
                    DBEXEC(TRUE, (g_timerService[i].idtWin = 0));
                    return TRUE;
                }
            }
            return 0;

        case WM_TIMER:      // wP=idtWin lP=0
            //TraceMsg(DM_UEMTRACE, "e.TM_TIMER: wP=0x%x lP=%x", wParam, lParam);
            if (EVAL(IDT_SERVICE0 <= wParam && wParam <= IDT_SERVICE0 + TSVC_NTIMER - 1)) {
                i = (int)TSVC_IDToIndex(wParam);
                pfn = g_timerService[i].pfnSvc;
                if (EVAL(IS_VALID_CODE_PTR(pfn, TIMERPROC)))
                    (*pfn)(_hwnd, WM_TIMER, wParam, GetTickCount());
            }
            return 0;
    }

    ASSERT(0);      /*NOTREACHED*/
    return 0;
}

void CTray::RealityCheck()
{
    //
    // Make sure that the tray's actual z-order position agrees with what we think
    // it is.  We need to do this because there's a recurring bug where the tray
    // gets bumped out of TOPMOST position.  (Lots of things, like a tray-owned
    // window moving itself to non-TOPMOST or a random app messing with the tray
    // window position, can cause this.)
    //
    _ResetZorder(0);
}

#define DELAY_STARTUPTROUBLESHOOT   (15 * 1000)

void CTray::LogFailedStartupApp()
{
    if (_hwnd)
    {
        PostMessage(_hwnd, TM_HANDLESTARTUPFAILED, 0, 0);
    }
    else
    {
        _fEarlyStartupFailure = TRUE;
    }
}

void WINAPI CTray::TroubleShootStartupCB(HWND hwnd, UINT uMsg, UINT_PTR idTimer, DWORD dwTime)
{
    KillTimer(hwnd, idTimer);

    if (!c_tray._fStartupTroubleshooterLaunched)
    {
        TCHAR szCmdLine[MAX_PATH];
        DWORD cb;

        c_tray._fStartupTroubleshooterLaunched = TRUE;
        cb = sizeof(szCmdLine);
        if (SHGetValue(HKEY_LOCAL_MACHINE, REGSTR_PATH_EXPLORER,
            TEXT("StartupTroubleshoot"), NULL,
            szCmdLine, &cb) == ERROR_SUCCESS)
        {
            ShellExecuteRegApp(szCmdLine, RRA_NOUI | RRA_DELETE);
        }
    }
}

void CTray::_OnHandleStartupFailed()
{
    /*
     *  Don't launch the troubleshooter until we have gone
     *  DELAY_STARTUPTROUBLESHOOT milliseconds without a startup problem.
     *  This gives time for the system to settle down before starting to
     *  annoy the user all over again.
     *
     *  (And, of course, don't launch it more than once.)
     */
    if (!_fStartupTroubleshooterLaunched)
    {
        SetTimer(_hwnd, IDT_STARTUPFAILED, DELAY_STARTUPTROUBLESHOOT, TroubleShootStartupCB);
    }
}

DEFINE_GUID(CLSID_SysTray, 0x35CEC8A3, 0x2BE6, 0x11D2, 0x87, 0x73, 0x92, 0xE2, 0x20, 0x52, 0x41, 0x53);

// EXEX-VISTA: Implementation changed since XP. Inspect later.
void CTray::_HandleDelayBootStuff()
{
    // This posted message is the last one processed by the primary
    // thread (tray thread) when we boot.  At this point we will
    // want to load the shell services (which usually create threads)
    // and resume both the background start menu thread and the fs_notfiy
    // thread.

    if (!_fHandledDelayBootStuff)
    {
        if (GetShellWindow() == nullptr)
        {
            // The desktop browser hasn't finished navigating yet.
            SetTimer(_hwnd, IDT_HANDLEDELAYBOOTSTUFF, 3 * 1000, NULL);
            return;
        }

        _fHandledDelayBootStuff = TRUE;

        if (SUCCEEDED(CoCreateInstance(CLSID_SysTray, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pSysTray))))
        {
            _pSysTray->Exec(&CGID_ShellServiceObject, SSOCMDID_OPEN, 0, nullptr, nullptr);
        }

        //if (g_dwStopWatchMode)
        //{
        //    StopWatch_StartTimed(SWID_STARTUP, TEXT("_DelayedBootStuff"), SPMODE_SHELL | SPMODE_DEBUGOUT, GetPerfTime());
        //}

        PostMessage(_hwnd, TM_SHELLSERVICEOBJECTS, 0, 0);

        BandSite_HandleDelayBootStuff(_ptbs);

        //check to see if there are any files or folders that could interfere
        //with the fact that Program Files has a space in it.  An example would
        //be a folder called "C:\Program" or a file called "C:\Program.exe"
        _CheckForRogueProgramFile();

        // Create a named event and fire it so that the services can
        // go to work, reducing contention during boot.
        _hShellReadyEvent = CreateEvent(0, TRUE, TRUE, TEXT("ShellReadyEvent"));
        if (_hShellReadyEvent)
        {
            // Set the event in case it was already created and our "create
            // signaled" parameter to CreateEvent got ignored.
            SetEvent(_hShellReadyEvent);
        }

        TBOOL(WinStationRegisterConsoleNotification(SERVERNAME_CURRENT, _hwnd, NOTIFY_FOR_THIS_SESSION));
#ifdef EXEX_DLL
        //_spTaskmanWnd.Attach(new (std::nothrow) CTaskmanWindow());
        InitializeImmersiveShell();
#endif
        //if (g_dwStopWatchMode)
        //{
        //    StopWatch_StopTimed(SWID_STARTUP, TEXT("_DelayedBootStuff"), SPMODE_SHELL | SPMODE_DEBUGOUT, GetPerfTime());
        //}
    }
}

LRESULT CTray::_OnDeviceChange(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    switch (wParam)
    {
        case DBT_CONFIGCHANGED:
            // We got an update. Refresh.
            _RefreshStartMenu();
            break;

        case DBT_QUERYCHANGECONFIG:
            //
            // change to registry settings
            //
            ChangeDisplaySettings(NULL, 0);
            break;

        case DBT_MONITORCHANGE:
            //
            // handle monitor change
            //
            HandleDisplayChange(LOWORD(lParam), HIWORD(lParam), TRUE);
            break;

        case DBT_CONFIGCHANGECANCELED:
            //
            // if the config change was canceled go back
            //
            HandleDisplayChange(0, 0, FALSE);
            break;
    }

    Mixer_DeviceChange(wParam, lParam);
    return 0;
}

//
//  The "resizable" edge of the taskbar is the edge adjacent to
//  the desktop, i.e. opposite the stuck place.
//
//  returns HTXXX if on a resizable edge, else HTBORDER
//
DWORD CTray::_PtOnResizableEdge(POINT pt, LPRECT prcClient)
{
    RECT rc;
    GetWindowRect(_hwnd, &rc);

    DWORD dwHit = HTBORDER;

    switch (_uStuckPlace)
    {
        case STICK_LEFT:    rc.left = prcClient->right; dwHit = HTRIGHT;    break;
        case STICK_TOP:     rc.top = prcClient->bottom; dwHit = HTBOTTOM;   break;
        case STICK_RIGHT:   rc.right = prcClient->left; dwHit = HTLEFT;     break;
        case STICK_BOTTOM:  rc.bottom = prcClient->top; dwHit = HTTOP;      break;
    }

    return PtInRect(&rc, pt) ? dwHit : HTBORDER;
}

//
//  _OnFactoryMessage
//
//  The OPK "factory.exe" tool sends us this message to tell us that
//  it has dorked some setting or other and we should refresh so the
//  OEM can see the effect immediately and feel confident that it
//  actually worked.  This is not technically necessary but it cuts
//  down on OEM support calls when they ask us why their setting didn't
//  work.  (It did work, they just have to log off and back on to see
//  it.)
//

int CTray::_OnFactoryMessage(WPARAM wParam, LPARAM lParam)
{
    switch (wParam)
    {
        case 0:         // FACTORY_OEMLINK:  factory.exe has dorked the OEM link
            _stb.CloseStartMenu();
            _stb.BuildStartMenu();  // Force a rebuild
            return 1;

        case 1:         // FACTORY_MFU:  factory.exe has written a new MFU
            HandleFirstTime();      // Rebuild the default MFU
            _stb.CloseStartMenu();
            _stb.BuildStartMenu();  // Force a rebuild
            return 1;
    }

    return 0;
}

#define CX_OFFSET   g_cxEdge
#define CY_OFFSET   g_cyEdge

//
// _MapNCToClient
//
// see comments in _TryForwardNCToClient
//
BOOL CTray::_MapNCToClient(LPARAM* plParam)
{
    POINT pt = { GET_X_LPARAM(*plParam), GET_Y_LPARAM(*plParam) };

    RECT rcClient;
    GetClientRect(_hwnd, &rcClient);
    MapWindowPoints(_hwnd, NULL, (LPPOINT)&rcClient, 2);

    //
    // point must be outside the client area and not on the
    // resizable edge of the taskbar
    //

    if (!PtInRect(&rcClient, pt) && _PtOnResizableEdge(pt, &rcClient) == HTBORDER)
    {
        //
        // fudge it over onto the client edge and return TRUE
        //
        if (pt.x < rcClient.left)
            pt.x = rcClient.left + CX_OFFSET;
        else if (pt.x > rcClient.right)
            pt.x = rcClient.right - CX_OFFSET;

        if (pt.y < rcClient.top)
            pt.y = rcClient.top + CY_OFFSET;
        else if (pt.y > rcClient.bottom)
            pt.y = rcClient.bottom - CY_OFFSET;

        *plParam = MAKELONG(pt.x, pt.y);
        return TRUE;
    }

    //
    // didn't pass the test.  leave the point alone and return FALSE.
    //
    return FALSE;
}

HWND _TopChildWindowFromPoint(HWND hwnd, POINT pt)
{
    HWND hwndLast = NULL;
    hwnd = ChildWindowFromPoint(hwnd, pt);
    while (hwnd && hwnd != hwndLast)
    {
        hwndLast = hwnd;
        hwnd = ChildWindowFromPoint(hwnd, pt);
    }
    return hwndLast;
}

//
// _TryForwardNCToClient
//
// Hack!  This exists to solve a usability problem.  When you slam your
// mouse into the bottom corner of the screen and click, we want that to
// activate the start button.  Similarly, when you slam your mouse below
// a Quick Launch button or task button and click, we want that to
// activate the button.
//
// We hack this by remapping the coordinate of NC mouse messages and
// manually forwarding to the appropriate window.  We only do this for
// clicks on the edges non-resizable edge of the taskbar.
//
// We also warp the mouse cursor over to the new position.  This is needed
// because e.g. if toolbar is the client window we're forwarding
// to, it will set capture and receive subsequent mouse messages.  (And
// once it gets a mouse message outside the window, it will deselect the
// button and so the button won't get activated.)
//
// _MapNCToClient has the rules for figuring out if the click is on
// one of the edges we want and for remapping the coordinate into the
// client area.
//
BOOL CTray::_TryForwardNCToClient(UINT uMsg, LPARAM lParam)
{
    if (_MapNCToClient(&lParam))
    {
        // see if this is over one of our windows
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        MapWindowPoints(NULL, _hwnd, &pt, 1);
        HWND hwnd = _TopChildWindowFromPoint(_hwnd, pt);

        if (hwnd)
        {
            // warp the mouse cursor to this screen coord
            SetCursorPos(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

            // map to window coords
            MapWindowPoints(_hwnd, hwnd, &pt, 1);

            // set lparam to window coords
            lParam = MAKELONG(pt.x, pt.y);

            // map to client message
            ASSERT(InRange(uMsg, WM_NCMOUSEFIRST, WM_NCMOUSELAST));
            uMsg += (WM_LBUTTONDOWN - WM_NCLBUTTONDOWN);

            // never send double click message
            if (uMsg == WM_LBUTTONDBLCLK)
            {
                return FALSE;
            }

            // forward it
            SendMessage(hwnd, uMsg, 0, lParam);
            return TRUE;
        }
    }
    return FALSE;
}

//  --------------------------------------------------------------------------
//  CTray::CountOfRunningPrograms
//
//  Arguments:  <none>
//
//  Returns:    DWORD
//
//  Purpose:    Iterates the window list. Looks for windows that are visible
//              with a non-zero length window title. Gets that window process
//              ID and keeps the IDs in a list. For each window iterated it
//              checks against the list to see if the process is already
//              tagged and if so doesn't add it again. Finally it returns the
//              count of the unique processes handling open visible windows
//              in the user's desktop.
//
//              The list is fixed at 1000 processes (using stack space).
//
//  History:    2000-06-29  vtan        created
//  --------------------------------------------------------------------------

static  const int   MAXIMUM_PROCESS_COUNT = 1000;

typedef struct
{
    DWORD   dwCount;
    DWORD   dwProcessIDs[MAXIMUM_PROCESS_COUNT];
} tProcessIDList;

bool    FoundProcessID(tProcessIDList* pProcessIDList, DWORD dwProcessID)

{
    bool    fFound;
    DWORD   dwIndex;

    for (fFound = false, dwIndex = 0; !fFound && (dwIndex < pProcessIDList->dwCount); ++dwIndex)
    {
        fFound = (pProcessIDList->dwProcessIDs[dwIndex] == dwProcessID);
    }
    return(fFound);
}

void    AddProcessID(tProcessIDList* pProcessIDList, DWORD dwProcessID)

{
    if (pProcessIDList->dwCount < MAXIMUM_PROCESS_COUNT)
    {
        pProcessIDList->dwProcessIDs[pProcessIDList->dwCount++] = dwProcessID;
    }
}

BOOL    CALLBACK    CountOfRunningProgramsEnumWindowsProc(HWND hwnd, LPARAM lParam)

{
    if ((GetShellWindow() != hwnd) && IsWindowVisible(hwnd))
    {
        DWORD   dwThreadID, dwProcessID;
        TCHAR   szWindowTitle[256];

        dwThreadID = GetWindowThreadProcessId(hwnd, &dwProcessID);
        if ((InternalGetWindowText(hwnd, szWindowTitle, ARRAYSIZE(szWindowTitle)) > 0) &&
            (szWindowTitle[0] != TEXT('\0')))
        {
            if (!FoundProcessID(reinterpret_cast<tProcessIDList*>(lParam), dwProcessID))
            {
                AddProcessID(reinterpret_cast<tProcessIDList*>(lParam), dwProcessID);
            }
        }
    }
    return(TRUE);
}

void CTray::GetStuckMonitorRect(RECT* prcStuck)
{
    GetMonitorRects(_hmonStuck, prcStuck, NULL);
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
BOOL CTray::IsMouseOverStartButton()    // TODO: revise
{
    BOOL bRet = FALSE;

    if ((_uAutoHide & AH_HIDING) == 0)
    {
        RECT rc;

        rc.top = _arStuckRects[_uStuckPlace].top;
        if ((_uStuckPlace & 1) != 0)
            rc.bottom = _arStuckRects[_uStuckPlace].bottom;
        else
            rc.bottom = _stb._sizeStart.cy + _arStuckRects[_uStuckPlace].top;
        if (IsBiDiLocalizedSystem())
        {
            if ((_uStuckPlace & 1) != 0)
                rc.left = _arStuckRects[_uStuckPlace].right - _stb._sizeStart.cx;
            else
                rc.left = _arStuckRects[_uStuckPlace].left;
            rc.right = _arStuckRects[_uStuckPlace].right;
        }
        else
        {
            rc.left = _arStuckRects[_uStuckPlace].left;
            if ((_uStuckPlace & 1) != 0)
                rc.right = _stb._sizeStart.cx + _arStuckRects[_uStuckPlace].left;
            else
                rc.right = _arStuckRects[_uStuckPlace].right;
        }
        if (_fCanSizeMove)
        {
            if (!_uStuckPlace)
            {
                rc.right = rc.right - _sizeSizingBar.cy;
            }
            else
            {
                if (!(_uStuckPlace - 1))
                {
                    rc.bottom -= _sizeSizingBar.cy;
                }
                else
                {
                    if (_uStuckPlace - 1 == 1)
                        rc.left = _sizeSizingBar.cy + rc.left;
                    else
                        rc.top += _sizeSizingBar.cy;
                }
            }
        }
        POINT pt;
        GetCursorPos(&pt);
        bRet = PtInRect(&rc, pt);
    }

    return bRet;
}

BOOL CTray::IsMouseOverClock()  // @TODO: Cleanup
{
    POINT v8;
    RECT Rect;

    BOOL result = 0;
    if ((_uAutoHide & 2) == 0)
    {
        GetWindowRect(_GetClockWindow(), &Rect);
        if (_fCanSizeMove)
        {
            if (_uStuckPlace)
            {
                UINT v5 = _uStuckPlace - 1;
                if (v5)
                {
                    bool v6 = v5 == 1;
                    LONG v7 = _cyClockMargin;
                    if (v6)
                        Rect.left += v7;
                    else
                        Rect.top += v7;
                }
                else
                {
                    Rect.bottom -= _cyClockMargin;
                }
            }
            else
            {
                Rect.right -= _cyClockMargin;
            }
        }
        LPARAM MessagePos = GetMessagePos();
        v8.y = GET_Y_LPARAM(MessagePos);
        v8.x = GET_X_LPARAM(MessagePos);
        return PtInRect(&Rect, v8);
    }
    return result;
}

// EXEX-VISTA: Validated. But needs clean up.
BOOL CTray::ShowClockFlyoutAsNeeded(LPARAM lParam)
{
    BOOL v7;
    POINT v9;
    RECT rc;
    RECT rcClockWnd;

    GetWindowRect(_GetClockWindow(), &rcClockWnd);
    if (lParam == -1)
    {
        v7 = 1;
    }
    else
    {
        rc = rcClockWnd;

        if (_fCanSizeMove)
        {
            switch (_uStuckPlace)
            {
                case STICK_TOP:
                {
                    rc.bottom = rcClockWnd.bottom - _cyClockMargin;
                    break;
                }

                case STICK_BOTTOM:
                {
                    rc.top = rcClockWnd.top + _cyClockMargin;
                    break;
                }

                case STICK_LEFT:
                {
                    rc.right = rcClockWnd.right - _cyClockMargin;
                    break;
                }

                case STICK_RIGHT:
                {
                    rc.left = rcClockWnd.left + _cyClockMargin;
                    break;
                }
            }
        }

        v9.y = GET_Y_LPARAM(lParam);
        v9.x = GET_X_LPARAM(lParam);
        v7 = PtInRect(&rc, v9);
    }
    if (v7)
    {
        SendMessageW(_GetClockWindow(), WM_SHOWCLOCKFLYOUT, TRUE, 0);
    }
    return v7;
}

DWORD CTray::CountOfRunningPrograms()
{
    tProcessIDList processIDList = { 0 };
    TBOOL(EnumWindows(CountOfRunningProgramsEnumWindowsProc, reinterpret_cast<LPARAM>(&processIDList)));
    return processIDList.dwCount;
}

//  --------------------------------------------------------------------------
//  CTray::_OnSessionChange
//
//  Arguments:  wParam  =   WTS_xxx notification.
//              lParam  =   WTS_SESSION_NOTIFICATION struct pointer.
//
//  Returns:    LRESULT
//
//  Purpose:    Handles console/remote dis/reconnects.
//
//  History:    2000-07-12  vtan        created
//  --------------------------------------------------------------------------

// EXEX-VISTA: SLIGHTLY MODIFIED. Validate later.
LRESULT CTray::_OnSessionChange(WPARAM wParam, LPARAM lParam)
{
    ASSERTMSG(DWORD(lParam) == NtCurrentPeb()->SessionId, "Session ID mismatch in CTray::_OnSessionChange");

    if ((WTS_CONSOLE_CONNECT == wParam) || (WTS_REMOTE_CONNECT == wParam) || (WTS_SESSION_UNLOCK == wParam))
    {
        _fIsDesktopConnected = TRUE;
    }
    else if ((WTS_CONSOLE_DISCONNECT == wParam) || (WTS_REMOTE_DISCONNECT == wParam) || (WTS_SESSION_LOCK == wParam))
    {
        _fIsDesktopConnected = FALSE;
    }

    if ((WTS_CONSOLE_CONNECT == wParam) || (WTS_REMOTE_CONNECT == wParam))
    {
        _RefreshStartMenu();
        SHUpdateRecycleBinIcon();
    }
    else if ((WTS_SESSION_LOCK == wParam) || (WTS_SESSION_UNLOCK == wParam))
    {
        if (IsOS(OS_FASTUSERSWITCHING))
        {
            if (wParam == WTS_SESSION_LOCK)
            {
                ExplorerPlaySound(TEXT("WindowsLogoff"));
            }
            else if (wParam == WTS_SESSION_UNLOCK)
            {
                ExplorerPlaySound(TEXT("WindowsLogon"));
            }
        }

        PostMessage(_hwnd, TM_WORKSTATIONLOCKED, (WTS_SESSION_LOCK == wParam), 0);
    }
    else if (WTS_SESSION_REMOTE_CONTROL == wParam)
    {
        // optimization not needed on remote sessions
        if (!GetSystemMetrics(SM_REMOTESESSION)) {
            _stb.BuildStartMenu();
        }
    }

    return 1;
}

#define SEN_FIRST       (0U-550U)       // ;Internal
#define SEN_LAST        (0U-559U)       // ;Internal

#define SEN_DDEEXECUTE (SEN_FIRST-0)

#define DTM_THREADEXIT              (WM_USER + 76)
#define DTM_SAVESTATE               (WM_USER + 77)
#define DTM_SHELLSERVICEOBJECTS     (WM_USER + 79)
#define DTM_STARTWAIT               (WM_USER + 80)
#define DTM_ENDWAIT                 (WM_USER + 81)

LRESULT CTray::v_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static  UINT uDDEExec = 0;
    LRESULT lres = 0;
    MSG msg;

    msg.hwnd = hwnd;
    msg.message = uMsg;
    msg.wParam = wParam;
    msg.lParam = lParam;

    if (_stb.TranslateMenuMessage(&msg, &lres) == S_OK)
        return lres;

    if (_pmbTasks &&
        _pmbTasks->TranslateMenuMessage(&msg, &lres) == S_OK)
        return lres;

    wParam = msg.wParam;
    lParam = msg.lParam;

    INSTRUMENT_WNDPROC(SHCNFI_TRAY_WNDPROC, hwnd, uMsg, wParam, lParam);

    switch (uMsg)
    {
        case WMTRAY_REGISTERHOTKEY:
            return _RegisterHotkey(hwnd, (int)wParam);

        case WMTRAY_UNREGISTERHOTKEY:
            return _UnregisterHotkey(hwnd, (int)wParam);

        case WMTRAY_SCREGISTERHOTKEY:
            return _ShortcutRegisterHotkey(hwnd, (WORD)wParam, (ATOM)lParam);

        case WMTRAY_SCUNREGISTERHOTKEY:
            return _ShortcutUnregisterHotkey(hwnd, (WORD)wParam);

        case WMTRAY_SETHOTKEYENABLE:
            return _SetHotkeyEnable(hwnd, (BOOL)wParam);

        // EXEX-VISTA: LAZILY MODIFIED. Redo later.
        case WMTRAY_QUERY_MENU:
            //return (LRESULT)_stb._hwndStart;
            return 0;

        case WMTRAY_QUERY_VIEW:
            return (LRESULT)_hwndTasks;

        case WMTRAY_TOGGLEQL:
            return _ToggleQL((int)lParam);

        case WM_USER + 240:
            return (LRESULT)_stb._hwndStart;

        case WM_COPYDATA:
            // Check for NULL it can happen if user runs out of selectors or memory...
            if (lParam)
            {
                switch (((PCOPYDATASTRUCT)lParam)->dwData) {
                    case TCDM_NOTIFY:
                    {
                        BOOL bRefresh = FALSE;

                        lres = _trayNotify.TrayNotify(_hwndNotify, (HWND)wParam, (PCOPYDATASTRUCT)lParam, &bRefresh);
                        if (bRefresh)
                        {
                            SizeWindows();
                        }
                        return(lres);
                    }

                    case TCDM_APPBAR:
                        return _OnAppBarMessage((PCOPYDATASTRUCT)lParam);

                    case TCDM_LOADINPROC:
                        return (UINT)_LoadInProc((PCOPYDATASTRUCT)lParam);
                }
            }
            return FALSE;

        // EXEX-VISTA: Validated.
        case WM_NCCALCSIZE:
        {
            if (_iPaddedBorderWidth > 0)
            {
                DefWindowProc(hwnd, uMsg, wParam, lParam);
                InflateRect((LPRECT)lParam, _iPaddedBorderWidth, _iPaddedBorderWidth);
                return 0;
            }

            goto DoDefault;
        }

        // EXEX-VISTA: Validated.
        case WM_NCLBUTTONDBLCLK:
        {
            // @NOTE: this is how it was done, but this doesn't solve any of the iffiness relating to the clock flyout
            if (!_TryForwardNCToClient(uMsg, lParam) && IsPosInHwnd(lParam, _hwndNotify))
            {
                ShowClockFlyoutAsNeeded(lParam);
            }

            break;
        }

        // EXEX-VISTA: Validated.
        case WM_NCLBUTTONDOWN:
        {
            if (IsMouseOverStartButton())
            {
                if (!_stb.IsButtonPushed())
                {
                    SendMessageW(_stb._hwndStart, BM_SETSTATE, 1, 0);
                    return 0;
                }

                goto DoDefault;
            }

            if (ShowClockFlyoutAsNeeded(lParam))
            {
                return 0;
            }

            break;
        }

        // EXEX-VISTA: Validated.
        case WM_NCMOUSEMOVE:
        {
            TRACKMOUSEEVENT mouseEvent;
            if (IsMouseOverClock())
            {
                if (_fMouseInTaskbar || _stb.IsButtonPushed())
                {
                    goto DoDefault;
                }

                mouseEvent.cbSize = sizeof(mouseEvent);
                mouseEvent.dwFlags = TME_HOVER | TME_LEAVE | TME_NONCLIENT;
                mouseEvent.hwndTrack = _hwnd;
                mouseEvent.dwHoverTime = GetDoubleClickTime();
            }
            else
            {
                //
                // Track so we know to draw the start button's hover state.
                //

                if (_fMouseInTaskbar || _stb.IsButtonPushed())
                {
                    goto DoDefault;
                }

                if (!_hTheme || !IsMouseOverStartButton())
                {
                    // The hover state is not used without a theme.
                    // If the mouse is not over the start button, then we skip
                    // this logic.
                    goto DoDefault;
                }

                mouseEvent.cbSize = sizeof(mouseEvent);
                mouseEvent.dwFlags = TME_LEAVE | TME_NONCLIENT;
                mouseEvent.hwndTrack = _hwnd;
                mouseEvent.dwHoverTime = NULL;
            }

            TrackMouseEvent(&mouseEvent);
            _fMouseInTaskbar = true;
            return 0;
        }

        // EXEX-VISTA: Validated.
        case WM_NCMOUSELEAVE:
        {
            if (!IsPosInHwnd(lParam, _GetClockWindow()))
            {
                SendMessageW(_GetClockWindow(), WM_SHOWCLOCKTOOLTIP, FALSE, 0);
                if (_hTheme)
                {
                    if (!_stb.IsButtonPushed() && !IsMouseOverStartButton())
                    {
                        _stb.DrawStartButton(PBS_NORMAL, true);
                    }
                }
            }
            _fMouseInTaskbar = false;
            break;
        }

        // EXEX-VISTA: Validated.
        case WM_CREATE:
            return _OnCreate(hwnd);

        // EXEX-VISTA: Validated.
        case WM_DESTROY:
            return _HandleDestroy();

            #ifdef DEBUG
        case WM_QUERYENDSESSION:
            //TraceMsg(DM_SHUTDOWN, "Tray.wp WM_QUERYENDSESSION");
            goto DoDefault;
            #endif

        // EXEX-VISTA: Validated.
        case WM_ENDSESSION:
            // save our settings if we are shutting down
            if (wParam)
            {
                if (lParam | ENDSESSION_LOGOFF)
                {
                    _fIsLogoff = TRUE;
                    _RecomputeAllWorkareas();
                }

                _SaveTrayAndDesktop();

                ShowWindow(_hwnd, SW_HIDE);
                ShowWindow(v_hwndDesktop, SW_HIDE);

                DestroyWindow(_hwnd);
            }
            break;

        // EXEX-VISTA: Validated.
        case WM_PRINTCLIENT:
        case WM_PAINT:
        {
            RECT rc;
            PAINTSTRUCT ps;
            HDC hdc = (HDC)wParam;
            if (hdc == 0)
                hdc = BeginPaint(hwnd, &ps);

            GetClientRect(hwnd, &rc);

            if (_hTheme)
            {
                RECT rcClip;
                if (GetClipBox(hdc, &rcClip) == NULLREGION)
                    rcClip = rc;

                if (IsCompositionActive())
                {
                    SHFillRectClr(hdc, &rcClip, RGB(0, 0, 0));
                }

                DrawThemeBackground(_hTheme, hdc, _GetPart(_uStuckPlace), 0, &rc, &rcClip);
            }
            else
            {
                FillRect(hdc, &rc, (HBRUSH)(COLOR_3DFACE + 1));

                // draw etched line around on either side of the bandsite
                MapWindowPoints(HWND_DESKTOP, hwnd, (LPPOINT)&rc, 2);
                InflateRect(&rc, g_cxEdge, g_cyEdge);
                DrawEdge(hdc, &rc, EDGE_ETCHED, BF_TOPLEFT);
            }

            if ((_uAutoHide & AH_ON) == 0)
            {
                // EXEX-VISTA TODO: Uncomment when implemented _DrawBackupStartButton.
                //_DrawBackupStartButton(hdc);
            }

            if (wParam == 0)
                EndPaint(hwnd, &ps);
        }
        break;
    case TM_FIREDESKTOPSWITCH:
    {
        _hShellDesktopSwitch = CreateEvent(0, TRUE, TRUE, TEXT("ShellDesktopSwitchEvent"));
        if (_hShellDesktopSwitch)
        {
            SetEvent(_hShellDesktopSwitch);
        }
        break;
    }

        // EXEX-VISTA: Validated.
        case WM_ERASEBKGND:
        {
            if (_hTheme)
            {
                if (!_fSkipErase)
                {
                    RECT rc;
                    GetClientRect(hwnd, &rc);

                    if (IsCompositionActive())
                    {
                        SHFillRectClr((HDC)wParam, &rc, RGB(0, 0, 0));
                    }

                    DrawThemeBackground(_hTheme, (HDC)wParam, _GetPart(_uStuckPlace), 0, &rc, NULL);

                    // Only draw the first time to prevent piece-mail taskbar painting
                    _fSkipErase = TRUE;
                }

                return 0;
            }
            else
            {
                goto DoDefault;
            }
            break;
        }

        // EXEX-VISTA: Validated.
        case WM_NCPAINT:
        {
            if (_hTheme)
            {
                UpdateWindow(_hwnd);
                return 0;
            }
            else
            {
                goto DoDefault;
            }
            break;
        }

        // EXEX-VISTA: Validated.
        case WM_POWER:
        case WM_POWERBROADCAST:
            _PropagateMessage(hwnd, uMsg, wParam, lParam);
            _HandlePowerStatus(uMsg, wParam, lParam);
            goto DoDefault;

        // EXEX-VISTA: Validated.
        case WM_DEVICECHANGE:
        {
            // EXEX-VISTA TODO: Uncomment when implemented CSystemMixer:
            #if 0
            if (_pSystemMixer)
            {
                _pSystemMixer->ForwardWindowMessage(uMsg, wParam, lParam);
            }
            #endif

            lres = _OnDeviceChange(hwnd, wParam, lParam);
            if (lres == 0)
            {
                goto DoDefault;
            }
            break;
        }

        case WM_NOTIFY:
        {
            NMHDR* pnm = (NMHDR*)lParam;
            if (!BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, &lres)) {
                switch (pnm->code)
                {
                    case SEN_DDEEXECUTE:
                        if (((LPNMHDR)lParam)->idFrom == 0)
                        {
                            LPNMVIEWFOLDER pnmPost = DDECreatePostNotify((LPNMVIEWFOLDER)pnm);

                            if (pnmPost)
                            {
                                PostMessage(hwnd, GetDDEExecMsg(), 0, (LPARAM)pnmPost);
                                return TRUE;
                            }
                        }
                        break;

                    case NM_STARTWAIT:
                    case NM_ENDWAIT:
                        _OnWaitCursorNotify((NMHDR*)lParam);
                        PostMessage(v_hwndDesktop, ((NMHDR*)lParam)->code == NM_STARTWAIT ? DTM_STARTWAIT : DTM_ENDWAIT,
                            0, 0); // forward it along
                        break;

                    // EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
                    case NM_THEMECHANGED:
                        SizeWindows();
                        break;

                    case TTN_NEEDTEXT:
                        //
                        //  Make the clock manage its own tooltip.
                        //
                        return SendMessage(_GetClockWindow(), WM_NOTIFY, wParam, lParam);

                    case TTN_SHOW:
                        //SetWindowZorder(_hwndTrayTips, HWND_TOP);
                        break;

                }
            }
            break;
        }

        case WM_CLOSE:
            _DoExitWindows(v_hwndDesktop, FALSE, FALSE);
            break;

        case WM_NCHITTEST:
        {
            RECT r1;
            POINT pt;

            GetClientRect(hwnd, &r1);
            MapWindowPoints(hwnd, NULL, (LPPOINT)&r1, 2);

            pt.x = GET_X_LPARAM(lParam);
            pt.y = GET_Y_LPARAM(lParam);

            _SetUnhideTimer(pt.x, pt.y);

            // If the user can't size or move the taskbar, then just say
            // they hit something useless
            if (!_fCanSizeMove)
            {
                return HTBORDER;
            }
            else if (PtInRect(&r1, pt))
            {
                // allow dragging if mouse is in client area of _hwnd
                return HTCAPTION;
            }
            else
            {
                return _PtOnResizableEdge(pt, &r1);
            }
        }
        break;

        case WM_WINDOWPOSCHANGING:
            _HandleWindowPosChanging((LPWINDOWPOS)lParam);
            break;

        case WM_ENTERSIZEMOVE:
            //DebugMsg(DM_TRAYDOCK, TEXT("Tray -- WM_ENTERSIZEMOVE"));
            g_fInSizeMove = TRUE;
            GetCursorPos((LPPOINT)&_rcSizeMoveIgnore);
            _rcSizeMoveIgnore.right = _rcSizeMoveIgnore.left;
            _rcSizeMoveIgnore.bottom = _rcSizeMoveIgnore.top;
            InflateRect(&_rcSizeMoveIgnore, GetSystemMetrics(SM_CXICON),
                GetSystemMetrics(SM_CYICON));

            //
            // unclip the tray from the current monitor.
            // keeping the tray properly clipped in the MoveSize loop is extremely
            // hairy and provides almost no benefit.  we'll re-clip when it lands.
            //
            _ClipWindow(FALSE);

            // Remember the old monitor we were on
            _hmonOld = _hmonStuck;

            // set up for WM_MOVING/WM_SIZING messages
            _uMoveStuckPlace = (UINT)-1;
            _fSysSizing = TRUE;

            if (!g_fDragFullWindows)
            {
                SendMessage(_hwndRebar, WM_SETREDRAW, FALSE, 0);
            }
            break;

        case WM_EXITSIZEMOVE:
            //DebugMsg(DM_TRAYDOCK, TEXT("Tray -- WM_EXITSIZEMOVE"));

            // done sizing
            _fSysSizing = FALSE;
            _fDeferedPosRectChange = FALSE;

            if (!g_fDragFullWindows)
            {
                SendMessage(_hwndRebar, WM_SETREDRAW, TRUE, 0);
            }

            //
            // kick the size code one last time after the loop is done.
            // NOTE: we rely on the WM_SIZE code re-clipping the tray.
            //
            PostMessage(hwnd, WM_SIZE, 0, 0L);
            g_fInSizeMove = FALSE;
            break;

        case WM_MOVING:
            _HandleMoving(wParam, (LPRECT)lParam);
            break;

        case WM_ENTERMENULOOP:
            // //DebugMsg(DM_TRACE, "c.twp: Enter menu loop.");
            _HandleEnterMenuLoop();
            break;

        case WM_EXITMENULOOP:
            // DebugMsg(DM_TRACE, "c.twp: Exit menu loop.");
            _HandleExitMenuLoop();
            break;

        case WM_TIMER:
            if (IDT_SERVICE0 <= wParam && wParam <= IDT_SERVICELAST)
                return _OnTimerService(uMsg, wParam, lParam);
            _HandleTimer(wParam);
            break;

        case WM_NCMOUSEHOVER:
            SendMessageW(_GetClockWindow(), WM_SHOWCLOCKTOOLTIP, TRUE, 0);
            break;

        case WM_SIZING:
            _HandleSizing(wParam, (LPRECT)lParam, _uStuckPlace, _fSelfSizing);
            break;

        // EXEX-VISTA: Validated.
        case WM_MOVE:
            if (_bBool679)
            {
                _stb.UpdateStartButton(false);
            }
            break;

        // EXEX-VISTA: Validated.
        case WM_SIZE:
            if (lParam && _bBool679)
            {
                _stb.UpdateStartButton(false);
            }

            _HandleSize();
            break;

        case WM_DISPLAYCHANGE:
            // NOTE: we get WM_DISPLAYCHANGE in the below two situations
            // 1. a display size changes (HMON will not change in USER)
            // 2. a display goes away or gets added (HMON will change even if
            // the monitor that went away has nothing to do with our hmonStuck)

            // In the above two situations we actually need to do different things
            // because in 1, we do not want to update our hmonStuck because we might
            // end up on another monitor, but in 2 we do want to update hmonStuck because
            // our hmon is invalid

            // The way we handle this is to call GetMonitorInfo on our old HMONITOR
            // and see if it's still valid, if not, we update it by calling _SetStuckMonitor
            // all these code is in _ScreenSizeChange;

            _ScreenSizeChange(hwnd);

            // Force the Start Pane to rebuild because a change in color depth
            // causes themes to run around destroying fonts (ruining the OOBE
            // text) and we need to reload our bitmaps for the new color depth
            // anyway.
            ::PostMessage(_hwnd, SBM_REBUILDMENU, 0, 0);

            break;

        // Don't go to default wnd proc for this one...
        // EXEX-VISTA: Validated.
        case WM_INPUTLANGCHANGEREQUEST:
            return(LRESULT)0L;

        // EXEX-VISTA: Validated.
        case WM_GETMINMAXINFO:
            ((MINMAXINFO*)lParam)->ptMinTrackSize.x = g_cxFrame;
            ((MINMAXINFO*)lParam)->ptMinTrackSize.y = g_cyFrame;
            break;

        // EXEX-VISTA: Validated. New.
        case WM_SHOWWINDOW:
        {
            _StuckTrayChange();
            break;
        }

        case WM_WININICHANGE:
            if (!_fFromStart)
            {
                if (lParam && (0 == lstrcmpi((LPCTSTR)lParam, TEXT("SaveTaskbar"))))
                {
                    _SaveTrayAndDesktop();
                }
                else
                {
                    BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, NULL);
                    _PropagateMessage(hwnd, uMsg, wParam, lParam);
                    _OnWinIniChange(hwnd, wParam, lParam);
                }

                if (lParam)
                {
                    //TraceMsg(TF_TRAY, "Tray Got: lParam=%s", (LPCSTR)lParam);
                }
            }
            break;

        // EXEX-VISTA: Validated.
        case WM_TIMECHANGE:
            _PropagateMessage(hwnd, uMsg, wParam, lParam);
            break;

        case WM_SYSCOLORCHANGE:
            if (!_fFromStart)
            {
                _OnNewSystemSizes();
                BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, nullptr);
            }
            _PropagateMessage(hwnd, uMsg, wParam, lParam);
            break;

        // EXEX-VISTA: Validated.
        case WM_SETCURSOR:
            if (_iWaitCount) {
                SetCursor(LoadCursor(NULL, IDC_APPSTARTING));
                return TRUE;
            }
            else
                goto DoDefault;

        // EXEX-VISTA: Validated.
        case WM_SETFOCUS:
            IUnknown_UIActivateIO(_ptbs, TRUE, NULL);
            break;

        // EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
        case WM_SYSCHAR:
            if (wParam == TEXT(' ')) {
                HMENU hmenu;
                int idCmd;

                SHSetWindowBits(hwnd, GWL_STYLE, WS_SYSMENU, WS_SYSMENU);
                hmenu = GetSystemMenu(hwnd, FALSE);
                if (hmenu) {
                    EnableMenuItem(hmenu, SC_RESTORE, MFS_GRAYED | MF_BYCOMMAND);
                    EnableMenuItem(hmenu, SC_MAXIMIZE, MFS_GRAYED | MF_BYCOMMAND);
                    EnableMenuItem(hmenu, SC_MINIMIZE, MFS_GRAYED | MF_BYCOMMAND);
                    EnableMenuItem(hmenu, SC_MOVE, (_fCanSizeMove ? MFS_ENABLED : MFS_GRAYED) | MF_BYCOMMAND);
                    EnableMenuItem(hmenu, SC_SIZE, (_fCanSizeMove ? MFS_ENABLED : MFS_GRAYED) | MF_BYCOMMAND);

                    idCmd = _stb.TrackMenu(hmenu);
                    if (idCmd)
                        SendMessage(_hwnd, WM_SYSCOMMAND, idCmd, 0L);
                }
                SHSetWindowBits(hwnd, GWL_STYLE, WS_SYSMENU, 0L);
            }
            break;

        case WM_SYSCOMMAND:

            // if we are sizing, make the full screen accessible
            switch (wParam & 0xFFF0) {
                case SC_CLOSE:
                    _DoExitWindows(v_hwndDesktop, FALSE, FALSE);
                    break;

                default:
                    goto DoDefault;
            }
            break;

        // EXEX-VISTA: Validated.
        case TM_DESKTOPSTATE:
            _OnDesktopState(lParam);
            break;

        // EXEX-VISTA: Validated.
        case TM_RAISEDESKTOP:
            _RaiseDesktop((BOOL)wParam, FALSE);
            break;

            #ifdef DEBUG
        case TM_NEXTCTL:
            #endif

        // EXEX-VISTA: Validated. (both)
        case TM_UIACTIVATEIO:
        case TM_ONFOCUSCHANGEIS:
            _OnFocusMsg(uMsg, wParam, lParam);
            break;

        // EXEX-VISTA: Validated.
        case TM_MARSHALBS:  // wParam=IID lRes=pstm
            return BandSite_OnMarshallBS(wParam, lParam);

        // EXEX-VISTA: Validated.
        case TM_SETTIMER:
        case TM_KILLTIMER:
            return _OnTimerService(uMsg, wParam, lParam);
            break;

        case TM_FACTORY:
            return _OnFactoryMessage(wParam, lParam);

        case TM_ACTASTASKSW:
            _ActAsSwitcher();
            break;

        case TM_RELAYPOSCHANGED:
            _AppBarNotifyAll((HMONITOR)lParam, ABN_POSCHANGED, (HWND)wParam, 0);
            break;

        // EXEX-VISTA: Validated.
        case TM_BRINGTOTOP:
            SetWindowPos((HWND)wParam, NULL, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
            break;

        // EXEX-VISTA: Validated.
        case TM_WARNNOAUTOHIDE:
            //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.twp collision UI request"));

            //
            // this may look a little funny but what we do is post this message all
            // over the place and ignore it when we think it is a bad time to put
            // up a message (like the middle of a fulldrag...)
            //
            // wParam tells us if we need to try to clear the state
            // the lowword of _SetAutoHideState's return tells if anything changed
            //
            if ((!_fSysSizing || !g_fDragFullWindows) &&
                (!wParam || LOWORD(_SetAutoHideState(FALSE))))
            {
                ShellMessageBox(g_hinstCabinet, hwnd,
                    MAKEINTRESOURCE(IDS_ALREADYAUTOHIDEBAR),
                    MAKEINTRESOURCE(IDS_TASKBAR), MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                //DebugMsg(DM_TRAYDOCK, TEXT("TRAYDOCK.twp blowing off extraneous collision UI request"));
            }
            break;

        // EXEX-Vista: Validated.
        case TM_PRIVATECOMMAND:
            _HandlePrivateCommand(lParam);
            break;

        // EXEX-VISTA: Validated.
        case TM_HANDLEDELAYBOOTSTUFF:
            _HandleDelayBootStuff();
            break;

        case TM_SHELLSERVICEOBJECTS:
            _ssomgr.LoadRegObjects();
            break;

        // EXEX-VISTA: Validated.
        case TM_CHANGENOTIFY:
            _HandleChangeNotify(wParam, lParam);
            break;

        // EXEX-VISTA: Validated.
        case TM_GETHMONITOR:
            *((HMONITOR*)lParam) = _hmonStuck;
            break;

        // EXEX-VISTA: Validated.
        case TM_DOTRAYPROPERTIES:
            switch (wParam)
            {
                case 1:
                    DoProperties(TPF_STARTMENUPAGE);
                    break;
                case 2:
                    DoProperties(TPF_NOTIFAREAPAGE);
                    break;
                case 3:
                    DoProperties(TPF_INVOKECUSTOMIZE);
                    break;
                case 4:
                    DoProperties(TPF_TOOLBARSPAGE);
                    break;

                default:
                    DoProperties(TPF_TASKBARPAGE);
                    break;
            }

            break;

        // EXEX-VISTA: Validated.
        case TM_STARTUPAPPSLAUNCHED:
            _fUseChangeNotifyTimer = TRUE;
            PostMessage(_hwndNotify, TNM_STARTUPAPPSLAUNCHED, 0, 0);
            break;

        case TM_LANGUAGEBAND:
            return _ToggleLanguageBand(lParam);

        case WM_NCRBUTTONUP:
            uMsg = WM_CONTEXTMENU;
            wParam = (WPARAM)_hwndTasks;
            goto L_WM_CONTEXTMENU;

        // EXEX-VISTA: Validated.
        case WM_CONTEXTMENU:
            L_WM_CONTEXTMENU:

            if (!SHRestricted(REST_NOTRAYCONTEXTMENU))
            {
#if 0
                // DOESNT WORK (do we need?, cstartbutton handles it fine when the button is big enough)
                if (((HWND)wParam) == _stb._hwndStart)
                {
                    // Don't display of the Start Menu is up.
                    if (SendMessage(_stb._hwndStart, BM_GETSTATE, 0, 0) & BST_PUSHED)
                        break;
                    _fFromStart = TRUE;
                    _stb.OnContextMenu(_stb._hwndStart, (DWORD)lParam);
                    _fFromStart = FALSE;
                }
#endif
                if (IsPosInHwnd(lParam, _hwndNotify) || SHIsChildOrSelf(_hwndNotify, GetFocus()) == S_OK)
                {
                    // if click was inthe clock, include
                    // the time
                    _ContextMenu((DWORD)lParam, TRUE);
                }
                else
                {
                    BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, &lres);
                }
            }
            break;

        case WM_INITMENUPOPUP:
        case WM_MEASUREITEM:
        case WM_DRAWITEM:
        case WM_MENUCHAR:
            // Don't call bandsite message handler when code path started via the start button context menu
            if (!_fFromStart)
            {
                BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, &lres);
            }
            break;

        // EXEX-VISTA: Validated.
        case TM_DOEXITWINDOWS:
            _DoExitWindows(v_hwndDesktop, FALSE, FALSE);
            break;

        // EXEX-VISTA: Validated.
        case TM_HANDLESTARTUPFAILED:
            _OnHandleStartupFailed();
            break;

        case WM_HOTKEY:
            if (wParam < GHID_FIRST)
            {
                _HandleHotKey((WORD)wParam);
            }
            else
            {
                _HandleGlobalHotkey(wParam);
            }
            break;

        case WM_COMMAND:
            if (!BandSite_HandleMessage(_ptbs, hwnd, uMsg, wParam, lParam, &lres))
                _Command(GET_WM_COMMAND_ID(wParam, lParam), lParam);
            break;

        // EXEX-VISTA: LAZILY MODIFIED. Redo later.
        case SBM_CANCELMENU:
            _stb.CloseStartMenu();
            break;

        // EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
        case SBM_REBUILDMENU:
            _stb.BuildStartMenu();
            break;

        case WM_WINDOWPOSCHANGED:
            _AppBarActivationChange2(hwnd, _uStuckPlace);
            SendMessage(_hwndNotify, TNM_TRAYPOSCHANGED, 0, 0);
            goto DoDefault;

        // EXEX-VISTA: LAZYILY MODIFIED. Redo later.
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_MBUTTONDOWN:
            if (_stb._hwndStartBalloon)
            {
                RECT rc;
                POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
                GetWindowRect(_stb._hwndStartBalloon, &rc);
                MapWindowRect(HWND_DESKTOP, _hwnd, &rc);

                if (PtInRect(&rc, pt))
                {

                    /*ShowWindow(_stb._hwndStartBalloon, SW_HIDE);
                    _DontShowTheStartButtonBalloonAnyMore();
                    _stb._DestroyStartButtonBalloon();*/
                }
            }
            break;

        case TM_SETPUMPHOOK:
            IUnknown_SafeReleaseAndNullPtr(&_pmbTasks);
            IUnknown_SafeReleaseAndNullPtr(&_pmpTasks);
            if (wParam && lParam)
            {
                _pmbTasks = (IMenuBand*)wParam;
                _pmbTasks->AddRef();
                _pmpTasks = (IMenuPopup*)lParam;
                _pmpTasks->AddRef();
            }
            break;

        // EXEX-VISTA: Validated.
        case WM_ACTIVATE:
            _AppBarActivationChange2(hwnd, _uStuckPlace);
            if (wParam != WA_INACTIVE)
            {
                Unhide();
            }
            else
            {
                // When tray is deactivated, remove our keyboard cues:
                //
                SendMessage(hwnd, WM_CHANGEUISTATE,
                    MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS | UISF_HIDEACCEL), 0);

                IUnknown_UIActivateIO(_ptbs, FALSE, NULL);
            }
            //
            // Tray activation is a good time to do a reality check
            // (make sure "always-on-top" agrees with actual window
            // position, make sure there are no ghost buttons, etc).
            //
            RealityCheck();
            goto L_default;

        case WM_WTSSESSION_CHANGE:
        {
            lres = _OnSessionChange(wParam, lParam);
            break;
        }

        // EXEX-VISTA: Validated.
        case WM_THEMECHANGED:
        {
            _OnThemeChanged();
            break;
        }

        // EXEX-VISTA: Validated.
        case TM_WORKSTATIONLOCKED:
        {
            // Desktop locked status changed...
            if (_fIsDesktopLocked != (BOOL)wParam)
            {
                _fIsDesktopLocked = (BOOL)wParam;
                _fIsLogoff = FALSE;
                _RecomputeAllWorkareas();
                PostMessage(_hwndNotify, TNM_WORKSTATIONLOCKED, wParam, 0);
            }
        }
        break;

        // EXEX-VISTA: Validated.
        case TM_SHOWTRAYBALLOON:
            PostMessage(_hwndNotify, TNM_SHOWTRAYBALLOON, wParam, 0);
            break;

        // EXEX-VISTA: Validated.
        case TM_STARTMENUDISMISSED:
            //  107561 - call CoFreeUnusedLibraries() peridically to free up dlls - ZekeL - 4-MAY-2001
            //  specifically to support MSONSEXT (webfolders) being used in RecentDocs
            //  after a file has been opened via webfolders.  we get the icon via
            //  their namespace but then COM holds on to the DLL for a while (forever?)
            //  calling CoFreeUnusedLibraries() does the trick
            SetTimer(_hwnd, IDT_COFREEUNUSED, 3 * 60 * 1000, NULL);
            break;

        case MM_MIXM_CONTROL_CHANGE:
            Mixer_ControlChange(wParam, lParam);
            break;

        // EXEX-VISTA: LAZILY MODIFIED. Redo later.
        default:
            L_default:
            if (uMsg == GetDDEExecMsg())
            {
                ASSERT(lParam && 0 == ((LPNMHDR)lParam)->idFrom);
                DDEHandleViewFolderNotify(NULL, _hwnd, (LPNMVIEWFOLDER)lParam);
                LocalFree((LPNMVIEWFOLDER)lParam);
                return TRUE;
            }
            else if (uMsg == _uStartButtonBalloonTip)
            {
                // _ShowStartButtonToolTip();
            }
            else if (uMsg == _uLogoffUser)
            {
                // Log off the current user (message from U&P control panel)
                ExitWindowsEx(EWX_LOGOFF, 0);
            }
            else if (uMsg == _uMsgEnableUserTrackedBalloonTips)
            {
                PostMessage(_hwndNotify, TNM_ENABLEUSERTRACKINGINFOTIPS, wParam, 0);
            }
            else if (uMsg == _uWinMM_DeviceChange)
            {
                Mixer_MMDeviceChange();
            }


            DoDefault:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return lres;
}

HWND TermSoundWindow()
{
    // EXEX-VISTA: Uncomment when implemented g_hwndSound
    return nullptr;
    /*if (g_hwndSound)
    {
        PostMessageW(g_hwndSound, WM_CLOSE, 0, 0);
        g_hwndSound = 0;
    }*/
}

static BOOL g_fShellShutdown = FALSE;

void CTray::_DoExitExplorer()
{
    // User cancelled...
    // The shift key means exit the tray...
    // ??? - Used to destroy all cabinets...
    // PostQuitMessage(0);
    g_fFakeShutdown = TRUE; // Don't blow away session state; the session will survive
    //TraceMsg(TF_TRAY, "c.dew: Posting quit message for tid=%#08x hwndDesk=%x(IsWnd=%d) hwndTray=%x(IsWnd=%d)", GetCurrentThreadId(),
    //    v_hwndDesktop, IsWindow(v_hwndDesktop), _hwnd, IsWindow(_hwnd));
    // 1 means close all the shell windows too
    PostMessage(v_hwndDesktop, WM_QUIT, 0, 1);
    PostMessage(_hwnd, WM_QUIT, 0, 0);

    g_fShellShutdown = TRUE;
}

// EXEX-VISTA: CHANGED since XP. Some new arguments were added. Revalidate later.
void CTray::_DoExitWindows(HWND hwnd, BOOL fIsRestarting, BOOL fFromNotifArea)
{
    if (!g_fShellShutdown)
    {
        if (_Restricted(hwnd, REST_NOCLOSE) && !fIsRestarting)
            return;

        {
            UEMFireEvent(&UEMIID_SHELL, UEME_CTLSESSION, UEMF_XEVENT, FALSE, -1);
            // really #ifdef DEBUG, but want for testing
            // however can't do unconditionally due to perf
            if (ERROR_SUCCESS == SHGetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("StartMenuForceRefresh"),
                NULL, NULL, NULL) || GetAsyncKeyState(VK_SHIFT) < 0)
            {
                _RefreshStartMenu();
            }
        }

        _SaveTrayAndDesktop();

        if (!fIsRestarting)
        {
            _uModalMode = MM_SHUTDOWN;
            ExitWindowsDialog(hwnd);
        }

        // NB User can have problems if the focus is forcebly changed to the desktop while
        // shutting down since it tries to serialize the whole process by making windows sys-modal.
        // If we hit this code at just the wrong moment (ie just after the sharing dialog appears)
        // the desktop will become sys-modal so you can't switch back to the sharing dialog
        // and you won't be able to shutdown.
        // SetForegroundWindow(hwnd);
        // SetFocus(hwnd);

        _uModalMode = 0;
        if ((GetKeyState(VK_SHIFT) < 0) && (GetKeyState(VK_CONTROL) < 0) && (GetKeyState(VK_MENU) < 0))
        {
            TermSoundWindow();
            _DoExitExplorer();
        }
    }
}

void CTray::_SaveTray(void)
{
    if (SHRestricted(REST_NOSAVESET))
        return;

    if (SHRestricted(REST_CLEARRECENTDOCSONEXIT))
        ClearRecentDocumentsAndMRUStuff(FALSE);

    //
    // Don't persist tray stuff if in safe mode.  We want this
    // to be a temporary mode where the UI settings don't stick.
    //
    if (GetSystemMetrics(SM_CLEANBOOT) == 0)
    {
        _SaveTrayStuff();
    }
}

DWORD WINAPI CTray::PropertiesThreadProc(void* pv)
{
    return c_tray._PropertiesThreadProc(PtrToUlong(pv));
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Definitely different than Vista.
DWORD CTray::_PropertiesThreadProc(DWORD dwFlags)
{
    HWND hwnd;
    RECT rc;
    DWORD dwExStyle = WS_EX_TOOLWINDOW;

    GetWindowRect(_stb._hwndStart, &rc);
    dwExStyle |= IS_BIDI_LOCALIZED_SYSTEM() ? dwExStyleRTLMirrorWnd : 0L;

    _hwndProp = hwnd = CreateWindowEx(dwExStyle, TEXT("static"), NULL, 0,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, NULL, NULL, g_hinstCabinet, NULL);

    #define IDI_STTASKBR 40         // stolen from shell32\ids.h
    if (_hwndProp)
    {
        // Get the Alt+Tab icon right
        HICON hicoStub = LoadIcon(GetModuleHandle(TEXT("SHELL32")), MAKEINTRESOURCE(IDI_STTASKBR));
        SendMessage(_hwndProp, WM_SETICON, ICON_BIG, (LPARAM)hicoStub);

        // SwitchToThisWindow(hwnd, TRUE);
        // SetForegroundWindow(hwnd);

        DoTaskBarProperties(hwnd, dwFlags);

        _hwndProp = NULL;
        DestroyWindow(hwnd);
        if (hicoStub)
            DestroyIcon(hicoStub);
    }

    return TRUE;
}

// EXEX-VISTA: Move later.
typedef struct tagTRAYPROPTHREADDATA
{
    DWORD dwFlags;
    RECT rcStartButton;
    LPSTREAM pCatBandManager;
} TRAYPROPTHREADDATA, * PTRAYPROPTHREADDATA;

#define RUNWAITSECS 5
void CTray::DoProperties(DWORD dwFlags)
{
    if (!_Restricted(_hwnd, REST_NOSETTASKBAR))
    {
        int i = RUNWAITSECS;
        while (_hwndProp == ((HWND)-1) && i--)
        {
            // we're in the process of coming up. wait
            Sleep(1000);
        }

        // failed!  blow it off.
        if (_hwndProp == (HWND)-1)
        {
            _hwndProp = NULL;
        }

        if (_hwndProp)
        {
            // there's a window out there... activate it
            SwitchToThisWindow(GetLastActivePopup(_hwndProp), TRUE);
        }
        else
        {
            TRAYPROPTHREADDATA *pParams = new TRAYPROPTHREADDATA();
            if (pParams)
            {
                _stb.GetRect(&pParams->rcStartButton);
                pParams->dwFlags = dwFlags;

                // EXEX-VISTA: Weird thing to understand. Temporary code for now:
                #if 1 // TEMP
                pParams->pCatBandManager = nullptr;
                #endif
                /*
                v8 = *(IUnknown **)&this->GAP_67A[2];
                if ( v8 )
                  CoMarshalInterThreadInterfaceInStream(&IID_ICatBandManager, v8, &params->pCatBandManager);
                */

                _hwndProp = (HWND)-1;
                if (!SHCreateThread(PropertiesThreadProc, pParams, CTF_COINIT, NULL))
                {
                    delete pParams;
                    _hwndProp = NULL;
                }
            }
        }
    }
}

BOOL CTray::TileEnumProc(HWND hwnd, LPARAM lParam)
{
    CTray* ptray = (CTray*)lParam;

    if (IsWindowVisible(hwnd) && !IsIconic(hwnd) &&
        ((GetWindowLong(hwnd, GWL_STYLE) & WS_CAPTION) == WS_CAPTION) &&
        (hwnd != ptray->_hwnd) && hwnd != v_hwndDesktop)
    {
        return FALSE;   // we *can* tile this guy
    }
    return TRUE;    // we *cannot* tile this guy
}

HMENU CTray::BuildContextMenu(BOOL fIncludeTime)
{
    HMENU hmContext = LoadMenuPopup(MAKEINTRESOURCE(MENU_TRAYCONTEXT));
    if (!hmContext)
        return NULL;

    if (fIncludeTime)
    {
        if (_trayNotify.GetIsNoTrayItemsDisplayPolicyEnabled())
        {
            // We know the position of IDM_NOTIFYCUST from the menu resource...
            DeleteMenu(hmContext, 1, MF_BYPOSITION);
        }
        else
        {
            UINT uEnable = MF_BYCOMMAND;
            if (_trayNotify.GetIsNoAutoTrayPolicyEnabled() || !_trayNotify.GetIsAutoTrayEnabledByUser())
            {
                uEnable |= MFS_DISABLED;
            }
            else
            {
                uEnable |= MFS_ENABLED;
            }
            EnableMenuItem(hmContext, IDM_NOTIFYCUST, uEnable);
        }
    }
    else
    {
        INSTRUMENT_STATECHANGE(SHCNFI_STATE_TRAY_CONTEXT);
        for (int i = 2; i >= 0; i--)    // separator, IDM_SETTIME, IDM_NOTIFYCUST,
        {
            DeleteMenu(hmContext, i, MF_BYPOSITION);
        }
    }

    CheckMenuItem(hmContext, IDM_LOCKTASKBAR,
        MF_BYCOMMAND | (_fCanSizeMove ? MF_UNCHECKED : MF_CHECKED));

    // Don't let users accidentally check lock the taskbar when the taskbar is zero height
    RECT rc;
    GetClientRect(_hwnd, &rc);
    EnableMenuItem(hmContext, IDM_LOCKTASKBAR,
        MF_BYCOMMAND | ((_IsSizeMoveRestricted() || (RECTHEIGHT(rc) == 0)) ? MFS_DISABLED : MFS_ENABLED));

    if (!_fUndoEnabled || !_pPositions)
    {
        DeleteMenu(hmContext, IDM_UNDO, MF_BYCOMMAND);
    }
    else
    {
        TCHAR szTemplate[30];
        TCHAR szCommand[30];
        TCHAR szMenu[64];
        LoadString(g_hinstCabinet, IDS_UNDOTEMPLATE, szTemplate, ARRAYSIZE(szTemplate));
        LoadString(g_hinstCabinet, _pPositions->idRes, szCommand, ARRAYSIZE(szCommand));
        StringCchPrintf(szMenu, ARRAYSIZE(szMenu), szTemplate, szCommand);
        ModifyMenu(hmContext, IDM_UNDO, MF_BYCOMMAND | MF_STRING, IDM_UNDO, szMenu);
    }

    if (g_fDesktopRaised)
    {
        TCHAR szHideDesktop[64];
        LoadString(g_hinstCabinet, IDS_HIDEDESKTOP, szHideDesktop, ARRAYSIZE(szHideDesktop));
        ModifyMenu(hmContext, IDM_TOGGLEDESKTOP, MF_BYCOMMAND | MF_STRING, IDM_TOGGLEDESKTOP, szHideDesktop);
    }

    if (!_CanTileAnyWindows())
    {
        EnableMenuItem(hmContext, IDM_CASCADE, MFS_GRAYED | MF_BYCOMMAND);
        EnableMenuItem(hmContext, IDM_HORIZTILE, MFS_GRAYED | MF_BYCOMMAND);
        EnableMenuItem(hmContext, IDM_VERTTILE, MFS_GRAYED | MF_BYCOMMAND);
    }

    HKEY hKeyPolicy;
    if (RegOpenKeyEx(HKEY_CURRENT_USER,
        TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System"),
        0, KEY_READ, &hKeyPolicy) == ERROR_SUCCESS)
    {
        DWORD dwType, dwData = 0, dwSize = sizeof(dwData);
        RegQueryValueEx(hKeyPolicy, TEXT("DisableTaskMgr"), NULL,
            &dwType, (LPBYTE)&dwData, &dwSize);
        RegCloseKey(hKeyPolicy);
        if (dwData)
            EnableMenuItem(hmContext, IDM_SHOWTASKMAN, MFS_GRAYED | MF_BYCOMMAND);
    }

    return hmContext;
}

void CTray::ContextMenuInvoke(int idCmd, BOOL fFromNotifArea)
{
    if (idCmd)
    {
        if (idCmd < IDM_TRAYCONTEXTFIRST)
        {
            BandSite_HandleMenuCommand(_ptbs, idCmd);
            AsyncSaveSettings();
        }
        else
        {
            _Command(idCmd, fFromNotifArea);
        }
    }
}

//
//  CTray::AsyncSaveSettings
//
//  We need to save our tray settings, but there may be a bunch
//  of these calls coming, (at startup time or when dragging
//  items in the task bar) so gather them up into one save that
//  will happen in at least 2 seconds.
//
void CTray::AsyncSaveSettings()
{
    if (!_fHandledDelayBootStuff)        // no point in saving if we're not done booting
        return;

    KillTimer(_hwnd, IDT_SAVESETTINGS);
    SetTimer(_hwnd, IDT_SAVESETTINGS, 2000, NULL);
}


void CTray::_ContextMenu(DWORD dwPos, BOOL fFromNotifArea)
{
#ifdef DEAD_CODE
    POINT pt = { LOWORD(dwPos), HIWORD(dwPos) };

    SwitchToThisWindow(_hwnd, TRUE);
    SetForegroundWindow(_hwnd);
    SendMessage(_hwndTrayTips, TTM_ACTIVATE, FALSE, 0L);

    if (dwPos != (DWORD)-1 &&
        IsChildOrHWND(_hwndRebar, WindowFromPoint(pt)))
    {
        // if the context menu came from below us, reflect down
        BandSite_HandleMessage(_ptbs, _hwnd, WM_CONTEXTMENU, 0, dwPos, NULL);

    }
    else
    {
        HMENU hmenu;

        if (dwPos == (DWORD)-1)
        {
            HWND hwnd = GetFocus();
            pt.x = pt.y = 0;
            ClientToScreen(hwnd, &pt);
            dwPos = MAKELONG(pt.x, pt.y);
        }

        hmenu = BuildContextMenu(fFromNotifArea);
        if (hmenu)
        {
            int idCmd;

            BandSite_AddMenus(_ptbs, hmenu, 0, 0, IDM_TRAYCONTEXTFIRST);

            idCmd = TrackPopupMenu(hmenu, TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_LEFTALIGN,
                GET_X_LPARAM(dwPos), GET_Y_LPARAM(dwPos), 0, _hwnd, NULL);

            DestroyMenu(hmenu);

            ContextMenuInvoke(idCmd, fFromNotifArea);
        }
    }

    SendMessage(_hwndTrayTips, TTM_ACTIVATE, TRUE, 0L);
#endif
    int idCmd;
    HMENU hmenu;

    POINT pt = { LOWORD(dwPos), HIWORD(dwPos) };

    SwitchToThisWindow(_hwnd, TRUE);
    SetForegroundWindow(_hwnd);
    EnableTooltips(FALSE);

    if (dwPos != -1)
    {
        if (IsChildOrHWND(_hwndRebar, WindowFromPoint(pt)))
        {
            BandSite_HandleMessage(_ptbs, _hwnd, 0x7Bu, 0, dwPos, nullptr);
            goto LABEL_7;
        }
    }
    else
    {
        HWND hwnd = GetFocus();
        pt.x = pt.y = 0;
        ClientToScreen(hwnd, &pt);
        dwPos = MAKELONG(pt.x, pt.y);
    }

    hmenu = BuildContextMenu(fFromNotifArea);
    if (hmenu)
    {
        BandSite_AddMenus(_ptbs, hmenu, 0, 0, 0x190u);
        idCmd = TrackPopupMenu(hmenu, 0x102u, GET_X_LPARAM(dwPos), GET_Y_LPARAM(dwPos), 0, _hwnd, nullptr);
        DestroyMenu(hmenu);
        ContextMenuInvoke(idCmd, fFromNotifArea);
    }

LABEL_7:
    EnableTooltips(TRUE);
}

#define RFD_NOBROWSE		0x00000001
#define RFD_NODEFFILE		0x00000002
#define RFD_USEFULLPATHDIR	0x00000004
#define RFD_NOSHOWOPEN          0x00000008
#define RFD_CONSENTHOTKEY   0x00000100

void _RunFileDlg(HWND hwnd, UINT idIcon, LPCITEMIDLIST pidlWorkingDir, UINT idTitle, UINT idPrompt, DWORD dwFlags)
{
    HICON hIcon;
    LPCTSTR lpszTitle;
    LPCTSTR lpszPrompt;
    TCHAR szTitle[256];
    TCHAR szPrompt[256];
    TCHAR szWorkingDir[MAX_PATH];

    dwFlags |= RFD_USEFULLPATHDIR | RFD_CONSENTHOTKEY;
    szWorkingDir[0] = 0;

    hIcon = idIcon ? LoadIcon(g_hinstCabinet, MAKEINTRESOURCE(idIcon)) : NULL;

    if (!pidlWorkingDir || !SHGetPathFromIDList(pidlWorkingDir, szWorkingDir))
    {
        // This is either the Tray, or some non-file system folder, so
        // we will "suggest" the Desktop as a working dir, but if the
        // user types a full path, we will use that instead.  This is
        // what WIN31 Progman did (except they started in the Windows
        // dir instead of the Desktop).
        goto UseDesktop;
    }

    // if it's a removable dir, make sure it's still there
    if (szWorkingDir[0])
    {
        int idDrive = PathGetDriveNumber(szWorkingDir);
        if ((idDrive != -1))
        {
            UINT dtype = DriveType(idDrive);
            if (((dtype == DRIVE_REMOVABLE) || (dtype == DRIVE_CDROM))
                && !PathFileExists(szWorkingDir))
            {
                goto UseDesktop;
            }
        }
    }

    //
    // Check if this is a directory. Notice that it could be a in-place
    // navigated document.
    //
    if (PathIsDirectory(szWorkingDir)) {
        goto UseWorkingDir;
    }

    UseDesktop:
    SHGetSpecialFolderPath(hwnd, szWorkingDir, CSIDL_DESKTOPDIRECTORY, FALSE);

    UseWorkingDir:

    if (idTitle)
    {
        LoadString(g_hinstCabinet, idTitle, szTitle, ARRAYSIZE(szTitle));
        lpszTitle = szTitle;
    }
    else
        lpszTitle = NULL;
    if (idPrompt)
    {
        LoadString(g_hinstCabinet, idPrompt, szPrompt, ARRAYSIZE(szPrompt));
        lpszPrompt = szPrompt;
    }
    else
        lpszPrompt = NULL;

    RunFileDlg(hwnd, hIcon, szWorkingDir, lpszTitle, lpszPrompt, dwFlags);
}

// EXEX-VISTA: Validated.
BOOL CTray::SavePosEnumProc(HWND hwnd, LPARAM lParam)
{
    // dont need to entercritical here since we are only ever
    // called from SaveWindowPositions, which as already entered the critical secion
    // for _pPositions

    //ASSERTCRITICAL;

    CTray* ptray = (CTray*)lParam;

    ASSERT(ptray->_pPositions);

    if (IsWindowVisible(hwnd) &&
        (hwnd != ptray->_hwnd) &&
        (hwnd != v_hwndDesktop))
    {
        HWNDANDPLACEMENT hap;

        hap.wp.length = sizeof(WINDOWPLACEMENT);
        GetWindowPlacement(hwnd, &hap.wp);

        if (hap.wp.showCmd != SW_SHOWMINIMIZED)
        {
            hap.hwnd = hwnd;
            hap.fRestore = TRUE;

            DSA_AppendItem(ptray->_pPositions->hdsaWP, &hap);
        }
    }
    return TRUE;
}

// EXEX-VISTA: Validated.
void CTray::SaveWindowPositions(UINT idRes)
{
    ENTERCRITICAL;
    if (_pPositions)
    {
        if (_pPositions->hdsaWP)
            DSA_DeleteAllItems(_pPositions->hdsaWP);
    }
    else
    {
        _pPositions = (LPWINDOWPOSITIONS)LocalAlloc(LPTR, sizeof(WINDOWPOSITIONS));
        if (_pPositions)
        {
            _pPositions->hdsaWP = DSA_Create(sizeof(HWNDANDPLACEMENT), 4);
        }
    }

    if (_pPositions)
    {
        _pPositions->idRes = idRes;

        // CheckWindowPositions tested for these...
        ASSERT(idRes == IDS_MINIMIZEALL || idRes == IDS_CASCADE || idRes == IDS_SHOW_STACKED || idRes == IDS_SHOW_SIDE_BY_SIDE);
        EnumWindows(SavePosEnumProc, (LPARAM)this);
    }
    LEAVECRITICAL;
}

typedef struct
{
    LPWINDOWPOSITIONS pPositions;
    HWND hwndDesktop;
    HWND hwndTray;
    BOOL fPostLowerDesktop;
} RESTOREWNDDATA, * PRESTOREWNDDATA;

DWORD WINAPI RestoreWndPosThreadProc(void* pv)
{
    PRESTOREWNDDATA pWndData = (PRESTOREWNDDATA)pv;
    if (pWndData && pWndData->pPositions)
    {
        LPHWNDANDPLACEMENT phap;
        LONG iAnimate;
        ANIMATIONINFO ami;

        ami.cbSize = sizeof(ANIMATIONINFO);
        SystemParametersInfo(SPI_GETANIMATION, sizeof(ami), &ami, FALSE);
        iAnimate = ami.iMinAnimate;
        ami.iMinAnimate = FALSE;
        SystemParametersInfo(SPI_SETANIMATION, sizeof(ami), &ami, FALSE);

        if (pWndData->pPositions->hdsaWP)
        {
            for (int i = DSA_GetItemCount(pWndData->pPositions->hdsaWP) - 1; i >= 0; i--)
            {
                phap = (LPHWNDANDPLACEMENT)DSA_GetItemPtr(pWndData->pPositions->hdsaWP, i);
                if (IsWindow(phap->hwnd))
                {

                    #ifndef WPF_ASYNCWINDOWPLACEMENT
                    #define WPF_ASYNCWINDOWPLACEMENT 0x0004
                    #endif
                                        // pass this async.
                    if (!IsHungAppWindow(phap->hwnd))
                    {
                        phap->wp.length = sizeof(WINDOWPLACEMENT);
                        phap->wp.flags |= WPF_ASYNCWINDOWPLACEMENT;
                        if (phap->fRestore)
                        {
                            // only restore those guys we've actually munged.
                            SetWindowPlacement(phap->hwnd, &phap->wp);
                        }
                    }
                }
            }
        }

        ami.iMinAnimate = iAnimate;
        SystemParametersInfo(SPI_SETANIMATION, sizeof(ami), &ami, FALSE);

        _DestroySavedWindowPositions(pWndData->pPositions);

        if (pWndData->fPostLowerDesktop)
        {
            PostMessage(pWndData->hwndDesktop, DTM_RAISE, (WPARAM)pWndData->hwndTray, DTRF_LOWER);
        }

        delete pWndData;
    }

    return 1;
}


BOOL CTray::_RestoreWindowPositions(BOOL fPostLowerDesktop)
{
    BOOL fRet = FALSE;

    ENTERCRITICAL;
    if (_pPositions)
    {
        PRESTOREWNDDATA pWndData = new RESTOREWNDDATA;
        if (pWndData)
        {
            pWndData->pPositions = _pPositions;
            pWndData->fPostLowerDesktop = fPostLowerDesktop;
            pWndData->hwndDesktop = v_hwndDesktop;
            pWndData->hwndTray = _hwnd;

            if (SHCreateThread(RestoreWndPosThreadProc, pWndData, 0, NULL))
            {
                fRet = TRUE;
                _pPositions = NULL;
            }
            else
            {
                delete pWndData;
            }
        }
    }
    LEAVECRITICAL;

    return fRet;
}

void _DestroySavedWindowPositions(LPWINDOWPOSITIONS pPositions)
{
    ENTERCRITICAL;
    if (pPositions)
    {
        // free the global struct
        DSA_Destroy(pPositions->hdsaWP);
        LocalFree(pPositions);
    }
    LEAVECRITICAL;
}

void CTray::HandleWindowDestroyed(HWND hwnd)
{
    // enter critical section so we dont corrupt the hdsaWP
    ENTERCRITICAL;
    if (_pPositions)
    {
        int i = DSA_GetItemCount(_pPositions->hdsaWP) - 1;
        for (; i >= 0; i--) {
            LPHWNDANDPLACEMENT phap = (LPHWNDANDPLACEMENT)DSA_GetItemPtr(_pPositions->hdsaWP, i);
            if (phap->hwnd == hwnd || !IsWindow(phap->hwnd)) {
                DSA_DeleteItem(_pPositions->hdsaWP, i);
            }
        }

        if (!DSA_GetItemCount(_pPositions->hdsaWP))
        {
            _DestroySavedWindowPositions(_pPositions);
            _pPositions = NULL;
        }
    }
    LEAVECRITICAL;
}

// Allow us to bump the activation of the run dlg hidden window.
// Certain apps (Norton Desktop setup) use the active window at RunDlg time
// as the parent for their dialogs. If that window disappears then they fault.
// We don't want the tray to get the activation coz it will cause it to appeare
// if you're in auto-hide mode.
LRESULT WINAPI RunDlgStaticSubclassWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_ACTIVATE:
            if (wParam == WA_ACTIVE)
            {
                // Bump the activation to the desktop.
                if (v_hwndDesktop)
                {
                    SetForegroundWindow(v_hwndDesktop);
                    return 0;
                }
            }
            break;

        case WM_NOTIFY:
            // relay it to the tray
            return SendMessage(v_hwndTray, uMsg, wParam, lParam);
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

DWORD WINAPI CTray::RunDlgThreadProc(void* pv)
{
    return c_tray._RunDlgThreadProc((HANDLE)pv);
}

BOOL _IsBrowserWindow(HWND hwnd)
{
    static const TCHAR* c_szClasses[] =
    {
        TEXT("ExploreWClass"),
        TEXT("CabinetWClass"),
        TEXT("IEFrame"),
    };

    TCHAR szClass[32];
    GetClassName(hwnd, szClass, ARRAYSIZE(szClass));

    for (int i = 0; i < ARRAYSIZE(c_szClasses); i++)
    {
        if (lstrcmpi(szClass, c_szClasses[i]) == 0)
        {
            return TRUE;
        }
    }
    return FALSE;
}

__inline HBRUSH SelectBrush(HDC hdc, HBRUSH hbr) { return (HBRUSH)SelectObject(hdc, (HGDIOBJ)(HBRUSH)hbr); }
__inline HFONT SelectFont(HDC hdc, HFONT hfont) { return (HFONT)SelectObject(hdc, (HGDIOBJ)(HFONT)hfont); }
__inline HBITMAP SelectBitmap(HDC hdc, HBITMAP hbm) { return (HBITMAP)SelectObject(hdc, (HGDIOBJ)(HBITMAP)hbm); }
__inline WNDPROC SubclassWindow(HWND hwnd, WNDPROC lpfn) { return (WNDPROC)SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LPARAM)lpfn); }

#define CWM_COMMANDLINE                 (WM_USER + 11)
// global clone your current pidl
#define CWM_CLONEPIDL                   (WM_USER + 12)
// See if the root of the instance is as specified
#define CWM_COMPAREROOT                 (WM_USER + 13)

// EXEX-VISTA: SLIGHTLY MODIFIED. Definitely different than Vista.
DWORD CTray::_RunDlgThreadProc(HANDLE hdata)
{
    RECT            rc, rcTemp;
    HRESULT hrInit = SHCoInitialize();

    // 99/04/12 #316424 vtan: Get the rectangle for the "Start" button.
    // If this is off the screen the tray is probably in auto hide mode.
    // In this case offset the rectangle into the monitor where it should
    // belong. This may be up, down, left or right depending on the
    // position of the tray.

    // First thing to do is to establish the dimensions of the monitor on
    // which the tray resides. If no monitor can be found then use the
    // primary monitor.

    MONITORINFO monitorInfo;
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (GetMonitorInfo(_hmonStuck, &monitorInfo) == 0)
    {
        TBOOL(SystemParametersInfo(SPI_GETWORKAREA, 0, &monitorInfo.rcMonitor, 0));
    }

    // Get the co-ordinates of the "Start" button.

    GetWindowRect(_stb._hwndStart, &rc);

    // Look for an intersection in the monitor.

    if (IntersectRect(&rcTemp, &rc, &monitorInfo.rcMonitor) == 0)
    {
        LONG    lDeltaX, lDeltaY;

        // Does not exist in the monitor. Move the co-ordinates by the
        // width or height of the tray so that it does.

        // This bizarre arithmetic is used because _ComputeHiddenRect()
        // takes into account the frame and that right/bottom of RECT
        // is exclusive in GDI.

        lDeltaX = _sStuckWidths.cx - g_cxFrame - g_cxPaddedBorder;
        lDeltaY = _sStuckWidths.cy - g_cyFrame - g_cxPaddedBorder;
        if (rc.left < monitorInfo.rcMonitor.left)
        {
            --lDeltaX;
            lDeltaY = 0;
        }
        else if (rc.top < monitorInfo.rcMonitor.top)
        {
            lDeltaX = 0;
            --lDeltaY;
        }
        else if (rc.right > monitorInfo.rcMonitor.right)
        {
            lDeltaX = -lDeltaX;
            lDeltaY = 0;
        }
        else if (rc.bottom > monitorInfo.rcMonitor.bottom)
        {
            lDeltaX = 0;
            lDeltaY = -lDeltaY;
        }
        TBOOL(OffsetRect(&rc, lDeltaX, lDeltaY));
    }

    HWND hwnd = CreateWindowEx(WS_EX_TOOLWINDOW, TEXT("static"), NULL, 0,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, NULL, NULL, g_hinstCabinet, NULL);
    if (hwnd)
    {
        BOOL fSimple = FALSE;
        HANDLE hMemWorkDir = NULL;
        LPITEMIDLIST pidlWorkingDir = NULL;

        // Subclass it.
        SubclassWindow(hwnd, RunDlgStaticSubclassWndProc);

        if (hdata)
            SetProp(hwnd, TEXT("WaitingThreadID"), hdata);

        if (!SHRestricted(REST_STARTRUNNOHOMEPATH))
        {
            // On NT, we like to start apps in the HOMEPATH directory.  This
            // should be the current directory for the current process.
            TCHAR szDir[MAX_PATH];
            TCHAR szPath[MAX_PATH];

            GetEnvironmentVariable(TEXT("HOMEDRIVE"), szDir, ARRAYSIZE(szDir));
            GetEnvironmentVariable(TEXT("HOMEPATH"), szPath, ARRAYSIZE(szPath));

            if (PathAppend(szDir, szPath) && PathIsDirectory(szDir))
            {
                pidlWorkingDir = SHSimpleIDListFromPath(szDir);
                if (pidlWorkingDir)
                {
                    // free it the "simple" way...
                    fSimple = TRUE;
                }
            }
        }

        if (!pidlWorkingDir)
        {
            // If the last active window was a folder/explorer window with the
            // desktop as root, use its as the current dir
            if (_hwndLastActive)
            {
                ENTERCRITICAL;
                if (_hwndLastActive && !IsMinimized(_hwndLastActive) && _IsBrowserWindow(_hwndLastActive))
                {
                    SendMessageTimeout(_hwndLastActive, CWM_CLONEPIDL, GetCurrentProcessId(), 0, SMTO_ABORTIFHUNG | SMTO_BLOCK, 500, (DWORD_PTR*)&hMemWorkDir);
                    pidlWorkingDir = (LPITEMIDLIST)SHLockShared(hMemWorkDir, GetCurrentProcessId());
                }
                LEAVECRITICAL;
            }
        }

        _RunFileDlg(hwnd, 0, pidlWorkingDir, 0, 0, 0);
        if (pidlWorkingDir)
        {
            if (fSimple)
            {
                ILFree(pidlWorkingDir);
            }
            else
            {
                SHUnlockShared(pidlWorkingDir);
            }
        }

        if (hMemWorkDir)
        {
            ASSERT(fSimple == FALSE);
            SHFreeShared(hMemWorkDir, GetCurrentProcessId());
        }

        if (hdata)
        {
            RemoveProp(hwnd, TEXT("WaitingThreadID"));
        }

        DestroyWindow(hwnd);
    }

    SHCoUninitialize(hrInit);
    return TRUE;
}

//
// TPS_DEMANDTHREAD - always create a new thread if none currently available.
// Used in situations where immediate response required
//

#define TPS_DEMANDTHREAD    0x00000004

//
// TPS_LONGEXECTIME - the work item will take relatively long time to execute.
// Used as management hint to TPS
//

#define TPS_LONGEXECTIME    0x00000008

typedef struct tagRUNDLGTHREADDATA
{
    DWORD dwThreadId;
    RECT rcStartButton;
} RUNDLGTHREADDATA, * PRUNDLGTHREADDATA;

// EXEX-VISTA: Validated.
void CTray::_RunDlg(BOOL fCreateEvent)
{
    HANDLE hEvent;
    void* pvThreadParam;

    if (!_Restricted(_hwnd, REST_NORUN))
    {
        TCHAR szRunDlgTitle[MAX_PATH];
        HWND  hwndOldRun;
        LoadString(g_hinstCabinet, IDS_RUNDLGTITLE, szRunDlgTitle, ARRAYSIZE(szRunDlgTitle));

        // See if there is already a run dialog up, and if so, try to activate it

        hwndOldRun = FindWindow(WC_DIALOG, szRunDlgTitle);
        if (hwndOldRun)
        {
            DWORD dwPID;

            GetWindowThreadProcessId(hwndOldRun, &dwPID);
            if (dwPID == GetCurrentProcessId())
            {
                if (IsWindowVisible(hwndOldRun))
                {
                    SetForegroundWindow(hwndOldRun);
                    return;
                }
            }
        }

        RUNDLGTHREADDATA *pThreadParam = new RUNDLGTHREADDATA();
        if (!pThreadParam)
        {
            return;
        }

        // Create an event so we can wait for the run dlg to appear before
        // continue - this allows it to capture any type-ahead.
        if (fCreateEvent)
        {
            hEvent = CreateEvent(NULL, TRUE, FALSE, TEXT("MSShellRunDlgReady"));
        }
        else
        {
            hEvent = NULL;
        }

        if (hEvent)
        {
            pThreadParam->dwThreadId = GetCurrentThreadId();
        }
        else
        {
            pThreadParam->dwThreadId = NULL;
        }

        _stb.GetRect(&pThreadParam->rcStartButton);

        if (SHCreateThread(RunDlgThreadProc, pThreadParam, CTF_COINIT_STA, NULL))
        {
            if (!hEvent)
            {
                return;
            }

            WaitForSingleObject(hEvent, 500);
        }

        if (hEvent)
            CloseHandle(hEvent);
    }
}

void CTray::_ExploreCommonStartMenu(BOOL bExplore)
{
    TCHAR szPath[MAX_PATH];
    TCHAR szCmdLine[MAX_PATH + 50];

    //
    // Get the common start menu path.
    //
    // we want to force the directory to exist, but not on W95 machines
    if (!SHGetSpecialFolderPath(NULL, szPath, CSIDL_COMMON_STARTMENU, FALSE))
    {
        return;
    }

    //
    // If we are starting in explorer view, then the command line
    // has a "/e, " before the quoted diretory.
    //

    if (bExplore)
    {
        StringCchCopy(szCmdLine, ARRAYSIZE(szCmdLine), TEXT("explorer.exe /e, \""));
    }
    else
    {
        StringCchCopy(szCmdLine, ARRAYSIZE(szCmdLine), TEXT("explorer.exe \""));
    }

    StringCchCat(szCmdLine, ARRAYSIZE(szCmdLine), szPath);
    StringCchCat(szCmdLine, ARRAYSIZE(szCmdLine), TEXT("\""));

    // Initialize process startup info
    STARTUPINFO si = { 0 };
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_SHOWNORMAL;

    // Start explorer
    PROCESS_INFORMATION pi = { 0 };
    if (CreateProcess(NULL, szCmdLine, NULL, NULL, FALSE, NORMAL_PRIORITY_CLASS, NULL, NULL, &si, &pi))
    {
        // Close the process and thread handles
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
}

int CTray::_GetQuickLaunchID()
{
    int iQLBandID = -1;

    DWORD dwBandID;
    for (int i = 0; (iQLBandID == -1) && SUCCEEDED(_ptbs->EnumBands(i, &dwBandID)); i++)
    {
        if (BandSite_TestBandCLSID(_ptbs, dwBandID, CLSID_ISFBand) == S_OK)
        {
            IUnknown* punk;
            if (SUCCEEDED(_ptbs->GetBandObject(dwBandID, IID_PPV_ARG(IUnknown, &punk))))
            {
                VARIANTARG v = { 0 };
                v.vt = VT_I4;
                if (SUCCEEDED(IUnknown_Exec(punk, &CLSID_ISFBand, 1, 0, NULL, &v)))
                {
                    if ((v.vt == VT_I4) && (CSIDL_APPDATA == (DWORD)v.lVal))
                    {
                        iQLBandID = (int)dwBandID;
                    }
                }
                punk->Release();
            }
        }
    }

    return iQLBandID;
}

int CTray::_ToggleQL(int iVisible)
{
    int iQLBandID = _GetQuickLaunchID();

    bool fOldVisible = (-1 != iQLBandID);
    bool fNewVisible = (0 != iVisible);

    if ((iVisible != -1) && (fNewVisible != fOldVisible))
    {
        if (fNewVisible)
        {
            LPITEMIDLIST pidl;
            if (SUCCEEDED(SHGetSpecialFolderLocation(NULL, CSIDL_APPDATA, &pidl)))
            {
                TCHAR szPath[MAX_PATH];
                SHGetPathFromIDList(pidl, szPath);
                PathCombine(szPath, szPath, L"Microsoft\\Internet Explorer\\Quick Launch");
                ILFree(pidl);
                pidl = ILCreateFromPath(szPath);

                if (pidl)
                {
                    IFolderBandPriv* pfbp;
                    // create an ISF band to show folders as hotlinks
                    if (SUCCEEDED(CoCreateInstance(CLSID_ISFBand, NULL, CLSCTX_INPROC, IID_PPV_ARG(IFolderBandPriv, &pfbp))))
                    {
                        IShellFolderBand* psfb;
                        if (SUCCEEDED(pfbp->QueryInterface(IID_PPV_ARG(IShellFolderBand, &psfb))))
                        {
                            if (SUCCEEDED(psfb->InitializeSFB(NULL, pidl)))
                            {
                                pfbp->SetNoText(TRUE);

                                VARIANTARG v;
                                v.vt = VT_I4;
                                v.lVal = CSIDL_APPDATA;
                                IUnknown_Exec(psfb, &CLSID_ISFBand, 1, 0, &v, NULL);

                                v.lVal = UEMIND_SHELL;  // UEMIND_SHELL/BROWSER
                                IUnknown_Exec(psfb, &CGID_ShellDocView, SHDVID_UEMLOG, 0, &v, NULL);

                                IDeskBand* ptb;
                                if (SUCCEEDED(pfbp->QueryInterface(IID_PPV_ARG(IDeskBand, &ptb))))
                                {
                                    HRESULT hr = _ptbs->AddBand(ptb);
                                    if (SUCCEEDED(hr))
                                    {
                                        _ptbs->SetBandState(ShortFromResult(hr), BSSF_NOTITLE, BSSF_NOTITLE);
                                    }
                                    ptb->Release();
                                }
                            }
                            psfb->Release();
                        }
                        pfbp->Release();
                    }

                    ILFree(pidl);
                }
            }
        }
        else
        {
            int iBandID;
            do {
                iBandID = _GetQuickLaunchID();
                if (iBandID != -1)
                {
                    _ptbs->RemoveBand(iBandID);
                }
            } while (iBandID != -1);
        }
    }

    return iQLBandID;
}

void GiveDesktopFocus()
{
    SetForegroundWindow(v_hwndDesktop);
    SendMessage(v_hwndDesktop, DTM_UIACTIVATEIO, (WPARAM)TRUE, /*dtb*/0);
}

/*----------------------------------------------------------
Purpose: loads the given resource string and executes it.
         The resource string should follow this format:

         "program.exe>parameters"

         If there are no parameters, the format should simply be:

         "program.exe"

*/
void _ExecResourceCmd(UINT ids)
{
    TCHAR szCmd[2 * MAX_PATH];

    if (LoadString(g_hinstCabinet, ids, szCmd, SIZECHARS(szCmd)))
    {
        SHELLEXECUTEINFO sei = { 0 };

        // Find list of parameters (if any)
        LPTSTR pszParam = StrChr(szCmd, TEXT('>'));

        if (pszParam)
        {
            // Replace the '>' with a null terminator
            *pszParam = 0;
            pszParam++;
        }

        sei.cbSize = sizeof(sei);
        sei.nShow = SW_SHOWNORMAL;
        sei.lpFile = szCmd;
        sei.lpParameters = pszParam;
        ShellExecuteEx(&sei);
    }
}

void CTray::_RefreshStartMenu()
{
    _stb.ExecRefresh();
    _RefreshSettings();
    _UpdateBandSiteStyle();
}

BOOL CTray::_CanMinimizeAll()
{
    return (_hwndTasks && SendMessage(_hwndTasks, TBC_CANMINIMIZEALL, 0, 0));
}

BOOL CTray::_MinimizeAll(BOOL fPostRaiseDesktop)
{
    BOOL fRet = FALSE;

    if (_hwndTasks)
    {
        fRet = (BOOL)SendMessage(_hwndTasks, TBC_MINIMIZEALL, (WPARAM)_hwnd, (LPARAM)fPostRaiseDesktop);
    }

    return fRet;
}

extern void _UpdateNotifySetting(BOOL fNotifySetting);

//
//  Due to the weirdness of PnP, if the eject request occurs on a thread
//  that contains windows, the eject stalls for 15 seconds.  So do it
//  on its own thread.
//
DWORD CALLBACK _EjectThreadProc(LPVOID lpThreadParameter)
{
    CM_Request_Eject_PC();
    return 0;
}

// EXEX-VISTA: Validated.
BOOL _AllowLockWorkStation()
{
    DWORD dwData;
    BOOL fSuccess = FALSE;

    HKEY hkey;
    if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\system"), 0, KEY_QUERY_VALUE, &hkey))
    {
        DWORD cbData = sizeof(DWORD);
        if (!RegQueryValueEx(hkey, TEXT("DisableLockWorkstation"), 0, 0, (LPBYTE)&dwData, &cbData))
        {
            fSuccess = dwData != 0;
        }
        RegCloseKey(hkey);
    }

    if (!RegOpenKeyEx(HKEY_CURRENT_USER, REGSTR_POLICIES_EXPLORER, 0, KEY_QUERY_VALUE, &hkey))
    {
        DWORD cbData = sizeof(DWORD);
        if (!RegQueryValueEx(hkey, TEXT("DisableLockWorkstation"), 0, 0, (LPBYTE)&dwData, &cbData))
        {
            fSuccess = fSuccess || dwData;
        }
        RegCloseKey(hkey);
    }

    return fSuccess ? FALSE : TRUE;
}

// EXEX-VISTA: Validated.
HRESULT DisplayAltTab()
{
    HWND hWnd = FindWindow(TEXT("AltTab_KeyHookWnd"), NULL);

    if (!hWnd)
    {
        return E_FAIL;
    }

    // EXEX-VISTA TODO: Document.
    if (!PostMessageW(hWnd, WM_USER, 0x50494C46, 0))
    {
        return E_FAIL;
    }

    return S_OK;
}

// EXEX-VISTA: Validated.
void CTray::_ActivateWindowSwitcher()
{
    if (!IsCompositionActive() || FAILED(DwmpStartOrStopFlip3D()))
    {
        DisplayAltTab();
    }
}

#define SLS_OFF (0)
#define SLS_ON  (1)
#define SLS_TOGGLE (2)

// EXEX-VISTA: Validated.
void CTray::_SetLockState(UINT newVal)
{
    bool fNewVal = newVal == SLS_TOGGLE
        ? !_fCanSizeMove
        : newVal == SLS_OFF;

    newVal = fNewVal;

    SHSetValueW(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("TaskbarSizeMove"),
        REG_DWORD, &newVal, sizeof(DWORD));

    _RefreshSettings();
    _UpdateBandSiteStyle();
    _AccountAllBandsForTaskbarSizingBar();
}

// EXEX-VISTA: Validated.
void CTray::_InvokeSearch()
{
    SHELLEXECUTEINFOW execInfo = { 0 };

    execInfo.hwnd = _hwnd;
    execInfo.cbSize = sizeof(execInfo);
    execInfo.nShow = SW_SHOWNORMAL;
    execInfo.lpFile = L"search-ms:";
    ShellExecuteExW(&execInfo);
}

// EXEX-VISTA: Validated.
void CTray::_Command(UINT idCmd, BOOL fFromNotifArea)
{
#ifdef DEAD_CODE
    INSTRUMENT_ONCOMMAND(SHCNFI_TRAYCOMMAND, _hwnd, idCmd);

    switch (idCmd) {

        // EXEX-VISTA: Validated.
    case IDM_CONTROLS:
    case IDM_PRINTERS:
        _ShowFolder(_hwnd,
            idCmd == IDM_CONTROLS ? CSIDL_CONTROLS : CSIDL_PRINTERS, FALSE);
        break;

        // EXEX-VISTA: Validated.
    case IDM_WINDOWSWITCHER:
        _ActivateWindowSwitcher();
        break;

        // EXEX-VISTA: Validated.
    case IDM_EJECTPC:
        // Must use SHCreateThread and not a queued workitem because
        // a workitem might inherit a thread that has windows on it.
        // CTF_INSIST: In emergency, eject synchronously.  This stalls
        // for 15 seconds but it's better than nothing.
        SHCreateThread(_EjectThreadProc, NULL, CTF_INSIST, NULL);
        break;

        // EXEX-VISTA: Validated.
    case IDM_LOGOFF:
        // Let the desktop get a chance to repaint to get rid of the
        // start menu bits before bringing up the logoff dialog box.
        UpdateWindow(_hwnd);
        Sleep(100);

        _SaveTrayAndDesktop();
        LogoffWindowsDialog(v_hwndDesktop);
        break;

        // EXEX-VISTA: Validated.
    case IDM_MU_DISCONNECT:
        // Do the same sleep as above for the same reason.
        UpdateWindow(_hwnd);
        Sleep(100);
        DisconnectWindowsDialog(v_hwndDesktop);
        break;

        // EXEX-VISTA: Validated.
    case IDM_EXITWIN:
        // Do the same sleep as above for the same reason.
        UpdateWindow(_hwnd);
        Sleep(100);

        _DoExitWindows(v_hwndDesktop, FALSE, fFromNotifArea);
        break;

        // EXEX-VISTA: Validated.
    case IDM_TOGGLEDESKTOP:
        _RaiseDesktop(!g_fDesktopRaised, TRUE);
        break;

        // EXEX-VISTA: Validated.
    case IDM_FILERUN:
        _RunDlg(FALSE);
        break;

        // EXEX-VISTA: Validated.
    case IDM_MINIMIZEALLHOTKEY:
        _HandleGlobalHotkey(GHID_MINIMIZEALL);
        break;

#ifdef DEBUG
    case IDM_SIZEUP:
    {
        RECT rcView;
        GetWindowRect(_hwndRebar, &rcView);
        MapWindowPoints(HWND_DESKTOP, _hwnd, (LPPOINT)&rcView, 2);
        rcView.bottom -= 18;
        SetWindowPos(_hwndRebar, NULL, 0, 0, RECTWIDTH(rcView), RECTHEIGHT(rcView), SWP_NOMOVE | SWP_NOZORDER);
    }
    break;

    case IDM_SIZEDOWN:
    {
        RECT rcView;
        GetWindowRect(_hwndRebar, &rcView);
        MapWindowPoints(HWND_DESKTOP, _hwnd, (LPPOINT)&rcView, 2);
        rcView.bottom += 18;
        SetWindowPos(_hwndRebar, NULL, 0, 0, RECTWIDTH(rcView), RECTHEIGHT(rcView), SWP_NOMOVE | SWP_NOZORDER);
    }
    break;
#endif

    // EXEX-VISTA: Validated.
    case IDM_MINIMIZEALL:
        // minimize all window
        _MinimizeAll(FALSE);
        _fUndoEnabled = TRUE;
        break;

        // EXEX-VISTA: Validated.
    case IDM_UNDO:
        _RestoreWindowPositions(FALSE);
        break;

        // EXEX-VISTA: Validated.
    case IDM_SETTIME:
        // run the default applet in timedate.cpl
        ShellExecuteW(_hwnd, NULL, TEXT("timedate.cpl"), NULL, NULL, SW_SHOWNORMAL);
        break;

        // EXEX-VISTA: Validated.
    case IDM_NOTIFYCUST:
        DoProperties(TPF_INVOKECUSTOMIZE);
        break;

        // EXEX-VISTA: Validated.
    case IDM_LOCKTASKBAR:
    {
        // @MOD Skipped telemetry ShellTraceId_Taskbar_LockState_ChangeNotify_Start

        _SetLockState(SLS_TOGGLE);

        // @MOD Skipped telemetry ShellTraceId_Taskbar_LockState_ChangeNotify_Stop
    }
    break;

    // EXEX-VISTA: Validated.
    case IDM_SHOWTASKMAN:
        RunSystemMonitor();
        break;

        // EXEX-VISTA: Validated. (all 3 cases)
    case IDM_CASCADE:
    case IDM_VERTTILE:
    case IDM_HORIZTILE:
        if (_CanTileAnyWindows())
        {
            UINT idRes;
            switch (idCmd)
            {
            case IDM_CASCADE:
                idRes = IDS_CASCADE;
                break;
            case IDM_HORIZTILE:
                idRes = IDS_SHOW_STACKED;
                break;
            case IDM_VERTTILE:
                idRes = IDS_SHOW_SIDE_BY_SIDE;
                break;
            }

            SaveWindowPositions(idRes);

            _AppBarNotifyAll(NULL, ABN_WINDOWARRANGE, NULL, TRUE);

            if (idCmd == IDM_CASCADE)
            {
                CascadeWindows(GetDesktopWindow(), 0, NULL, 0, NULL);
            }
            else
            {
                TileWindows(GetDesktopWindow(), ((idCmd == IDM_VERTTILE) ?
                    MDITILE_VERTICAL : MDITILE_HORIZONTAL), NULL, 0, NULL);
            }

            // do it *before* ABN_xxx so don't get 'indirect' moves
            // REVIEW or should it be after?
            // CheckWindowPositions();
            _fUndoEnabled = FALSE;
            SetTimer(_hwnd, IDT_ENABLEUNDO, 500, NULL);

            _AppBarNotifyAll(NULL, ABN_WINDOWARRANGE, NULL, FALSE);
        }
        break;

        // EXEX-VISTA: Validated.
    case IDM_TRAYPROPERTIES:
        if (fFromNotifArea)
        {
            DoProperties(TPF_NOTIFAREAPAGE);
        }
        else
        {
            DoProperties(TPF_TASKBARPAGE);
        }
        break;

        // EXEX-VISTA: Validated.
    case IDM_HELPSEARCH:
    {
        CoInitialize(NULL);

        IHxHelpPane* pHelpPane = NULL;
        if (SUCCEEDED(CoCreateInstance(CLSID_HxHelpPane, NULL, CLSCTX_SERVER, IID_PPV_ARGS(&pHelpPane))))
        {
            CComBSTR bstrUri(TEXT("mshelp://help/?id=home"));

            if (bstrUri)
            {
                pHelpPane->DisplayTask(bstrUri.m_str);
            }
        }

        CoUninitialize();

        if (pHelpPane)
        {
            pHelpPane->Release();
        }

        break;
    }

    // NB The Alt-s comes in here.
    // EXEX-VISTA: Validated.
    case IDC_KBSTART:
        SetForegroundWindow(_hwnd);
        // This pushes the start button and causes the start menu to popup.
        SendMessage(_stb._hwndStart, BM_SETSTATE, TRUE, 0);
        // This forces the button back up.
        SendMessage(_stb._hwndStart, BM_SETSTATE, FALSE, 0);
        break;

        // EXEX-VISTA: Validated.
    case IDC_ASYNCSTART:
#if 0 // (for testing UAssist locking code)
        UEMFireEvent(&UEMIID_SHELL, UEME_DBSLEEP, UEMF_XEVENT, -1, (LPARAM)10000);
#endif
#ifdef DEBUG
        if (GetAsyncKeyState(VK_SHIFT) < 0)
        {
            UEMFireEvent(&UEMIID_SHELL, UEME_CTLSESSION, UEMF_XEVENT, TRUE, -1);
            _RefreshStartMenu();
        }
#endif

        // Make sure the button is down.
        // DebugMsg(DM_TRACE, "c.twp: IDC_START.");

        // Make sure the Start button is down.
        if (!_bMainMenuInit && _stb.IsButtonPushed())
        {
            // DebugMsg(DM_TRACE, "c.twp: Start button down.");
            // Set the focus.
            _SetFocus(_stb._hwndStart);
            _ToolbarMenu();
        }
        break;

        // NB LButtonDown on the Start button come in here.
        // Space-bar stuff also comes in here.
        // EXEX-VISTA: Validated.
    case IDC_START:
        // User gets a bit confused with space-bar tuff (the popup ends up
        // getting the key-up and beeps).
        PostMessage(_hwnd, WM_COMMAND, IDC_ASYNCSTART, 0);
        break;

        // EXEX-VISTA: Validated.
    case FCIDM_FINDFILES:
        _InvokeSearch();
        break;

        // EXEX-VISTA: Validated.
    case FCIDM_FINDCOMPUTER:
        SHFindComputer(NULL, NULL);
        break;

        // EXEX-VISTA: Validated.
    case FCIDM_REFRESH:
        _RefreshStartMenu();
        break;

        // EXEX-VISTA: Validated.
    case FCIDM_NEXTCTL:
    {
        MSG msg = { 0, WM_KEYDOWN, VK_TAB };
        HWND hwndFocus = GetFocus();

        // Since we are Tab or Shift Tab we should turn the focus rect on.
        //
        // Note: we don't need to do this in the GiveDesktopFocus cases below,
        // but in those cases we're probably already in the UIS_CLEAR UISF_HIDEFOCUS
        // state so this message is cheap to send.
        //
        SendMessage(_hwnd, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR,
            UISF_HIDEFOCUS), 0);

        BOOL fShift = GetAsyncKeyState(VK_SHIFT) < 0;

        if (hwndFocus && (IsChildOrHWND(_stb._hwndStart, hwndFocus)))
        {
            if (fShift)
            {
                // gotta deactivate manually
                GiveDesktopFocus();
            }
            else
            {
                IUnknown_UIActivateIO(_ptbs, TRUE, &msg);
            }
        }
        else if (hwndFocus && (IsChildOrHWND(_hwndNotify, hwndFocus)))
        {
            if (fShift)
            {
                IUnknown_UIActivateIO(_ptbs, TRUE, &msg);
            }
            else
            {
                GiveDesktopFocus();
            }
        }
        else
        {
            if (IUnknown_TranslateAcceleratorIO(_ptbs, &msg) != S_OK)
            {
                if (fShift)
                {
                    _SetFocus(_stb._hwndStart);
                }
                else
                {
                    // if you tab forward out of the bands, the next focus guy is the tray notify set
                    _SetFocus(_hwndNotify);
                }
            }
        }
    }
    break;

    // EXEX-VISTA: Validated.
    case IDM_MU_SECURITY:
        MuSecurity();
        break;

        // EXEX-VISTA: Validated.
    case IDM_LOCKWORKSTATION:
    {
        if (_AllowLockWorkStation())
        {
            LockWorkStation();
        }

        break;
    }
    }
#else
    unsigned int v5; // eax
    // eax
    LONG v7; // eax
    HWND Focus; // edi
    // [esp-10h] [ebp-38h]
    RECT rcView; // [esp+18h] [ebp-10h] BYREF

    if (idCmd <= 0x1A5)
    {
        if (idCmd == 0x1A5)
        {
            DoProperties(4u);
        }
        else if (idCmd > 0x198)
        {
            switch (idCmd)
            {
                case 0x19Au:
                    SHCreateThread(_EjectThreadProc, nullptr, 1u, nullptr);
                    break;
                case 0x19Bu:
                    _ActivateWindowSwitcher();
                    break;
                case 0x19Du:
                    if (fFromNotifArea)
                        DoProperties(8);
                    else
                        DoProperties(1);
                    break;
                case 0x19Fu:
                    _MinimizeAll(0);
                    _fUndoEnabled = 1;
                    break;
                case 0x1A0u:
                    _RestoreWindowPositions(0);
                    break;
                case 0x1A3u:
                    _HandleGlobalHotkey(0x1F5u);
                    break;
                case 0x1A4u:
                    RunSystemMonitor();
                    break;
            }
        }
        else if (idCmd == 408)
        {
            ShellExecuteW(_hwnd, nullptr, L"timedate.cpl", nullptr, nullptr, 1);
        }
        else if (idCmd > 0x191)
        {
            if (idCmd == 402)
            {
                UpdateWindow(_hwnd);
                Sleep(0x64u);
                //g_fEndSessionInitiated = 1;
                _SaveTrayAndDesktop();
                LogoffWindowsDialog(v_hwndDesktop);
            }
            else if (idCmd <= 0x195)
            {
                if (_CanTileAnyWindows())
                {
                    if (idCmd == 403)
                        v5 = 535;
                    else
                        v5 = 2 * (idCmd != 404) + 536;

                    SaveWindowPositions(v5);
                    _AppBarNotifyAll(nullptr, 3u, nullptr, 1);
                    if (idCmd == 403)
                    {
                        CascadeWindows(GetDesktopWindow(), 0, nullptr, 0, nullptr);
                    }
                    else
                    {
                        TileWindows(GetDesktopWindow(), idCmd != 404, nullptr, 0, nullptr);
                    }
                    this->_fUndoEnabled = 0;
                    SetTimer(this->_hwnd, 0x12u, 0x1F4u, nullptr);
                    _AppBarNotifyAll(nullptr, 3u, nullptr, 0);
                }
            }
            else if (idCmd == 407)
            {
                _RaiseDesktop(g_fDesktopRaised == 0, 1);
            }
        }
        else
        {
            switch (idCmd)
            {
            case 0x191u:
                _RunDlg(0);
                break;
            case 0x130u:
                PostMessageW(_hwnd, 0x111u, 0x132u, 0);
                break;
            case 0x131u:
                SetForegroundWindow(_hwnd);
                SendMessageW(_stb._hwndStart, 0xF3u, 1u, 0);
                SendMessageW(_stb._hwndStart, 0xF3u, 0, 0);
                break;
            case 0x132u:
                if (GetAsyncKeyState(VK_SHIFT) < 0)
                {
                    UEMFireEvent(&CLSID_ActiveDesktop, 40, 0, 1, -1);
                    _RefreshStartMenu();
                }
                if (!_bMainMenuInit && _stb.IsButtonPushed())
                {
                    _SetFocus(this->_stb._hwndStart);
                    _ToolbarMenu();
                }
                break;
            }
        }
        return;
    }

    if (idCmd > 0x205)
    {
        if (idCmd == 5000)
        {
            UpdateWindow(this->_hwnd);
            Sleep(0x64u);
            DisconnectWindowsDialog(v_hwndDesktop);
            return;
        }
        if (idCmd == 5001)
        {
            if (SHGetMachineInfo(3))
            {
                WinStationSetInformationW(nullptr, -1, WinStationNtSecurity, nullptr, 0);
            }
            return;
        }
        if (idCmd == 41008)
        {
            MSG msg = { nullptr, WM_KEYDOWN, VK_TAB };

            memset(&rcView, 0, sizeof(rcView));
            Focus = GetFocus();
            SendMessageW(this->_hwnd, 0x128u, 0x10002u, 0);
            fFromNotifArea = GetAsyncKeyState(16) < 0;
            if (Focus)
            {
                if (IsChildOrHWND(this->_stb._hwndStart, Focus))
                {
                    if (!fFromNotifArea)
                    {
                    LABEL_84:
                        IUnknown_UIActivateIO(this->_ptbs, 1, &msg);
                        return;
                    }
                LABEL_87:
                    GiveDesktopFocus();
                    return;
                }
                if (IsChildOrHWND(this->_hwndNotify, Focus))
                {
                    if (fFromNotifArea)
                    {
                        goto LABEL_84;
                    }
                    goto LABEL_87;
                }
            }
            if (IUnknown_TranslateAcceleratorIO(this->_ptbs, &msg))
            {
                if (fFromNotifArea)
                    _SetFocus(this->_stb._hwndStart);
                else
                    _SetFocus(this->_hwndNotify);
            }
            return;
        }

        switch (idCmd)
        {
            case 0xA065u:
                CTray::_RefreshStartMenu();
                break;
            case 0xA085u:
                CTray::_InvokeSearch();
                break;
            case 0xA086u:
                SHFindComputer(nullptr, nullptr);
                break;
        }
        return;
    }

    switch (idCmd)
    {
        case 0x205u:
            if (_AllowLockWorkStation())
            {
                LockWorkStation();
            }
            break;
        case 0x1A8u:
            // Skipped telemetry LockState_ChangeNotify_Start
            _SetLockState(2u);
            // Skipped telemetry LockState_ChangeNotify_Stop
            return;
        case 0x1ABu:
            GetWindowRect(_hwndRebar, &rcView);
            MapWindowRect(nullptr, _hwnd, &rcView);
            rcView.bottom -= 18;
            SetWindowPos(_hwndRebar, nullptr, 0, 0, rcView.right - rcView.left, rcView.bottom - rcView.top, 6u);
            return;
        case 0x1ACu:
            GetWindowRect(_hwndRebar, &rcView);
            MapWindowRect(nullptr, _hwnd, &rcView);
            rcView.bottom += 18;
            SetWindowPos(_hwndRebar, nullptr, 0, 0, rcView.right - rcView.left, rcView.bottom - rcView.top, 6u);
            return;
    }

    if (idCmd == 503)
    {
        CoInitialize(nullptr);

        IHxHelpPane* pxhp = nullptr;
        if (CoCreateInstance(CLSID_HxHelpPane, nullptr, 0x17u, IID_PPV_ARGS(&pxhp)) >= 0)
        {
            CComBSTR bstrUri(TEXT("mshelp://help/?id=home"));
            if (bstrUri)
            {
                pxhp->DisplayTask(bstrUri.m_str);
            }
        }
        CoUninitialize();
        if (pxhp)
        {
            pxhp->Release();
        }
        return;
    }
    if (idCmd != 505)
    {
        if (idCmd == 506)
        {
            UpdateWindow(_hwnd);
            Sleep(0x64u);
            //g_fEndSessionInitiated = 1;
            _DoExitWindows(v_hwndDesktop, 0, fFromNotifArea);
            return;
        }
        if (idCmd != 510)
            return;
    }
    _ShowFolder(_hwnd, idCmd == 505 ? CSIDL_CONTROLS : CSIDL_PRINTERS, 0);
#endif
}

//// Start menu/Tray tab as a drop target

HRESULT CStartDropTarget::_GetStartMenuDropTarget(IDropTarget** pptgt)
{
    HRESULT hr = E_FAIL;
    *pptgt = NULL;

    LPITEMIDLIST pidlStart = SHCloneSpecialIDList(NULL, CSIDL_STARTMENU, TRUE);

    if (pidlStart)
    {
        IShellFolder* psf = BindToFolder(pidlStart);
        if (psf)
        {
            hr = psf->CreateViewObject(_ptray->_hwnd, IID_PPV_ARG(IDropTarget, pptgt));
            psf->Release();
        }

        ILFree(pidlStart);
    }
    return hr;
}


STDMETHODIMP CDropTargetBase::QueryInterface(REFIID riid, void** ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CDropTargetBase, IDropTarget),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

STDMETHODIMP_(ULONG) CDropTargetBase::AddRef()
{
    return 2;
}

STDMETHODIMP_(ULONG) CDropTargetBase::Release()
{
    return 1;
}

STDMETHODIMP CDropTargetBase::DragEnter(IDataObject* pdtobj, DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    _ptray->_SetUnhideTimer(ptl.x, ptl.y);

    HWND hwndLock = _ptray->_hwnd;  // no clippy

    _DragEnter(hwndLock, ptl, pdtobj);

    return S_OK;
}

// EXEX-VISTA: SLIGHTLY MODIFIED. Revalidate later.
STDMETHODIMP CDropTargetBase::DragOver(DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    _ptray->_SetUnhideTimer(ptl.x, ptl.y);
    _DragMove(_ptray->_stb._hwndStart, ptl);

    return S_OK;
}

STDMETHODIMP CDropTargetBase::DragLeave()
{
    DAD_DragLeave();
    return S_OK;
}

STDMETHODIMP CDropTargetBase::Drop(IDataObject* pdtobj, DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    DAD_DragLeave();
    return S_OK;
}

//
//  There are two different policies for the Start Button depending on
//  whether we are in Classic mode or Personal (New Start Pane) mode.
//
//  Classic mode:  Drops onto the Start Button are treated as if they
//  were drops into the CSIDL_STARTMENU folder.
//
//  Personal mode:  Drops onto the Start Button are treated as if they
//  were drops into the pin list.
//

CStartDropTarget::CStartDropTarget() : CDropTargetBase(IToClass(CTray, _dtStart, this))
{
}

CTrayDropTarget::CTrayDropTarget() : CDropTargetBase(IToClass(CTray, _dtTray, this))
{
}

STDMETHODIMP CTrayDropTarget::DragEnter(IDataObject* pdtobj, DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    *pdwEffect = DROPEFFECT_NONE;
    return CDropTargetBase::DragEnter(pdtobj, grfKeyState, ptl, pdwEffect);
}

STDMETHODIMP CTrayDropTarget::DragOver(DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    *pdwEffect = DROPEFFECT_NONE;
    return CDropTargetBase::DragOver(grfKeyState, ptl, pdwEffect);
}

void CStartDropTarget::_StartAutoOpenTimer(POINTL* pptl)
{
    POINT pt = { pptl->x, pptl->y };
    RECT rc;
    //Make sure it really is in the start menu..
    GetWindowRect(_ptray->_stb._hwndStart, &rc);
    if (PtInRect(&rc, pt))
    {
        SetTimer(_ptray->_hwnd, IDT_STARTMENU, 1000, NULL);
    }
}

STDMETHODIMP CStartDropTarget::DragEnter(IDataObject* pdtobj, DWORD grfKeyState, POINTL ptl, DWORD* pdwEffect)
{
    HRESULT hr = S_OK;

    // default to not allowing drops
    _dwEffectsAllowed = DROPEFFECT_NONE;

    if (Tray_StartPanelEnabled())
    {
        // if we've disabled dragging and dropping, don't do anything, but if only the pin list is restricted, then still start the timer
        if (!IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOCHANGESTARMENU, TEXT("Advanced"), TEXT("Start_EnableDragDrop"), ROUS_KEYALLOWS | ROUS_DEFAULTALLOW))
        {
            // Personal mode: Treat it as an add to the pin list.
            // IsPinnable checks REST_NOSMPINNEDLIST
            if (_ptray->_psmpin && _ptray->_psmpin->IsPinnable(pdtobj, SMPINNABLE_REJECTSLOWMEDIA) == S_OK)
            {
                _dwEffectsAllowed = DROPEFFECT_LINK;
            }

            *pdwEffect &= _dwEffectsAllowed;

            // Always start the AutoOpen timer because once we open, the user
            // can drop onto other things which may have different drop policies
            // from the pin list.
            _StartAutoOpenTimer(&ptl);
        }
    }
    else
    {
        if (!IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOCHANGESTARMENU, TEXT("Advanced"), TEXT("StartMenuChange"), ROUS_KEYALLOWS | ROUS_DEFAULTALLOW))
        {
            // Classic mode: Treat it as a drop on the Start Menu folder.
            IDropTarget* ptgt;
            _dwEffectsAllowed = DROPEFFECT_LINK;

            hr = _GetStartMenuDropTarget(&ptgt);
            if (SUCCEEDED(hr))
            {
                // Check to make sure that we're going to accept the drop before we expand the start menu.
                ptgt->DragEnter(pdtobj, grfKeyState, ptl,
                    pdwEffect);

                // DROPEFFECT_NONE means it ain't gonna work, so don't popup the Start Menu.
                if (*pdwEffect != DROPEFFECT_NONE)
                {
                    _StartAutoOpenTimer(&ptl);
                }

                ptgt->DragLeave();
                ptgt->Release();
            }
        }
    }

    CDropTargetBase::DragEnter(pdtobj, grfKeyState, ptl, pdwEffect);
    return hr;
}

STDMETHODIMP CStartDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    *pdwEffect = (_dwEffectsAllowed & DROPEFFECT_LINK);
    return CDropTargetBase::DragOver(grfKeyState, pt, pdwEffect);
}

STDMETHODIMP CStartDropTarget::DragLeave()
{
    KillTimer(_ptray->_hwnd, IDT_STARTMENU);
    return CDropTargetBase::DragLeave();
}

STDMETHODIMP CStartDropTarget::Drop(IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    KillTimer(_ptray->_hwnd, IDT_STARTMENU);

    HRESULT hr;

    if (Tray_StartPanelEnabled())
    {
        // Personal mode: Treat it as an add to the pin list.
        LPITEMIDLIST pidl;
        if (_ptray->_psmpin && IsPinnable(pdtobj, SMPINNABLE_REJECTSLOWMEDIA, &pidl) == S_OK)
        {
            // Delete it from the pin list if it's already there because
            // we want to move it to the bottom.
            _ptray->_psmpin->Modify(pidl, NULL);
            // Now add it to the bottom.
            _ptray->_psmpin->Modify(NULL, pidl);
            ILFree(pidl);
            hr = S_OK;
        }
        else
        {
            hr = E_FAIL;
        }
    }
    else
    {
        IDropTarget* pdrop;
        hr = _GetStartMenuDropTarget(&pdrop);
        if (SUCCEEDED(hr))
        {
            if (!Tray_StartPanelEnabled())
            {
                POINTL ptDrop = { 0, 0 };
                DWORD grfKeyStateDrop = 0;

                *pdwEffect &= DROPEFFECT_LINK;

                pdrop->DragEnter(pdtobj, grfKeyStateDrop, ptDrop, pdwEffect);
                hr = pdrop->Drop(pdtobj, grfKeyStateDrop, ptDrop, pdwEffect);

                pdrop->DragLeave();
            }
            pdrop->Release();
        }
    }

    DAD_DragLeave();

    return hr;
}

void CTray::_RegisterDropTargets()
{
    THR(RegisterDragDrop(_stb._hwndStart, &_dtStart));
    THR(RegisterDragDrop(_hwnd, &_dtTray));

    //// It is not a serious error if this fails; it just means that
    //// drag/drop to the Start Button will not add to the pin list
    //CoCreateInstance(CLSID_StartMenuPin, NULL, CLSCTX_INPROC_SERVER,
    //    IID_PPV_ARGS(&_psmpin));
}

void CTray::_RevokeDropTargets()
{
    RevokeDragDrop(_stb._hwndStart);
    RevokeDragDrop(_hwnd);
    ATOMICRELEASET(_psmpin, IStartMenuPin);
}

void CTray::_HandleGlobalHotkey(WPARAM wParam)
{
    INSTRUMENT_HOTKEY(SHCNFI_GLOBALHOTKEY, wParam);

    switch (wParam)
    {
        // EXEX-VISTA: Validated.
        case GHID_RUN:
            _RunDlg(TRUE);
            break;

        // EXEX-VISTA: Validated.
        case GHID_MINIMIZEALL:
            _MinimizeAll(FALSE);
            SetForegroundWindow(v_hwndDesktop);
            break;

        // EXEX-VISTA: Validated.
        case GHID_UNMINIMIZEALL:
            _RestoreWindowPositions(FALSE);
            break;

        // EXEX-VISTA: Validated.
        case GHID_HELP:
            _Command(IDM_HELPSEARCH, FALSE);
            break;

        // EXEX-VISTA: Validated.
        case GHID_DESKTOP:
            _RaiseDesktop(!g_fDesktopRaised, TRUE);
            break;

        // EXEX-VISTA: Validated.
        case GHID_TRAYNOTIFY:
            SwitchToThisWindow(_hwnd, TRUE);
            SetForegroundWindow(_hwnd);
            SendMessage(_hwndNotify, WM_CHANGEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEACCEL | UISF_HIDEFOCUS), 0);
            _SetFocus(_hwndNotify);
            break;

        // EXEX-VISTA: Validated.
        case GHID_EXPLORER:
            _ShowFolder(_hwnd, CSIDL_DRIVES, COF_CREATENEWWINDOW);
            break;

        // EXEX-VISTA: Validated.
        case GHID_FINDFILES:
            if (!SHRestricted(REST_NOFIND))
                _InvokeSearch();
            break;

        // EXEX-VISTA: Validated.
        case GHID_FINDCOMPUTER:
            if (!SHRestricted(REST_NOFIND))
                _Command(FCIDM_FINDCOMPUTER, FALSE);
            break;

        // EXEX-VISTA: Validated.
        case GHID_TASKTAB:
        case GHID_TASKSHIFTTAB:
            if (GetForegroundWindow() != _hwnd)
                SetForegroundWindow(_hwnd);
            SendMessage(_hwndTasks, TBC_TASKTAB, wParam == GHID_TASKTAB ? 1 : -1, 0L);
            break;

        // EXEX-VISTA: Validated.
        case GHID_SYSPROPERTIES:
            ShellExecute(_hwnd, NULL, TEXT("control.exe"), TEXT("system"), NULL, SW_SHOW);
            break;

        // EXEX-VISTA: Validated.
        default:
            // @NOTE Handles quick-launch hotkeys: Windows key + number to quickly open a program.
            // EXEX-VISTA TODO: Uncomment when implemented _HandleQLHotKey:
            // _HandleQLHotKey(wParam);
            break;
    }
}

void CTray::_UnregisterGlobalHotkeys()
{
    for (int i = GHID_FIRST; i < GHID_MAX; i++)
    {
        UnregisterHotKey(_hwnd, i);
    }
}

void CTray::_RegisterGlobalHotkeys()
{
    int i;
    // Are the Windows keys restricted?
    DWORD dwRestricted = SHRestricted(REST_NOWINKEYS);

    for (i = GHID_FIRST; i < GHID_MAX; i++)
    {
        // If the Windows Keys are Not restricted or it's not a Windows key
        if (!((HIWORD(GlobalKeylist[i - GHID_FIRST]) & MOD_WIN) && dwRestricted))
        {
            // Then register it.
            RegisterHotKey(_hwnd, i, HIWORD(GlobalKeylist[i - GHID_FIRST]), LOWORD(GlobalKeylist[i - GHID_FIRST]));
        }
    }
}

// EXEX-VISTA: Validated.
void CTray::_RaiseDesktop(BOOL fRaise, BOOL fRestoreWindows)
{
    if (v_hwndDesktop && (fRaise == !g_fDesktopRaised) && !_fProcessingDesktopRaise)
    {
        _fProcessingDesktopRaise = TRUE;
        BOOL fPostMessage = TRUE;

        if (fRaise)
        {
            HWND hwndFG = GetForegroundWindow();
            // If no window has focus then set focus to the tray
            if (hwndFG)
            {
                hwndFG = _hwnd;
            }

            if (!_hwndFocusBeforeRaise)
            {
                // See if the Foreground Window had a popup window
                _hwndFocusBeforeRaise = GetLastActivePopup(hwndFG);
            }
            if (!IsWindowVisible(_hwndFocusBeforeRaise))
            {
                _hwndFocusBeforeRaise = hwndFG;
            }

            // _MinimizeAll will save the windows positions synchronously, and will minimize the
            // the windows on a background thread
            _fMinimizedAllBeforeRaise = _CanMinimizeAll();
            if (_fMinimizedAllBeforeRaise)
            {
                fPostMessage = !_MinimizeAll(TRUE);
            }
        }
        else
        {
            if (fRestoreWindows)
            {
                HWND hwnd = _hwndFocusBeforeRaise;
                if (_fMinimizedAllBeforeRaise)
                {
                    // Since the windows are restored on a seperate thread, I want the make that the
                    // desktop is not raised until they are done, so the window restore thread will
                    // actually post the message for raising the desktop
                    fPostMessage = !_RestoreWindowPositions(TRUE);
                }

                SetForegroundWindow(hwnd);
                if (hwnd == _hwnd)
                {
                    _SetFocus(_stb._hwndStart);
                }
            }

            _hwndFocusBeforeRaise = NULL;
        }

        if (fPostMessage)
            PostMessage(v_hwndDesktop, DTM_RAISE, (WPARAM)_hwnd, fRaise ? DTRF_RAISE : DTRF_LOWER);
    }
}

// EXEX-VISTA: Validated.
void CTray::_OnDesktopState(LPARAM lParam)
{
    g_fDesktopRaised = (!(lParam & DTRF_LOWER));

    DAD_ShowDragImage(FALSE);       // unlock the drag sink if we are dragging.

    if (!g_fDesktopRaised)
    {
        HandleFullScreenApp(NULL);
    }
    else
    {
        // if the desktop is raised, we need to force the tray to be always on top
        // until it's lowered again
        _ResetZorder(0);
    }

    DAD_ShowDragImage(TRUE);       // unlock the drag sink if we are dragging.
    _fProcessingDesktopRaise = FALSE;
}

BOOL CTray::_ToggleLanguageBand(BOOL fShowIt)
{
    HRESULT hr = E_FAIL;

    DWORD dwBandID;
    BOOL fFound = FALSE;
    for (int i = 0; !fFound && SUCCEEDED(_ptbs->EnumBands(i, &dwBandID)); i++)
    {
        if (BandSite_TestBandCLSID(_ptbs, dwBandID, CLSID_MSUTBDeskBand) == S_OK)
        {
            fFound = TRUE;
        }
    }

    BOOL fShow = fFound;

    if (fShowIt && !fFound)
    {
        IDeskBand* pdb;
        HRESULT hr = CoCreateInstance(CLSID_MSUTBDeskBand, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARG(IDeskBand, &pdb));
        if (SUCCEEDED(hr))
        {
            hr = _ptbs->AddBand(pdb);
            fShow = TRUE;
            pdb->Release();
        }
    }
    else if (!fShowIt && fFound)
    {
        hr = _ptbs->RemoveBand(dwBandID);
        if (SUCCEEDED(hr))
        {
            fShow = FALSE;
        }
    }

    return fShow;
}

// Process the message by propagating it to all of our child windows
typedef struct
{
    UINT uMsg;
    WPARAM wP;
    LPARAM lP;
    CTray* ptray;
} CABPM;

BOOL CTray::PropagateEnumProc(HWND hwnd, LPARAM lParam)
{
    CABPM* ppm = (CABPM*)lParam;

    if (SHIsChildOrSelf(ppm->ptray->_hwndRebar, hwnd) == S_OK)
    {
        return TRUE;
    }
    SendMessage(hwnd, ppm->uMsg, ppm->wP, ppm->lP);
    return TRUE;
}

void CTray::_PropagateMessage(HWND hwnd, UINT uMessage, WPARAM wParam, LPARAM lParam)
{
    CABPM pm = { uMessage, wParam, lParam, this };

    ASSERT(hwnd != _hwndRebar);
    EnumChildWindows(hwnd, PropagateEnumProc, (LPARAM)&pm);
}

//
// Used by desktop2\deskhost.cpp to get to the tray properties.
//
void WINAPI Tray_DoProperties(DWORD dwFlags)
{
    c_tray.DoProperties(dwFlags);
}

//
//  This function is called whenever we detect a change to the default
//  browser registration in HKCR\http\shell\open\command.
//
//  For compatibility with old browsers, if the default browser (URL handler)
//  is not XP-aware, then auto-generate a StartMenuInternet client
//  registration and set it as the default.
//

void CTray::_MigrateOldBrowserSettings()
{
    // We want only one person to do this work per machine (though it doesn't
    // hurt to have more than one person do it; it's just pointless), so try
    // to filter out people who clearly didn't instigate the key change.
    //
    if (!_fIsDesktopLocked && _fIsDesktopConnected)
    {
        // If the user does not have write access then we can't migrate the
        // setting... (In which case there was nothing to migrate anyway
        // since you need to be administrator to change the default browser...)

        HKEY hkBrowser;
        DWORD dwDisposition;

        if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Clients\\StartMenuInternet"),
            0, NULL, REG_OPTION_NON_VOLATILE,
            KEY_READ | KEY_WRITE, NULL, &hkBrowser, &dwDisposition) == ERROR_SUCCESS)
        {
            TCHAR szCommand[MAX_PATH];
            DWORD cch = ARRAYSIZE(szCommand);

            // It is important that we use AssocQueryString to parse the
            // executable, because Netscape has a habit of registering
            // their path incorrectly (they forget to quote the space in
            // "Program Files") but AssocQueryString has special recovery
            // code to detect and repair that case...
            if (SUCCEEDED(AssocQueryString(ASSOCF_NOUSERSETTINGS,
                ASSOCSTR_EXECUTABLE, L"http",
                L"open", szCommand, &cch)) &&
                szCommand[0])
            {
                TCHAR szAppName[MAX_PATH];
                StringCchCopy(szAppName, ARRAYSIZE(szAppName), PathFindFileName(szCommand));

                // You might think that we need to special-case MSN Explorer,
                // since they shipped before XP RTM'd, and convert MSN6.EXE
                // to "MSN Explorer", but that's not true because
                // they never take over as the default http handler, so we
                // will never see them here!

                // Create a registration for the new default browser if necessary.
                // We keep our hands off once we see a DefaultIcon key, since that
                // proves that the application is XP-aware.

                // When IE Access is turned off, StartMenuInternet\IExplore.exe is
                // also removed so that IE will not appear in "Customize Start Menu"
                // dialog box. Do not migrate IE.

                HKEY hkClient;

                if (0 != lstrcmpi(szAppName, TEXT("IEXPLORE.EXE")) &&
                    RegCreateKeyEx(hkBrowser, szAppName, 0, NULL, REG_OPTION_NON_VOLATILE,
                        KEY_WRITE, NULL, &hkClient, &dwDisposition) == ERROR_SUCCESS)
                {
                    if (dwDisposition == REG_CREATED_NEW_KEY)
                    {
                        TCHAR szFriendly[MAX_PATH];
                        cch = ARRAYSIZE(szFriendly);
                        if (SUCCEEDED(AssocQueryString(ASSOCF_NOUSERSETTINGS | ASSOCF_INIT_BYEXENAME | ASSOCF_VERIFY,
                            ASSOCSTR_FRIENDLYAPPNAME, szCommand,
                            NULL, szFriendly, &cch)))
                        {
                            // Set the friendly name
                            RegSetValueEx(hkClient, TEXT("LocalizedString"), 0, REG_SZ, (BYTE*)szFriendly, sizeof(TCHAR) * (cch + 1));

                            // Set the command string (properly quoted)
                            PathQuoteSpaces(szCommand);
                            SHSetValue(hkClient, TEXT("shell\\open\\command"), NULL,
                                REG_SZ, szCommand, sizeof(TCHAR) * (1 + lstrlen(szCommand)));
                        }
                    }

                    LONG l = 0;
                    if (RegQueryValue(hkClient, TEXT("DefaultIcon"), NULL, &l) == ERROR_FILE_NOT_FOUND)
                    {
                        // Set it as the system default
                        RegSetValueEx(hkBrowser, NULL, 0, REG_SZ, (BYTE*)szAppName, sizeof(TCHAR) * (lstrlen(szAppName) + 1));

                        // Now tell everybody about the change
                        SHSendMessageBroadcast(WM_SETTINGCHANGE, 0, (LPARAM)TEXT("Software\\Clients\\StartMenuInternet"));
                    }
                    RegCloseKey(hkClient);
                }

            }
            RegCloseKey(hkBrowser);
        }
    }

    // Restart the monitoring of the registry...
    // (RegNotifyChangeKeyValue is good for only one shot.)
    // Some apps (like Opera) delete the key as part of their registration,
    // which causes our HKEY to go bad, so close it and make a new one.

    if (_hkHTTP)
    {
        RegCloseKey(_hkHTTP);
        _hkHTTP = NULL;
    }

    //
    //  Note!  We have to register on HKCR\http\shell recursively
    //  even though we only care about HKCR\http\shell\open\command.
    //  The reason is that shell\open\command might not exist (IE
    //  deletes it as part of its uninstall) and you can't register
    //  a wait on a key that doesn't exist.  We don't want to create
    //  a blank key on our own, because that means "To launch a web
    //  browser, run the null string as a command," which doesn't work
    //  too great.
    //
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, TEXT("http\\shell"),
        0, NULL, REG_OPTION_NON_VOLATILE,
        KEY_ENUMERATE_SUB_KEYS |
        KEY_QUERY_VALUE | KEY_SET_VALUE | KEY_NOTIFY,
        NULL, &_hkHTTP, NULL) == ERROR_SUCCESS)
    {
        RegNotifyChangeKeyValue(_hkHTTP, TRUE,
            REG_NOTIFY_CHANGE_NAME |
            REG_NOTIFY_CHANGE_LAST_SET,
            _hHTTPEvent, TRUE);
    }
}

void CTray::_MigrateOldBrowserSettingsCB(PVOID lpParameter, BOOLEAN)
{
    //
    //  Sleep a little while so the app can finish installing all the
    //  registry keys it wants before we start cleaning up behind it.
    //
    Sleep(1000);

    CTray* self = (CTray*)lpParameter;
    self->_MigrateOldBrowserSettings();
}

//
// *** WARNING ***
//
//  This is a private interface EXPLORER.EXE exposes to SHDOCVW, which
// allows SHDOCVW (mostly desktop) to access tray. All member must be
// thread safe!
//

CDeskTray::CDeskTray()
{
    _ptray = IToClass(CTray, _desktray, this);
}

HRESULT CDeskTray::QueryInterface(REFIID riid, void** ppvObj)
{
    #if 0   // no IID_IDeskTray yet defined
    static const QITAB qit[] =
    {
        QITABENT(CDeskTray, IDeskTray),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvObj);
    #else
    return E_NOTIMPL;
    #endif
}

ULONG CDeskTray::AddRef()
{
    return 2;
}

ULONG CDeskTray::Release()
{
    return 1;
}

HRESULT CDeskTray::GetTrayWindow(HWND* phwndTray)
{
    ASSERT(_ptray->_hwnd);
    *phwndTray = _ptray->_hwnd;
    return S_OK;
}

HRESULT CDeskTray::SetDesktopWindow(HWND hwndDesktop)
{
    ASSERT(v_hwndDesktop == NULL);
    v_hwndDesktop = hwndDesktop;
    return S_OK;
}

UINT CDeskTray::AppBarGetState()
{
    return (_ptray->_uAutoHide ? ABS_AUTOHIDE : 0) |
        (_ptray->_fAlwaysOnTop ? ABS_ALWAYSONTOP : 0);
}

//***   CDeskTray::SetVar -- set an explorer variable (var#i := value)
// ENTRY/EXIT
//  var     id# of variable to be changed
//  value   value to be assigned
// NOTES
//  WARNING: thread safety is up to caller!
//  notes: currently only called in 1 place, but extra generality is cheap
// minimal cost
#define SVTRAY_EXITEXPLORER     0   // g_fExitExplorer
HRESULT CDeskTray::SetVar(int var, DWORD value)
{
    extern BOOL g_fExitExplorer;

    //TraceMsg(DM_TRACE, "c.cdt_sv: set var(%d):=%d", var, value);
    switch (var) {
        case SVTRAY_EXITEXPLORER:
            //TraceMsg(DM_TRACE, "c.cdt_sv: set g_fExitExplorer:=%d", value);
            g_fExitExplorer = value;
            WriteCleanShutdown(1);
            break;

        default:
            ASSERT(0);
            return S_FALSE;
    }
    return S_OK;
}