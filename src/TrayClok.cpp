#include "pch.h"
#include "trayclok.h"

ULONG CClockCtl::AddRef()
{
    return InterlockedIncrement(&_cRef);
}

ULONG CClockCtl::Release()
{
    return InterlockedDecrement(&_cRef);
}

void CClockCtl::_UpdateLastHour()
{
    SYSTEMTIME st;
    // Grab the time so we don't try to refresh the timezone information now.
    GetLocalTime(&st);
    _wLastHour = st.wHour;
    _wLastMinute = st.wMinute;
    _wLastSecond = st.wSecond;
}

void CClockCtl::_EnableTimer(HWND hwnd, DWORD dtNextTick)
{
    if (dtNextTick)
    {
        SetTimer(hwnd, 0, dtNextTick, NULL);
        _fClockRunning = TRUE;
    }
    else if (_fClockRunning)
    {
        _fClockRunning = FALSE;
        KillTimer(hwnd, 0);
    }
}

LRESULT CClockCtl::_HandleCreate(HWND hwnd)
{
    _UpdateLastHour();

    _fShowSeconds = SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("ShowSecondsInSystemClock"), FALSE, FALSE);

    return 1;
}

LRESULT CClockCtl::_HandleDestroy(HWND hwnd)
{
    _EnableTimer(hwnd, 0);
    return 1;
}

DWORD CClockCtl::_RecalcCurTime(HWND hwnd)
{
    SYSTEMTIME st;
    
    // Current time.
    GetLocalTime(&st);
    
    // Don't recalc the text if the time hasn't changed yet.
    if ((st.wMinute != _wLastMinute) || (st.wHour != _wLastHour) ||
        (!_fShowSeconds || st.wSecond != _wLastSecond) ||
        !*_szCurTime)
    {
        _wLastMinute = st.wMinute;
        _wLastSecond = st.wSecond;
        
        // Text for the current time.
        _cchCurTime = GetTimeFormat(LOCALE_USER_DEFAULT, _fShowSeconds ? 0 : TIME_NOSECONDS,
                                    &st, _szTimeFmt, _szCurTime, ARRAYSIZE(_szCurTime));
        // Don't count the NULL terminator.
        if (_cchCurTime > 0)
        {
            _cchCurTime--;
        }

        // Update our window text so accessibility apps can see.  Since we
        // don't have a caption USER won't try to paint us or anything, it
        // will just set the text and fire an event if any accessibility
        // clients are listening...
        SetWindowText(hwnd, _szCurTime);
        
        // Do a timezone check about once an hour (Win95).
        if (st.wHour != _wLastHour)
        {
#ifndef WINNT
            PostMessage(hwnd, TCM_TIMEZONEHACK, 0, 0L);
#endif
            _wLastHour = st.wHour;
        }
    }
    // Return number of milliseconds till we need to be called again.
    if (_fShowSeconds)
        return 1000UL - st.wMilliseconds;
    return 1000UL * (60 - st.wSecond);
}

void CClockCtl::_EnsureFontsInitialized()
{
    if (!_hfontCapNormal)
    {
        NONCLIENTMETRICS ncm;
        ncm.cbSize = sizeof(ncm);
        if (SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
        {
            // Create the normal font
            ncm.lfCaptionFont.lfWeight = FW_NORMAL;
            HFONT hfont = CreateFontIndirect(&ncm.lfCaptionFont);
            if (hfont)
            {
                if (_hfontCapNormal)
                    DeleteFont(_hfontCapNormal);

                _hfontCapNormal = hfont;
            }
        }
    }
}

LRESULT CClockCtl::_DoPaint(HWND hwnd, BOOL fPaint)
{
    PAINTSTRUCT ps;
    RECT rcClient, rcClip = { 0 };
    DWORD dtNextTick = 0;
    HDC hdc;
    
    // If we are asked to paint and the clock is not running then start it.
    // Otherwise wait until we get a clock tick to recompute the time etc.
    BOOL fDoTimer = !fPaint || !_fClockRunning;
    
    // Get a DC to paint with.
    if (fPaint)
    {
        hdc = BeginPaint(hwnd, &ps);
    }
    else
    {
        hdc = GetDC(hwnd);
    }

    _EnsureFontsInitialized();


    // Update the time if we need to.
    if (fDoTimer || !*_szCurTime)
    {
        dtNextTick = _RecalcCurTime(hwnd);

        ASSERT(dtNextTick);
    }
    
    // Paint the clock face if we are not clipped or if we got a real
    // paint message for the window.  We want to avoid turning off the
    // timer on paint messages (regardless of clip region) because this
    // implies the window is visible in some way. If we guessed wrong, we
    // will turn off the timer next timer tick anyway so no big deal.

    if (GetClipBox(hdc, &rcClip) != NULLREGION || fPaint)
    {
        HFONT hfontOld = 0;
        SIZE size;
        
        // Draw the text centered in the window.
        GetClientRect(hwnd, &rcClient);
        if (_hfontCapNormal)
        {
            hfontOld = (HFONT)SelectObject(hdc, _hfontCapNormal);
        }

        SetBkColor(hdc, GetSysColor(COLOR_3DFACE));
        SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));

        GetTextExtentPoint(hdc, _szCurTime, _cchCurTime, &size);

        int x = max((rcClient.right - size.cx)/2, 0);
        int y = max((rcClient.bottom - size.cy)/2, 0);
        ExtTextOut(hdc, x, y, ETO_OPAQUE,
                   &rcClient, _szCurTime, _cchCurTime, NULL);

        //  figure out if the time is clipped
        _fClockClipped = (size.cx > rcClient.right || size.cy > rcClient.bottom);

        if (_hfontCapNormal)
        {
            SelectObject(hdc, hfontOld);
        }
    }
    else
    {
        // We are obscured so make sure we turn off the clock.

        dtNextTick = 0;
        fDoTimer = TRUE;
    }
    
    // Release our paint DC.
    if (fPaint)
    {
        EndPaint(hwnd, &ps);
    }
    else
    {
        ReleaseDC(hwnd, hdc);
    }
    
    // Reset/Kill the timer.
    if (fDoTimer)
    {
        _EnableTimer(hwnd, dtNextTick);


        // If we just killed the timer becuase we were clipped when it arrived,
        // make sure that we are really clipped by invalidating ourselves once.

        if (!dtNextTick && !fPaint)
        {
            InvalidateRect(hwnd, NULL, FALSE);
        }
    }

    return 0;
}

void CClockCtl::_Reset(HWND hwnd)
{
    // Reset the clock by killing the timer and invalidating.
    // Everything will be updated when we try to paint.
    _EnableTimer(hwnd, 0);
    InvalidateRect(hwnd, NULL, FALSE);
}

LRESULT CClockCtl::_HandleTimeChange(HWND hwnd)
{
    *_szCurTime = 0; // Force a text recalc.
    _UpdateLastHour();
    _Reset(hwnd);
    return 1;
}

LRESULT CClockCtl::_CalcMinSize(HWND hWnd)
{
    RECT rc;
    HFONT hfontOld = 0;
    SYSTEMTIME st = { 0 }; // Initialize to 0...
    SIZE sizeAM;
    SIZE sizePM;
    TCHAR szTime[40];

    if (!(GetWindowLong(hWnd, GWL_STYLE) & WS_VISIBLE))
    {
        return 0L;
    }

    if (_szTimeFmt[0] == TEXT('\0'))
    {
        if (GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, _szTimeFmt,
                          ARRAYSIZE(_szTimeFmt)) == 0)
        {
            wprintf(L"c.ccms: GetLocalInfo Failed %d.\n", GetLastError());
        }

        *_szCurTime = 0; // Force the text to be recomputed.
    }

    HDC hdc = GetDC(hWnd);
    if (!hdc)
    {
        return (0L);
    }

    _EnsureFontsInitialized();

    if (_hfontCapNormal)
    {
        hfontOld = (HFONT)SelectObject(hdc, _hfontCapNormal);
    }

    // We need to get the AM and the PM sizes...
    // We append Two 0s and end to add slop into size
    st.wHour = 11;
    int cch = GetTimeFormat(LOCALE_USER_DEFAULT, _fShowSeconds ? 0 : TIME_NOSECONDS, &st,
                            _szTimeFmt, szTime, ARRAYSIZE(szTime) - ARRAYSIZE(szSlop));
    lstrcat(szTime, szSlop);
    GetTextExtentPoint(hdc, szTime, cch + 2, &sizeAM);

    st.wHour = 23;
    cch = GetTimeFormat(LOCALE_USER_DEFAULT, _fShowSeconds ? 0 : TIME_NOSECONDS, &st,
                        _szTimeFmt, szTime, ARRAYSIZE(szTime) - ARRAYSIZE(szSlop));
    lstrcat(szTime, szSlop);
    GetTextExtentPoint(hdc, szTime, cch + 2, &sizePM);

    if (_hfontCapNormal)
    {
        SelectObject(hdc, hfontOld);
    }

    ReleaseDC(hWnd, hdc);

    // Now lets set up our rectangle...
    // The width is 6 digits (a digit slop on both ends + size of
    // : or sep and max AM or PM string...)
    SetRect(&rc, 0, 0, max(sizeAM.cx, sizePM.cx),
            max(sizeAM.cy, sizePM.cy) + 4 * g_cyBorder);

    AdjustWindowRectEx(&rc, GetWindowLong(hWnd, GWL_STYLE), FALSE,
                       GetWindowLong(hWnd, GWL_EXSTYLE));

    // make sure we're at least the size of other buttons:
    if (rc.bottom - rc.top < g_cySize + g_cyEdge)
    {
        rc.bottom = rc.top + g_cySize + g_cyEdge;
    }
    return MAKELRESULT((rc.right - rc.left),
                       (rc.bottom - rc.top));
}

LRESULT CClockCtl::_HandleIniChange(HWND hWnd, WPARAM wParam, LPTSTR pszSection)
{
    // Only process certain sections...
    if ((pszSection == NULL) || (lstrcmpi(pszSection, TEXT("intl")) == 0) ||
        (wParam == SPI_SETICONTITLELOGFONT))
    {
        TOOLINFO ti;

        _szTimeFmt[0] = TEXT('\0'); // Go reread the format.

        // And make sure we have it recalc...
        _CalcMinSize(hWnd);

        ti.cbSize = sizeof(ti);
        ti.uFlags = 0;
        ti.hwnd = v_hwndTray;
        ti.uId = (UINT_PTR)hWnd;
        ti.lpszText = LPSTR_TEXTCALLBACK;
        SendMessage(c_tray.GetTrayTips(), TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

        _Reset(hWnd);
    }

    return 0;
}

LRESULT CClockCtl::_OnNeedText(HWND hWnd, LPNMTTDISPINFOW lpttt)
{
    int iDateFormat = DATE_LONGDATE;

    //  This code is really squirly.  We don't know if the time has been
    //  clipped until we actually try to paint it, since the clip logic
    //  is in the WM_PAINT handler...  Go figure...
    if (!*_szCurTime)
    {
        InvalidateRect(hWnd, NULL, FALSE);
        UpdateWindow(hWnd);
    }

    // If the current user locale is any BiDi locale, then
    // Make the date reading order it RTL. SetBiDiDateFlags only adds
    // DATE_RTLREADING if the locale is BiDi. (see apithk.c for more info). [samera]
    SetBiDiDateFlags(&iDateFormat);

    if (_fClockClipped)
    {
        // we need to put the time in here too
        TCHAR sz[80];
        GetDateFormat(LOCALE_USER_DEFAULT, iDateFormat, NULL, NULL, sz, ARRAYSIZE(sz));
        wnsprintf(lpttt->szText, ARRAYSIZE(lpttt->szText), TEXT("%s %s"), _szCurTime, sz);
    }
    else
    {
        GetDateFormat(LOCALE_USER_DEFAULT, iDateFormat, NULL, NULL, lpttt->szText, ARRAYSIZE(lpttt->szText));
    }

    return TRUE;
}

LRESULT CClockCtl::v_WndProc(HWND hWnd, UINT wMsg, WPARAM wParam, LPARAM lParam)
{
    switch (wMsg)
    {
        case WM_CALCMINSIZE:
            return _CalcMinSize(hWnd);

        case WM_CREATE:
            return _HandleCreate(hWnd);

        case WM_DESTROY:
            return _HandleDestroy(hWnd);

        case WM_PAINT:
        case WM_TIMER:
            return _DoPaint(hWnd, (wMsg == WM_PAINT));

        case WM_GETTEXT:

            // Update the text if we are not running and somebody wants it.

            if (!_fClockRunning)
                _RecalcCurTime(hWnd);
            goto DoDefault;

        case WM_WININICHANGE:
            return _HandleIniChange(hWnd, wParam, (LPTSTR)lParam);

#ifndef WINNT
        case WM_POWERBROADCAST:
            switch (wParam)
            {
            case PBT_APMRESUMECRITICAL:
            case PBT_APMRESUMESTANDBY:
            case PBT_APMRESUMESUSPEND:
                    goto TimeChanged;
            }
            break;
#endif

        case WM_POWER:

            // a critical resume does not generate a WM_POWERBROADCAST
            // to windows for some reason, but it does generate a old
            // WM_POWER message.

            if (wParam == PWR_CRITICALRESUME)
                goto TimeChanged;
            break;

            TimeChanged:
            case WM_TIMECHANGE:
            return _HandleTimeChange(hWnd);
        case WM_NCHITTEST:
            return(HTTRANSPARENT);
        case WM_SHOWWINDOW:
            if (wParam)
                break;
            // fall through
        case TCM_RESET:
            _Reset(hWnd);
            break;
        case WM_NOTIFY:
        {
            NMHDR *pnm = (NMHDR*)lParam;
            switch (pnm->code)
            {
                case TTN_NEEDTEXT:
                    return _OnNeedText(hWnd, (LPTOOLTIPTEXT)lParam);
                    break;
            }
            break;
        }
        default:
            DoDefault:
            return (DefWindowProc(hWnd, wMsg, wParam, lParam));
    }

    return 0;
}

BOOL ClockCtl_Class(HINSTANCE hinst)
{
    WNDCLASS wc = {0};

    wc.lpszClassName = WC_TRAYCLOCK;
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = CClockCtl::s_WndProc;
    wc.hInstance = hinst;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.cbWndExtra = sizeof(CClockCtl*);

    return RegisterClass(&wc);
}

HWND ClockCtl_Create(HWND hwndParent, UINT uID, HINSTANCE hInst)
{
    HWND hwnd = NULL;

    CClockCtl* pcc = new CClockCtl();
    if (pcc)
    {
        hwnd = CreateWindowEx(0, WC_TRAYCLOCK,
            NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE, 0, 0, 0, 0,
            hwndParent, IntToPtr_(HMENU, uID), hInst, pcc);

        pcc->Release();
    }
    return hwnd;
}
