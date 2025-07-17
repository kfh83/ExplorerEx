//---------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation 1991-1992
//
// Put a clock in a window.
//---------------------------------------------------------------------------

#include "pch.h"
#include "cabinet.h"
#include "tray.h"
#include "util.h"


BOOL ClockCtl_Class(HINSTANCE hinst);
HWND ClockCtl_Create(HWND hwndParent, UINT uID, HINSTANCE hInst);

#define WC_TRAYCLOCK TEXT("TrayClockWClass")
#define szSlop TEXT("00")

// Message to calculate the minimum size
#define  WM_CALCMINSIZE  (WM_USER + 100)
#define TCM_RESET        (WM_USER + 101)
#define TCM_TIMEZONEHACK (WM_USER + 102)

class CClockCtl : public CImpWndProc
{
public:
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    CClockCtl() : _cRef(1) {}
protected:
    void _UpdateLastHour(void);
    void _EnableTimer(HWND hwnd, DWORD dtNextTick);
    LRESULT _HandleCreate(HWND hwnd);
    LRESULT _HandleDestroy(HWND hwnd);
    DWORD _RecalcCurTime(HWND hwnd);
    void _EnsureFontsInitialized();
    LRESULT _DoPaint(HWND hwnd, BOOL fPaint);
    void _Reset(HWND hwnd);
    LRESULT _HandleTimeChange(HWND hwnd);
    LRESULT _CalcMinSize(HWND hWnd);
    LRESULT _HandleIniChange(HWND hWnd, WPARAM wParam, LPTSTR pszSection);
    LRESULT _OnNeedText(HWND hWnd, LPTOOLTIPTEXT lpttt);
    LRESULT v_WndProc(HWND hWnd, UINT wMsg, WPARAM wParam, LPARAM lParam);
    // Register the clock class.
    // Create the window.
private:
    ULONG           _cRef;
    
    TCHAR _szTimeFmt[40] = TEXT(""); // The format string to pass to GetFormatTime
    TCHAR _szCurTime[40] = TEXT(""); // The current string.
    int   _cchCurTime = 0;
    WORD  _wLastHour;                // wHour from local time of last clock tick
    WORD  _wLastMinute;              // wMinute from local time of last clock tick
    BOOL  _fClockRunning = FALSE;
    BOOL  _fClockClipped = FALSE;

    HFONT _hfontCapNormal;

    friend BOOL ClockCtl_Class(HINSTANCE hinst);
};
