#include "pch.h"
#include "cocreateinstancehook.h"
#include "stdafx.h"
#include "tray.h"
#include "startmnu.h"
#include "hostutil.h"
#include "deskhost.h"
#include "util.h"

#include <propvarutil.h>


#define TF_DV2HOST  0
// #define TF_DV2HOST TF_CUSTOM1

#define TF_DV2DIALOG  0
// #define TF_DV2DIALOG TF_CUSTOM1

EXTERN_C HINSTANCE g_hinstCabinet;
HRESULT StartMenuHost_Create(IMenuPopup** ppmp, IMenuBand** ppmb, IUnknown** ppunkSite);
void RegisterDesktopControlClasses();

const WCHAR c_wzStartMenuTheme[] = L"StartMenu";

//*****************************************************************

CPopupMenu::~CPopupMenu()
{
    IUnknown_SetSite(_pmp, NULL);
    ATOMICRELEASE(_pmp);
    ATOMICRELEASE(_pmb);
    ATOMICRELEASE(_psm);
}

HRESULT CPopupMenu::Popup(RECT* prcExclude, DWORD dwFlags)
{
    COMPILETIME_ASSERT(sizeof(RECT) == sizeof(RECTL));
    return _pmp->Popup((POINTL*)prcExclude, (RECTL*)prcExclude, dwFlags);
}


HRESULT CPopupMenu::Initialize(IShellMenu* psm, IUnknown* punkSite, HWND hwnd)
{
    HRESULT hr;

    // We should have been zero-initialized
    ASSERT(_pmp == NULL);
    ASSERT(_pmb == NULL);
    ASSERT(_psm == NULL);

    hr = CoCreateInstance(CLSID_MenuDeskBar, NULL, CLSCTX_INPROC_SERVER,
        IID_PPV_ARG(IMenuPopup, &_pmp));
    if (SUCCEEDED(hr))
    {
        IUnknown_SetSite(_pmp, punkSite);

        IBandSite* pbs;
        hr = CoCreateInstance(CLSID_MenuBandSite, NULL, CLSCTX_INPROC_SERVER,
            IID_PPV_ARG(IBandSite, &pbs));
        if (SUCCEEDED(hr))
        {
            hr = _pmp->SetClient(pbs);
            if (SUCCEEDED(hr))
            {
                IDeskBand* pdb;
                if (SUCCEEDED(psm->QueryInterface(IID_PPV_ARG(IDeskBand, &pdb))))
                {
                    hr = pbs->AddBand(pdb);
                    if (SUCCEEDED(hr))
                    {
                        DWORD dwBandID;
                        hr = pbs->EnumBands(0, &dwBandID);
                        if (SUCCEEDED(hr))
                        {
                            hr = pbs->GetBandObject(dwBandID, IID_PPV_ARG(IMenuBand, &_pmb));
                        }
                    }
                    pdb->Release();
                }
            }
            pbs->Release();
        }
    }

    if (SUCCEEDED(hr))
    {
        // Failure to set the theme is nonfatal
        IShellMenu2* psm2;
        if (SUCCEEDED(psm->QueryInterface(IID_PPV_ARGS(&psm2))))
        {
            BOOL fThemed = IsAppThemed();
            psm2->SetTheme(fThemed ? c_wzStartMenuTheme : NULL);
            psm2->SetNoBorder(fThemed ? TRUE : FALSE);
            psm2->Release();
        }

        // Tell the popup that we are the window to parent UI on
        // This will fail on purpose so don't freak out
        psm->SetMenu(NULL, hwnd, 0);
    }

    if (SUCCEEDED(hr))
    {
        _psm = psm;
        psm->AddRef();
        hr = S_OK;
    }

    return hr;
}

HRESULT CPopupMenu_CreateInstance(IShellMenu* psm,
    IUnknown* punkSite,
    HWND hwnd,
    CPopupMenu** ppmOut)
{
    HRESULT hr;
    *ppmOut = NULL;
    CPopupMenu* ppm = new CPopupMenu();
    if (ppm)
    {
        hr = ppm->Initialize(psm, punkSite, hwnd);
        if (FAILED(hr))
        {
            ppm->Release();
        }
        else
        {
            *ppmOut = ppm;  // transfer ownership to called
        }
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }
    return hr;
}

//*****************************************************************

const STARTPANELMETRICS g_spmDefault = {
    {400,410},
    {
        {L"Desktop User Pane",      0x10000000,     SPP_USERPANE,   {150,   75}, NULL, NULL, FALSE, NULL},
        {L"Desktop Open Pane Host", 0x12000000,     SPP_PROGLIST,   {250,  330}, NULL, NULL, FALSE, NULL},
        {L"Desktop OpenBox Host",   0x10010000,     SPP_OPENBOX,    {250,   40}, NULL, NULL, FALSE, NULL},
        {L"DesktopSFTBarHost",      0x12000000,     SPP_PLACESLIST, {150,  295}, NULL, NULL, FALSE, NULL},
        {L"DesktopLogoffPane",      0x10000000,     SPP_LOGOFF,     {150,   40}, NULL, NULL, FALSE, NULL},
    }
};

// EXEX-VISTA(allison): Validated.
HRESULT
CDesktopHost::Initialize(HWND hwndParent)
{
	_hwndParent = hwndParent;
    ASSERT(_hwnd == NULL);

    //
    //  Load some settings.
    //
    _fAutoCascade = _SHRegGetBoolValueFromHKCUHKLM(REGSTR_EXPLORER_ADVANCED, TEXT("Start_AutoCascade"), TRUE);

    return S_OK;
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::QueryInterface(REFIID riid, void** ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CDesktopHost, IMenuPopup),
        QITABENT(CDesktopHost, IDeskBar),       // IMenuPopup derives from IDeskBar
        QITABENTMULTI(CDesktopHost, IOleWindow, IMenuPopup),  // IDeskBar derives from IOleWindow

        QITABENT(CDesktopHost, IMenuBand),
        QITABENT(CDesktopHost, IServiceProvider),
        QITABENT(CDesktopHost, IOleCommandTarget),
        QITABENT(CDesktopHost, IObjectWithSite),

        QITABENT(CDesktopHost, ITrayPriv),      // going away
        QITABENT(CDesktopHost, ITrayPriv2),     // going away

        QITABENTMULTI(CDesktopHost, IDispatch, IAccessible),    // Vista - New
		QITABENT(CDesktopHost, IEnumVARIANT),                   // Vista - New
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::SetSite(IUnknown* punkSite)
{
    CObjectWithSite::SetSite(punkSite);
    if (!_punkSite)
    {
        // This is our cue to break the recursive reference loop
        // The _ppmpPrograms contains multiple backreferences to
        // the CDesktopHost (we are its site, it also references
        // us via CDesktopShellMenuCallback...)
        IUnknown_SafeReleaseAndNullPtr(&_ppmPrograms);

        for (int i = 0; i < ARRAYSIZE(_spm.panes); ++i)
        {
            if (_spm.panes[i].punk)
            {
                IUnknown_SetSite(_spm.panes[i].punk, NULL);
				IUnknown_SafeReleaseAndNullPtr(&_spm.panes[i].punk);
            }
        }

        if (_hwnd)
        {
            ASSERT(GetWindowThreadProcessId(_hwnd, NULL) == GetCurrentThreadId()) // 211
            DestroyWindow(_hwnd);
        }
    }
    return S_OK;
}

// EXEX-VISTA(allison): Validated.
CDesktopHost::~CDesktopHost()
{
    if (_hbmCachedSnapshot)
    {
        DeleteObject(_hbmCachedSnapshot);
    }

	IUnknown_SafeReleaseAndNullPtr(&_ppmPrograms);
	IUnknown_SafeReleaseAndNullPtr(&_ppmTracking);

    if (_hwnd)
    {
        SetWindowLongPtr(_hwnd, GWLP_USERDATA, NULL);
    }
}

// EXEX-VISTA(allison): Validated.
BOOL CDesktopHost::Register()
{
    _wmDragCancel = RegisterWindowMessage(TEXT("CMBDragCancel"));

    WNDCLASSEX  wndclass;

    wndclass.cbSize = sizeof(wndclass);
    wndclass.style = CS_DROPSHADOW;
    wndclass.lpfnWndProc = WndProc;
    wndclass.cbClsExtra = 0;
    wndclass.cbWndExtra = 0;
    wndclass.hInstance = g_hinstCabinet;
    wndclass.hIcon = NULL;
    wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground = GetStockBrush(HOLLOW_BRUSH);
    wndclass.lpszMenuName = NULL;
    wndclass.lpszClassName = WC_DV2;
    wndclass.hIconSm = NULL;

    return (0 != RegisterClassEx(&wndclass));
}

// EXEX-VISTA(allison): Validated.
inline int _ClipCoord(int x, int xMin, int xMax)
{
    if (x < xMin) x = xMin;
    if (x > xMax) x = xMax;
    return x;
}

// EXEX-VISTA(allison): Validated.
void ClipRect(RECT *prcDst, const RECT *prcMax)
{
    prcDst->left = _ClipCoord(prcDst->left, prcMax->left, prcMax->right);
    prcDst->right = _ClipCoord(prcDst->right, prcMax->left, prcMax->right);
    prcDst->top = _ClipCoord(prcDst->top, prcMax->top, prcMax->bottom);
    prcDst->bottom = _ClipCoord(prcDst->bottom, prcMax->top, prcMax->bottom);
}

//
//  Everybody conspires against us.
//
//  CTray does not pass us any MPPF_POS_MASK flags to tell us where we
//  need to pop up relative to the point, so there's no point looking
//  at the dwFlags parameter.  Which is for the better, I guess, because
//  the MPPF_* flags are not the same as the TPM_* flags.  Go figure.
//
//  And then the designers decided that the Start Menu should pop up
//  in a location different from the location that the standard
//  TrackPopupMenuEx function chooses, so we need a custom positioning
//  algorithm anyway.
//
//  And finally, the AnimateWindow function takes AW_* flags, which are
//  not the same as TPM_*ANIMATE flags.  Go figure.  But since we gave up
//  on trying to map IMenuPopup::Popup to TrackPopupMenuEx anyway, we
//  don't have to do any translation here anyway.
//
//  Returns animation direction.
//

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void CDesktopHost::_ChoosePopupPosition(POINT *ppt, LPCRECT prcExclude, LPRECT prcWindow, DWORD dwFlags)
{
#ifdef DEAD_CODE
    //
    // Calculate the monitor BEFORE we adjust the point.  Otherwise, we might
    // move the point offscreen.  In which case, we will end up pinning the
    // popup to the primary display, which is wron_
    //
    HMONITOR hmon = MonitorFromPoint(*ppt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO minfo;
    minfo.cbSize = sizeof(minfo);
    GetMonitorInfo(hmon, &minfo);

    // Clip the exclude rectangle to the monitor

    RECT rcExclude;
    if (prcExclude)
    {
        // We can't use IntersectRect because it turns the rectangle
        // into (0,0,0,0) if the intersection is empty (which can happen if
        // the taskbar is autohide) but we want to glue it to the nearest
        // valid edge.
        rcExclude.left = _ClipCoord(prcExclude->left, minfo.rcMonitor.left, minfo.rcMonitor.right);
        rcExclude.right = _ClipCoord(prcExclude->right, minfo.rcMonitor.left, minfo.rcMonitor.right);
        rcExclude.top = _ClipCoord(prcExclude->top, minfo.rcMonitor.top, minfo.rcMonitor.bottom);
        rcExclude.bottom = _ClipCoord(prcExclude->bottom, minfo.rcMonitor.top, minfo.rcMonitor.bottom);
    }
    else
    {
        rcExclude.left = rcExclude.right = ppt->x;
        rcExclude.top = rcExclude.bottom = ppt->y;
    }

    _ComputeActualSize(&minfo, &rcExclude);

    // initialize the height and width from what the layout asked for
    int cy = RECTHEIGHT(_rcActual);
    int cx = RECTWIDTH(_rcActual);

    ASSERT(cx && cy); // we're in trouble if these are zero

    int x, y;

    //
    //  First: Determine whether we are going to pop upwards or downwards.
    //

    BOOL fSide = FALSE;

    if (rcExclude.top - cy >= minfo.rcMonitor.top)
    {
        // There is room above.
        y = rcExclude.top - cy;
    }
    else if (rcExclude.bottom - cy >= minfo.rcMonitor.top)
    {
        // There is room above if we slide to the side.
        y = rcExclude.bottom - cy;
        fSide = TRUE;
    }
    else if (rcExclude.bottom + cy <= minfo.rcMonitor.bottom)
    {
        // There is room below.
        y = rcExclude.bottom;
    }
    else if (rcExclude.top + cy <= minfo.rcMonitor.bottom)
    {
        // There is room below if we slide to the side.
        y = rcExclude.top;
        fSide = TRUE;
    }
    else
    {
        // We don't fit anywhere.  Pin to the appropriate edge of the screen.
        // And we have to go to the side, too.
        fSide = TRUE;

        if (rcExclude.top - minfo.rcMonitor.top < minfo.rcMonitor.bottom - rcExclude.bottom)
        {
            // Start button at top of screen; pin to top
            y = minfo.rcMonitor.top;
        }
        else
        {
            // Start button at bottom of screen; pin to bottom
            y = minfo.rcMonitor.bottom - cy;
        }
    }

    //
    //  Now choose whether we will pop left or right.  Try right first.
    //

    x = fSide ? rcExclude.right : rcExclude.left;
    if (x + cx > minfo.rcMonitor.right)
    {
        // Doesn't fit to the right; pin to the right edge.
        // Notice that we do *not* try to pop left.  For some reason,
        // the start menu never pops left.

        x = minfo.rcMonitor.right - cx;
    }

    SetRect(prcWindow, x, y, x + cx, y + cy);
#else
    POINT pt = {0};
    if (ppt)
    {
        pt.x = ppt->x;
        pt.y = ppt->y;
    }

    HMONITOR hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO minfo;
    minfo.cbSize = sizeof(minfo);
    GetMonitorInfo(hmon, &minfo);

    RECT rcExclude;
    if (prcExclude)
    {
        CopyRect(&rcExclude, prcExclude);
        ClipRect(&rcExclude, &minfo.rcMonitor);
    }
    else if (ppt)
    {
        rcExclude.right = ppt->x;
        rcExclude.left = rcExclude.right;
        rcExclude.bottom = ppt->y;
        rcExclude.top = rcExclude.bottom;
    }
    else
    {
        CopyRect(&rcExclude, &field_94);
    }

    CopyRect(&field_94, &rcExclude);
    _ComputeActualSize(&minfo, &rcExclude);

    int cy = RECTHEIGHT(_rcActual);
    int cx = RECTWIDTH(_rcActual);

    ASSERT(cx && cy); // 323

    int x, y;

    DWORD v11 = dwFlags & 0xE0000000;
    if ((dwFlags & 0xE0000000) != 0x40000000 && v11 != 0x60000000)
    {
        BOOL fSide = 0;
        y = rcExclude.top - cy;
        if (rcExclude.top - cy < minfo.rcMonitor.top)
        {
            y = rcExclude.bottom - cy;
            if (rcExclude.bottom - cy < minfo.rcMonitor.top)
            {
                if (cy + rcExclude.bottom > minfo.rcMonitor.bottom)
                {
                    if (cy + rcExclude.top > minfo.rcMonitor.bottom)
                    {
                        fSide = 1;
                        if (rcExclude.top - minfo.rcMonitor.top >= minfo.rcMonitor.bottom - rcExclude.bottom)
                        {
                            y = minfo.rcMonitor.bottom - cy;
                        }
                        else
                        {
                            y = minfo.rcMonitor.top;
                        }
                    }
                    else
                    {
                        y = rcExclude.top;
                    }
                }
                else
                {
                    y = rcExclude.bottom;
                }
            }
        }

        if (fSide)
        {
            x = rcExclude.right;
        }
        else
        {
            x = rcExclude.left;
        }

        if (x + cx > minfo.rcMonitor.right)
        {
            x = minfo.rcMonitor.right - cx;
        }
    }
    else
    {
        y = rcExclude.top;
        if (v11 == 0x60000000)
            x = rcExclude.right;
        else
            x = rcExclude.left - cx;
        if (cy + rcExclude.top > minfo.rcMonitor.bottom)
            y = minfo.rcMonitor.bottom - cy;
        if (y < minfo.rcMonitor.top)
            y = minfo.rcMonitor.top;
        if (x + cx > minfo.rcMonitor.right)
            x = minfo.rcMonitor.right - cx;
        if (x < minfo.rcMonitor.left)
            x = minfo.rcMonitor.left;
    }
    SetRect(prcWindow, x, y, x + cx, y + cy);
#endif
}

// EXEX-VISTA(allison): Validated.
int GetDesiredHeight(HWND hwndHost, SMPANEDATA* psmpd, SIZE* psizOut)
{
    SMNGETMINSIZE nmgms = { 0 };
    nmgms.hdr.hwndFrom = hwndHost;
    nmgms.hdr.code = SMN_GETMINSIZE;
    nmgms.siz = psmpd->size;

    SendMessage(psmpd->hwnd, WM_NOTIFY, nmgms.hdr.idFrom, (LPARAM)&nmgms);

    if (psizOut)
		*psizOut = nmgms.field_14;

    return nmgms.siz.cy;
}

extern int g_iLPX;
extern int g_iLPY;
extern void InitDPI();

void SHLogicalToPhysicalDPI(int *a1, int *a2)
{
    InitDPI();
    if (a1)
        *a1 = MulDiv(*a1, g_iLPX, 96);
    if (a2)
        *a2 = MulDiv(*a2, g_iLPY, 96);
}

//
// Query each item to see if it has any size requirements.
// Position all the items at their final locations.
//

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void CDesktopHost::_ComputeActualSize(MONITORINFO *pminfo, LPCRECT prcExclude)
{
#ifdef DEAD_CODE
    // Compute the maximum permissible space above/below the Start Menu.
    // Designers don't want the Start Menu to slide horizontally; it must
    // fit entirely above or below.

    int cxMax = RECTWIDTH(pminfo->rcWork);
    int cyMax = max(prcExclude->top - pminfo->rcMonitor.top,
        pminfo->rcMonitor.bottom - prcExclude->bottom);

    // Start at the minimum size and grow as necesary
    _rcActual = _rcDesired;

    // Ask the windows if they wants any adjustments
    int iMFUHeight = GetDesiredHeight(_hwnd, &_spm.panes[SMPANETYPE_MFU], 0);
    int iPlacesHeight = GetDesiredHeight(_hwnd, &_spm.panes[SMPANETYPE_PLACES], 0);
    int iMoreProgHeight = _spm.panes[SMPANETYPE_MOREPROG].size.cy;

    // Figure out the maximum size for each pane
    int cyPlacesMax = cyMax - (_spm.panes[SMPANETYPE_USER].size.cy + _spm.panes[SMPANETYPE_LOGOFF].size.cy);
    int cyMFUMax = cyPlacesMax - _spm.panes[SMPANETYPE_MOREPROG].size.cy;


    //TraceMsg(TF_DV2HOST, "MFU Desired Height=%d(cur=%d,max=%d), Places Desired Height=%d(cur=%d,max=%d)",
    //    iMFUHeight, _spm.panes[SMPANETYPE_MFU].size.cy, cyMFUMax,
    //    iPlacesHeight, _spm.panes[SMPANETYPE_PLACES].size.cy, cyPlacesMax);

    // Clip each pane to its max - the smaller of (The largest possible or The largest we want to be)
    _fClipped = FALSE;
    if (iMFUHeight > cyMFUMax)
    {
        iMFUHeight = cyMFUMax;
        _fClipped = TRUE;
    }

    if (iPlacesHeight > cyPlacesMax)
    {
        iPlacesHeight = cyPlacesMax;
        _fClipped = TRUE;
    }

    // ensure that places == mfu + moreprog by growing the smaller of the two.
    if (iPlacesHeight > iMFUHeight + iMoreProgHeight)
        iMFUHeight = iPlacesHeight - iMoreProgHeight;
    else
        iPlacesHeight = iMFUHeight + iMoreProgHeight;

    //
    // move the actual windows
    // See diagram of layout in deskhost.h for the hardcoded assumptions here.
    //  this could be made more flexible/variable, but we want to lock in this layout
    //

    // helper variables...
    DWORD dwUserBottomEdge = _spm.panes[SMPANETYPE_USER].size.cy;
    DWORD dwMFURightEdge = _spm.panes[SMPANETYPE_MFU].size.cx;
    DWORD dwMFUBottomEdge = dwUserBottomEdge + iMFUHeight;
    DWORD dwMoreProgBottomEdge = dwMFUBottomEdge + iMoreProgHeight;

    // set the size of the overall pane
    _rcActual.right = _spm.panes[SMPANETYPE_USER].size.cx;
    _rcActual.bottom = dwMoreProgBottomEdge + _spm.panes[SMPANETYPE_LOGOFF].size.cy;

    HDWP hdwp = BeginDeferWindowPos(5);
    const DWORD dwSWPFlags = SWP_NOACTIVATE | SWP_NOZORDER;
    DeferWindowPos(hdwp, _spm.panes[SMPANETYPE_USER].hwnd, NULL, 0, 0, _rcActual.right, dwUserBottomEdge, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[SMPANETYPE_MFU].hwnd, NULL, 0, dwUserBottomEdge, dwMFURightEdge, iMFUHeight, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[SMPANETYPE_MOREPROG].hwnd, NULL, 0, dwMFUBottomEdge, dwMFURightEdge, iMoreProgHeight, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[SMPANETYPE_PLACES].hwnd, NULL, dwMFURightEdge, dwUserBottomEdge, _rcActual.right - dwMFURightEdge, iPlacesHeight, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[SMPANETYPE_LOGOFF].hwnd, NULL, 0, dwMoreProgBottomEdge, _rcActual.right, _spm.panes[SMPANETYPE_LOGOFF].size.cy, dwSWPFlags);
    EndDeferWindowPos(hdwp);
#else
    LONG cyTopHeight; // eax

    int v4 = pminfo->rcMonitor.bottom - prcExclude->bottom;
    if (prcExclude->top - pminfo->rcMonitor.top > v4)
        v4 = prcExclude->top - pminfo->rcMonitor.top;

    int v5 = 0;
    if (_hTheme && !field_C4)
    {
        BOOL v25 = 0;
        DwmIsCompositionEnabled(&v25);
        if (v25)
        {
            SIZE v26;
            v26.cx = 70;
            SHLogicalToPhysicalDPI(nullptr, (int *)&v26);
            v5 = v26.cx - _spm.panes[0].size.cy;
        }
    }
    int v8 = v4 - v5;

    _rcActual = _rcDesired;

    SIZE v27;
    v27.cx = 0;
    v27.cy = 0;
    int DesiredHeight = GetDesiredHeight(_hwnd, &_spm.panes[1], NULL);
    LONG cy = _spm.panes[2].size.cy;
    
    _fClipped = FALSE;
    int v13 = GetDesiredHeight(_hwnd, &_spm.panes[3], &v27);
    int v14 = v8 - _spm.panes[4].size.cy - _spm.panes[0].size.cy;

    //CcshellDebugMsgW(
    //    0,
    //    (__int64)"MFU Desired Height=%d(cur=%d,max=%d), Places Desired Height=%d(cur=%d,max=%d)",
    //    DesiredHeight,
    //    _spm.panes[1].size.cy,
    //    v8 - cy,
    //    GetDesiredHeight(_hwnd, &_spm.panes[3], &v27),
    //    _spm.panes[3].size.cy,
    //    v14);
    
    if (DesiredHeight > v8 - cy)
    {
        DesiredHeight = v8 - cy;
        _fClipped = TRUE;
    }

    if (v13 > v14)
    {
        int v16 = v5 + v8;
        field_C4 = 1;

        MARGINS margins;
        margins.cxLeftWidth = 0;
        margins.cxRightWidth = 0;
        margins.cyTopHeight = 0;
        margins.cyBottomHeight = 0;
        
        if (_hTheme)
        {
            GetThemeMargins(_hTheme, NULL, SPP_PROGLIST, 0, TMT_CONTENTMARGINS, NULL, &margins);
            cyTopHeight = margins.cyTopHeight;
        }
        else
        {
            cyTopHeight = 2 * SHGetSystemMetricsScaled(SM_CXEDGE);
        }
        
        _spm.panes[0].size.cy = cyTopHeight;
        int v18 = v16 - cyTopHeight - _spm.panes[4].size.cy;
        if (v13 > v18)
        {
            v13 = v18;
            _fClipped = TRUE;
        }
    }
    
    if (field_C4)
        IUnknown_QueryServiceExec(static_cast<IMenuBand*>(this), SID_SM_UserPane, &SID_SM_DV2ControlHost, 323, 0, NULL, NULL);
    
    LONG v19 = _spm.panes[4].size.cy;
    LONG v20 = _spm.panes[0].size.cy;
    LONG cx = _spm.panes[1].size.cx;

    int v22 = DesiredHeight + cy - v19;
    if (v20 + v13 > v22)
    {
        v22 = v20 + v13;
    }
    
    int v23 = v22 + v19 - cy;
    if (_spm.panes[4].size.cx < v27.cx)
    {
        _spm.panes[4].size.cx = v27.cx;
    }

    _rcActual.right = cx + _spm.panes[4].size.cx;
    _rcActual.bottom = v19 + v22;

    HDWP hdwp = BeginDeferWindowPos(5);
    const DWORD dwSWPFlags = SWP_NOACTIVATE | SWP_NOZORDER;
    DeferWindowPos(hdwp, _spm.panes[0].hwnd, nullptr, cx, 0, _rcActual.right - cx, v20, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[1].hwnd, nullptr, 0, 0, cx, v23, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[2].hwnd, nullptr, 0, v23, cx, cy, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[3].hwnd, nullptr, cx, v20, _rcActual.right - cx, v22 - v20, dwSWPFlags);
    DeferWindowPos(hdwp, _spm.panes[4].hwnd, nullptr, cx, v22, _rcActual.right - cx, _spm.panes[4].size.cy, dwSWPFlags);
    EndDeferWindowPos(hdwp);
#endif
}

// EXEX-VISTA(allison): Validated.
HWND CDesktopHost::_Create()
{
    TCHAR szTitle[MAX_PATH];

    LoadString(g_hinstCabinet, IDS_STARTMENU, szTitle, MAX_PATH);

    Register();

    // Must load metrics early to determine whether we are themed or not
    LoadPanelMetrics();

    DWORD dwExStyle = WS_EX_TOOLWINDOW;
    if (IS_BIDI_LOCALIZED_SYSTEM())
    {
        dwExStyle |= WS_EX_LAYOUTRTL;
    }

    DWORD dwStyle = WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;  // We will make it visible as part of animation
    if (!_hTheme)
    {
        // Normally the theme provides the border effects, but if there is
        // no theme then we have to do it ourselves.
        dwStyle |= WS_DLGFRAME;
    }

    _hwnd = SHFusionCreateWindowEx(
        dwExStyle,
        WC_DV2,
        szTitle,
        dwStyle,
        0, 0,
        0, 0,
        NULL,
        NULL,
        g_hinstCabinet,
        this);
    if (_hwnd)
    {
        v_hwndStartPane = _hwnd;
        if (_hwnd)
        {
            SetAccessibleSubclassWindow(_hwnd);
        }
    }

    return _hwnd;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_ReapplyRegion()
{
    SMNMAPPLYREGION ar;

    // If we fail to create a rectangular region, then remove the region
    // entirely so we don't carry the old (bad) region around.
    // Yes it means you get ugly black corners, but it's better than
    // clipping away huge chunks of the Start Menu!

    if (_hTheme)
    {
        ar.hrgn = CreateRectRgn(0, 0, _sizWindowPrev.cx, _sizWindowPrev.cy);
        if (ar.hrgn)
        {
            // Let all the clients take a bite out of it
            ar.hdr.hwndFrom = _hwnd;
            ar.hdr.idFrom = 0;
            ar.hdr.code = SMN_APPLYREGION;

            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&ar, SPM_SEND | SPM_ONELEVEL);

            RECT rc;
            RECT rcNew = {0};
            GetWindowRect(_spm.panes[0].hwnd, &rc);
            MapWindowRect(NULL, _hwnd, &rc);
            UnionRect(&rcNew, &rcNew, &rc);

            GetWindowRect(_spm.panes[3].hwnd, &rc);
            MapWindowRect(NULL, _hwnd, &rc);
            UnionRect(&rcNew, &rcNew, &rc);

            GetWindowRect(_spm.panes[4].hwnd, &rc);
            MapWindowRect(NULL, _hwnd, &rc);
            UnionRect(&rcNew, &rcNew, &rc);

            HandleApplyRegionFromRect(rcNew, _hTheme, &ar, SPP_PLACESLIST, 0);
        }

        if (SetWindowRgn(_hwnd, ar.hrgn, FALSE))
        {
            _RegisterForGlass(TRUE, ar.hrgn);
        }
        else if (ar.hrgn)
        {
            // SetWindowRgn takes ownership on success
            // On failure we need to free it ourselves
            DeleteObject(ar.hrgn);
        }
    }
}


//
//  We need to use PrintWindow because WM_PRINT messes up RTL.
//  PrintWindow requires that the window be visible.
//  Making the window visible causes the shadow to appear.
//  We don't want the shadow to appear until we are ready.
//  So we have to do a lot of goofy style mangling to suppress the
//  shadow until we're ready.
//

// EXEX-VISTA(allison): Validated.
BOOL ShowCachedWindow(HWND hwnd, SIZE sizWindow, HBITMAP hbmpSnapshot, BOOL fRepaint)
{
    BOOL fSuccess = FALSE;
    if (hbmpSnapshot)
    {
        // Turn off the shadow so it won't get triggered by our SetWindowPos
        DWORD dwClassStylePrev = GetClassLong(hwnd, GCL_STYLE);
        SetClassLong(hwnd, GCL_STYLE, dwClassStylePrev & ~CS_DROPSHADOW);

        // Show the window and tell it not to repaint; we'll do that
        SetWindowPos(hwnd, NULL, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER |
            SWP_NOREDRAW | SWP_SHOWWINDOW);

        // Turn the shadow back on
        SetClassLong(hwnd, GCL_STYLE, dwClassStylePrev);

        // Disable WS_CLIPCHILDREN because we need to draw over the kids for our BLT
        DWORD dwStylePrev = SHSetWindowBits(hwnd, GWL_STYLE, WS_CLIPCHILDREN, 0);

        HDC hdcWindow = GetDCEx(hwnd, NULL, DCX_WINDOW | DCX_CACHE | DCX_CLIPSIBLINGS);
        if (hdcWindow)
        {
            HDC hdcMem = CreateCompatibleDC(hdcWindow);
            if (hdcMem)
            {
                HBITMAP hbmPrev = (HBITMAP)SelectObject(hdcMem, hbmpSnapshot);

                // PrintWindow only if fRepaint says it's necessary
                if (!fRepaint || PrintWindow(hwnd, hdcMem, 0))
                {
                    // Do this horrible dance because sometimes GDI takes a long
                    // time to do a BitBlt so you end up seeing the shadow for
                    // a half second before the bits show up.
                    //
                    // So show the bits first, then show the shadow.

                    if (BitBlt(hdcWindow, 0, 0, sizWindow.cx, sizWindow.cy, hdcMem, 0, 0, SRCCOPY))
                    {
                        // Tell USER to attach the shadow
                        // Do this by hiding the window and then showing it
                        // again, but do it in this goofy way to avoid flicker.
                        // (If we used ShowWindow(SW_HIDE), then the window
                        // underneath us would repaint pointlessly.)

                        SHSetWindowBits(hwnd, GWL_STYLE, WS_VISIBLE, 0);
                        SetWindowPos(hwnd, NULL, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER |
                            SWP_NOREDRAW | SWP_SHOWWINDOW);

                        // Validate the window now that we've drawn it
                        RedrawWindow(hwnd, NULL, NULL, RDW_NOERASE | RDW_NOFRAME |
                            RDW_NOINTERNALPAINT | RDW_VALIDATE);
                        fSuccess = TRUE;
                    }
                }

                SelectObject(hdcMem, hbmPrev);
                DeleteDC(hdcMem);
            }
            ReleaseDC(hwnd, hdcWindow);
        }

        SetWindowLong(hwnd, GWL_STYLE, dwStylePrev);
    }

    if (!fSuccess)
    {
        // re-hide the window so USER knows it's all invalid again
        ShowWindow(hwnd, SW_HIDE);
    }
    return fSuccess;
}

// EXEX-VISTA(allison): Validated.
BOOL CDesktopHost::_TryShowBuffered()
{
    BOOL fSuccess = FALSE;
    BOOL fRepaint = FALSE;

    if (!_hbmCachedSnapshot)
    {
        HDC hdcWindow = GetDCEx(_hwnd, NULL, DCX_WINDOW | DCX_CACHE | DCX_CLIPSIBLINGS);
        if (hdcWindow)
        {
            _hbmCachedSnapshot = CreateBitmap(hdcWindow, _sizWindowPrev.cx, _sizWindowPrev.cy);
            fRepaint = TRUE;
            ReleaseDC(_hwnd, hdcWindow);
        }
    }
    if (_hbmCachedSnapshot)
    {
        fSuccess = ShowCachedWindow(_hwnd, _sizWindowPrev, _hbmCachedSnapshot, fRepaint);
        if (!fSuccess)
        {
            DeleteObject(_hbmCachedSnapshot);
            _hbmCachedSnapshot = NULL;
        }
    }
    return fSuccess;
}

// EXEX-VISTA(allison): Validated.
LRESULT CDesktopHost::OnNeedRepaint()
{
    if (_hwnd && _hbmCachedSnapshot)
    {
        // This will force a repaint the next time the window is shown
        DeleteObject(_hbmCachedSnapshot);
        _hbmCachedSnapshot = NULL;
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CDesktopHost::OnNeedRebuild()
{
    if (IsWindowVisible(_hwnd))
        field_D0 = 1;
    else
        PostMessage(v_hwndTray, SBM_REBUILDMENU, 0, 0);
    return 0;
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::_Popup(POINT *ppt, RECT *prcExclude, DWORD dwFlags)
{
#ifdef DEAD_CODE
    if (_hwnd)
    {
        RECT rcWindow;
        _ChoosePopupPosition(ppt, prcExclude, &rcWindow, dwFlags);
        SIZE sizWindow = { RECTWIDTH(rcWindow), RECTHEIGHT(rcWindow) };

        MoveWindow(_hwnd, rcWindow.left, rcWindow.top,
            sizWindow.cx, sizWindow.cy, TRUE);

        if (sizWindow.cx != _sizWindowPrev.cx ||
            sizWindow.cy != _sizWindowPrev.cy)
        {
            _sizWindowPrev = sizWindow;
            _ReapplyRegion();
            // We need to repaint since our size has changed
            OnNeedRepaint();
        }

        _RegisterForGlass(TRUE, NULL);

        // If the user toggles the tray between topmost and nontopmost
        // our own topmostness can get messed up, so re-assert it here.
        // SetWindowZorder(_hwnd, HWND_TOPMOST);

        if (GetSystemMetrics(SM_REMOTESESSION) || GetSystemMetrics(SM_REMOTECONTROL))
        {
            // If running remotely, then don't cache the Start Menu
            // or double-buffer.  Show the keyboard cues accurately
            // from the start (to avoid flicker).

            SendMessage(_hwnd, WM_CHANGEUISTATE, UIS_INITIALIZE, 0);
            if (dwFlags & MPPF_KEYBOARD)
            {
                _EnableKeyboardCues();
            }
            ShowWindow(_hwnd, SW_SHOW);
        }
        else
        {
            // If running locally, then force keyboard cues off so our
            // cached bitmap won't have underlines.  Then draw the
            // Start Menu, then turn on keyboard cues if necessary.

            SendMessage(_hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS | UISF_HIDEACCEL), 0);

            if (!_TryShowBuffered())
            {
                ShowWindow(_hwnd, SW_SHOW);
            }

            if (dwFlags & MPPF_KEYBOARD)
            {
                _EnableKeyboardCues();
            }
        }

        NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPSTART, _hwnd, OBJID_CLIENT, CHILDID_SELF);

        // Tell tray that the start pane is active, so it knows to eat
        // mouse clicks on the Start Button.
        IStartButton *pstb = _GetIStartButton();
        if (pstb)
        {
            pstb->SetStartPaneActive(TRUE);
            pstb->Release();
        }

        _fOpen = TRUE;
        _fMenuBlocked = FALSE;
        _fMouseEntered = FALSE;
        _fOfferedNewApps = FALSE;

        _MaybeOfferNewApps();
        _MaybeShowClipBalloon();

        // Tell all our child windows it's time to maybe revalidate
        NMHDR nm = { _hwnd, 0, SMN_POSTPOPUP };
        SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);

        ExplorerPlaySound(TEXT("MenuPopup"));


        return S_OK;
    }
    else
    {
        return E_FAIL;
    }
#else
    if (_hwnd)
    {
        //SHTracePerf(&ShellTraceId_Explorer_StartMenu_Scenario_Start);

        RECT rcWindow;
        _ChoosePopupPosition(ppt, prcExclude, &rcWindow, dwFlags);

        SIZE sizWindow = { RECTWIDTH(rcWindow), RECTHEIGHT(rcWindow) };
        MoveWindow(_hwnd, rcWindow.left, rcWindow.top, RECTWIDTH(rcWindow), RECTHEIGHT(rcWindow), TRUE);

        if (sizWindow.cx != _sizWindowPrev.cx || sizWindow.cy != _sizWindowPrev.cy)
        {
            _sizWindowPrev.cx = sizWindow.cx;
            _sizWindowPrev.cy = sizWindow.cy;
            _ReapplyRegion();
            OnNeedRepaint();
        }

        _RegisterForGlass(TRUE, NULL);

        NMHDR nm;
        nm.hwndFrom = _hwnd;
        nm.idFrom = 0;
        nm.code = 221;
        SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);

        if (GetSystemMetrics(SM_REMOTESESSION) || GetSystemMetrics(SM_REMOTECONTROL))
        {
            SendMessage(_hwnd, WM_CHANGEUISTATE, 0x30001u, 0);
            ShowWindow(_hwnd, SW_SHOW);
        }
        else
        {
            SendMessage(this->_hwnd, WM_CHANGEUISTATE, 0x30001u, 0);

            if (!_TryShowBuffered())
            {
                ShowWindow(_hwnd, SW_SHOW);
            }
        }

        IStartButton *pstb = _GetIStartButton();
        if (pstb)
        {
            pstb->SetStartPaneActive(TRUE);
            pstb->Release();
        }

        _hwndLastMouse = 0;
        _lParamLastMouse = 0;
        _fOpen = TRUE;

        _SetFocusToOpenBox();

        _fMenuBlocked = FALSE;
        _fMouseEntered = FALSE;
        _fOfferedNewApps = FALSE;

        _MaybeOfferNewApps();
        _MaybeShowClipBalloon();

        NMHDR nm2;
        nm2.idFrom = 0;
        nm2.code = 208;
        SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm2, SPM_SEND | SPM_ONELEVEL);

        //SHPlaySound(L"MenuPopup", 2);
        NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPSTART, _hwnd, OBJID_CLIENT, CHILDID_SELF);
        _LockStartPane();
        return S_OK;
    }
    else
    {
        return E_FAIL;
    }
#endif
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::Popup(POINTL* pptl, RECTL* prclExclude, MP_POPUPFLAGS dwFlags)
{
    COMPILETIME_ASSERT(sizeof(POINTL) == sizeof(POINT));
    POINT* ppt = reinterpret_cast<POINT*>(pptl);

    COMPILETIME_ASSERT(sizeof(RECTL) == sizeof(RECT));
    RECT* prcExclude = reinterpret_cast<RECT*>(prclExclude);

    if (_hwnd == NULL)
    {
        _hwnd = _Create();
    }

    return _Popup(ppt, prcExclude, dwFlags);
}

// EXEX-VISTA(allison): Validated.
LRESULT CDesktopHost::OnHaveNewItems(NMHDR* pnm)
{
    PSMNMHAVENEWITEMS phni = (PSMNMHAVENEWITEMS)pnm;

    _hwndNewHandler = pnm->hwndFrom;

    // We have a new "new app" list, so tell the cached Programs menu
    // its cache is no longer valid and it should re-query us
    // so we can color the new apps appropriately.

    if (_ppmPrograms)
    {
        _ppmPrograms->Invalidate();
    }

    //
    //  Were any apps in the list installed since the last time the
    //  user acknowledged a new app?
    //

    FILETIME ftBalloon = { 0, 0 };      // assume never
    DWORD dwSize = sizeof(ftBalloon);

    _SHRegGetValueFromHKCUHKLM(DV2_REGPATH, DV2_NEWAPP_BALLOON_TIME, SRRF_RT_ANY, NULL, &ftBalloon, &dwSize);

    if (CompareFileTime(&ftBalloon, &phni->ftNewestApp) < 0)
    {
        _iOfferNewApps = NEWAPP_OFFER_COUNT;
        _MaybeOfferNewApps();
    }

    return 1;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_MaybeOfferNewApps()
{
    // Display the balloon tip only once per pop-open,
    // and only if there are new apps to offer
    // and only if we're actually visible
    if (_fOfferedNewApps || !_iOfferNewApps || !IsWindowVisible(_hwnd) ||
        !_SHRegGetBoolValueFromHKCUHKLM(REGSTR_EXPLORER_ADVANCED, REGSTR_VAL_DV2_NOTIFYNEW, TRUE))
    {
        return;
    }

    _fOfferedNewApps = TRUE;
    _iOfferNewApps--;

    SMNMBOOL nmb = { { _hwnd, 0, SMN_SHOWNEWAPPSTIP }, TRUE };
    SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nmb, SPM_SEND | SPM_ONELEVEL);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnSeenNewItems()
{
    _iOfferNewApps = 0; // Do not offer More Programs balloon tip again

    // Remember the time the user acknowledged the balloon so we only
    // offer the balloon if there is an app installed after this point.

    FILETIME ftNow;
    GetSystemTimeAsFileTime(&ftNow);
    SHSetValue(HKEY_CURRENT_USER, DV2_REGPATH, DV2_NEWAPP_BALLOON_TIME, REG_BINARY,
        &ftNow, sizeof(ftNow));
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_MaybeShowClipBalloon()
{
    if (_fClipped && !_fWarnedClipped)
    {
        _fWarnedClipped = TRUE;

        RECT rc;
        GetWindowRect(_spm.panes[3].hwnd, &rc);    // show the clipped ballon pointing to the bottom of the MFU

        _hwndClipBalloon = CreateBalloonTip(_hwnd,
            (rc.right + rc.left) / 2, rc.bottom,
            NULL,
            IDS_STARTPANE_CLIPPED_TITLE,
            IDS_STARTPANE_CLIPPED_TEXT);
        if (_hwndClipBalloon)
        {
            SetProp(_hwndClipBalloon, PROP_DV2_BALLOONTIP, DV2_BALLOONTIP_CLIP);
        }
    }
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnContextMenu(LPARAM lParam)
{
    if (!IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOTRAYCONTEXTMENU, TEXT("Advanced"), TEXT("TaskbarContextMenu"), ROUS_KEYALLOWS | ROUS_DEFAULTALLOW))
    {
        HMENU hmenu = SHLoadMenuPopup(g_hinstCabinet, MENU_STARTPANECONTEXT);
        if (hmenu)
        {
            if (GetAsyncKeyState(VK_SHIFT) < 0 && GetAsyncKeyState(VK_CONTROL) < 0)
            {
                DWORD dwValue;
                DWORD cbData = sizeof(dwValue);
                if (SHRegGetValue(
                    HKEY_CURRENT_USER,
                    TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System"),
                    TEXT("DisableTaskMgr"),
                    SRRF_RT_REG_DWORD,
                    nullptr,
                    &dwValue,
                    &cbData) == ERROR_SUCCESS)
                {
                    if (dwValue)
                    {
                        DeleteMenu(hmenu, 32756, 0);
                    }
                }
            }

            if (GetAsyncKeyState(VK_SHIFT) >= 0 || GetAsyncKeyState(VK_CONTROL) >= 0)
                DeleteMenu(hmenu, 32756, 0);

            POINT pt;
            if (IS_WM_CONTEXTMENU_KEYBOARD(lParam))
            {
                pt.x = pt.y = 0;
                MapWindowPoints(_hwnd, HWND_DESKTOP, &pt, 1);
            }
            else
            {
                pt.x = GET_X_LPARAM(lParam);
                pt.y = GET_Y_LPARAM(lParam);
            }

            AddRef();
            int idCmd = TrackPopupMenuEx(hmenu, TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_LEFTALIGN,
                pt.x, pt.y, _hwnd, NULL);
            if (idCmd == IDSYSPOPUP_STARTMENUPROP)
            {
                DesktopHost_Dismiss(_hwnd);
                Tray_DoProperties(TPF_STARTMENUPAGE);
            }
			else if (idCmd == 32756) // EXEX-Vista(allison): TODO: Check what this command is
            {
                DesktopHost_Dismiss(_hwnd);
                PostMessage(v_hwndTray, 0x5B4u, 0, 0);
            }
            DestroyMenu(hmenu);
            Release();
        }
    }
}

// EXEX-VISTA(allison): Validated.
BOOL CDesktopHost::_ShouldIgnoreFocusChange(HWND hwndFocusRecipient)
{
    // Ignore focus changes when a popup menu is up
    if (_ppmTracking)
    {
        return TRUE;
    }

    // If a focus change from a special balloon, this means that the
    // user is clicking a tooltip.  So dismiss the ballon and not the Start Menu.
    HANDLE hProp = GetProp(hwndFocusRecipient, PROP_DV2_BALLOONTIP);
    if (hProp)
    {
        SendMessage(hwndFocusRecipient, TTM_POP, 0, 0);
        if (hProp == DV2_BALLOONTIP_MOREPROG)
        {
            OnSeenNewItems();
        }
        return TRUE;
    }

    HWND hwndCurrent = hwndFocusRecipient;
    while (TRUE)
    {
        HWND hwndParent = GetParent(hwndCurrent);
        if (!hwndParent)
        {
            hwndParent = ::GetWindow(hwndCurrent, GW_OWNER);
            if (!hwndParent)
                break;
        }
        hwndCurrent = hwndParent;
        if (hwndParent == _hwnd)
            return TRUE;
    }

    // Otherwise, dismiss ourselves
    return FALSE;
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::TranslatePopupMenuMessage(MSG* pmsg, LRESULT* plres)
{
    BOOL fDismissOnlyPopup = FALSE;

    // If the user drags an item off of a popup menu, the popup menu
    // will autodismiss itself.  If the user is over our window, then
    // we only want it to dismiss up to our level.

    // (under low memory conditions, _wmDragCancel might be WM_NULL)
    if (pmsg->message == _wmDragCancel && pmsg->message != WM_NULL)
    {
        RECT rc;
        POINT pt;
        if (GetWindowRect(_hwnd, &rc) &&
            GetCursorPos(&pt) &&
            PtInRect(&rc, pt))
        {
            fDismissOnlyPopup = TRUE;
        }
    }

    if (fDismissOnlyPopup)
        _fDismissOnlyPopup++;

    HRESULT hr = _ppmTracking->TranslateMenuMessage(pmsg, plres);

    if (fDismissOnlyPopup)
        _fDismissOnlyPopup--;

    return hr;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnThemeChanged(UINT a2)
{
    if (_hTheme)
    {
        CloseThemeData(_hTheme);
        _hTheme = NULL;
    }

    if (a2 && SHGetCurColorRes() > 8)
    {
        _hTheme = _GetStartMenuTheme();
    }
}

// EXEX-VISTA(allison): Validated. Still needs heavy cleanup.
LRESULT CALLBACK CDesktopHost::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    CDesktopHost* pdh = (CDesktopHost*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    LPCREATESTRUCT pcs;

    if (pdh && pdh->_ppmTracking)
    {
        MSG msg = { hwnd, uMsg, wParam, lParam };
        LRESULT lres;
        if (pdh->TranslatePopupMenuMessage(&msg, &lres) == S_OK)
        {
            return lres;
        }
        wParam = msg.wParam;
        lParam = msg.lParam;
    }

    switch (uMsg)
    {
    case WM_NCCREATE:
        pcs = (LPCREATESTRUCT)lParam;
        pdh = (CDesktopHost*)pcs->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pdh);
        break;

    case WM_CREATE:
        pdh->OnCreate(hwnd);
        break;

    case WM_ACTIVATEAPP:
        if (!wParam)
        {
            DesktopHost_Dismiss(hwnd);
        }
        break;

    case WM_ACTIVATE:
        if (pdh)
        {
            if (LOWORD(wParam) == WA_INACTIVE)
            {
                pdh->_SaveChildFocus();
                HWND hwndAncestor = GetAncestor((HWND)lParam, GA_ROOTOWNER);
                if (hwnd != hwndAncestor &&
                    !(hwndAncestor == v_hwndTray && pdh->_ShouldIgnoreFocusChange((HWND)lParam)) &&
                    !pdh->_ppmTracking)
                    // Losing focus to somebody unrelated to us = dismiss
                {
#ifdef FULL_DEBUG
                    if (!(GetAsyncKeyState(VK_SHIFT) < 0))
#endif
                        DesktopHost_Dismiss(hwnd);
                }
            }
            else
            {
                pdh->_RestoreChildFocus();
            }
        }
        break;

    case WM_DESTROY:
        pdh->OnDestroy();
        break;

    case WM_SHOWWINDOW:
        /*
         *  If hiding the window, save the focus for restoration later.
         */
        if (!wParam)
        {
            pdh->_SaveChildFocus();
        }
        break;

    case WM_SETFOCUS:
        pdh->OnSetFocus((HWND)wParam);
        break;

    case WM_ERASEBKGND:
        pdh->OnPaint((HDC)wParam, TRUE);
        return TRUE;

#if 0
        // currently, the host doesn't do anything on WM_PAINT
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC         hdc;

        if (hdc = BeginPaint(hwnd, &ps))
        {
            pdh->OnPaint(hdc, FALSE);
            EndPaint(hwnd, &ps);
        }
    }

    break;
#endif

    case WM_NOTIFY:
    {
        LPNMHDR pnm = (LPNMHDR)lParam;
        switch (pnm->code)
        {
        case SMN_HAVENEWITEMS:
            return pdh->OnHaveNewItems(pnm);
        case SMN_COMMANDINVOKED:
            return pdh->OnCommandInvoked(pnm);
        case SMN_FILTEROPTIONS:
            return pdh->OnFilterOptions(pnm);

        case SMN_NEEDREPAINT:
            return pdh->OnNeedRepaint();

        case SMN_TRACKSHELLMENU:
            pdh->OnTrackShellMenu(pnm);
            return 0;

        case SMN_BLOCKMENUMODE:
            pdh->_fMenuBlocked = ((SMNMBOOL*)pnm)->f;
            break;
        case SMN_SEENNEWITEMS:
            pdh->OnSeenNewItems();
            break;

        case SMN_CANCELSHELLMENU:
            pdh->_DismissTrackShellMenu();
            break;
        }
    }
    break;

    case WM_CONTEXTMENU:
        pdh->OnContextMenu(lParam);
        return 0;                   // do not bubble up

    case WM_SETTINGCHANGE:
        if ((wParam == SPI_ICONVERTICALSPACING) ||
            ((wParam == 0) && (lParam != 0) && (StrCmpIC((LPCTSTR)lParam, TEXT("Policy")) == 0)))
        {
            // A change in icon vertical spacing is how the themes control
            // panel tells us that it changed the Large Icons setting (!)
            ::PostMessage(v_hwndTray, SBM_REBUILDMENU, 0, 0);
        }

        // Toss our cached bitmap because the user may have changed something
        // that affects our appearance (e.g., toggled keyboard cues)
        pdh->OnNeedRepaint();

        SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL); // forward to kids
        break;


    case WM_DISPLAYCHANGE:
    case WM_SYSCOLORCHANGE:
        // Toss our cached bitmap because these settings may affect our
        // appearance (e.g., color changes)
        pdh->OnNeedRepaint();

        SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL); // forward to kids
        break;

    case WM_TIMER:
        switch (wParam)
        {
        case IDT_MENUCHANGESEL:
            pdh->_OnMenuChangeSel();
            return 0;
        }
        break;

    case DHM_DISMISS:
        pdh->_OnDismiss((BOOL)wParam);
        break;

        // Alt+F4 dismisses the window, but doesn't destroy it
    case WM_CLOSE:
        pdh->_OnDismiss(FALSE);
        return 0;

    case WM_SYSCOMMAND:
        switch (wParam & ~0xF) // must ignore bottom 4 bits
        {
        case SC_SCREENSAVE:
            DesktopHost_Dismiss(hwnd);
            break;
        }
        break;

    case WM_WINDOWPOSCHANGING:
        LPWINDOWPOS pwp = reinterpret_cast<LPWINDOWPOS>(lParam);
        if (pdh && (pwp->flags & SWP_HIDEWINDOW))
        {
            IUnknown_QueryServiceExec(static_cast<IMenuBand*>(pdh), SID_SM_UserPane, &stru_1013554, 323, 0, 0, 0);
        }
        break;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
#else
    HWND Ancestor; // eax
    IStartButton* IStartButton; // eax MAPDST
    HWND i; // eax
    bool v12; // zf
    struct IStartButton* v14; // edi MAPDST
    bool bSkipRebuild; // zf
    tagGUITHREADINFO gui; // [esp+Ch] [ebp-34h] BYREF
    DWORD v25; // [esp+3Ch] [ebp-4h] BYREF

    CDesktopHost* pdh = (CDesktopHost*)GetWindowLongPtr(hwnd, -21);
    LPCREATESTRUCT pcs;

    LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);
    if (pdh && pdh->_ppmTracking)
    {
        MSG msg = { hwnd, uMsg, wParam, lParam };
        LRESULT lres;
        if (pdh->TranslatePopupMenuMessage(&msg, &lres) == S_OK)
        {
            return lres;
        }
        wParam = msg.wParam;
        lParam = msg.lParam;
    }
    
    if (uMsg > 0x47)
    {
        if (uMsg > 0x112)
        {
            if (uMsg != 275)
            {
                if (uMsg == 794)
                {
                    if (pdh)
                    {
                        pdh->OnThemeChanged(wParam);
                    }
                }
                else if (uMsg == 798)
                {
                    if (pdh)
                    {
                        BOOL fEnabled = FALSE;
                        DwmIsCompositionEnabled(&fEnabled);
                        pdh->OnThemeChanged(fEnabled);
                        pdh->_RegisterForGlass(fEnabled, 0);
                        PostMessage(v_hwndTray, 0x40Du, 0, 0);
                    }
                }
                else if (uMsg == 0x8000 && pdh)
                {
                    pdh->_OnDismiss(wParam);
                }
                return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
            }
            if (wParam != 1 || !pdh)
            {
                return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
            }
            pdh->_OnMenuChangeSel();
        }
        else
        {
            if (uMsg == 274)
            {
                v12 = (wParam & 0xFFFFFFF0) == 61760;
                goto LABEL_60;
            }
            if (uMsg != 78)
            {
                if (uMsg == 123)
                {
                    if (pdh)
                    {
                        pdh->OnContextMenu(lParam);
                    }
                    return 0;
                }
                if (uMsg != 126)
                {
                    if (uMsg == 0x81) // WM_NCCREATE
                    {
                        pcs = (LPCREATESTRUCT)lParam;
                        pdh = (CDesktopHost*)pcs->lpCreateParams;
                        if (pdh)
                        {
                            pdh->AddRef();
                        }
                        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pdh);
                    }
                    else if (uMsg == 0x82) // WM_NCDESTROY
                    {
                        if (pdh)
                        {
                            pdh->Release();
                        }
                        SetWindowLongW(hwnd, -21, 0);
                    }
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                }
                goto LABEL_98;
            }

            if (!pdh)
            {
                return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
            }

            if (pnm->code <= 211)
            {
                if (pnm->code == 211) // SMN_CANCELSHELLMENU
                {
                    pdh->_DismissTrackShellMenu();
                }
                else
                {
                    if (pnm->code == 202) // SMN_HAVENEWITEMS
                    {
                        return pdh->OnHaveNewItems(pnm);
                    }
                    if (pnm->code == 204) // SMN_COMMANDINVOKED
                    {
                        return pdh->OnCommandInvoked(pnm);
                    }
                    if (pnm->code == 207) // SMN_SEENNEWITEMS
                    {
                        pdh->OnSeenNewItems();
                    }
                    else if (pnm->code == 209) // SMN_NEEDREPAINT
                    {
                        return pdh->OnNeedRepaint();
                    }
                }
                return DefWindowProc(hwnd, uMsg, wParam, lParam);
            }

            if (pnm->code == 212)
            {
                pdh->_fMenuBlocked = ((SMNMBOOL*)pnm)->f;
                return DefWindowProc(hwnd, uMsg, wParam, lParam);
            }

            if (pnm->code == 216)
            {
                pdh->OnTrackShellMenu(pnm);
            }
            else
            {
                if (pnm->code == 218)
                {
                    pdh->_OnGetIStartButton(pnm);
                }
                else if (pnm->code == 225)
                {
                    return pdh->OnNeedRebuild();
                }
                return DefWindowProc(hwnd, uMsg, wParam, lParam);
            }
        }
        return 0;
    }
    
    if (uMsg == 71)                             // WM_WINDOWPOSCHANGING
    {
        LPWINDOWPOS pwp = reinterpret_cast<LPWINDOWPOS>(lParam);
        if (pdh && (pwp->flags & SWP_HIDEWINDOW))
        {
            IUnknown_QueryServiceExec(static_cast<IMenuBand*>(pdh), SID_SM_UserPane, &SID_SM_DV2ControlHost, 323, 0, 0, 0);
        }
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    
    if (uMsg > 0x14)
    {
        if (uMsg != 21)
        {
            if (uMsg == 24)
            {
                if (!wParam && pdh)
                {
                    pdh->_SaveChildFocus();
                }
                return DefWindowProcW(hwnd, uMsg, wParam, lParam);
            }
            if (uMsg != 26)
            {
                if (uMsg != 0x1C)
                {
                    if (uMsg == 0x46)
                    {
                        if (pdh)
                        {
                            LPWINDOWPOS pwp = reinterpret_cast<LPWINDOWPOS>(lParam);
                            if ((pwp->flags & 4) == 0)
                            {
                                IStartButton = pdh->_GetIStartButton();
                                if (IStartButton)
                                {
                                    DWORD dwPos;
                                    if (SUCCEEDED(IStartButton->GetPopupPosition(&dwPos)))
                                    {
                                        IStartButton->Release();
                                        if (dwPos == 0x20000000 || dwPos == 0x80000000)
                                        {
                                            pwp->hwndInsertAfter = v_hwndTray;
                                            return 0;
                                        }
                                        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                                    }
                                    return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                                }
                                return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                            }
                            return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                        }
                        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                    }
                    return DefWindowProcW(hwnd, uMsg, wParam, lParam);
                }

                if (wParam)
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                if (!pdh)
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                if (pdh->field_1A8)
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                gui.cbSize = 48;
                if (!GetGUIThreadInfo(0, &gui))
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                for (i = gui.hwndActive; i != hwnd; i = GetParent(i))
                {
                    if (!i)
                        goto LABEL_61;
                }
                v12 = i == 0;
            LABEL_60:
                if (!v12)
                    return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
                goto LABEL_61;
            }
            if (!pdh)
                return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
            if (wParam == 47)
            {
                v25 = 0;
                v14 = pdh->_GetIStartButton();
                if (v14)
                {
                    v14->GetPopupPosition(&v25);
					v14->Release();
                }
                bSkipRebuild = v25 == pdh->field_1AC;
            }
            else
            {
                if (wParam != 24)
                {
                    if (wParam || !pnm || StrCmpICW((LPCWSTR)lParam, L"Policy"))
                        goto LABEL_76;
                    goto LABEL_75;
                }
                if (!pnm)
                {
                LABEL_75:
                    pdh->OnNeedRebuild();
                    goto LABEL_76;
                }
                bSkipRebuild = StrCmpICW((LPCWSTR)lParam, L"Desktop") == 0;
            }
            if (bSkipRebuild)
            {
            LABEL_76:
                pdh->OnNeedRepaint();
                pnm = (LPNMHDR)lParam;
                SHPropagateMessage(hwnd, 26, wParam, lParam, 3);
                return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
            }
            goto LABEL_75;
        }
    LABEL_98:
        if (pdh)
        {
            pdh->OnNeedRepaint();
            SHPropagateMessage(hwnd, uMsg, wParam, (LPARAM)pnm, 3);
        }
        return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
    }
    switch (uMsg)
    {
    case 0x14u:
        if (pdh)
            pdh->OnPaint((HDC)wParam, 1);
        return 1;
    case 1u:
        if (pdh)
            pdh->OnCreate(hwnd);
        else
            return -1;
        break;
    case 2u:
        if (pdh)
            pdh->OnDestroy();
        return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
    case 6u:
        if (!pdh)
            return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);

        if (LOWORD(wParam) != WA_INACTIVE)
        //if ((_WORD)wParam)
        {
            pdh->_RestoreChildFocus();
            return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
        }

        pdh->_SaveChildFocus();
        Ancestor = GetAncestor((HWND)pnm, 3u);
        if (pdh->field_1A8
            || hwnd == Ancestor
            || Ancestor == pdh->_hwndParent && pdh->_ShouldIgnoreFocusChange((HWND)pnm)
            || pdh->_ppmTracking
            || !IsWindowEnabled(hwnd))
        {
            return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
        }
    LABEL_61:
        SendMessageW(hwnd, 0x8000u, 0, 0);
        return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
    case 7u:
        if (pdh)
            pdh->OnSetFocus((HWND)wParam);
        break;
    case 0x10u:
        if (pdh)
            pdh->_OnDismiss(0);
        return 0;
    default:
        return DefWindowProcW(hwnd, uMsg, wParam, (LPARAM)pnm);
    }
    return 0;
#endif
}

//
//  If the user executes something or cancels out, we dismiss ourselves.
//

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::OnSelect(DWORD dwSelectType)
{
    HRESULT hr = E_NOTIMPL;

    switch (dwSelectType)
    {
        case MPOS_EXECUTE:
        case MPOS_CANCELLEVEL:
        {
            _DismissMenuPopup();
            hr = S_OK;
            break;
        }
        case MPOS_FULLCANCEL:
        {
            if (!_fDismissOnlyPopup && !field_1A8)
            {
                _DismissMenuPopup();
            }
            hr = S_OK;
            break;
        }
        case MPOS_SELECTLEFT:
        {
            _DismissTrackShellMenu();

            NMHDR nm;
            nm.hwndFrom = _hwnd;
            nm.idFrom = 0;
            nm.code = 224;
            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);
            hr = S_OK;
            break;
        }
    }
    return hr;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_DismissTrackShellMenu()
{
    if (_ppmTracking)
    {
        _fDismissOnlyPopup++;
        _ppmTracking->OnSelect(MPOS_FULLCANCEL);
        _fDismissOnlyPopup--;
    }
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_CleanupTrackShellMenu()
{
    IUnknown_SafeReleaseAndNullPtr(&_ppmTracking);
    _hwndTracking = NULL;
    _hwndAltTracking = NULL;
    KillTimer(_hwnd, IDT_MENUCHANGESEL);

    NMHDR nm = { _hwnd, 0, SMN_SHELLMENUDISMISSED };
    SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_DismissMenuPopup()
{
    DesktopHost_Dismiss(_hwnd);
}

BOOL CDesktopHost::_FilterMouseWheel(MSG *pmsg, HWND hwndTarget)
{
    BOOL fRet = 0;
    if (hwndTarget)
    {
        SMNDIALOGMESSAGE pdm;
        pdm.pmsg = pmsg;
        pdm.pt.x = GET_X_LPARAM(pmsg->lParam);
        pdm.pt.y = GET_Y_LPARAM(pmsg->lParam);
        return _FindChildItem(hwndTarget, &pdm, 0xBu);
    }
    return fRet;
}

BOOL CDesktopHost::_FilterMouseButtonDown(MSG *pmsg, HWND hwndTarget)
{
    SMNDIALOGMESSAGE pdm;
    pdm.hwnd = pmsg->hwnd;
    pdm.pmsg = pmsg;
    pdm.pt.y = GET_Y_LPARAM(pmsg->lParam);
    pdm.pt.x = GET_X_LPARAM(pmsg->lParam);
    return _FindChildItem(hwndTarget, &pdm, 0x117u) == 0;
}

BOOL CDesktopHost::_FilterLMouseButtonUp(MSG *pmsg, HWND hwndTarget)
{
    SMNDIALOGMESSAGE pdm;
    pdm.hwnd = pmsg->hwnd;
    pdm.pmsg = pmsg;
    pdm.pt.y = GET_Y_LPARAM(pmsg->lParam);
    pdm.pt.x = GET_X_LPARAM(pmsg->lParam);
    return _FindChildItem(hwndTarget, &pdm, 6u);
}

BOOL CDesktopHost::_DlgNavigateTab(HWND hwndStart, struct tagMSG *pmsg)
{
    SHORT KeyState; // ax
    SMNDIALOGMESSAGE pdm; // [esp+4h] [ebp-28h] BYREF

    KeyState = GetKeyState(16);
    pdm.pmsg = pmsg;
    CDesktopHost::_DlgFindItem(
        hwndStart,
        &pdm,
        ((KeyState < 0) + 3) | 0x500,
        GetNextDlgGroupItem,
        (KeyState < 0) | 2);
    return 1;
}

void CDesktopHost::_RemoveKeyboardCues()
{
    SendMessage(this->_hwnd, 0x127u, 0x30001u, 0);
}

//
//  The PMs want custom keyboard navigation behavior on the Start Panel,
//  so we have to do it all manually.
//
BOOL CDesktopHost::_IsDialogMessage(MSG *pmsg)
{
#ifdef DEAD_CODE
    //
    //  If the menu isn't even open or if menu mode is blocked, then
    //  do not mess with the message.
    //
    if (!_fOpen || _fMenuBlocked) {
        return FALSE;
    }

    //
    //  Tapping the ALT key dismisses menus.
    //
    if (pmsg->message == WM_SYSKEYDOWN && pmsg->wParam == VK_MENU)
    {
        DesktopHost_Dismiss(_hwnd);
        // For accessibility purposes, dismissing the
        // Start Menu should place focus on the Start Button.
        SetFocus(c_tray._stb._hwndStart);
        return TRUE;
    }

    if (SHIsChildOrSelf(_hwnd, pmsg->hwnd) != S_OK) {
        //
        //  If this is an uncaptured mouse move message, then eat it.
        //  That's what menus do -- they eat mouse moves.
        //  Let clicks go through, however, so the user
        //  can click away to dismiss the menu and activate
        //  whatever they clicked on.
        if (!GetCapture() && pmsg->message == WM_MOUSEMOVE) {
            return TRUE;
        }

        return FALSE;
    }

    //
    // Destination window must be a grandchild of us.  The child is the
    // host control; the grandchild is the real control.  Note also that
    // we do not attempt to modify the behavior of great-grandchildren,
    // because that would mess up inplace editing (which creates an
    // edit control as a child of the listview).

    HWND hwndTarget = GetParent(pmsg->hwnd);
    if (hwndTarget != NULL && GetParent(hwndTarget) != _hwnd)
    {
        hwndTarget = NULL;
    }

    //
    //  Intercept mouse messages so we can do mouse hot tracking goo.
    //  (But not if a client has blocked menu mode because it has gone
    //  into some modal state.)
    //
    switch (pmsg->message) {
        case WM_MOUSEMOVE:
            _FilterMouseMove(pmsg, hwndTarget);
            break;

        case WM_MOUSELEAVE:
            _FilterMouseLeave(pmsg, hwndTarget);
            break;

        case WM_MOUSEHOVER:
            _FilterMouseHover(pmsg, hwndTarget);
            break;

    }

    //
    //  Keyboard messages require a valid target.
    //
    if (hwndTarget == NULL) {
        return FALSE;
    }

    //
    //  Okay, hwndTarget is the host control that understands our
    //  wacky notification messages.
    //

    switch (pmsg->message)
    {
        case WM_KEYDOWN:
            _EnableKeyboardCues();

            switch (pmsg->wParam)
            {
                case VK_LEFT:
                case VK_RIGHT:
                case VK_UP:
                case VK_DOWN:
                    return _DlgNavigateArrow(hwndTarget, pmsg);

                case VK_ESCAPE:
                case VK_CANCEL:
                    DesktopHost_Dismiss(_hwnd);
                    // For accessibility purposes, hitting ESC to dismiss the
                    // Start Menu should place focus on the Start Button.
                    SetFocus(c_tray._stb._hwndStart);
                    return TRUE;

                case VK_RETURN:
                    _FindChildItem(hwndTarget, NULL, SMNDM_INVOKECURRENTITEM | SMNDM_KEYBOARD);
                    return TRUE;

                    // Eat space
                case VK_SPACE:
                    return TRUE;

                default:
                    break;
            }
            return FALSE;

            // Must dispatch there here so Tray's TranslateAccelerator won't see them
        case WM_SYSKEYDOWN:
        case WM_SYSKEYUP:
        case WM_SYSCHAR:
            DispatchMessage(pmsg);
            return TRUE;

        case WM_CHAR:
            return _DlgNavigateChar(hwndTarget, pmsg);

    }

    return FALSE;
#else
    unsigned int message; // eax
    HWND hwndTarget; // eax MAPDST
    HWND Parent; // eax
    BOOL v12; // edi
    int wParam; // eax
    LRESULT ChildItem; // eax
    HWND hwndFocus; // eax
    HWND v19; // eax
    SMNDIALOGMESSAGE lParam; // [esp+Ch] [ebp-50h] BYREF
    POINT pt; // [esp+54h] [ebp-8h] BYREF

    message = pmsg->message;
    if (message == 0x201 || message == 0x204)
        _UnlockStartPane();

    if (this->_fOpen && !this->_fMenuBlocked)
    {
        if (SHIsChildOrSelf(this->_hwnd, pmsg->hwnd))
        {
            return !GetCapture() && pmsg->message == 0x200;
        }

        hwndTarget = GetParent(pmsg->hwnd);
        if (hwndTarget)
        {
            Parent = GetParent(hwndTarget);
            if (Parent != this->_hwnd && Parent != this->_spm.panes[1].hwnd)
            {
                VARIANT vtIn;
                vtIn.vt = VT_BYREF;
                vtIn.byref = Parent;

                VARIANT vtOut;
                vtOut.vt = VT_BYREF;
                vtOut.byref = NULL;
                IUnknown_QueryServiceExec(static_cast<IMenuBand*>(this), SID_SM_OpenHost, &SID_SM_DV2ControlHost, 306, 0, &vtIn, &vtOut);
                hwndTarget = (HWND)vtOut.byref;

                if (!vtOut.byref)
                {
                    IUnknown_QueryServiceExec(static_cast<IMenuBand *>(this), SID_SM_OpenBox, &SID_SM_DV2ControlHost, 306, 0, &vtIn, &vtOut);
                    hwndTarget = (HWND)vtOut.byref;
                }
            }
        }

        if (pmsg->message > 0x205)
        {
            if (pmsg->message == WM_MOUSEWHEEL)
            {
                if (this->field_48 && _DoesOpenBoxHaveFocus())
                {
                    hwndTarget = this->field_48;
                    _FilterMouseWheel(pmsg, hwndTarget);
                }
            }
            else if (pmsg->message == WM_MOUSEHOVER)
            {
                _FilterMouseHover(pmsg, hwndTarget);
            }
            else if (pmsg->message == WM_MOUSELEAVE)
            {
                _FilterMouseLeave(pmsg, hwndTarget);
            }
        }
        else if (pmsg->message == WM_RBUTTONUP)
        {
            if (_FilterMouseButtonDown(pmsg, hwndTarget) || SHRestricted(REST_NOCHANGESTARMENU))
            {
                pt.x = GET_X_LPARAM(pmsg->lParam);
                pt.y = GET_Y_LPARAM(pmsg->lParam);
                MapWindowPoints(pmsg->hwnd, this->_hwnd, &pt, 1u);
                MAKELPARAM(pt.x, pt.y);
                pmsg->hwnd = this->_hwnd;
            }
        }
        else
        {
            if (pmsg->message == WM_MOUSEMOVE)
            {
                return _FilterMouseMove(pmsg, hwndTarget);
            }

            if (pmsg->message == WM_LBUTTONDOWN)
            {
                goto LABEL_21;
            }

            if (pmsg->message == WM_LBUTTONUP)
            {
                return _FilterLMouseButtonUp(pmsg, hwndTarget);
            }
            if (pmsg->message == WM_RBUTTONDOWN)
            {
            LABEL_21:
                v12 = _FilterMouseButtonDown(pmsg, hwndTarget);
                if (pmsg->message == 516 && SHRestricted(REST_NOCHANGESTARMENU))
                {
                    pt.x = GET_X_LPARAM(pmsg->lParam);
					pt.y = GET_Y_LPARAM(pmsg->lParam);
                    MapWindowPoints(pmsg->hwnd, this->_hwnd, &pt, 1u);
                    MAKELPARAM(pt.x, pt.y);
                    pmsg->hwnd = this->_hwnd;
                }
                return v12;
            }
        }
        
        v12 = 0;
        if (hwndTarget)
        {
            if (pmsg->message == WM_CONTEXTMENU)
            {
                if (HIDWORD(pmsg->wParam) == -1)
                {
                    if (this->field_48 && SHIsChildOrSelf(this->_spm.panes[2].hwnd, this->field_48))
                    {
                        v19 = this->field_48;
                        if (v19)
                        {
                            pmsg->hwnd = v19;
                        }
                    }
                    else
                    {
                        VARIANT vt;
                        vt.vt = VT_BYREF;
                        vt.byref = pmsg;
                        IUnknown_QueryServiceExec(static_cast<IMenuBand *>(this), SID_SM_OpenHost, &SID_SM_DV2ControlHost, 330, 0, &vt, &vt);
                    }
                    _UnlockStartPane();
                }
                return 0;
            }

            if (pmsg->message == WM_KEYDOWN)
            {
                if (this->field_48 && _DoesOpenBoxHaveFocus())
                {
                    hwndTarget = this->field_48;
                }

                wParam = pmsg->wParam;
                if (wParam != 3)
                {
                    if (wParam == 9)
                    {
                        ChildItem = _DlgNavigateTab(hwndTarget, pmsg);
                        goto LABEL_55;
                    }
                    if (wParam == 13)
                    {
                        _UnlockStartPane();
                        lParam.pmsg = pmsg;
                        _FindChildItem(hwndTarget, &lParam, 0x406u);
                        v12 = (lParam.flags & 0x1000) == 0;
                        goto LABEL_63;
                    }
                    
                    if (wParam != 27)
                    {
                        if (wParam != 32)
                        {
                            if ((unsigned int)(wParam - 37) > 3)
                            {
                                goto LABEL_63;
                            }
                            ChildItem = _DlgNavigateArrow(hwndTarget, pmsg);
                            goto LABEL_55;
                        }
                        if (!_DoesOpenBoxHaveFocus())
                        {
                            lParam.pmsg = pmsg;
                            ChildItem = _FindChildItem(hwndTarget, &lParam, 0x40Au);
                        LABEL_55:
                            v12 = ChildItem;
                        }
                    LABEL_63:
                        hwndFocus = GetFocus();
                        if (SHIsChildOrSelf(this->_spm.panes[2].hwnd, hwndFocus))
                            _EnableKeyboardCues();
                        else
                            _RemoveKeyboardCues();
                        return v12;
                    }
                }

                VARIANT vt;
                vt.iVal = -1;
                vt.vt = VT_BOOL;
                IUnknown_QueryServiceExec(static_cast<IMenuBand*>(this), SID_SM_OpenHost, &SID_SM_DV2ControlHost, 307, 0, 0, &vt);
                if (vt.iVal)
                {
                    SendMessage(this->_hwnd, 0x8000u, 0, 0);
                    _SetFocusToStartButton();
                }
                v12 = 1;
                goto LABEL_63;
            }
            
            if (pmsg->message == 258)
            {
                return _DlgNavigateChar(hwndTarget, pmsg);
            }
            if (pmsg->message - 260 <= 2)
            {
                TranslateMessage(pmsg);
                DispatchMessage(pmsg);
                return 1;
            }
            return 0;
        }
    }
    return 0;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup
LRESULT CDesktopHost::_FindChildItem(HWND hwnd, SMNDIALOGMESSAGE* pnmdm, UINT smndm)
{
#ifdef DEAD_CODE
    SMNDIALOGMESSAGE nmdm;
    if (!pnmdm)
    {
        pnmdm = &nmdm;
    }

    pnmdm->hdr.hwndFrom = _hwnd;
    pnmdm->hdr.idFrom = 0;
    pnmdm->hdr.code = SMN_FINDITEM;
    pnmdm->flags = smndm;

    LRESULT lres = ::SendMessage(hwnd, WM_NOTIFY, 0, (LPARAM)pnmdm);

    if (lres && (smndm & SMNDM_SELECT))
    {
        SetFocus(::GetWindow(hwnd, GW_CHILD));
    }

    return lres;
#else
    UINT flags; // eax
    HWND v9; // eax
    HWND v10; // eax

    SMNDIALOGMESSAGE nmdm; // [esp+Ch] [ebp-38h] BYREF
    if (!pnmdm)
    {
        pnmdm = &nmdm;
    }

    pnmdm->hdr.idFrom = 0;
    pnmdm->hdr.hwndFrom = _hwnd;
    pnmdm->hdr.code = 215;
    pnmdm->flags = smndm;
    pnmdm->field_24 =_hwndLastMouse;
    
    LRESULT lres = SendMessage(hwnd, WM_NOTIFY, 0, (LPARAM)pnmdm);
    if (lres)
    {
        flags = pnmdm->flags;
        if ((flags & 0x100) != 0)
        {
            if ((flags & 0x80000) == 0)
            {
                pnmdm->field_24 = ::GetWindow(hwnd, GW_CHILD);
            }

            // Logging the HWND and window text
            TCHAR szTitle[128] = { 0 };
            GetWindowText(pnmdm->field_24, szTitle, ARRAYSIZE(szTitle));
            TCHAR szLog[256];
            wsprintf(szLog, TEXT("SetFocus to HWND: 0x%p, Title: \"%s\"\n"), pnmdm->field_24, szTitle);
            OutputDebugString(szLog);

            SetFocus(pnmdm->field_24);
            field_48 = 0;
        }
    }
    
    if ((pnmdm->flags & 0x800) != 0)
    {
        if (!lres)
        {
            hwnd = _spm.panes[2].hwnd;
        }

        v9 = field_48;
        if (v9 != hwnd)
        {
            if (v9)
            {
                _RemoveSelection(field_48);
            }
            
            v10 = _spm.panes[2].hwnd;
            if (hwnd != v10 || field_48)
            {
                field_48 = hwnd;

                VARIANT vt;
                vt.vt = VT_INT;
                vt.lVal = pnmdm->flags;
                if (hwnd == v10)
                {
                    vt.lVal |= 0x20000u;
                }
                IUnknown_QueryServiceExec((IMenuBand*)this, SID_SM_OpenView, &SID_SM_DV2ControlHost, 304, 0, &vt, 0);
            }
        }
    }
    return lres;
#endif
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_EnableKeyboardCues()
{
    SendMessage(_hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEFOCUS | UISF_HIDEACCEL), 0);
}


//
//  _DlgFindItem does the grunt work of walking the group/tab order
//  looking for an item.
//
//  hwndStart = window after which to start searching
//  pnmdm = structure to receive results
//  smndm = flags for _FindChildItem call
//  GetNextDlgItem = GetNextDlgTabItem or GetNextDlgGroupItem
//  fl = flags (DFI_*)
//
//  DFI_INCLUDESTARTLAST:  Include hwndStart at the end of the search.
//                         Otherwise do not search in hwndStart.
//
//  Returns the found window, or NULL.
//

#define DFI_FORWARDS            0x0000
#define DFI_BACKWARDS           0x0001

#define DFI_INCLUDESTARTLAST    0x0002

// EXEX-VISTA(allison): Validated.
HWND CDesktopHost::_DlgFindItem(
    HWND hwndStart, SMNDIALOGMESSAGE* pnmdm, UINT smndm,
    GETNEXTDLGITEM GetNextDlgItem, UINT fl)
{
    HWND hwndT = hwndStart;
    int iLoopCount = 0;

    while ((hwndT = GetNextDlgItem(_hwnd, hwndT, fl & DFI_BACKWARDS)) != NULL)
    {
        if (!(fl & DFI_INCLUDESTARTLAST) && hwndT == hwndStart)
        {
            return NULL;
        }

        if (_FindChildItem(hwndT, pnmdm, smndm))
        {
            return hwndT;
        }

        if (hwndT == hwndStart)
        {
            ASSERT(fl & DFI_INCLUDESTARTLAST);
            return NULL;
        }

        if (++iLoopCount > 10)
        {
            // If this assert fires, it means that the controls aren't
            // playing nice with WS_TABSTOP and WS_GROUP and we got stuck.
            ASSERT(iLoopCount < 10);
            return NULL;
        }

    }
    return NULL;
}

BOOL CDesktopHost::_DlgNavigateArrow(HWND hwndStart, MSG* pmsg)
{
    HWND hwndT;
    SMNDIALOGMESSAGE nmdm;
    MSG msg;
    nmdm.pmsg = pmsg;   // other fields will be filled in by _FindChildItem

    //TraceMsg(TF_DV2DIALOG, "idm.arrow(%04x)", pmsg->wParam);

    // If RTL, then flip the left and right arrows
    UINT vk = (UINT)pmsg->wParam;
    BOOL fRTL = GetWindowLong(_hwnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL;
    if (fRTL)
    {
        if (vk == VK_LEFT) vk = VK_RIGHT;
        else if (vk == VK_RIGHT) vk = VK_LEFT;
        // Put the flipped arrows into the MSG structure so clients don't
        // have to know anything about RTL.
        msg = *pmsg;
        nmdm.pmsg = &msg;
        msg.wParam = vk;
    }
    BOOL fBackwards = vk == VK_LEFT || vk == VK_UP;
    BOOL fVerticalKey = vk == VK_UP || vk == VK_DOWN;


    //
    //  First see if the navigation can be handled by the control natively.
    //  We have to let the control get first crack because it might want to
    //  override default behavior (e.g., open a menu when VK_RIGHT is pressed
    //  instead of moving to the right).
    //

    //
    //  Holding the shift key while hitting the Right [RTL:Left] arrow
    //  suppresses the attempt to cascade.
    //

    DWORD dwTryCascade = 0;
    if (vk == VK_RIGHT && GetKeyState(VK_SHIFT) >= 0)
    {
        dwTryCascade |= SMNDM_TRYCASCADE;
    }

    if (_FindChildItem(hwndStart, &nmdm, dwTryCascade | SMNDM_FINDNEXTARROW | SMNDM_SELECT | SMNDM_KEYBOARD))
    {
        // That was easy
        return TRUE;
    }

    //
    //  If the arrow key is in alignment with the control's orientation,
    //  then walk through the other controls in the group until we find
    //  one that contains an item, or until we loop back.
    //

    ASSERT(nmdm.flags & (SMNDM_VERTICAL | SMNDM_HORIZONTAL));

    // Save this because subsequent callbacks will wipe it out.
    DWORD dwDirection = nmdm.flags;

    //
    //  Up/Down arrow always do prev/next.  Left/right arrow will
    //  work if we are in a horizontal control.
    //
    if (fVerticalKey || (dwDirection & SMNDM_HORIZONTAL))
    {
        // Search for next/prev control in group.

        UINT smndm = fBackwards ? SMNDM_FINDLAST : SMNDM_FINDFIRST;
        UINT fl = fBackwards ? DFI_BACKWARDS : DFI_FORWARDS;

        hwndT = _DlgFindItem(hwndStart, &nmdm,
            smndm | SMNDM_SELECT | SMNDM_KEYBOARD,
            GetNextDlgGroupItem,
            fl | DFI_INCLUDESTARTLAST);

        // Always return TRUE to eat the message
        return TRUE;
    }

    //
    //  Navigate to next column or row.  Look for controls that intersect
    //  the x (or y) coordinate of the current item and ask them to select
    //  the nearest available item.
    //
    //  Note that in this loop we do not want to let the starting point
    //  try again because it already told us that the navigation key was
    //  trying to leave the starting point.
    //

    //
    //  Note: For RTL compatibility, we must map rectangles.
    //
    RECT rcSrc = { nmdm.pt.x, nmdm.pt.y, nmdm.pt.x, nmdm.pt.y };
    MapWindowRect(hwndStart, HWND_DESKTOP, &rcSrc);
    hwndT = hwndStart;

    while ((hwndT = GetNextDlgGroupItem(_hwnd, hwndT, fBackwards)) != NULL &&
        hwndT != hwndStart)
    {
        // Does this window intersect in the desired direction?
        RECT rcT;
        BOOL fIntersect;

        GetWindowRect(hwndT, &rcT);
        if (dwDirection & SMNDM_VERTICAL)
        {
            rcSrc.left = rcSrc.right = fRTL ? rcT.right : rcT.left;
            fIntersect = rcSrc.top >= rcT.top && rcSrc.top < rcT.bottom;
        }
        else
        {
            rcSrc.top = rcSrc.bottom = rcT.top;
            fIntersect = rcSrc.left >= rcT.left && rcSrc.left < rcT.right;
        }

        if (fIntersect)
        {
            rcT = rcSrc;
            MapWindowRect(HWND_DESKTOP, hwndT, &rcT);
            nmdm.pt.x = rcT.left;
            nmdm.pt.y = rcT.top;
            if (_FindChildItem(hwndT, &nmdm,
                SMNDM_FINDNEAREST | SMNDM_SELECT | SMNDM_KEYBOARD))
            {
                return TRUE;
            }
        }
    }

    // Always return TRUE to eat the message
    return TRUE;
}

//
// Find the next/prev tabstop and tell it to select its first item.
// Keep doing this until we run out of controls or we find a control
// that is nonempty.
//

// EXEX-VISTA(allison): Validated.
HWND CDesktopHost::_FindNextDlgChar(HWND hwndStart, SMNDIALOGMESSAGE* pnmdm, UINT smndm)
{
    //
    //  See if there is a match in the hwndStart control.
    //
    if (_FindChildItem(hwndStart, pnmdm, SMNDM_FINDNEXTMATCH | SMNDM_KEYBOARD | smndm))
    {
        return hwndStart;
    }

    if ((pnmdm->flags & 0x1000) != 0)
    {
        return NULL;
    }

    //
    //  Oh well, look for some other control, possibly wrapping back around
    //  to the start.
    //
    return _DlgFindItem(hwndStart, pnmdm,
        SMNDM_FINDFIRSTMATCH | SMNDM_KEYBOARD | smndm,
        GetNextDlgGroupItem,
        DFI_FORWARDS | DFI_INCLUDESTARTLAST);

}

//
//  Find the next item that begins with the typed letter and
//  invoke it if it is unique.
//

// EXEX-VISTA(allison): Validated.
BOOL CDesktopHost::_DlgNavigateChar(HWND hwndStart, MSG *pmsg)
{
#ifdef DEAD_CODE
    SMNDIALOGMESSAGE nmdm;
    nmdm.pmsg = pmsg;   // other fields will be filled in by _FindChildItem

    //
    //  See if there is a match in the hwndStart control.
    //
    HWND hwndFound = _FindNextDlgChar(hwndStart, &nmdm, SMNDM_SELECT);
    if (hwndFound)
    {
        LRESULT idFound = nmdm.itemID;

        //
        //  See if there is another match for this character.
        //  We are only looking, so don't pass SMNDM_SELECT.
        //
        HWND hwndFound2 = _FindNextDlgChar(hwndFound, &nmdm, 0);
        if (hwndFound2 == hwndFound && nmdm.itemID == idFound)
        {
            //
            //  There is only one item that begins with this character.
            //  Invoke it!
            //
            UpdateWindow(_hwnd);
            _FindChildItem(hwndFound2, &nmdm, SMNDM_INVOKECURRENTITEM | SMNDM_KEYBOARD);
        }
    }

    return TRUE;
#else
    BOOL bRet = FALSE;
    if (SHIsChildOrSelf(_spm.panes[2].hwnd, pmsg->hwnd))
    {
        SMNDIALOGMESSAGE nmdm;
        nmdm.pmsg = pmsg;
        if (_FindNextDlgChar(hwndStart, &nmdm, SMNDM_SELECT))
            bRet = TRUE;
    }
    return bRet;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
int CDesktopHost::_FilterMouseMove(MSG *pmsg, HWND hwndTarget)
{
    if (!_fMouseEntered)
    {
        _fMouseEntered = TRUE;
        TRACKMOUSEEVENT tme;
        tme.hwndTrack = pmsg->hwnd;
        field_64 = tme.hwndTrack;
        tme.cbSize = sizeof(tme);
        tme.dwFlags = TME_LEAVE;
        TrackMouseEvent(&tme);
    }

    HWND hwndLastMouse = _hwndLastMouse;
    if (!hwndLastMouse && !_lParamLastMouse)
    {
        _hwndLastMouse = pmsg->hwnd;
        _lParamLastMouse = pmsg->lParam;
        return 1;
    }

    if (hwndLastMouse == pmsg->hwnd && _lParamLastMouse == pmsg->lParam)
        return 1;

    _hwndLastMouse = pmsg->hwnd;
    _lParamLastMouse = pmsg->lParam;

    int v6 = 0;
    LRESULT lres = 0;
    if (hwndTarget)
    {
        SMNDIALOGMESSAGE nmdm;
        nmdm.pmsg = pmsg;
        nmdm.hwnd = pmsg->hwnd;
        nmdm.pt.x = GET_X_LPARAM(pmsg->lParam);
        nmdm.pt.y = GET_Y_LPARAM(pmsg->lParam);
        lres = _FindChildItem(hwndTarget, &nmdm, (_DoesOpenBoxHaveFocus() ? 2048 : 256) | 7);

        v6 = nmdm.flags & 0x2000;
    }

    if (!v6)
    {
        if (!lres)
        {
            _RemoveSelection(hwndTarget);
        }
        else
        {
            if (_fAutoCascade)
            {
                TRACKMOUSEEVENT tme;
                tme.cbSize = sizeof(tme);
                tme.dwFlags = TME_HOVER;
                tme.hwndTrack = pmsg->hwnd;
                if (!SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0, &tme.dwHoverTime, 0))
                {
                    tme.dwHoverTime = HOVER_DEFAULT;
                }
                TrackMouseEvent(&tme);
            }
        }
    }
    return 0;
}

// EXEX-VISTA(allison): Validated. Still needs very minor cleanup.
void CDesktopHost::_FilterMouseLeave(MSG* pmsg, HWND hwndTarget)
{
#ifdef DEAD_CODE
    _fMouseEntered = FALSE;
    _hwndLastMouse = NULL;

    // If we got a WM_MOUSELEAVE due to a menu popping up, don't
    // give up the focus since it really didn't leave yet.
    if (!_ppmTracking)
    {
        _RemoveSelection(hwndTarget);
    }
#else
    _fMouseEntered = FALSE;

    if (!_ppmTracking && field_64 == pmsg->hwnd)
    {
        NMHDR nm;
        nm.hwndFrom = hwndTarget;
        nm.idFrom = GetDlgCtrlID(hwndTarget);
        nm.code = 225;
        if (!SendMessage(hwndTarget, WM_NOTIFY, nm.idFrom, (LPARAM)&nm))
        {
            _RemoveSelection(hwndTarget);
            if (hwndTarget != _spm.panes[2].hwnd && _DoesOpenBoxHaveFocus())
            {
                VARIANT vt;
                vt.vt = VT_INT;
                vt.lVal = 0x20000;
                IUnknown_QueryServiceExec(static_cast<IMenuBand*>(this), SID_SM_OpenView, &SID_SM_DV2ControlHost, 304, 0, &vt, NULL);
            }
        }
    }
    
    _hwndLastMouse = NULL;
    _lParamLastMouse = 0;
    field_64 = NULL;
#endif
}

// EXEX-Vista(allison): Validated.
void CDesktopHost::_FilterMouseHover(MSG *pmsg, HWND hwndTarget)
{
    SMNDIALOGMESSAGE nmdm;
    nmdm.hwnd = pmsg->hwnd;
    nmdm.pmsg = pmsg;
    nmdm.pt.y = GET_Y_LPARAM(pmsg->lParam);
    nmdm.pt.x = GET_X_LPARAM(pmsg->lParam);
    CDesktopHost::_FindChildItem(hwndTarget, &nmdm, SMNDM_OPENCASCADE);
}

//
//  Remove the menu selection and put it in the "dead space" above
//  the first visible item.
//

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_RemoveSelection(HWND hwnd)
{
    if (hwnd)
    {
        NMHDR nm;
        nm.hwndFrom = hwnd;
        nm.idFrom = GetDlgCtrlID(hwnd);
        nm.code = NM_KILLFOCUS;
        SendMessage(hwnd, WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
    }

    if (!_DoesOpenBoxHaveFocus())
    {
        HWND hwndChild = GetNextDlgTabItem(_hwnd, NULL, FALSE);
        if (hwndChild)
        {
            HWND hwndInner = ::GetWindow(hwndChild, GW_CHILD);
            SetFocus(hwndInner);

            NMHDR nm;
            nm.hwndFrom = hwndInner;
            nm.idFrom = GetDlgCtrlID(hwndInner);
            nm.code = NM_KILLFOCUS;
            SendMessage(hwndChild, WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
        }
    }
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::IsMenuMessage(MSG *pmsg)
{
    if (!_hwnd)
        return E_FAIL;

    if (_ppmTracking)
    {
        HRESULT hr = _ppmTracking->IsMenuMessage(pmsg);
        if (hr == E_FAIL)
        {
            _CleanupTrackShellMenu();
            hr = S_FALSE;
        }

        if (hr == S_OK)
        {
            return hr;
        }
    }
    return _IsDialogMessage(pmsg) ? S_OK : S_FALSE;
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::TranslateMenuMessage(MSG* pmsg, LRESULT* plres)
{
    if (_ppmTracking)
    {
        return _ppmTracking->TranslateMenuMessage(pmsg, plres);
    }
    return E_NOTIMPL;
}

// IServiceProvider::QueryService
// EXEX-VISTA(allison): Validated.
STDMETHODIMP CDesktopHost::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    HRESULT hr = E_FAIL;

    if (IsEqualGUID(guidService, SID_SMenuPopup))
    {
        hr = QueryInterface(riid, ppvObject);
    }
    else if (IsEqualGUID(guidService, SID_SM_OpenView)
        || IsEqualGUID(guidService, SID_SM_TopMatch)
        || IsEqualGUID(guidService, SID_SM_OpenHost))
    {
        ASSERT(_spm.panes[SMPANETYPE_OPENVIEWHOST].punk != NULL); // 2410
        hr = IUnknown_QueryService(_spm.panes[SMPANETYPE_OPENVIEWHOST].punk, guidService, riid, ppvObject);
    }
    else if (IsEqualGUID(guidService, SID_SM_OpenBox))
    {
        ASSERT(_spm.panes[SMPANETYPE_OPENBOX].punk != NULL); // 2415
        hr = IUnknown_QueryService(_spm.panes[SMPANETYPE_OPENBOX].punk, guidService, riid, ppvObject);
    }
    else if (IsEqualGUID(guidService, SID_SM_UserPane))
    {
        ASSERT(_spm.panes[SMPANETYPE_USER].punk != NULL); // 2420
        hr = IUnknown_QueryService(_spm.panes[SMPANETYPE_USER].punk, guidService, riid, ppvObject);
    }
    else if (IsEqualGUID(guidService, IID_IFolderView))
    {
        ASSERT(_spm.panes[SMPANETYPE_KNOWNFOLDER].punk != NULL); // 2425
        hr = IUnknown_QueryService(_spm.panes[SMPANETYPE_KNOWNFOLDER].punk, guidService, riid, ppvObject);
    }

    if (FAILED(hr))
    {
        return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
    }
    return hr;
}

// *** IOleCommandTarget ***
STDMETHODIMP  CDesktopHost::QueryStatus(const GUID* pguidCmdGroup,
    ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pcmdtext)
{
    return E_NOTIMPL;
}

void CDesktopHost::_SetFocusToOpenBox()
{
    if (this->_fOpen)
        _FindChildItem(this->_spm.panes[2].hwnd, 0, 0x103u);
}

BOOL CDesktopHost::_DoesOpenBoxHaveFocus()
{
    return SHIsChildOrSelf(this->_spm.panes[2].hwnd, GetFocus()) == 0;
}

// EXEX-VISTA(allison): Validated. Still needs slight cleanup.
STDMETHODIMP  CDesktopHost::Exec(const GUID *pguidCmdGroup,
    DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut)
{
    if (IsEqualGUID(CLSID_MenuBand, *pguidCmdGroup))
    {
        if (nCmdID == MBANDCID_REFRESH)
        {
            NMHDR nm = { _hwnd, 0, SMN_REFRESHLOGOFF };
            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);
            OnNeedRepaint();
        }
        return 0;
    }
    
    if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
    {
        switch (nCmdID)
        {
            case 300u:
            {
                if (field_48)
                    _RemoveSelection(field_48);
                field_48 = 0;

                VARIANT vt;
                vt.vt = VT_INT;
                vt.lVal = 0x20000;
                IUnknown_QueryServiceExec(SAFECAST(this, IMenuBand *), SID_SM_OpenView, &SID_SM_DV2ControlHost, 304, 0, &vt, 0);
                return 0;
            }
            case 309u:
            {
                pvarargOut->vt = VT_BOOL;
                pvarargOut->iVal = -_DoesOpenBoxHaveFocus();
                return 0;
            }
            case 310u:
            {
                if (field_48)
                    _RemoveSelection(field_48);
                field_48 = 0;
                _SetFocusToOpenBox();
                return 0;
            }
            case 312u:
            {
                field_1A8 = VariantToBooleanWithDefault(*pvarargIn, FALSE);
                return 0;
            }
        }
        
        if (nCmdID == 324)
        {
            ASSERT(pvarargIn->vt == VT_BYREF); // 2495
			//return _HandleOpenBoxArrowKey(pvarargIn); // EXEX-VISTA(allison): TODO: Uncomment when implemented
        }
        if (nCmdID == 326)
        {
            _LockStartPane();
        }
        else if (nCmdID == 327)
        {
            pvarargOut->vt = 11;
            pvarargOut->iVal = -(field_C4 != 0);
        }
        return 0;
    }
    return 0;
}

// ITrayPriv2::ModifySMInfo
// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::ModifySMInfo(IN LPSMDATA psmd, IN OUT SMINFO* psminfo)
{
    if (_hwndNewHandler)
    {
        SMNMMODIFYSMINFO nmsmi;
        nmsmi.hdr.hwndFrom = _hwnd;
        nmsmi.hdr.idFrom = 0;
        nmsmi.hdr.code = SMN_MODIFYSMINFO;
        nmsmi.psmd = psmd;
        nmsmi.psminfo = psminfo;
        SendMessage(_hwndNewHandler, WM_NOTIFY, 0, (LPARAM)&nmsmi);
    }

    return S_OK;
}

// EXEX-VISTA(allison): Validated. Still need to figure out the magic numbers.
BOOL CDesktopHost::AddWin32Controls()
{
    RegisterDesktopControlClasses();

    // we create the controls with an arbitrary size, since we won't know how big we are until we pop up...

    // Note that we do NOT set WS_EX_CONTROLPARENT because we want the
    // dialog manager to think that our child controls are the interesting
    // objects, not the inner grandchildren.
    //
    // Setting the control ID equal to the internal index number is just
    // for the benefit of the test automation tools.

    for (int i = 0; i < ARRAYSIZE(_spm.panes); i++)
    {
        DWORD dwExStyle = 0;
        if (i == 1)
        {
            dwExStyle = WS_EX_CONTROLPARENT; // If the pane is the Desktop User Pane, give it WS_EX_NOPARENTNOTIFY
        }

        HWND hwndPane = CreateWindowEx(
            dwExStyle,
            _spm.panes[i].pszClassName,
            NULL,
            _spm.panes[i].dwStyle | 0x44000000,
            0,
            0,
            _spm.panes[i].size.cx,
            _spm.panes[i].size.cy,
            _hwnd,
            IntToPtr_(HMENU, i),
            NULL,
            &_spm.panes[i]);

        if (!hwndPane || !GetWindowLongPtr(_hwnd, GWLP_USERDATA))
            break;

        _spm.panes[i].hwnd = hwndPane;

        SMNSETSITE nmss;
        nmss.hdr.hwndFrom = _hwnd;
        nmss.hdr.code = 223;
        nmss.punkSite = static_cast<IMenuBand *>(this);
        SendMessage(hwndPane, WM_NOTIFY, 0, (LPARAM)&nmss);
    }

    return TRUE;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnPaint(HDC hdc, BOOL bBackground)
{
    if (IsCompositionActive())
    {
        RECT rc;
        GetClientRect(_hwnd, &rc);

        rc.left = _spm.panes[SMPANETYPE_OPENVIEWHOST].size.cx;
        SHFillRectClr(hdc, &rc, 0);
        DrawThemeBackground(_hTheme, hdc, SPP_PROGLIST, 0, &rc, NULL);
    }
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_ReadPaneSizeFromTheme(SMPANEDATA* psmpd)
{
    RECT rc;
    if (SUCCEEDED(GetThemeRect(psmpd->hTheme, psmpd->iPartId, 0, TMT_DEFAULTPANESIZE, &rc)))
    {
        // semi-hack to take care of the fact that if one the start panel parts is missing a property, 
        // themes will use the next level up (to the panel itself)
        if ((rc.bottom != _spm.sizPanel.cy) || (rc.right != _spm.sizPanel.cx))
        {
            psmpd->size.cx = RECTWIDTH(rc);
            psmpd->size.cy = RECTHEIGHT(rc);
        }
    }
}

void RemapSizeForHighDPI(SIZE* psiz)
{
    static int iLPX, iLPY;

    if (!iLPX || !iLPY)
    {
        HDC hdc = GetDC(NULL);
        iLPX = GetDeviceCaps(hdc, LOGPIXELSX);
        iLPY = GetDeviceCaps(hdc, LOGPIXELSY);
        ReleaseDC(NULL, hdc);
    }

    // 96 DPI is small fonts, so scale based on the multiple of that.
    psiz->cx = (psiz->cx * iLPX) / 96;
    psiz->cy = (psiz->cy * iLPY) / 96;
}



void CDesktopHost::LoadResourceInt(UINT ids, LONG* pl)
{
    TCHAR sz[64];
    if (LoadString(g_hinstCabinet, ids, sz, ARRAYSIZE(sz)))
    {
        int i = StrToInt(sz);
        if (i)
        {
            *pl = i;
        }
    }
}

// EXEX-VISTA(allison): Validated. Still needs a bit of cleanup.
void CDesktopHost::LoadPanelMetrics()
{
    _spm = g_spmDefault;

    LoadResourceInt(0x2040, &_spm.sizPanel.cy);
    LoadResourceInt(0x2041, &_spm.sizPanel.cx);
    LoadResourceInt(0x2042, &_spm.panes[SMPANETYPE_USER].size.cy);
    LoadResourceInt(0x2043, &_spm.panes[SMPANETYPE_OPENBOX].size.cy);
    LoadResourceInt(0x2044, &_spm.panes[SMPANETYPE_LOGOFF].size.cy);

    for (int i = 0; i < ARRAYSIZE(_spm.panes); i++)
    {
        _spm.panes[i].size.cx = MulDiv(g_spmDefault.panes[i].size.cx, _spm.sizPanel.cx, g_spmDefault.sizPanel.cx);
    }

    _spm.panes[SMPANETYPE_KNOWNFOLDER].size.cy = _spm.sizPanel.cy - _spm.panes[SMPANETYPE_USER].size.cy - _spm.panes[SMPANETYPE_LOGOFF].size.cy;
    _spm.panes[SMPANETYPE_OPENVIEWHOST].size.cy = _spm.sizPanel.cy - _spm.panes[SMPANETYPE_LOGOFF].size.cy;

    if (!_hTheme && SHGetCurColorRes() > 8)
        _hTheme = _GetStartMenuTheme();

    RECT rcT;
    if (_hTheme && SUCCEEDED(GetThemeRect(_hTheme, 0, 0, TMT_DEFAULTPANESIZE, &rcT)))
    {
        _spm.sizPanel.cx = RECTWIDTH(rcT);
        _spm.sizPanel.cy = RECTHEIGHT(rcT);

        for (int i = 0; i < ARRAYSIZE(_spm.panes); i++)
        {
            SMPANEDATA &paneData = _spm.panes[i];
            paneData.bPartDefined = IsThemePartDefined(_hTheme, paneData.iPartId, 0);
            if (paneData.bPartDefined)
            {
                paneData.hTheme = _hTheme;
                _ReadPaneSizeFromTheme(&paneData);
            }
            else
            {
                paneData.size.cx = 0;
                paneData.size.cy = 0;
            }
        }
    }

    ASSERT(_spm.sizPanel.cx == _spm.panes[SMPANETYPE_OPENVIEWHOST].size.cx + _spm.panes[SMPANETYPE_KNOWNFOLDER].size.cx); // 2719

    //CcshellDebugMsgW(
    //    0,
    //    "sizPanel.cy = %d, OpenViewHost =%d, openbox=%d, logoff=%d",
    //    _spm.sizPanel.cy,
    //    _spm.panes[2].size.cy,
    //    _spm.panes[2].size.cy,
    //    _spm.panes[4].size.cy);

    ASSERT(_spm.sizPanel.cy == _spm.panes[SMPANETYPE_USER].size.cy + _spm.panes[SMPANETYPE_KNOWNFOLDER].size.cy + _spm.panes[SMPANETYPE_LOGOFF].size.cy); // 2725

    RemapSizeForHighDPI(&_spm.sizPanel);
    for (int i = 0; i < ARRAYSIZE(_spm.panes); i++)
    {
        RemapSizeForHighDPI(&_spm.panes[i].size);
    }

    SetRect(&_rcDesired, 0, 0, _spm.sizPanel.cx, _spm.sizPanel.cy);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnCreate(HWND hwnd)
{
    _hwnd = hwnd;
    //TraceMsg(TF_DV2HOST, "Entering CDesktopHost::OnCreate");

    // Add the controls and background images
    AddWin32Controls();
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnDestroy()
{
    _hwnd = NULL;
    OnThemeChanged(0);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::OnSetFocus(HWND hwndLose)
{
    if (!_RestoreChildFocus())
    {
        _SetFocusToOpenBox();
    }
}

// EXEX-VISTA(allison): Validated.
LRESULT CDesktopHost::OnCommandInvoked(NMHDR *pnm)
{
    _UnlockStartPane();

	//SHPlaySound(L"MenuCommand", 2); // EXEX-VISTA(allison): TODO: Uncomment when implemented

    LRESULT lres = OnSelect(MPOS_EXECUTE);
    if (_DoesOpenBoxHaveFocus())
        _SetFocusToStartButton();
    return lres;
}

LRESULT CDesktopHost::OnFilterOptions(NMHDR* pnm)
{
    PSMNFILTEROPTIONS popt = (PSMNFILTEROPTIONS)pnm;

    if ((popt->smnop & SMNOP_LOGOFF) &&
        !_ShowStartMenuLogoff())
    {
        popt->smnop &= ~SMNOP_LOGOFF;
    }

    if ((popt->smnop & SMNOP_TURNOFF) &&
        !_ShowStartMenuShutdown())
    {
        popt->smnop &= ~SMNOP_TURNOFF;
    }

    if ((popt->smnop & SMNOP_DISCONNECT) &&
        !_ShowStartMenuDisconnect())
    {
        popt->smnop &= ~SMNOP_DISCONNECT;
    }

    if ((popt->smnop & SMNOP_EJECT) &&
        !_ShowStartMenuEject())
    {
        popt->smnop &= ~SMNOP_EJECT;
    }

    return 0;
}

LRESULT CDesktopHost::OnTrackShellMenu(NMHDR* pnm)
{
#ifdef DEAD_CODE
    PSMNTRACKSHELLMENU ptsm = CONTAINING_RECORD(pnm, SMNTRACKSHELLMENU, hdr);
    HRESULT hr;

    _hwndTracking = ptsm->hdr.hwndFrom;
    _itemTracking = ptsm->itemID;
    _hwndAltTracking = NULL;
    _itemAltTracking = 0;

    //
    //  Decide which direction we need to pop.
    //
    DWORD dwFlags;
    if (GetWindowLong(_hwnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL)
    {
        dwFlags = MPPF_LEFT;
    }
    else
    {
        dwFlags = MPPF_RIGHT;
    }

    // Don't _CleanupTrackShellMenu because that will undo some of the
    // work we've already done and make the client think that the popup
    // they requested got dismissed.

    //
    // ISSUE raymondc: actually this abandons the trackpopupmenu that
    // may already be in progress - its mouse UI gets messed up as a result.
    //
	if (_ppmTracking)
	{
		_ppmTracking->Release();
		_ppmTracking = NULL;
	}

    if (_hwndTracking == _spm.panes[SMPANETYPE_OPENBOX].hwnd)
    {
        if (_ppmPrograms && _ppmPrograms->IsSame(ptsm->psm))
        {
            // It's already in our cache, woo-hoo!
            hr = S_OK;
        }
        else
        {
            ATOMICRELEASE(_ppmPrograms);
            _SubclassTrackShellMenu(ptsm->psm);
            hr = CPopupMenu_CreateInstance(ptsm->psm, GetUnknown(), _hwnd, &_ppmPrograms);
        }

        if (SUCCEEDED(hr))
        {
            _ppmTracking = _ppmPrograms;
            _ppmTracking->AddRef();
        }
    }
    else
    {
        _SubclassTrackShellMenu(ptsm->psm);
        hr = CPopupMenu_CreateInstance(ptsm->psm, GetUnknown(), _hwnd, &_ppmTracking);
    }

    if (SUCCEEDED(hr))
    {
        hr = _ppmTracking->Popup(&ptsm->rcExclude, ptsm->dwFlags | dwFlags);
    }

    if (FAILED(hr))
    {
        // In addition to freeing any partially-allocated objects,
        // this also sends a SMN_SHELLMENUDISMISSED so the client
        // knows to remove the highlight from the item being cascaded
        _CleanupTrackShellMenu();
    }

    return 0;
#else
    HRESULT hr;
    PSMNTRACKSHELLMENU ptsm = CONTAINING_RECORD(pnm, SMNTRACKSHELLMENU, hdr);

    _hwndTracking = ptsm->hdr.hwndFrom;
    _hwndAltTracking = 0;
    _itemAltTracking = 0;
    _itemTracking = ptsm->itemID;

    DWORD dwFlags = ptsm->dwFlags & 0xE0000000;
    if (!dwFlags)
        dwFlags = ~(GetWindowLongPtr(_hwnd, GWL_EXSTYLE) << 7) & 0x20000000 | 0x40000000;

    _DestroyClipBalloon();
    if (_ppmTracking)
        _ppmTracking->OnSelect(2);
    IUnknown_SafeReleaseAndNullPtr(&_ppmTracking);

    int v11 = 0;
    if (_hwndTracking == _spm.panes[2].hwnd)
    {
        v11 = 1;
        if (_ppmPrograms && _ppmPrograms->IsSame(ptsm->psm))
        {
            hr = 0;
        }
        else
        {
            IUnknown_SafeReleaseAndNullPtr(&_ppmPrograms);
            _SubclassTrackShellMenu(ptsm->psm);
            hr = CPopupMenu_CreateInstance(ptsm->psm, GetUnknown(), this->_hwnd, &this->_ppmPrograms);
            if (hr < 0)
            {
            LABEL_17:
                _CleanupTrackShellMenu();
                return 0;
            }
        }
        _ppmTracking = _ppmPrograms;
        _ppmTracking->AddRef();
    }

    else
    {
        _SubclassTrackShellMenu(ptsm->psm);
        hr = CPopupMenu_CreateInstance(ptsm->psm, GetUnknown(), this->_hwnd, &this->_ppmTracking);
    }

    if (hr < 0)
        goto LABEL_17;
    hr = _ppmTracking->Popup(&ptsm->rcExclude, dwFlags | ptsm->dwFlags);

    //if (v11)
    //    SHTracePerf(&ShellTraceId_Explorer_StartPane_AllPrograms_Show_Stop);
    //else
    //    SHTracePerf(&ShellTraceId_Explorer_StartPane_Cascade_Show_Stop);

    if (hr < 0)
        goto LABEL_17;
    return 0;
#endif
}

HRESULT CDesktopHost::_MenuMouseFilter(LPSMDATA psmd, BOOL fRemove, LPMSG pmsg)
{
#ifdef DEAD_CODE
    HRESULT hr = S_FALSE;
    SMNDIALOGMESSAGE nmdm;

    enum {
        WHERE_IGNORE,               // ignore this message
        WHERE_OUTSIDE,              // outside the Start Menu entirely
        WHERE_DEADSPOT,             // a dead spot on the Start Menu
        WHERE_ONSELF,               // over the item that initiated the popup
        WHERE_ONOTHER,              // over some other item in the Start Menu
    } uiWhere;

    //
    //  Figure out where the mouse is.
    //
    //  Note: ChildWindowFromPointEx searches only immediate
    //  children; it does not search grandchildren. Fortunately, that's
    //  exactly what we want...
    //

    HWND hwndTarget = NULL;

    if (fRemove)
    {
        if (psmd->punk)
        {
            // Inside a menuband - mouse has left our window
            uiWhere = WHERE_OUTSIDE;
        }
        else
        {
            POINT pt = { GET_X_LPARAM(pmsg->lParam), GET_Y_LPARAM(pmsg->lParam) };
            ScreenToClient(_hwnd, &pt);

            hwndTarget = ChildWindowFromPointEx(_hwnd, pt, CWP_SKIPINVISIBLE);
            if (hwndTarget == _hwnd)
            {
                uiWhere = WHERE_DEADSPOT;
            }
            else if (hwndTarget)
            {
                LRESULT lres;
                nmdm.pt = pt;
                HWND hwndChild = ::GetWindow(hwndTarget, GW_CHILD);
                MapWindowPoints(_hwnd, hwndChild, &nmdm.pt, 1);
                lres = _FindChildItem(hwndTarget, &nmdm, SMNDM_HITTEST | SMNDM_SELECT);
                if (lres)
                {
                    // Mouse is over something; is it over the current item?

                    if (nmdm.itemID == _itemTracking &&
                        hwndTarget == _hwndTracking)
                    {
                        uiWhere = WHERE_ONSELF;
                    }
                    else
                    {
                        uiWhere = WHERE_ONOTHER;
                    }
                }
                else
                {
                    uiWhere = WHERE_DEADSPOT;
                }
            }
            else
            {
                // ChildWindowFromPoint failed - user has left the Start Menu
                uiWhere = WHERE_OUTSIDE;
            }
        }
    }
    else
    {
        // Ignore PM_NOREMOVE messages; we'll pay attention to them when
        // they are PM_REMOVE'd.
        uiWhere = WHERE_IGNORE;
    }

    //
    //  Now do appropriate stuff depending on where the mouse is.
    //
    switch (uiWhere)
    {
        case WHERE_IGNORE:
            break;

        case WHERE_OUTSIDE:
            //
            // If you've left the menu entirely, then we return the menu to
            // its original state, which is to say, as if you are hovering
            // over the item that caused the popup to open in the first place.
            // as being in a dead zone.
            //
            // FALL THROUGH
            goto L_WHERE_ONSELF_HOVER;

        case WHERE_DEADSPOT:
            // To avoid annoying flicker as the user wanders over dead spots,
            // we ignore mouse motion over them (but dismiss if they click
            // in a dead spot).
            if (pmsg->message == WM_LBUTTONDOWN ||
                pmsg->message == WM_RBUTTONDOWN)
            {
                // Must explicitly dismiss; if we let it fall through to the
                // default handler, then it will dismiss for us, causing the
                // entire Start Menu to go away instead of just the tracking
                // part.
                _DismissTrackShellMenu();
                hr = S_OK;
            }
            break;

        case WHERE_ONSELF:
            if (pmsg->message == WM_LBUTTONDOWN ||
                pmsg->message == WM_RBUTTONDOWN)
            {
                _DismissTrackShellMenu();
                hr = S_OK;
            }
            else
            {
            L_WHERE_ONSELF_HOVER:
                _hwndAltTracking = NULL;
                _itemAltTracking = 0;
                nmdm.itemID = _itemTracking;
                _FindChildItem(_hwndTracking, &nmdm, SMNDM_FINDITEMID | SMNDM_SELECT);
                KillTimer(_hwnd, IDT_MENUCHANGESEL);
            }
            break;

        case WHERE_ONOTHER:
            if (pmsg->message == WM_LBUTTONDOWN ||
                pmsg->message == WM_RBUTTONDOWN)
            {
                _DismissTrackShellMenu();
                hr = S_OK;
            }
            else if (hwndTarget == _hwndAltTracking && nmdm.itemID == _itemAltTracking)
            {
                // Don't restart the timer if the user wiggles the mouse
                // within a single item
            }
            else
            {
                _hwndAltTracking = hwndTarget;
                _itemAltTracking = nmdm.itemID;

                DWORD dwHoverTime;
                if (!SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0, &dwHoverTime, 0))
                {
                    dwHoverTime = 0;
                }
                SetTimer(_hwnd, IDT_MENUCHANGESEL, dwHoverTime, 0);
            }
            break;
    }

    return hr;
#else
    POINT pt;

    SMNDIALOGMESSAGE nmdm; // [esp+34h] [ebp-34h] BYREF
    HRESULT hr = 1;
    
    enum {
        WHERE_IGNORE,               // ignore this message
        WHERE_OUTSIDE,              // outside the Start Menu entirely
        WHERE_DEADSPOT,             // a dead spot on the Start Menu
        WHERE_ONSELF,               // over the item that initiated the popup
        WHERE_ONOTHER,              // over some other item in the Start Menu
    } uiWhere;

    HWND hwndTarget = 0;
    
    if (fRemove)
    {
        if (psmd->punk)
        {
            uiWhere = WHERE_OUTSIDE;
        }
        else
        {
            pt.y = GET_Y_LPARAM(pmsg->lParam);
            pt.x = GET_X_LPARAM(pmsg->lParam);
            ScreenToClient(_hwnd, &pt);
            
            hwndTarget = ChildWindowFromPointEx(_hwnd, pt, 1);
            
            HWND v8 = _spm.panes[1].hwnd;
            if (hwndTarget == v8)
            {
                hwndTarget = ChildWindowFromPointEx(v8, pt, 1);
            }
            
            HWND v9 = _hwnd;
            if (hwndTarget == v9 || hwndTarget == _spm.panes[1].hwnd)
            {
                uiWhere = WHERE_DEADSPOT;
            }
            else if (hwndTarget)
            {
                MapWindowPoints(v9, hwndTarget, &pt, 1u);
                
                nmdm.pt = pt;
                HWND fRemovea = ChildWindowFromPointEx(hwndTarget, pt, 1u);
                MapWindowPoints(hwndTarget, fRemovea, &nmdm.pt, 1u);
                nmdm.hwnd = fRemovea;
                if (_FindChildItem(hwndTarget, &nmdm, 0x107u))
                {
                    if (nmdm.itemID == _itemTracking && hwndTarget == _hwndTracking)
                    {
                        uiWhere = WHERE_ONSELF;
                    }
                    else
                    {
                        uiWhere = WHERE_ONOTHER;
                    }
                }
                else
                {
                    //uiWhere = ~(BYTE)(nmdm.flags >> 12) & 2;
                    uiWhere = (nmdm.flags & 0x2000) != 0 ? WHERE_IGNORE : WHERE_DEADSPOT;
                }
            }
            else
            {
                uiWhere = WHERE_OUTSIDE;
            }
        }
    }
    else
    {
        uiWhere = WHERE_IGNORE;
    }
    
    if (uiWhere == 1)
    {
        goto L_WHERE_ONSELF_HOVER;
    }

    if (uiWhere == 2)
    {
        if (pmsg->message == 513 || pmsg->message == 516)
        {
            _DismissTrackShellMenu();
            return 0;
        }
        return hr;
    }

    if (uiWhere == 3)
    {
        if (pmsg->message == 516)
        {
            _DismissTrackShellMenu();
            hr = 0;
        }
        if (pmsg->message == 513)
        {
            return 0;
        }
    L_WHERE_ONSELF_HOVER:
        _hwndAltTracking = 0;
        _itemAltTracking = 0;
        nmdm.itemID = _itemTracking;
        _FindChildItem(_hwndTracking, &nmdm, 0x109u);
        KillTimer(_hwnd, 1u);
        return hr;
    }

    if (uiWhere == 4)
    {
        if (pmsg->message == 513 || pmsg->message == 516)
        {
            _DismissTrackShellMenu();
            if (hwndTarget == _hwndTracking && pmsg->message == 513)
            {
                _FindChildItem(hwndTarget, &nmdm, 0x406u);
            }
            return 0;
        }

        if (hwndTarget != _hwndAltTracking || nmdm.itemID != _itemAltTracking)
        {
            _itemAltTracking = nmdm.itemID;
            _hwndAltTracking = hwndTarget;

            DWORD dwHoverTime;
            if (!SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0, &dwHoverTime, 0))
            {
                dwHoverTime = 0;
            }
            SetTimer(_hwnd, 1u, dwHoverTime, 0);
        }
        return hr;
    }
    
    return hr;

#endif
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_OnMenuChangeSel()
{
    KillTimer(_hwnd, IDT_MENUCHANGESEL);
    _DismissTrackShellMenu();
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_SaveChildFocus()
{
    if (!_hwndChildFocus)
    {
        HWND hwndFocus = GetFocus();
        if (hwndFocus && IsChild(_hwnd, hwndFocus))
        {
            _hwndChildFocus = hwndFocus;
        }
    }
}

// Returns non-NULL if focus was successfully restored
// EXEX-VISTA(allison): Validated.
HWND CDesktopHost::_RestoreChildFocus()
{
    HWND hwndRet = NULL;
    if (IsWindow(_hwndChildFocus))
    {
        HWND hwndT = _hwndChildFocus;
        _hwndChildFocus = NULL;
        hwndRet = SetFocus(hwndT);
    }
    return hwndRet;
}


// EXEX-VISTA(allison): Validated.
void CDesktopHost::_DestroyClipBalloon()
{
    if (_hwndClipBalloon)
    {
        DestroyWindow(_hwndClipBalloon);
        _hwndClipBalloon = NULL;
    }
}

// EXEX-VISTA(allison): Validated.
IStartButton* CDesktopHost::_GetIStartButton()
{
    IStartButton* pstb = NULL;
	IUnknown_QueryService(_punkSite, __uuidof(IStartButton), IID_PPV_ARGS(&pstb));
    return pstb;
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_LockStartPane()
{
    IStartButton *pstb = _GetIStartButton();
    if (pstb)
    {
        pstb->LockStartPane();
        pstb->Release();
    }
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_UnlockStartPane()
{
    IStartButton *pstb = _GetIStartButton();
    if (pstb)
    {
        pstb->UnlockStartPane();
        pstb->Release();
    }
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_SetFocusToStartButton()
{
    IStartButton *pstb = _GetIStartButton();
    if (pstb)
    {
        pstb->SetFocusToStartButton();
        pstb->Release();
	}
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
HTHEME CDesktopHost::_GetStartMenuTheme()
{
    field_1AC = 0;

    IStartButton *pstb = _GetIStartButton();
    if (pstb)
    {
        DWORD dwPopupPosition = 0;
        pstb->GetPopupPosition(&dwPopupPosition);
        pstb->Release();
        field_1AC |= dwPopupPosition;
    }

    LPCWSTR pszTheme;
    if (field_1AC == 0x80000000)
    {
        pszTheme = IsCompositionActive() ? L"StartPanelCompositedBottom::StartPanel" : L"StartPanelBottom::StartPanel";
    }
    else
    {
        pszTheme = IsCompositionActive() ? L"StartPanel" : L"StartPanel";
    }
    return OpenThemeData(_hwnd, pszTheme);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_RegisterForGlass(BOOL a2, HRGN hrgn)
{
    if (a2)
        a2 = IsCompositionActive() && _hTheme;

    DWM_BLURBEHIND bb;
    bb.fTransitionOnMaximized = FALSE;
    bb.fEnable = a2;
    bb.hRgnBlur = hrgn;
    bb.dwFlags = DWM_BB_ENABLE | DWM_BB_BLURREGION;
    DwmEnableBlurBehindWindow(_hwnd, &bb);
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
void CDesktopHost::_OnDismiss(BOOL bDestroy)
{
    // Break the recursion loop:  Call IMenuPopup::OnSelect only if the
    // window was previously visible.
    if (_fOpen || IsWindowVisible(_hwnd))
    {
        ShowWindow(_hwnd, SW_HIDE);
        _UnlockStartPane();

        if (_fOpen)
        {
            _fOpen = FALSE;

            // EXEX-VISTA(isabella): Figure out what is SMN_FIRST + 22.
            NMHDR nm2 = { _hwnd, 0, SMN_FIRST + 22 };
            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm2, SPM_SEND | SPM_ONELEVEL);

            if (_ppmTracking)
            {
                _ppmTracking->OnSelect(MPOS_FULLCANCEL);
            }

            OnSelect(MPOS_FULLCANCEL);

            NMHDR nm = { _hwnd, 0, SMN_DISMISS };
            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);

            _DestroyClipBalloon();
            _RegisterForGlass(FALSE, NULL);

            IStartButton *pstb = _GetIStartButton();
            if (pstb)
            {
                pstb->SetStartPaneActive(FALSE);
                pstb->OnStartMenuDismissed();
                pstb->Release();
            }

            // Don't try to preserve child focus across popups
            _hwndChildFocus = NULL;

            NotifyWinEvent(EVENT_SYSTEM_MENUPOPUPEND, _hwnd, OBJID_CLIENT, CHILDID_SELF);
           
            if (field_D0)
            {   
                field_D0 = 0;
                PostMessage(v_hwndTray, SBM_REBUILDMENU, 0, 0);
            }
        }
    }
    if (bDestroy)
    {
        v_hwndStartPane = NULL;
        ASSERT(GetWindowThreadProcessId(_hwnd, NULL) == GetCurrentThreadId());
        DestroyWindow(_hwnd);
    }
}

// EXEX-VISTA(allison): Validated.
HRESULT CDesktopHost::Build()
{
    HRESULT hr = S_OK;
    if (_hwnd == NULL)
    {
        _hwnd = _Create();

        if (_hwnd)
        {
            // Tell all our child windows it's time to reinitialize
            NMHDR nm = { _hwnd, 0, SMN_INITIALUPDATE };
            SHPropagateMessage(_hwnd, WM_NOTIFY, 0, (LPARAM)&nm, SPM_SEND | SPM_ONELEVEL);
        }
    }

    if (_hwnd == NULL)
    {
        hr = E_OUTOFMEMORY;
    }

    return hr;
}


//*****************************************************************
//
//  CDeskHostShellMenuCallback
//
//  Create a wrapper IShellMenuCallback that picks off mouse
//  messages.
//
class CDeskHostShellMenuCallback
    : public CUnknown
    , public IShellMenuCallback
    , public IServiceProvider
    , public CObjectWithSite
{
    friend class CDesktopHost;

public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    // *** IObjectWithSite ***
    STDMETHODIMP SetSite(IUnknown* punkSite);

    // *** IServiceProvider ***
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject);

private:
    CDeskHostShellMenuCallback(CDesktopHost* pdh)
    {
        _pdh = pdh; _pdh->AddRef();
    }

    ~CDeskHostShellMenuCallback()
    {
        IUnknown_SafeReleaseAndNullPtr(&_pdh);
        IUnknown_SetSite(_psmcPrev, NULL);
        IUnknown_SafeReleaseAndNullPtr(&_psmcPrev);
    }

    IShellMenuCallback* _psmcPrev;
    CDesktopHost* _pdh;
};

HRESULT CDeskHostShellMenuCallback::QueryInterface(REFIID riid, void** ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CDeskHostShellMenuCallback, IShellMenuCallback),
        QITABENT(CDeskHostShellMenuCallback, IObjectWithSite),
        QITABENT(CDeskHostShellMenuCallback, IServiceProvider),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

BOOL FeatureEnabledDeskHost(LPCTSTR pszFeature)
{
    return _SHRegGetBoolValueFromHKCUHKLM(REGSTR_EXPLORER_ADVANCED, pszFeature,
        FALSE); // Disable this cool feature.
}


HRESULT CDeskHostShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_MOUSEFILTER:
        if (_pdh)
            return _pdh->_MenuMouseFilter(psmd, (BOOL)wParam, (MSG*)lParam);

    case SMC_GETSFINFOTIP:
        if (!FeatureEnabledDeskHost(TEXT("ShowInfoTip")))
            return E_FAIL;  // E_FAIL means don't show. S_FALSE means show default
        break;

    }

    if (_psmcPrev)
        return _psmcPrev->CallbackSM(psmd, uMsg, wParam, lParam);

    return S_FALSE;
}

HRESULT CDeskHostShellMenuCallback::SetSite(IUnknown* punkSite)
{
    CObjectWithSite::SetSite(punkSite);
    // Each time our site changes, reassert ourselves as the site of
    // the inner object so he can try a new QueryService.
    IUnknown_SetSite(_psmcPrev, GetUnknown());

    // If the game is over, break our backreference
    if (!punkSite)
    {
        IUnknown_SafeReleaseAndNullPtr(&_pdh);
    }

    return S_OK;
}

HRESULT CDeskHostShellMenuCallback::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    if (IsEqualGUID(guidService, SID_SMenuPopup))
    {
        return _pdh->QueryInterface(riid, ppvObject);
    }
    return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

// EXEX-VISTA(allison): Validated.
void CDesktopHost::_SubclassTrackShellMenu(IShellMenu* psm)
{
    CDeskHostShellMenuCallback* psmc = new CDeskHostShellMenuCallback(this);
    if (psmc)
    {
        UINT uId, uIdAncestor;
        DWORD dwFlags;
        if (SUCCEEDED(psm->GetMenuInfo(&psmc->_psmcPrev, &uId, &uIdAncestor, &dwFlags)))
        {
            psm->Initialize(psmc, uId, uIdAncestor, dwFlags);
        }
        psmc->Release();
    }
}

STDAPI DesktopV2_Build(void* pvStartPane)
{
    HRESULT hr = E_POINTER;
    if (pvStartPane)
    {
        hr = reinterpret_cast<CDesktopHost*>(pvStartPane)->Build();
    }
    return hr;
}


STDAPI DesktopV2_Create(
    IMenuPopup **ppmp, IMenuBand **ppmb, void **ppvStartPane, IUnknown **ppunk, HWND hwnd)
{
    *ppmp = NULL;
    *ppmb = NULL;
    *ppunk = NULL;

    HRESULT hr;
    CDesktopHost *pdh = new CDesktopHost;
    if (pdh)
    {
        *ppvStartPane = pdh;
        hr = pdh->Initialize(hwnd);
        if (SUCCEEDED(hr))
        {
            hr = pdh->QueryInterface(IID_PPV_ARGS(ppmp));
            if (SUCCEEDED(hr))
            {
                hr = pdh->QueryInterface(IID_PPV_ARGS(ppmb));
                if (SUCCEEDED(hr))
                {
                    hr = pdh->QueryInterface(IID_PPV_ARGS(ppunk));
                }
            }
        }
        pdh->GetUnknown()->Release();
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }

    if (FAILED(hr))
    {
        ATOMICRELEASE(*ppmp);
        ATOMICRELEASE(*ppmb);
        ppvStartPane = NULL;
    }

    return hr;
}

DWORD Mirror_MirrorDC(HDC hdc)
{
	return Mirror_SetLayout(hdc, LAYOUT_RTL);
}

#define SET_DC_RTL_MIRRORED(hdc)         Mirror_MirrorDC(hdc)
HBITMAP CreateMirroredBitmap(HBITMAP hbmOrig)
{
    HDC     hdc, hdcMem1, hdcMem2;
    HBITMAP hbm = NULL, hOld_bm1, hOld_bm2;
    BITMAP  bm;
    int     IncOne = 0;

    if (!hbmOrig)
        return NULL;

    if (!GetObject(hbmOrig, sizeof(BITMAP), &bm))
        return NULL;

    // Grab the screen DC
    hdc = GetDC(NULL);

    if (hdc)
    {
        hdcMem1 = CreateCompatibleDC(hdc);

        if (!hdcMem1)
        {
            ReleaseDC(NULL, hdc);
            return NULL;
        }

        hdcMem2 = CreateCompatibleDC(hdc);
        if (!hdcMem2)
        {
            DeleteDC(hdcMem1);
            ReleaseDC(NULL, hdc);
            return NULL;
        }

        hbm = CreateCompatibleBitmap(hdc, bm.bmWidth, bm.bmHeight);

        if (!hbm)
        {
            ReleaseDC(NULL, hdc);
            DeleteDC(hdcMem1);
            DeleteDC(hdcMem2);
            return NULL;
        }

        //
        // Flip the bitmap
        //
        hOld_bm1 = (HBITMAP)SelectObject(hdcMem1, hbmOrig);
        hOld_bm2 = (HBITMAP)SelectObject(hdcMem2, hbm);

        SET_DC_RTL_MIRRORED(hdcMem2);
        BitBlt(hdcMem2, IncOne, 0, bm.bmWidth, bm.bmHeight, hdcMem1, 0, 0, SRCCOPY);

        SelectObject(hdcMem1, hOld_bm1);
        SelectObject(hdcMem1, hOld_bm2);

        DeleteDC(hdcMem1);
        DeleteDC(hdcMem2);

        ReleaseDC(NULL, hdc);
    }

    return hbm;
}