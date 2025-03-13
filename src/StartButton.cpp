#include "StartButton.h"

#include <shlwapi.h>
#include <windowsx.h>

#include "SHFusion.h"
#include "tray.h"
#include "Util.h"
#include "rcids.h"
#include <DeskHost.h>
#include <vssym32.h>

CStartButton::CStartButton(IStartButtonSite *pStartButtonSite)
    : _nSettingsChangeType(true)
    , _pStartButtonSite(pStartButtonSite)
{

}

HRESULT CStartButton::QueryInterface(REFIID riid, void** ppvObject)
{
    static const QITAB qit[] =
    {
        QITABENT(CStartButton, IStartButton),
        QITABENT(CStartButton, IServiceProvider),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObject);
}

ULONG CStartButton::AddRef()
{
    return 1; // Copilot generated body
}

ULONG CStartButton::Release()
{
    return 1; // Copilot generated body
}

HRESULT CStartButton::SetFocusToStartButton()  // taken from ep_taskbar 7-stuff
{
    SetFocus(_hwndStartBtn);
    return S_OK;
}

HRESULT CStartButton::OnContextMenu(HWND hWnd, LPARAM lParam)
{
    _nIsOnContextMenu = TRUE;
    _pStartButtonSite->HandleFullScreenApp(hWnd);
    SetForegroundWindow(hWnd);

    LPITEMIDLIST pidl = SHCloneSpecialIDList(hWnd, CSIDL_STARTMENU, TRUE);
    IShellFolder* shlFldr;
    LPCITEMIDLIST ppidlLast;
    if (SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARG(IShellFolder, &shlFldr), &ppidlLast)))
    {
        HMENU hMenu = CreatePopupMenu();
        IContextMenu* iContext;
        if (SUCCEEDED(shlFldr->GetUIObjectOf(hWnd, CSIDL_INTERNET, &ppidlLast, IID_IContextMenu, NULL, (void**)&iContext)))
        {
            if (SUCCEEDED(iContext->QueryContextMenu(hMenu, 0, 2, 32751, CMF_VERBSONLY)))
            {
                WCHAR Buffer[260];
                LoadStringW(GetModuleHandle(NULL), 720u, Buffer, 260);
                AppendMenuW(hMenu, 0, 32755u, Buffer);
                if (!SHRestricted(REST_NOCOMMONGROUPS))
                {
                    if (SHGetFolderPathW(0, CSIDL_COMMON_STARTMENU, 0, 0, Buffer) != S_OK)
                    {
                        // some LUA crap
                    }
                }
                int bResult;
                if (lParam == -1)
                {
                    bResult = TrackMenu(hMenu);
                }
                else
                {
                    _pStartButtonSite->EnableTooltips(FALSE);
                    UINT uFlag = TPM_RETURNCMD | TPM_RIGHTBUTTON;
                    if (IsBiDiLocalizedSystem())
                    {
                        uFlag |= TPM_LAYOUTRTL;
                    }
                    bResult = TrackPopupMenu(hMenu, uFlag,  GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) , 0, hWnd, 0);
                    _pStartButtonSite->EnableTooltips(TRUE);
                }
                if (bResult)
                {
                    LPITEMIDLIST ppidl;
                    switch (bResult)
                    {
                        case 32752:
                            _ExploreCommonStartMenu(FALSE);
                            break;
                        case 32753:
                            _ExploreCommonStartMenu(TRUE);
                            break;
                        case 32755:
                            Tray_DoProperties(TPF_STARTMENUPAGE);
                            break;
                        default:
                        {
                            WCHAR pszStr1[260];
                            ContextMenu_GetCommandStringVerb(iContext, bResult - 2, pszStr1, 260);
                            if (StrCmpICW(pszStr1, L"find"))
                            {
                                CHAR pszStr[260];
                                SHUnicodeToAnsi(pszStr1, pszStr, 260);

                                CMINVOKECOMMANDINFO cmInvokeInfo = {};
                                cmInvokeInfo.hwnd = hWnd;
                                cmInvokeInfo.lpVerb = pszStr;
                                cmInvokeInfo.nShow = SW_NORMAL;

                                CHAR pszPath[260];
                                SHGetPathFromIDListA(pidl, pszPath);
                                cmInvokeInfo.fMask |= SEE_MASK_UNICODE;
                                cmInvokeInfo.lpDirectory = pszPath;
                                iContext->InvokeCommand(&cmInvokeInfo);
                            }
                            else if (SUCCEEDED(SHGetKnownFolderIDList(FOLDERID_SearchHome, NULL, 0, &ppidl)))
                            {
                                SHELLEXECUTEINFO execInfo = {};
                                execInfo.lpVerb = L"open";
                                execInfo.fMask = SEE_MASK_IDLIST;
                                execInfo.nShow = SW_NORMAL;
                                execInfo.lpIDList = ppidl;
                                ShellExecuteEx(&execInfo);
                                ILFree(ppidl);
                            }
                            break;
                        }
                    }
                }
                iContext->Release();
            }
            DestroyMenu(hMenu);
        }
        shlFldr->Release();
    }
    ILFree(pidl);
    _nIsOnContextMenu = FALSE;
    return S_OK;
}

HRESULT CStartButton::CreateStartButtonBalloon(UINT a2, UINT uID)
{
    if (!_hwndStartBalloon)
    {
        _hwndStartBalloon = SHFusionCreateWindow(TOOLTIPS_CLASS, NULL, 
            WS_POPUP | TTS_BALLOON | TTS_NOPREFIX | TTS_ALWAYSTIP, 
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            _hwndStartBtn, NULL, GetModuleHandle(NULL), NULL);

        if (_hwndStartBalloon)
        {
            SendMessage(_hwndStartBalloon, 0x2007u, 6, 0);
            SendMessage(_hwndStartBalloon, TTM_SETMAXTIPWIDTH, 0, 300);
            SendMessage(_hwndStartBalloon, 0x200Bu, 0, (LPARAM)L"TaskBar");
            SetPropW(_hwndStartBalloon, L"StartMenuBalloonTip", (HANDLE)3);
        }
    }
    WCHAR Buffer[260];
    if (LoadString(GetModuleHandle(NULL), uID, Buffer, 260))
    {
        TOOLINFO toolInfo = { sizeof(toolInfo) };
        toolInfo.uFlags = TTF_TRANSPARENT | TTF_TRACK | TTF_IDISHWND; // 0x0100|0x0020|0x0001 -> 0x121 -> 289
        toolInfo.hwnd = _hwndStartBtn;
        toolInfo.uId = (UINT)_hwndStartBtn;  //TTF_IDISHWND
        
        SendMessage(_hwndStartBalloon, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);
        SendMessage(_hwndStartBalloon, TTM_TRACKACTIVATE, 0, 0);

        toolInfo.lpszText = Buffer;
        SendMessage(_hwndStartBalloon, TTM_UPDATETIPTEXT, 0, (LPARAM)&toolInfo);

        if (LoadString(GetModuleHandle(NULL), a2, Buffer, 260))
        {
            SendMessage(_hwndStartBalloon, TTM_SETTITLE, TTI_INFO, (LPARAM)&Buffer);
        }

        RECT Rect;
        GetWindowRect(_hwndStartBtn, &Rect);
        WORD xCoordinate = (WORD)((Rect.left + Rect.right) / 2);
        WORD yCoordinate = LOWORD(Rect.top);

        SendMessage(_hwndStartBalloon, TTM_TRACKPOSITION, 0, MAKELONG(xCoordinate, yCoordinate));
        SendMessage(_hwndStartBalloon, TTM_TRACKACTIVATE, TRUE, (LPARAM)&toolInfo);

        // show tooltip for 10 seconds
        SetTimer(_hwndStartBtn, 1u, 10000, 0);
    }
    return S_OK;
}

// dwRefData-> CTray+988 -> CStartButton
HWND CStartButton::CreateStartButton(HWND hWnd)
{
    DWORD dwStyleEx =  GetWindowLongW(hWnd, GWL_EXSTYLE);
    HWND hWndStartBtn = (HWND)SHFusionCreateWindowEx(
        dwStyleEx & (WS_EX_LAYOUTRTL | WS_EX_RTLREADING | WS_EX_RIGHT) | (WS_EX_LAYERED | WS_EX_TOOLWINDOW),
        L"Button",
        nullptr,
        WS_POPUP | WS_CLIPSIBLINGS | BS_VCENTER,
        0, 0, 1, 1,
        hWnd,
        nullptr,
        GetModuleHandle(NULL),
        nullptr);
    _hwndStartBtn = hWndStartBtn;

    if (hWndStartBtn)
    {
        SetPropW(hWndStartBtn, L"StartButtonTag", (HANDLE)0x130);
        SendMessageW(_hwndStartBtn, CCM_DPISCALE, TRUE, 0);
        SetWindowSubclass(
            _hwndStartBtn,
            (SUBCLASSPROC)CStartButton::s_StartButtonSubclassProc,
            0,
            (DWORD_PTR)this);
        LoadStringW(GetModuleHandle(NULL), /* TODO: Document. */ 0x253u, (LPWSTR)_pszWindowName, 50);
        SetWindowTextW(_hwndStartBtn, (LPCWSTR)_pszWindowName);
    }

    return _hwndStartBtn;
}

HRESULT CStartButton::SetStartPaneActive(BOOL bActive)
{
    if (bActive)
    {
        _nStartPaneActiveState = 1;
    }
    else if (_nStartPaneActiveState != 2)
    {
        _nStartPaneActiveState = 0;
        UnlockStartPane();
    }
    return S_OK;
}

HRESULT CStartButton::OnStartMenuDismissed()  // taken from ep_taskbar 7-stuff
{
    _pStartButtonSite->OnStartMenuDismissed();
    return S_OK;
}

HRESULT CStartButton::UnlockStartPane()  // taken from ep_taskbar 7-stuff
{
    if (_uLockCode)
    {
        _uLockCode = FALSE;
        LockSetForegroundWindow(LSFW_UNLOCK);
    }
    return S_OK;
}

HRESULT CStartButton::LockStartPane()  // taken from ep_taskbar 7-stuff
{
    if (!_uLockCode)
    {
        _uLockCode = TRUE;
        LockSetForegroundWindow(LSFW_LOCK);
    }
    return S_OK;
}

HRESULT CStartButton::GetPopupPosition(DWORD* out)  // taken from ep_taskbar 7-stuff
{
    if (!_pStartButtonSite)
        return E_FAIL;

    UINT stuckPlace = _pStartButtonSite->GetStartMenuStuckPlace();
    switch (stuckPlace)
    {
        case 0: *out = MPPF_LEFT; break;
        case 1: *out = MPPF_TOP; break;
        case 2: *out = MPPF_RIGHT; break;
        default: *out = MPPF_BOTTOM; break;
    }

    return S_OK;
}

HRESULT CStartButton::GetWindow(HWND* out)  // taken from ep_taskbar 7-stuff
{
    *out = _hwndStartBtn;
    return S_OK;
}

HRESULT CStartButton::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)  // taken from ep_taskbar 7-stuff
{
    *ppvObject = nullptr;
    HRESULT hr = E_FAIL;

    if (IsEqualGUID(guidService, __uuidof(IStartButton)))
    {
        hr = QueryInterface(riid, ppvObject);
    }

    return hr;
}

void CStartButton::BuildStartMenu() // from xp
{
    CloseStartMenu();
    _pStartButtonSite->PurgeRebuildRequests();
    DestroyStartMenu();
    if (Tray_StartPanelEnabled())
    {
        LPVOID pDeskHost;
        DesktopV2_Create(&_pNewStartMenu, &_pNewStartMenuBand, &pDeskHost);
        DesktopV2_Build(pDeskHost);
    }
    else
    {
        HRESULT hr = StartMenuHost_Create(&_pOldStartMenu, &_pOldStartMenuBand);
        if (SUCCEEDED(hr))
        {
            IBanneredBar* pbb;

            hr = _pOldStartMenu->QueryInterface(IID_PPV_ARG(IBanneredBar, &pbb));
            if (SUCCEEDED(hr))
            {
                pbb->SetBitmap(_hbmpStartBkg);
                if (_pStartButtonSite->ShouldUseSmallIcons())
                    pbb->SetIconSize(BMICON_SMALL);
                else
                    pbb->SetIconSize(BMICON_LARGE);

                pbb->Release();
            }
        }
    }
}

void CStartButton::CloseStartMenu()  // taken from ep_taskbar 7-stuff @MOD
{
    if (_pOldStartMenu)
    {
        _pOldStartMenu->OnSelect(MPOS_FULLCANCEL);
        // UnlockStartPane(); ?
    }
    if (_pNewStartMenu)
    {
        _pNewStartMenu->OnSelect(MPOS_FULLCANCEL);
    }
}

void CStartButton::DestroyStartMenu()
{
    IUnknown_SetSite(_pUnk1, NULL);
    ATOMICRELEASE(_pUnk1, IDeskBand);

    IUnknown_SetSite(_pOldStartMenu, NULL);
    ATOMICRELEASE(_pOldStartMenu, IMenuPopup);
    ATOMICRELEASE(_pOldStartMenuBand, IMenuBand);

    IUnknown_SetSite(_pNewStartMenu, NULL);
    ATOMICRELEASE(_pNewStartMenu, IMenuPopup);
    ATOMICRELEASE(_pNewStartMenuBand, IMenuBand);
}

void CStartButton::DisplayStartMenu() // xp
{
    RECTL    rcExclude;
    POINTL   ptPop;
    DWORD dwFlags = MPPF_KEYBOARD;      // Assume that we're popuping
    // up because of the keyboard
    // This is for the underlines on NT5

    if (_hwndStartBalloon)
    {
        _DontShowTheStartButtonBalloonAnyMore();
        _DestroyStartButtonBalloon();
    }

    if (GetKeyState(GetSystemMetrics(SM_SWAPBUTTON) ? VK_RBUTTON : VK_LBUTTON) < 0)
    {
        dwFlags = 0;    // Then set to the default
    }

    IMenuPopup** ppmpToDisplay = &_pOldStartMenu;
    if (_pNewStartMenu)
    {
        ppmpToDisplay = &_pNewStartMenu;
    }

    if (!*ppmpToDisplay)
    {
        BuildStartMenu();
    }

    //SetActiveWindow(_hwndStartBtn);

    // Exclude rect is the VISIBLE portion of the Start Button.
    _CalcExcludeRect(&rcExclude);
    ptPop.x = rcExclude.left;
    ptPop.y = rcExclude.top;

    if (*ppmpToDisplay && SUCCEEDED((*ppmpToDisplay)->Popup(&ptPop, &rcExclude, dwFlags)))
    {
        // All is well - the menu is up
        //TraceMsg(DM_MISC, "e.tbm: dwFlags=%x (0=mouse 1=key)", dwFlags);
    }
    else
    {
        //TraceMsg(TF_WARNING, "e.tbm: %08x->Popup failed", *ppmpToDisplay);
        // Start Menu failed to display -- reset the Start Button
        // so the user can click it again to try again
        Tray_OnStartMenuDismissed();
    }

    if (dwFlags == MPPF_KEYBOARD)
    {
        // Since the user has launched the start button by Ctrl-Esc, or some other worldly
        // means, then turn the rect on.
        SendMessage(_hwndStartBtn, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR,
            UISF_HIDEFOCUS), 0);
    }
}

void CStartButton::DrawStartButton(int iStateId, bool hdcSrc)
{
    POINT pptDst;
    HRGN hRgn;
    BOOL started = _CalcStartButtonPos(&pptDst, &hRgn);
    if (hdcSrc)
    {
        HDC DC = GetDC(_hwndStartBtn);
        if (DC)
        {
            HDC hdcSrca = CreateCompatibleDC(DC);
            if (hdcSrca)
            {
                RECT pRect;
                pRect = { 0, 0, _size.cx, _size.cy };

                BITMAPINFO pbmi;
                pbmi.bmiHeader.biWidth = pRect.right;
                pbmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                pbmi.bmiHeader.biHeight = -_size.cy;
                pbmi.bmiHeader.biPlanes = 1;
                pbmi.bmiHeader.biBitCount = 32;
                pbmi.bmiHeader.biCompression = BI_RGB;
                HBITMAP dibBmp = CreateDIBSection(DC, &pbmi, 0, 0, 0, 0);
                if (dibBmp)
                {
                    POINT pptSrc;
                    SIZE psize;
                    HBITMAP hOldBmp = (HBITMAP)SelectObject(hdcSrca, dibBmp);

                    bool isThemeEnabled = _hTheme == NULL;

                    HDC hdcDst = GetDC(NULL);
                    psize = {_size.cx, _size.cy};
                    pptSrc = {0, 0};

                    if (isThemeEnabled)
                    {
                        SendMessageW(_hwndStartBtn, WM_PRINTCLIENT, (WPARAM)hdcSrca, PRF_CLIENT);
                        UpdateLayeredWindow(_hwndStartBtn, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
                            0, NULL, ULW_OPAQUE);
                    }
                    else
                    {
                        SHFillRectClr(hdcSrca, &pRect, 0);
                        DrawThemeBackground(_hTheme, hdcSrca, 1, iStateId, &pRect, 0);

                        BLENDFUNCTION pblend;
                        pblend.BlendOp = 0;
                        pblend.BlendFlags = 0;
                        pblend.SourceConstantAlpha = 255;
                        pblend.AlphaFormat = 1;
                        UpdateLayeredWindow(_hwndStartBtn, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
                            0, &pblend, ULW_ALPHA);
                    }
                    ReleaseDC(0, hdcDst);
                    SelectObject(hdcSrca, hOldBmp);
                    DeleteObject(dibBmp);
                }
                DeleteDC(hdcSrca);
            }
            ReleaseDC(_hwndStartBtn, DC);
        }
    }

    UpdateLayeredWindow(
        _hwndStartBtn, NULL, &pptDst, NULL, NULL, NULL, 0, NULL, 0);

    if (started)
    {
        SetWindowRgn(_hwndStartBtn, hRgn, TRUE);
    }
}

void CStartButton::ExecRefresh()
{
    if (_pOldStartMenuBand)
    {
        IUnknown_Exec(_pOldStartMenuBand, &CLSID_MenuBand, 0x10000000, 0, 0, 0);
    }
    else
    {
        if (_pNewStartMenu)
        {
            IUnknown_Exec(_pNewStartMenu, &CLSID_MenuBand, 0x10000000, 0, 0, 0);
        }
    }
}

void CStartButton::ForceButtonUp()
{
    if (!_nStartBtnNotPressed)
    {
        MSG Msg;
        _nStartBtnNotPressed = 1;
        PeekMessageW(
            &Msg, _hwndStartBtn, WM_MOUSEFIRST + 1, WM_MOUSEFIRST + 1, PM_REMOVE);
        PeekMessageW(
            &Msg, _hwndStartBtn, WM_MOUSEFIRST + 1, WM_MOUSEFIRST + 1, PM_REMOVE);

        SendMessageW(
            _hwndStartBtn, WM_PASTE, 0, 0);

        PeekMessageW(
            &Msg, _hwndStartBtn, WM_MOUSEFIRST + 3, WM_MOUSEFIRST + 3, PM_REMOVE);
    }
}

void CStartButton::GetRect(RECT* lpRect)
{
    GetWindowRect(_hwndStartBtn, lpRect);
}

void CStartButton::GetSizeAndFont(HTHEME hTheme)
{
    if (hTheme)
    {
        HDC hdc = GetDC(_hwndStartBtn);
        GetThemePartSize(hTheme, hdc, BS_PUSHBUTTON, CBS_UNCHECKEDNORMAL, nullptr, TS_TRUE, &_size);
        DWORD dwLogPixelsX = GetDeviceCaps(hdc, LOGPIXELSX);

        // XXX(isabella): Looks to be DPI resolution?
        if (dwLogPixelsX >= 120)
        {
            if (dwLogPixelsX >= 144)
                _nSomeSize = dwLogPixelsX >= 192 ? 14 : 11;
            else
                _nSomeSize = 9;
        }
        else
        {
            _nSomeSize = 8;
        }

        ReleaseDC(_hwndStartBtn, hdc);
    }
    else // Classic theme:
    {
        int idResStartIcon = SHGetCurColorRes() <= 8 ? IDB_START16 : IDB_STARTCLASSIC;

        // @MOD (isabella): Bitmap in Explorer instead of ShellBrd
        HBITMAP hBmpStartIcon = LoadBitmap(hinstCabinet, MAKEINTRESOURCE(idResStartIcon));

        if (hBmpStartIcon)
        {
            BITMAP bitmap;

            if (GetObjectW(hBmpStartIcon, 24, &bitmap))
            {
                BUTTON_IMAGELIST buttonImageList = { 0 };

                if (_hIml)
                {
                    // Clean up any previously-existing image list:
                    ImageList_Destroy(_hIml);
                }

                HBITMAP hBmpStartIconMask = nullptr;

                DWORD ilcFlags = ILC_COLOR32;

                if (idResStartIcon == IDB_START16)
                {
                    ilcFlags = ILC_COLOR8 | ILC_MASK;

                    // @MOD (isabella): Bitmap in Explorer instead of ShellBrd
                    hBmpStartIconMask = LoadBitmap(hinstCabinet, MAKEINTRESOURCE(IDB_START16MASK));
                }

                if ((GetWindowLongW(_hwndStartBtn, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0)
                {
                    ilcFlags |= ILC_MIRROR;
                }

                HIMAGELIST hIml = ImageList_Create(bitmap.bmWidth, bitmap.bmHeight, ilcFlags, 1, 1);
                _hIml = hIml;
                buttonImageList.himl = hIml;
                ImageList_Add(hIml, hBmpStartIcon, hBmpStartIconMask);
                if (hBmpStartIconMask)
                    DeleteObject(hBmpStartIconMask);
                buttonImageList.uAlign = 0;
                SendMessageW(_hwndStartBtn, BCM_SETIMAGELIST, 0, (LPARAM)&buttonImageList);
            }
            DeleteObject(hBmpStartIcon);
        }

        if (_hStartFont)
        {
            DeleteObject(_hStartFont);
        }

        HFONT hFont = _CreateStartFont();
        _hStartFont = hFont;

        SendMessageW(_hwndStartBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Recalculate the size:
        _size = { 0 };
        SendMessageW(_hwndStartBtn, BCM_GETIDEALSIZE, 0, (LPARAM)&_size);
    }
}

BOOL CStartButton::InitBackgroundBitmap()
{
    _fBackgroundBitmapInitialized = TRUE;

    // @MOD (isabella): Vista loads this bitmap from ShellBrd, but we store the bitmap in our own
    // module. Vista's original code is such (link against WinBrand.dll):
    //    _hbmpStartBkg = BrandingLoadBitmap(L"Shellbrd", 1001);
    _hbmpStartBkg = LoadBitmap(hinstCabinet, MAKEINTRESOURCE(IDB_CLASSICSTARTBKG));

    return _hbmpStartBkg != nullptr;
}

void CStartButton::InitTheme()
{
    _pszCurrentThemeName = _GetCurrentThemeName();
    SetWindowTheme(_hwndStartBtn, _GetCurrentThemeName(), 0);
}

BOOL CStartButton::IsButtonPushed()
{
    return SendMessageW(_hwndStartBtn, BM_GETSTATE, 0, 0) & BST_PUSHED;
}

HRESULT CStartButton::IsMenuMessage(MSG* pmsg)  // taken from ep_taskbar 7-stuff @MOD
{
    HRESULT hr;
    if (_pOldStartMenuBand)
    {
        hr = _pOldStartMenuBand->IsMenuMessage(pmsg);
        if (hr != S_OK)
        {
            hr = S_FALSE;
        }
    }
    else if (_pNewStartMenuBand)
    {
        hr = _pNewStartMenuBand->IsMenuMessage(pmsg);
        if (hr != S_OK)
        {
            hr = S_FALSE;
        }
    }
    else
    {
        hr = S_FALSE;
    }
    return hr;
}

BOOL CStartButton::IsPopupMenuVisible()
{
    HWND phwnd;
    return SUCCEEDED(IUnknown_GetWindow(_pOldStartMenu, &phwnd)) && IsWindowVisible(phwnd)
        || SUCCEEDED(IUnknown_GetWindow(_pNewStartMenu, &phwnd)) && IsWindowVisible(phwnd);
}

int CStartButton::_CalcStartButtonPos(POINT *pPoint, HRGN *hRgn)
{
    int ptOrigin; // eax
    int height; // eax
    int v5; // eax
    struct tagPOINT *pPoint_3; // edi
    HMONITOR hMon; // eax
    int ptY; // eax
    int cyDlgFrame; // ebx
    int cyBorder; // edi
    LONG ptX; // eax
    LONG x; // eax
    LONG y; // edi
    LONG v21; // ecx
    int v22; // eax
    int v23; // ecx
    HRGN v24; // eax
    struct tagRECT rcRgn; // [esp+Ch] [ebp-54h] BYREF
    struct tagRECT rc; // [esp+1Ch] [ebp-44h] BYREF
    struct tagRECT rcDst; // [esp+2Ch] [ebp-34h] BYREF
    RECT rcSrc1; // [esp+3Ch] [ebp-24h] BYREF
    int cyBorder_2; // [esp+5Ch] [ebp-4h]
    int fRes; // [esp+68h] [ebp+8h]

    RECT rcTrayWnd;
    GetWindowRect(v_hwndTray, &rcTrayWnd);

    LONG cyFrameHalf = g_cyFrame / 2;

    if (_pszCurrentThemeName == L"StartTop")
    {
        pPoint_3 = pPoint;
        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _size.cx : rcTrayWnd.left;

        if (rcTrayWnd.bottom <= cyFrameHalf)
            ptY = rcTrayWnd.top - _size.cy - cyFrameHalf;
        else
            ptY = rcTrayWnd.bottom + _nSomeSize - _size.cy;
    }
    else if (_pszCurrentThemeName == L"StartBottom")
    {
        pPoint_3 = pPoint;
        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _size.cx : rcTrayWnd.left;

        hMon = MonitorFromRect(&rcTrayWnd, MONITOR_DEFAULTTONEAREST);
        GetMonitorRects(hMon, &rc, FALSE);

        if (rc.bottom - rcTrayWnd.top <= cyFrameHalf)
            ptY = cyFrameHalf + rcTrayWnd.top;
        else
            ptY = rcTrayWnd.top - _nSomeSize;
    }
    else
    {
        if (_hTheme)
        {
            if (g_offsetRECT == 1 || g_offsetRECT == 3)
            {
                if (IsBiDiLocalizedSystem())
                    ptOrigin = rcTrayWnd.right - _size.cx;
                else
                    ptOrigin = rcTrayWnd.left;
                pPoint_3 = pPoint;
                pPoint->x = ptOrigin;
                height = rcTrayWnd.bottom - _size.cy - rcTrayWnd.top;
            }
            else
            {
                pPoint_3 = pPoint;
                pPoint->x = rcTrayWnd.left + (rcTrayWnd.right - _size.cx - rcTrayWnd.left) / 2;
                height = g_cyTabSpace;
            }
            v5 = height / 2;
        }
        else
        {
            cyDlgFrame = GetSystemMetrics(SM_CYDLGFRAME);
            cyBorder = GetSystemMetrics(SM_CYBORDER);
            cyBorder_2 = cyBorder;
            if (IsBiDiLocalizedSystem() && (g_offsetRECT & 1) != 0) // compiler optimised 1 or 3
                ptX = rcTrayWnd.right - _size.cx - cyBorder - cyDlgFrame;
            else
                ptX = rcTrayWnd.left + cyBorder + cyDlgFrame;
            pPoint_3 = pPoint;
            pPoint->x = ptX;
            v5 = cyDlgFrame + cyBorder_2;
        }

        ptY = rcTrayWnd.top + v5;
    }
    fRes = 0;
    pPoint_3->y = ptY;
    if (hRgn)
    {
        if (GetSystemMetrics(SM_CMONITORS) == 1)
        {
            if (GetWindowRgnBox(_hwndStartBtn, &rc))
                SetWindowRgn(_hwndStartBtn, 0, 1);
        }
        else
        {
            c_tray._GetStuckDisplayRect(g_traystuckplace, &rc);
            x = pPoint_3->x;
            y = pPoint_3->y;
            v21 = x + _size.cx;
            rcSrc1.left = x;
            rcSrc1.bottom = y + _size.cy;
            rcSrc1.top = y;
            rcSrc1.right = v21;
            IntersectRect(&rcDst, &rcSrc1, &rc);
            if (EqualRect(&rcDst, &rcSrc1))
            {
                if (GetWindowRgnBox(_hwndStartBtn, &rcRgn))
                {
                    *hRgn = 0;
                    return 1;
                }
            }
            else
            {
                fRes = _ShouldDelayClip(&rcDst, &rc);
                v22 = -rcDst.left;
                v23 = -rcDst.top;
                if (g_offsetRECT)
                {
                    if (g_offsetRECT == 1)
                        v23 = rcSrc1.bottom + -rcDst.bottom - rcSrc1.top;
                }
                else
                {
                    v22 = rcSrc1.right + -rcDst.right - rcSrc1.left;
                }
                OffsetRect(&rcDst, v22, v23);
                v24 = CreateRectRgnIndirect(&rcDst);
                if (fRes)
                    *hRgn = v24;
                else
                    SetWindowRgn(_hwndStartBtn, v24, 1);
            }
        }
    }

    return fRes;
}

void CStartButton::RecalcSize()
{
    if (!_hTheme)
    {
        RECT rc;
        GetClientRect(v_hwndTray, &rc);
        LPCWSTR p_windowName = &WindowName;
        if (rc.right >= _size.cx)
        {
            p_windowName = _pszWindowName;
        }

        SetWindowTextW(_hwndStartBtn, p_windowName);

        int height = _pStartButtonSite->GetStartButtonMinHeight();
        if (!height && _size.cy >= 0)
        {
            height = _size.cy;
        }
        _size.cy = height;
    }
}

void CStartButton::RepositionBalloon()
{
    if (_hwndStartBalloon)
    {
        RECT rc;
        GetWindowRect(_hwndStartBtn, &rc);
        WORD xCoordinate = (WORD)((rc.left + rc.right) / 2);
        WORD yCoordinate = LOWORD(rc.top);
        SendMessageW(
            _hwndStartBalloon, 0x412u, 0, MAKELONG(xCoordinate, yCoordinate));
    }
}

void CStartButton::StartButtonReset()
{
    GetSizeAndFont(_hTheme);
    RecalcSize();
}

int CStartButton::TrackMenu(HMENU hMenu)
{
    TPMPARAMS tpm;
    RECT rc;
    tpm.cbSize = sizeof(tpm);
    GetClientRect(_hwndStartBtn, &tpm.rcExclude);
    HWND Parent = GetParent(_hwndStartBtn);
    GetClientRect(Parent, &rc);
    if (tpm.rcExclude.bottom >= rc.bottom)
    {
        tpm.rcExclude.bottom = rc.bottom;
    }
    MapWindowRect(_hwndStartBtn, NULL, &tpm.rcExclude);
    UINT uFlag = TPM_RETURNCMD | TPM_RIGHTBUTTON;
    if (IsBiDiLocalizedSystem())
    {
        uFlag |= TPM_LAYOUTRTL;
    }
    _pStartButtonSite->EnableTooltips(FALSE);
    int nMenuResult = TrackPopupMenuEx(hMenu, uFlag, tpm.rcExclude.left, tpm.rcExclude.bottom, _hwndStartBtn, &tpm);
    _pStartButtonSite->EnableTooltips(TRUE);
    return nMenuResult;
}

BOOL CStartButton::TranslateMenuMessage(MSG* pmsg, LRESULT* plRet)  // taken from ep_taskbar 7-stuff @MOD
{
    BOOL result = TRUE;
    if (_pOldStartMenuBand)
    {
        // S_FALSE is same as TRUE
        // S_OK is same as FALSE
        // in pseudocode, it checks if TranslateMenuMessage is TRUE (S_FALSE)
        result = (_pOldStartMenuBand->TranslateMenuMessage(pmsg, plRet) == S_FALSE) ? TRUE : FALSE;

        if (result && _pNewStartMenuBand)
        {
            result = (_pNewStartMenuBand->TranslateMenuMessage(pmsg, plRet) == S_FALSE) ? TRUE : FALSE;
        }
    }
    return result;
}

void CStartButton::UpdateStartButton(bool a2)
{
    if (_hTheme && _GetCurrentThemeName() != _pszCurrentThemeName)
    {
        _pszCurrentThemeName = _GetCurrentThemeName();
        SetWindowTheme(_hwndStartBtn, _GetCurrentThemeName(), 0);
    }
    else
    {
        DrawStartButton(1, a2);
    }
}

void CStartButton::_DestroyStartButtonBalloon()
{
    if (_hwndStartBalloon)
    {
        DestroyWindow(_hwndStartBalloon);
        _hwndStartBalloon = nullptr;
    }
    KillTimer(_hwndStartBtn, 1);
}

void CStartButton::_DontShowTheStartButtonBalloonAnyMore()
{
    DWORD dwData = 2;
    SHSetValueW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced",
        L"StartButtonBalloonTip", REG_DWORD, &dwData, sizeof(dwData));
}

LRESULT CStartButton::OnMouseClick(HWND hWndTo, LPARAM lParam)
{
    LRESULT lRes = S_OK;
    if (_hwndStartBalloon)
    {
        RECT rcBalloon;
        GetWindowRect(_hwndStartBalloon, &rcBalloon);
        MapWindowRect(nullptr, hWndTo, &rcBalloon);
        if (PtInRect(&rcBalloon, {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}))
        {
            ShowWindow(_hwndStartBalloon, SW_HIDE);
            _DontShowTheStartButtonBalloonAnyMore();
            _DestroyStartButtonBalloon();
            lRes = S_FALSE;
        }
    }

    return lRes;
}

void CStartButton::_CalcExcludeRect(RECTL* lprcDst) // from xp
{
    RECTL rcExclude;
    RECT rcParent;

    GetClientRect(_hwndStartBtn, (RECT*)&rcExclude);
    MapWindowRect(_hwndStartBtn, HWND_DESKTOP, &rcExclude);

    GetClientRect(v_hwndTray, &rcParent);
    MapWindowRect(v_hwndTray, HWND_DESKTOP, &rcParent);

    IntersectRect((RECT*)&rcExclude, (RECT*)&rcExclude, &rcParent);

    *lprcDst = rcExclude;
}

HFONT CStartButton::_CreateStartFont()  // taken from xp
{
    HFONT hfontStart = NULL;
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

    return hfontStart;
}

void CStartButton::_ExploreCommonStartMenu(BOOL a2)
{
    LPITEMIDLIST ppidl;
    if (SUCCEEDED(SHGetFolderLocation(nullptr, CSIDL_COMMON_STARTMENU, nullptr, KF_FLAG_DEFAULT, &ppidl)))
    {
        SHELLEXECUTEINFOW execInfo = { sizeof(execInfo) };
        execInfo.fMask = SEE_MASK_IDLIST | SEE_MASK_ASYNCOK;
        execInfo.lpVerb = a2 ? L"explore" : L"open";
        execInfo.nShow = SW_NORMAL;
        execInfo.lpIDList = ppidl;
        ShellExecuteExW(&execInfo);
        ILFree(ppidl);
    }
}

LPCWSTR CStartButton::_GetCurrentThemeName()
{
    RECT rc;
    GetWindowRect(v_hwndTray, &rc);
    LPCWSTR pszResult = L"StartTop";

    if (g_traystuckplace != 1)
    {
        if (g_traystuckplace == 3 && RECTHEIGHT(rc) < _size.cy)
        {
            return L"StartBottom";
        }
        return L"StartMiddle";
    }

    if (RECTHEIGHT(rc) >= _size.cy)
    {
        return L"StartMiddle";
    }

    return pszResult;
}

void CStartButton::_HandleDestroy()
{
    _fBackgroundBitmapInitialized = 0;
    _DestroyStartButtonBalloon();

    if (_hbmpStartBkg)
    {
        DeleteObject(_hbmpStartBkg);
    }
    if (_hStartFont)
    {
        DeleteObject(_hStartFont);
    }
    if (_hIml)
    {
        ImageList_Destroy(_hIml);
    }

    RemovePropW(_hwndStartBtn, L"StartButtonTag");
}

void CStartButton::_OnSettingChanged(UINT a2)
{
    if (!_hTheme && a2 != 47)
    {
        bool v3 = !_nSettingsChangeType;
        if (_nSettingsChangeType)
        {
            PostMessageW(_hwndStartBtn, 0x31Au, 0, 0);
            v3 = !_nSettingsChangeType;
        }
        _nSettingsChangeType = v3;
    }
}

bool CStartButton::_OnThemeChanged(bool bForceUpdate)
{
    if (_hTheme)
    {
        CloseThemeData(_hTheme);
        _hTheme = NULL;
    }

    bool bThemeApplied = false;
    HTHEME hTheme = OpenThemeData(_hwndStartBtn, L"Button");
    _hTheme = hTheme;
    if (hTheme)
    {
        StartButtonReset();
        DrawStartButton(1, true);
    }
    else if (!bForceUpdate)
    {
        bool bSettingsChanged = !_nSettingsChangeType;
        _pszCurrentThemeName = NULL;
        if (bSettingsChanged)
        {
            StartButtonReset();
            DrawStartButton(1, true);
            bThemeApplied = true;
        }
        else
        {
            PostMessageW(_hwndStartBtn, 0x31Au, 0, 0);
        }
        _nSettingsChangeType = !_nSettingsChangeType;
    }

    return bThemeApplied;
}

BOOL CStartButton::_ShouldDelayClip(const RECT* a2, const RECT* lprcSrc2)
{
    RECT rc1;
    RECT rcClip;
    RECT rcDst;

    GetWindowRect(_hwndStartBtn, &rcClip);
    IntersectRect(&rcDst, &rcClip, lprcSrc2);
    IntersectRect(&rc1, &rcDst, a2);

    return EqualRect(&rc1, &rcDst);
}

LRESULT CStartButton::s_StartButtonSubclassProc(
    HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, CStartButton* dwRefData)
{
    return dwRefData->_StartButtonSubclassProc(hWnd, uMsg, wParam, lParam);
}
