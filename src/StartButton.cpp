#include "pch.h"
#include "StartButton.h"
#include "SHFusion.h"
#include "tray.h"
#include "Util.h"
#include "DeskHost.h"

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
    SetFocus(_hwndStart);
    return S_OK;
}

HRESULT CStartButton::OnContextMenu(HWND hWnd, LPARAM lParam)
{
    _nIsOnContextMenu = TRUE;
    _pStartButtonSite->HandleFullScreenApp2(hWnd);
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
            _hwndStart, NULL, GetModuleHandle(NULL), NULL);

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
        toolInfo.hwnd = _hwndStart;
        toolInfo.uId = (UINT_PTR)_hwndStart;  //TTF_IDISHWND
        
        SendMessage(_hwndStartBalloon, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);
        SendMessage(_hwndStartBalloon, TTM_TRACKACTIVATE, 0, 0);

        toolInfo.lpszText = Buffer;
        SendMessage(_hwndStartBalloon, TTM_UPDATETIPTEXT, 0, (LPARAM)&toolInfo);

        if (LoadString(GetModuleHandle(NULL), a2, Buffer, 260))
        {
            SendMessage(_hwndStartBalloon, TTM_SETTITLE, TTI_INFO, (LPARAM)&Buffer);
        }

        RECT Rect;
        GetWindowRect(_hwndStart, &Rect);
        WORD xCoordinate = (WORD)((Rect.left + Rect.right) / 2);
        WORD yCoordinate = LOWORD(Rect.top);

        SendMessage(_hwndStartBalloon, TTM_TRACKPOSITION, 0, MAKELONG(xCoordinate, yCoordinate));
        SendMessage(_hwndStartBalloon, TTM_TRACKACTIVATE, TRUE, (LPARAM)&toolInfo);

        // show tooltip for 10 seconds
        SetTimer(_hwndStart, 1u, 10000, 0);
    }
    return S_OK;
}

// dwRefData-> CTray+988 -> CStartButton
HWND CStartButton::CreateStartButton(HWND hWnd)
{
    wprintf(L"CStartButton::CreateStartButton zero\n");

    DWORD dwStyleEx =  GetWindowLongPtrW(hWnd, GWL_EXSTYLE);
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
    _hwndStart = hWndStartBtn;

    wprintf(L"CStartButton::CreateStartButton parent HWND = %p\n", hWnd);

    wprintf(L"CStartButton::CreateStartButton one\n");

    if (hWndStartBtn)
    {
        SetPropW(hWndStartBtn, L"StartButtonTag", (HANDLE)0x130);
        SendMessageW(_hwndStart, CCM_DPISCALE, TRUE, 0);
        SetWindowSubclass(
            _hwndStart,
            (SUBCLASSPROC)CStartButton::s_StartButtonSubclassProc,
            0,
            (DWORD_PTR)this);
        LoadStringW(GetModuleHandle(NULL), IDS_STARTCLASSIC, (LPWSTR)_szWindowName, 50);
        SetWindowTextW(_hwndStart, (LPCWSTR)_szWindowName);
    }

    wprintf(L"CStartButton::CreateStartButton two\n");

    return _hwndStart;
}

HRESULT CStartButton::SetStartPaneActive(BOOL bActive)
{
    if (bActive)
    {
        _uStartButtonState = 1;

        // @MOD check if this happens
        _uDown = 1;
    }
    else if (_uStartButtonState != 2)
    {
        _uStartButtonState = 0;

        // @MOD check if this happens

        _uDown = 0;
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
    *out = _hwndStart;
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
        IUnknown_SetSite(_pUnk1, static_cast<IServiceProvider *>(this));
        DesktopV2_Build(pDeskHost);
    }
    else
    {
        HRESULT hr = StartMenuHost_Create(&_pOldStartMenu, &_pOldStartMenuBand);
        if (SUCCEEDED(hr))
        {
            IUnknown_SetSite(_pUnk1, static_cast<IServiceProvider *>(this));
            HWND hwnd;
            if (SUCCEEDED(IUnknown_GetWindow(_pOldStartMenu, &hwnd)))
            {
                SetWindowSubclass(hwnd, s_StartMenuSubclassProc, 0, (DWORD_PTR)this);
            }

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
    ATOMICRELEASE(_pUnk1);

    IUnknown_SetSite(_pOldStartMenu, NULL);
    ATOMICRELEASE(_pOldStartMenu);
    ATOMICRELEASE(_pOldStartMenuBand);

    IUnknown_SetSite(_pNewStartMenu, NULL);
    ATOMICRELEASE(_pNewStartMenu);
    ATOMICRELEASE(_pNewStartMenuBand);
}

// EXEX-VISTA: REVALIDATE. Partially reversed from Vista.
void CStartButton::DisplayStartMenu()
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

    // Recalculate the position of the taskbar and start button:
    bool fMoveTaskbar = false;
    if ((GetWindowLongPtrW(_hwndStart, GWL_EXSTYLE) & WS_EX_TOPMOST) == 0
        && _pszCurrentThemeName == L"StartBottom"
        || _pszCurrentThemeName == L"StartTop")
    {
        fMoveTaskbar = true;
        SetWindowPos(v_hwndTray, HWND_TOPMOST, 0, 0, 0, 0, (SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER));
        SetWindowPos(_hwndStart, HWND_TOPMOST, 0, 0, 0, 0, (SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER));
    }

    HWND hWndMenu;
    if (SUCCEEDED(IUnknown_GetWindow(*ppmpToDisplay, &hWndMenu)))
    {
        SetWindowPos(hWndMenu, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        SetWindowPos(hWndMenu, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    // Exclude rect is the VISIBLE portion of the Start Button.
    _CalcExcludeRect(&rcExclude);
    ptPop.x = rcExclude.left;
    ptPop.y = rcExclude.top;

    if (*ppmpToDisplay && SUCCEEDED((*ppmpToDisplay)->Popup(&ptPop, &rcExclude, dwFlags)))
    {
        // All is well - the menu is up
        //TraceMsg(DM_MISC, "e.tbm: dwFlags=%x (0=mouse 1=key)", dwFlags);

        if (dwFlags == MPPF_KEYBOARD)
        {
            SendMessage(_hwndStart, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEFOCUS), 0);
        }
    }
    else
    {
        //TraceMsg(TF_WARNING, "e.tbm: %08x->Popup failed", *ppmpToDisplay);
        if (dwFlags == MPPF_KEYBOARD)
        {
            // Since the user has launched the start button by Ctrl-Esc, or some other worldly
            // means, then turn the rect on.
            SendMessage(_hwndStart, WM_UPDATEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEFOCUS), 0);
        }

        // Start Menu failed to display -- reset the Start Button
        // so the user can click it again to try again
        OnStartMenuDismissed();
    }

    if (fMoveTaskbar)
    {
        SetWindowPos(v_hwndTray, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        SetWindowPos(_hwndStart, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        SetWindowPos(v_hwndTray, _hwndStart, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }
}

void CStartButton::DrawStartButton(int iStateId, bool hdcSrc)
{
    POINT pptDst;
    HRGN hRgn;
    BOOL started = _CalcStartButtonPos(&pptDst, &hRgn);

    if (hdcSrc)
    {
        HDC DC = GetDC(_hwndStart);
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
                        SendMessageW(_hwndStart, WM_PRINTCLIENT, (WPARAM)hdcSrca, PRF_CLIENT);
                        UpdateLayeredWindow(_hwndStart, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
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
                        UpdateLayeredWindow(_hwndStart, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
                            0, &pblend, ULW_ALPHA);
                    }
                    ReleaseDC(0, hdcDst);
                    SelectObject(hdcSrca, hOldBmp);
                    DeleteObject(dibBmp);
                }
                DeleteDC(hdcSrca);
            }
            ReleaseDC(_hwndStart, DC);
        }
    }

    UpdateLayeredWindow(
        _hwndStart, NULL, &pptDst, NULL, NULL, NULL, 0, NULL, 0);


    if (started)
    {
        SetWindowRgn(_hwndStart, hRgn, TRUE);
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
            &Msg, _hwndStart, WM_MOUSEFIRST + 1, WM_MOUSEFIRST + 1, PM_REMOVE);
        PeekMessageW(
            &Msg, _hwndStart, WM_MOUSEFIRST + 1, WM_MOUSEFIRST + 1, PM_REMOVE);

        SendMessageW(
            _hwndStart, WM_PASTE, 0, 0);

        PeekMessageW(
            &Msg, _hwndStart, WM_MOUSEFIRST + 3, WM_MOUSEFIRST + 3, PM_REMOVE);
    }
}

void CStartButton::GetRect(RECT* lpRect)
{
    GetWindowRect(_hwndStart, lpRect);
}

void CStartButton::GetSizeAndFont(const HTHEME hTheme)
{
    if (hTheme)
    {
        HDC hdc = GetDC(_hwndStart);
        GetThemePartSize(hTheme, hdc, BP_PUSHBUTTON, CBS_UNCHECKEDNORMAL, nullptr, TS_TRUE, &_size);
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

        ReleaseDC(_hwndStart, hdc);
    }
    else // Classic theme:
    {
        int idResStartIcon = SHGetCurColorRes() <= 8 ? IDB_START16 : IDB_STARTCLASSIC;

        // @MOD (isabella): Bitmap in Explorer instead of ShellBrd
        HBITMAP hBmpStartIcon = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(idResStartIcon));

        if (hBmpStartIcon)
        {
            BITMAP bitmap{};

            if (GetObjectW(hBmpStartIcon, sizeof(bitmap), &bitmap))
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
                    hBmpStartIconMask = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_START16MASK));
                }

                if ((GetWindowLongPtrW(_hwndStart, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0)
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
                SendMessageW(_hwndStart, BCM_SETIMAGELIST, 0, (LPARAM)&buttonImageList);
            }
            DeleteObject(hBmpStartIcon);
        }

        if (_hStartFont)
        {
            DeleteObject(_hStartFont);
        }

        HFONT hFont = _CreateStartFont();
        _hStartFont = hFont;

        SendMessageW(_hwndStart, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Recalculate the size:
        _size = { 0 };
        SendMessageW(_hwndStart, BCM_GETIDEALSIZE, 0, (LPARAM)&_size);
    }
}

BOOL CStartButton::InitBackgroundBitmap()
{
    _fBackgroundBitmapInitialized = TRUE;

    // @MOD (isabella): Vista loads this bitmap from ShellBrd, but we store the bitmap in our own
    // module. Vista's original code is such (link against WinBrand.dll):
    //    _hbmpStartBkg = BrandingLoadBitmap(L"Shellbrd", 1001);
    _hbmpStartBkg = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_CLASSICSTARTBKG));

    return _hbmpStartBkg != nullptr;
}

void CStartButton::InitTheme()
{
    _pszCurrentThemeName = _GetCurrentThemeName();
    SetWindowTheme(_hwndStart, _GetCurrentThemeName(), 0);
}

BOOL CStartButton::IsButtonPushed()
{
    return SendMessageW(_hwndStart, BM_GETSTATE, 0, 0) & BST_PUSHED;
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

BOOL CStartButton::_CalcStartButtonPos(POINT *pPoint, HRGN *phRgn)
{
    //*pPoint = { 0, 0 };
    //return TRUE;

    RECT rcTrayWnd;
    GetWindowRect(v_hwndTray, &rcTrayWnd);

    LONG cyFrameHalf = g_cyFrame / 2;

    wprintf(L"CStartButton::_CalcStartButtonPos zero\n");

    if (_pszCurrentThemeName == L"StartTop")
    {
        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _size.cx : rcTrayWnd.left;

        if (rcTrayWnd.bottom <= cyFrameHalf)
            pPoint->y = rcTrayWnd.top - _size.cy - cyFrameHalf;
        else
            pPoint->y = rcTrayWnd.bottom + _nSomeSize - _size.cy;
    }
    else if (_pszCurrentThemeName == L"StartBottom")
    {
        RECT rc;

        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _size.cx : rcTrayWnd.left;

        HMONITOR hMon = MonitorFromRect(&rcTrayWnd, MONITOR_DEFAULTTONEAREST);
        GetMonitorRects(hMon, &rc, FALSE);

        if (rc.bottom - rcTrayWnd.top <= cyFrameHalf)
            pPoint->y = cyFrameHalf + rcTrayWnd.top;
        else
            pPoint->y = rcTrayWnd.top - _nSomeSize;
    }
    else if (_hTheme)
    {
        int height;

        if (STUCK_HORIZONTAL(c_tray._uStuckPlace))
        {
            pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _size.cx : rcTrayWnd.left;
            height = rcTrayWnd.bottom - _size.cy - rcTrayWnd.top;
        }
        else
        {
            pPoint->x = rcTrayWnd.left + (rcTrayWnd.right - _size.cx - rcTrayWnd.left) / 2;
            height = g_cyTabSpace;
        }

        pPoint->y = rcTrayWnd.top +  height / 2;
    }
    else
    {
        int cyDlgFrame = GetSystemMetrics(SM_CYDLGFRAME);
        int cyBorder = GetSystemMetrics(SM_CYBORDER);
        if (IsBiDiLocalizedSystem() && (c_tray._uStuckPlace == 1 || c_tray._uStuckPlace == 3) != 0)
            pPoint->x = rcTrayWnd.right - _size.cx - cyBorder - cyDlgFrame;
        else
            pPoint->x = rcTrayWnd.left + cyBorder + cyDlgFrame;
        pPoint->y = rcTrayWnd.top + cyDlgFrame + cyBorder;
    }

    // XXX(isabella): Inlined function? New result variable in the middle of the call may be indicative.
    BOOL fRes = FALSE;
    if (phRgn)
    {
        RECT rc;
        RECT rcDst;
        RECT rcRgn;

        if (GetSystemMetrics(SM_CMONITORS) == 1)
        {
            if (GetWindowRgnBox(_hwndStart, &rc))
            {
                SetWindowRgn(_hwndStart, nullptr, TRUE);
            }
        }
        else
        {
            c_tray.GetStuckMonitorRect(&rc);

            RECT rcSrc;
            rcSrc.left = pPoint->x;
            rcSrc.bottom = pPoint->y + _size.cy;
            rcSrc.top = pPoint->y;
            rcSrc.right = pPoint->x + _size.cx;
            IntersectRect(&rcDst, &rcSrc, &rc);

            if (EqualRect(&rcDst, &rcSrc))
            {
                if (GetWindowRgnBox(_hwndStart, &rcRgn))
                {
                    *phRgn = 0;
                    return 1;
                }
            }
            else
            {
                fRes = _ShouldDelayClip(&rcDst, &rc);

                int dx = -rcDst.left;
                int dy = -rcDst.top;

                if (c_tray._uStuckPlace == 1)
                {
                    dy = rcSrc.bottom + -rcDst.bottom - rcSrc.top;
                }
                else
                {
                    dx = rcSrc.right + -rcDst.right - rcSrc.left;
                }

                OffsetRect(&rcDst, dx, dy);
                HRGN hRgn = CreateRectRgnIndirect(&rcDst);
                if (fRes)
                {
                    *phRgn = hRgn;
                }
                else
                {
                    SetWindowRgn(_hwndStart, hRgn, TRUE);
                }
            }
        }
    }

    wprintf(L"CStartButton::_CalcStartButtonPos one\n");

    return fRes;
}

void CStartButton::RecalcSize()
{
    if (!_hTheme)
    {
        RECT rc;
        GetClientRect(v_hwndTray, &rc);
        LPCWSTR pszWindowName = L"";
        if (rc.right >= _size.cx)
        {
            pszWindowName = _szWindowName;
        }

        SetWindowTextW(_hwndStart, pszWindowName);

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
        GetWindowRect(_hwndStart, &rc);
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
    GetClientRect(_hwndStart, &tpm.rcExclude);
    HWND Parent = GetParent(_hwndStart);
    GetClientRect(Parent, &rc);
    if (tpm.rcExclude.bottom >= rc.bottom)
    {
        tpm.rcExclude.bottom = rc.bottom;
    }
    MapWindowRect(_hwndStart, NULL, &tpm.rcExclude);
    UINT uFlag = TPM_RETURNCMD | TPM_RIGHTBUTTON;
    if (IsBiDiLocalizedSystem())
    {
        uFlag |= TPM_LAYOUTRTL;
    }
    _pStartButtonSite->EnableTooltips(FALSE);
    int nMenuResult = TrackPopupMenuEx(hMenu, uFlag, tpm.rcExclude.left, tpm.rcExclude.bottom, _hwndStart, &tpm);
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
        SetWindowTheme(_hwndStart, _GetCurrentThemeName(), 0);
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
    KillTimer(_hwndStart, 1);
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
#if 0
    RECTL rcExclude;
    RECT rcParent;

    GetClientRect(_hwndStart, (RECT*)&rcExclude);
    MapWindowRect(_hwndStart, HWND_DESKTOP, &rcExclude);

    GetClientRect(v_hwndTray, &rcParent);
    MapWindowRect(v_hwndTray, HWND_DESKTOP, &rcParent);

    IntersectRect((RECT*)&rcExclude, (RECT*)&rcExclude, &rcParent);

    *lprcDst = rcExclude;
#else
    RECT rcStart;
    GetWindowRect(_hwndStart, &rcStart);

    RECT rcMonitor;
    GetMonitorRect(MonitorFromRect(&rcStart, 0), &rcMonitor);

    UINT uStuckPlace = c_tray.GetStartMenuStuckPlace();

    RECT rcStuck = c_tray._arStuckRects[uStuckPlace];
    if (IsBiDiLocalizedSystem() && STUCK_HORIZONTAL(uStuckPlace))
        rcStart.left = rcStuck.right - _size.cx;
    else
        rcStart.left = rcStuck.left;

    if (uStuckPlace == 1 || uStuckPlace == 3)
    {
        rcStart.top = rcStuck.top;
        rcStart.bottom = rcStuck.bottom;
    }

    IntersectRect((RECT*)lprcDst, &rcMonitor, &rcStart);
#endif
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

    if (c_tray._uStuckPlace == STICK_BOTTOM && RECTHEIGHT(rc) < _size.cy)
    {
        return L"StartBottom";
    }
    else if (c_tray._uStuckPlace == STICK_TOP && RECTHEIGHT(rc) < _size.cy)
    {
        return L"StartTop";
    }
    else
    {
        return L"StartMiddle";
    }
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

    RemovePropW(_hwndStart, L"StartButtonTag");
}

void CStartButton::_OnSettingChanged(UINT a2)
{
    if (!_hTheme && a2 != 0x2F)
    {
        bool v3 = !_nSettingsChangeType;
        if (_nSettingsChangeType)
        {
            PostMessageW(_hwndStart, 0x31Au, 0, 0);
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
        _hTheme = nullptr;
    }

    bool bThemeApplied = false;
    _hTheme = OpenThemeData(_hwndStart, L"Button");

    if (_hTheme)
    {
        StartButtonReset();
        DrawStartButton(PBS_NORMAL, true);
    }
    else if (!bForceUpdate)
    {
        _pszCurrentThemeName = nullptr;
        if (!_nSettingsChangeType)
        {
            StartButtonReset();
            DrawStartButton(PBS_NORMAL, true);
            bThemeApplied = true;
        }
        else
        {
            PostMessageW(_hwndStart, 0x31Au, 0, 0);
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

    GetWindowRect(_hwndStart, &rcClip);
    IntersectRect(&rcDst, &rcClip, lprcSrc2);
    IntersectRect(&rc1, &rcDst, a2);

    return EqualRect(&rc1, &rcDst);
}

LRESULT CStartButton::_StartButtonSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (uMsg == BM_SETSTATE)
    {
        // this part is mixed with xp code, 
        if (wParam)
        {
            if (!_uDown)
            {
                // Nope.
                INSTRUMENT_STATECHANGE(SHCNFI_STATE_START_DOWN);
                _uDown = 1;

                _fAllowUp = FALSE;
                _pStartButtonSite->EnableTooltips(FALSE);

                // Show the button down.
                LRESULT lRet = DefWindowProc(hWnd, uMsg, wParam, lParam);
                SendMessage(GetParent(hWnd), WM_COMMAND, (WPARAM)LOWORD(GetDlgCtrlID(hWnd)), (LPARAM)hWnd);
                _pStartButtonSite->StartButtonClicked();
                DrawStartButton(PBS_PRESSED, true);
                _tmOpen = GetTickCount();
                return lRet;
            }
            else
            {
                // Yep. Do nothing.
                // fDown = FALSE;
                return DefWindowProc(hWnd, uMsg, wParam, lParam);
            }
        }
        else
        {
            if (_uDown == 1 || !_fAllowUp)
            {
                INSTRUMENT_STATECHANGE(SHCNFI_STATE_START_UP);

                _uDown = 2;
                return DefWindowProc(hWnd, uMsg, wParam, lParam);
            }
            else
            {
                _pStartButtonSite->EnableTooltips(TRUE);

                // Nope, Forward it on.
                _uDown = 0;
                DrawStartButton(PBS_NORMAL, true);
                return DefSubclassProc(hWnd, BM_SETSTATE, 0, lParam);
            }
        }
    }
    else
    {
        // all vista reimpl'd code

        if (_uStartButtonState == 2 && uMsg >= WM_MOUSEFIRST && uMsg <= WM_MOUSELAST)
        {
            _uStartButtonState = 0;
            CloseStartMenu();
        }

        switch (uMsg)
        {
            case WM_WINDOWPOSCHANGING:
            {
                WINDOWPOS* pWindowPos = (WINDOWPOS*)lParam;

                if (!(pWindowPos->flags & SWP_NOMOVE))
                {
                    POINT pt;
                    _CalcStartButtonPos(&pt, nullptr);

                    if (pWindowPos->x != pt.x || pWindowPos->y != pt.y)
                    {
                        pWindowPos->x = pt.x;
                        pWindowPos->y = pt.y;
                        return 0;
                    }
                }
                break;
            }

            case WM_SETFOCUS:
            {
                if (_hTheme && !_uDown)
                {
                    DrawStartButton(PBS_HOT, true);
                }
                break;
            }

            case WM_KILLFOCUS:
            {
                if (!_fAllowUp && _hTheme && !_uDown)
                {
                    DrawStartButton(PBS_NORMAL, true);
                }
                break;
            }

            case WM_CLOSE:
            {
                _pStartButtonSite->OnStartButtonClosing();
                return 0;
            }

            case WM_SETTINGCHANGE:
            {
                _OnSettingChanged(wParam);
                break;
            }

            case WM_MOUSEMOVE:
            {
                if (_hTheme && !_mouseOver && !_uDown)
                {
                    DrawStartButton(PBS_HOT, true);

                    TRACKMOUSEEVENT evt;
                    evt.dwHoverTime = HOVER_DEFAULT;
                    evt.hwndTrack = _hwndStart;
                    evt.cbSize = sizeof(TRACKMOUSEEVENT);
                    evt.dwFlags = TME_LEAVE;

                    TrackMouseEvent(&evt);

                    _mouseOver = 1;
                }
                break;
            }

            case WM_MOUSELEAVE:
            {
                if (_hTheme && !_uDown)
                {
                    DrawStartButton(PBS_NORMAL, true);
                }

                _mouseOver = 0;
                return 0;
            }

            case WM_THEMECHANGED:
            case WM_DWMCOMPOSITIONCHANGED:
            {
                if (_OnThemeChanged(uMsg == WM_DWMCOMPOSITIONCHANGED))
                {
                    return 0;
                }

                break;
            }

            case WM_LBUTTONDOWN:
            case WM_RBUTTONDOWN:
            case WM_MBUTTONDOWN:
            {
                if (OnMouseClick(hWnd, lParam))
                {
                    return 0;   
                }
                if (uMsg == WM_LBUTTONDOWN)
                {
                    SendMessageW(GetAncestor(hWnd, GA_ROOTOWNER), WM_UPDATEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
                    LRESULT lr = DefSubclassProc(hWnd, uMsg, wParam, lParam);
                    SetCapture(0);
                    return lr;
                }

                break;
            }

            // _nStatePaneActiveState... HERE

            case WM_TIMER:
            {
                if (wParam == 1) // Timer ID
                {
                    _DestroyStartButtonBalloon();
                }

                return DefSubclassProc(hWnd, uMsg, wParam, lParam);
            }

            case WM_APP: // probably like "recalc size"
            {
                RecalcSize();
                return 0;
            }

            case WM_APP + 1: // probably like "get size"
            {
                return MAKELRESULT(_size.cx, _size.cy);
            }

            case WM_KEYDOWN:
            {
                SendMessageW(GetAncestor(hWnd, GA_ROOTOWNER), WM_UPDATEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEACCEL), 0);

                if (wParam == VK_RETURN)
                {
                    PostMessageW(GetAncestor(hWnd, GA_ROOTOWNER), WM_COMMAND, 0x131, 0);
                }

                LRESULT lr = DefSubclassProc(hWnd, uMsg, wParam, lParam);
                return lr;
            }

            case WM_DESTROY:
            {
                _HandleDestroy();
                return DefSubclassProc(hWnd, uMsg, wParam, lParam);
            }

            case WM_CONTEXTMENU:
            {
                if (!SHRestricted(REST_NOTRAYCONTEXTMENU))
                    OnContextMenu(hWnd, lParam);
                return 0;
            }
            case WM_NCDESTROY:
            {
                RemoveWindowSubclass(hWnd, s_StartButtonSubclassProc, 0);
                break;
            }
            case WM_NCHITTEST:
            {
                if (GetTickCount() - _tmOpen < GetDoubleClickTime())
                {
                    return 0;
                }

                _pStartButtonSite->SetUnhideTimer(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
                break;
            }

            case WM_MOUSEACTIVATE:
            {
                if (!_uStartButtonState)
                {
                    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
                }

                _uStartButtonState = 2;
                return 2;
            }
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }
    }

    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CStartButton::s_StartButtonSubclassProc(
    HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    return ((CStartButton *)dwRefData)->_StartButtonSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CStartButton::s_StartMenuSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    if (uMsg != WM_WINDOWPOSCHANGING)
    {
        if (uMsg == WM_NCDESTROY)
            RemoveWindowSubclass(hWnd, s_StartMenuSubclassProc, uIdSubclass);
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    }

    WINDOWPOS* wndPos = (WINDOWPOS*)lParam;
    if ((wndPos->flags & SWP_NOZORDER) != 0 || c_tray._uStuckPlace != 1 && c_tray._uStuckPlace != 3)
    {
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    }
    wndPos->hwndInsertAfter = v_hwndTray;

    return 0;
}
