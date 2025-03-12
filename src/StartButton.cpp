#include "StartButton.h"

#include <shlwapi.h>
#include <windowsx.h>

#include "SHFusion.h"
#include "tray.h"
#include "Util.h"

CStartButton::CStartButton()
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
                BITMAPINFO pbmi;
                pRect = { 0, 0, _lWidth, _lHeight };
                pbmi.bmiHeader.biWidth = pRect.right;
                pbmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                pbmi.bmiHeader.biHeight = -_lHeight;
                pbmi.bmiHeader.biPlanes = 1;
                pbmi.bmiHeader.biBitCount = 32;
                pbmi.bmiHeader.biCompression = 0;
                HBITMAP v6 = CreateDIBSection(DC, &pbmi, 0, 0, 0, 0);
                HGDIOBJ ho = v6;
                if (v6)
                {
                    POINT pptSrc;
                    SIZE psize;
                    HGDIOBJ h = SelectObject(hdcSrca, v6);
                    bool isThemeEnabled = _hTheme == NULL;
                    HDC hdcDst = GetDC(NULL);
                    psize = {_lWidth, _lHeight};
                    pptSrc = {0, 0};
                    if (isThemeEnabled)
                    {
                        SendMessageW(_hwndStartBtn, WM_PRINTCLIENT, (WPARAM)hdcSrca, 4);
                        UpdateLayeredWindow(_hwndStartBtn, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
                            0, NULL, ULW_OPAQUE);
                    }
                    else
                    {
                        BLENDFUNCTION pblend;
                        SHFillRectClr(hdcSrca, &pRect, 0);
                        DrawThemeBackground(_hTheme, hdcSrca, 1, iStateId, &pRect, 0);
                        pblend.BlendOp = 0;
                        pblend.BlendFlags = 0;
                        pblend.SourceConstantAlpha = 255;
                        pblend.AlphaFormat = 1;
                        UpdateLayeredWindow(_hwndStartBtn, hdcDst, NULL, &psize, hdcSrca, &pptSrc,
                            0, &pblend, ULW_ALPHA);
                    }
                    ReleaseDC(0, hdcDst);
                    SelectObject(hdcSrca, h);
                    DeleteObject(ho);
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
        SetWindowRgn(_hwndStartBtn, hRgn, 1);
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

BOOL CStartButton::InitBackgroundBitmap()
{
    // NOTICE: Do like XP (shellbrd images are gone)
    return TRUE;
}

void CStartButton::InitTheme()
{
    pszCurrentThemeName = _GetCurrentThemeName();
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

void CStartButton::RecalcSize()
{
    if (!_hTheme)
    {
        RECT rc;
        GetClientRect(v_hwndTray, &rc);
        LPCWSTR p_windowName = &WindowName;
        if (rc.right >= _lWidth)
        {
            p_windowName = _pszWindowName;
        }

        SetWindowTextW(_hwndStartBtn, p_windowName);

        int height = _pStartButtonSite->GetStartButtonMinHeight();
        if (!height && _lHeight >= 0)
        {
            height = _lHeight;
        }
        _lHeight = height;
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
    if (_hTheme && _GetCurrentThemeName() != pszCurrentThemeName)
    {
        pszCurrentThemeName = _GetCurrentThemeName();
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

void CStartButton::_ExploreCommonStartMenu(BOOL bExplore)
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

    if (g_nReplaceMe1 != 1)
    {
        if (g_nReplaceMe1 == 3 && RECTHEIGHT(rc) < _lHeight)
        {
            return L"StartBottom";
        }
        return L"StartMiddle";
    }

    if (RECTHEIGHT(rc) >= _lHeight)
    {
        return L"StartMiddle";
    }

    return pszResult;
}

void CStartButton::_HandleDestroy()
{
    _nBackgroundBitmapInitialized = 0;
    _DestroyStartButtonBalloon();

    if (_nShellBrdBitmap1001)
    {
        DeleteObject(_nShellBrdBitmap1001);
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
        pszCurrentThemeName = NULL;
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
