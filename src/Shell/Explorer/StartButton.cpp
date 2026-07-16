#include "pch.h"

#include "StartButton.h"
#include "SHFusion.h"
#include "tray.h"
#include "Util.h"
#include "DeskHost.h"

CStartButton::CStartButton(IStartButtonSite* pStartButtonSite)
    : field_4C(true)
    , _pStartButtonSite(pStartButtonSite)
{
}

HRESULT CStartButton::QueryInterface(REFIID riid, void** ppvObject)
{
    static const QITAB qit[] =
    {
        QITABENT(CStartButton, IStartButton),
        QITABENT(CStartButton, IServiceProvider),
        {},
    };
    return QISearch(this, qit, riid, ppvObject);
}

ULONG CStartButton::AddRef()
{
    return 2;
}

ULONG CStartButton::Release()
{
    return 1;
}

HRESULT CStartButton::SetFocusToStartButton()
{
    SetFocus(_hwndStart);
    return S_OK;
}

#include <ntstatus.h>

enum LUAUSERTYPE
{
    ADMINTOKEN = 0x0,
    SPLITADMINTOKEN = 0x1,
    LUATOKEN = 0x2,
    LUAUSERTYPE_MAX_NON_UIA = 0x3,

    ADMINTOKEN_UIA = 0x10,
    SPLITADMINTOKEN_UIA = 0x11,
    LUATOKEN_UIA = 0x12,
    LUAUSERTYPE_MAX = 0x13,
};

extern "C" NTSTATUS NTAPI NtQueryInformationToken(
    HANDLE TokenHandle, TOKEN_INFORMATION_CLASS TokenInformationClass, PVOID TokenInformation,
    ULONG TokenInformationLength, PULONG ReturnLength);

EXTERN_C NTSTATUS LUAIsElevatedToken(HANDLE hToken, BOOL* pfElevatedToken, BOOL* pfSplitToken)
{
    *pfSplitToken = TRUE;
    *pfElevatedToken = FALSE;

    TOKEN_ELEVATION_TYPE ElevationType = TokenElevationTypeDefault;
    ULONG cbSize;
    NTSTATUS hr = NtQueryInformationToken(hToken, TokenElevationType, &ElevationType, sizeof(ElevationType), &cbSize);
    if (hr >= 0)
    {
        TOKEN_ELEVATION Elevation;
        if (ElevationType == TokenElevationTypeFull)
        {
            *pfElevatedToken = TRUE;
        }
        else if (ElevationType == TokenElevationTypeDefault)
        {
            *pfSplitToken = FALSE;

            hr = NtQueryInformationToken(hToken, TokenElevation, &Elevation, sizeof(Elevation), &cbSize);
            if (hr >= 0 && Elevation.TokenIsElevated)
            {
                *pfElevatedToken = TRUE;
            }
        }
    }

    return hr;
}

EXTERN_C NTSTATUS LUAIsUIAToken(HANDLE hToken, BOOL* pfUIAToken)
{
    *pfUIAToken = FALSE;

    DWORD dwValue;
    DWORD cbSize = sizeof(DWORD);
    NTSTATUS hr = NtQueryInformationToken(hToken, TokenUIAccess, &dwValue, sizeof(dwValue), &cbSize);
    if (SUCCEEDED(hr))
    {
        *pfUIAToken = dwValue != 0;
    }

    return hr;
}

#define NtCurrentProcess() ((HANDLE)(LONG_PTR)-1)
#define NtCurrentThread() ((HANDLE)(LONG_PTR)-2)

EXTERN_C
NTSYSCALLAPI
NTSTATUS
NTAPI
NtOpenThreadToken(
    _In_ HANDLE ThreadHandle,
    _In_ ACCESS_MASK DesiredAccess,
    _In_ BOOLEAN OpenAsSelf,
    _Out_ PHANDLE TokenHandle
    );

EXTERN_C
NTSYSCALLAPI
NTSTATUS
NTAPI
NtOpenProcessToken(
    _In_ HANDLE ProcessHandle,
    _In_ ACCESS_MASK DesiredAccess,
    _Out_ PHANDLE TokenHandle
    );

EXTERN_C NTSTATUS LUAGetUserType(HANDLE hToken, LUAUSERTYPE* pLuaUserType)
{
    NTSTATUS hr = STATUS_SUCCESS;
    HANDLE hToClose = nullptr;

    if (!hToken)
    {
        hr = NtOpenThreadToken(NtCurrentThread(), TOKEN_QUERY, FALSE, &hToken);
        if (hr == 0xC000007C)
        {
            hr = NtOpenProcessToken(NtCurrentProcess(), TOKEN_QUERY, &hToken);
        }
        if (hr >= 0)
        {
            hToClose = hToken;
        }
        else
        {
            return hr;
        }
    }

    if (hr >= 0)
    {
        BOOL fElevatedToken;
        BOOL fSplitToken;
        hr = LUAIsElevatedToken(hToken, &fElevatedToken, &fSplitToken);
        if (hr >= 0)
        {
            *pLuaUserType = fElevatedToken ? ADMINTOKEN : fSplitToken ? SPLITADMINTOKEN : LUATOKEN;

            BOOL fUIAToken;
            hr = LUAIsUIAToken(hToken, &fUIAToken);
            if (hr >= 0 && fUIAToken)
            {
                LUAUSERTYPE luaUserType = *pLuaUserType;
                if (*pLuaUserType < ADMINTOKEN_UIA || luaUserType >= LUAUSERTYPE_MAX)
                {
                    luaUserType = (LUAUSERTYPE)(luaUserType + 0x10);
                }
                *pLuaUserType = luaUserType;
            }
        }
    }

    if (hToClose)
        NtClose(hToClose);
    return hr;
}

HRESULT CStartButton::OnContextMenu(HWND hwnd, LPARAM lParam)
{
    field_40 = 1;
    _pStartButtonSite->HandleFullScreenApp(nullptr);
    SetForegroundWindow(hwnd);

    ITEMIDLIST* pidlStart = SHCloneSpecialIDList(hwnd, CSIDL_STARTMENU, TRUE);
    const ITEMIDLIST* pidlLast;
    IShellFolder* psf;
    if (SUCCEEDED(SHBindToParent(pidlStart, IID_PPV_ARGS(&psf), &pidlLast)))
    {
        HMENU hmenu = CreatePopupMenu();
        if (hmenu)
        {
            IContextMenu* pcm;
            if (SUCCEEDED(psf->GetUIObjectOf(hwnd, 1, &pidlLast, IID_IContextMenu, nullptr, (void**)&pcm)))
            {
                if (SUCCEEDED(pcm->QueryContextMenu(hmenu, 0, 2, 32751, CMF_VERBSONLY)))
                {
                    WCHAR szCommon[260];
                    LoadStringW(g_hinstCabinet, 720, szCommon, ARRAYSIZE(szCommon));
                    AppendMenuW(hmenu, 0, 32755, szCommon);

                    if (!SHRestricted(REST_NOCOMMONGROUPS))
                    {
                        BOOL fAddCommon = SHGetFolderPathW(nullptr, CSIDL_COMMON_STARTMENU, nullptr, 0, szCommon) == S_OK;
                        if (fAddCommon)
                        {
                            LUAUSERTYPE luaUserType;
                            if (LUAGetUserType(nullptr, &luaUserType) == STATUS_SUCCESS
                                && (luaUserType == ADMINTOKEN
                                || luaUserType == SPLITADMINTOKEN
                                || luaUserType == ADMINTOKEN_UIA
                                || luaUserType == SPLITADMINTOKEN_UIA))
                            {
                                fAddCommon = TRUE;
                            }
                            if (fAddCommon)
                            {
                                AppendMenuW(hmenu, MFT_SEPARATOR, 0, nullptr);
                                LoadStringW(g_hinstCabinet, 718, szCommon, ARRAYSIZE(szCommon));
                                AppendMenuW(hmenu, MFT_STRING, 32752, szCommon);
                                LoadStringW(g_hinstCabinet, 719, szCommon, ARRAYSIZE(szCommon));
                                AppendMenuW(hmenu, MFT_STRING, 32753, szCommon);
                            }
                        }
                    }

                    int idCmd;
                    if (lParam == -1)
                    {
                        idCmd = TrackMenu(hmenu);
                    }
                    else
                    {
                        _pStartButtonSite->EnableTooltips(FALSE);
                        UINT uFlags = TPM_RIGHTBUTTON | TPM_RETURNCMD;
                        if (IsBiDiLocalizedSystem())
                        {
                            uFlags |= TPM_LAYOUTRTL;
                        }
                        idCmd = TrackPopupMenu(hmenu, uFlags, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), 0, hwnd, nullptr);
                        _pStartButtonSite->EnableTooltips(TRUE);
                    }
                    if (idCmd)
                    {
                        switch (idCmd)
                        {
                            case 32752:
                                _ExploreCommonStartMenu(FALSE);
                                break;
                            case 32753:
                                _ExploreCommonStartMenu(TRUE);
                                break;
                            case 32755:
                                Tray_DoProperties(2);
                                break;
                            default:
                                WCHAR szVerb[260];
                                ContextMenu_GetCommandStringVerb(pcm, idCmd - 2, szVerb, ARRAYSIZE(szVerb));
                                if (StrCmpICW(szVerb, L"find"))
                                {
                                    CMINVOKECOMMANDINFOEX ici = {};
                                    ici.cbSize = sizeof(ici);
                                    ici.hwnd = hwnd;
                                    ici.lpVerb = (const CHAR*)MAKEINTRESOURCE(idCmd - 2);
                                    ici.nShow = SW_SHOWNORMAL;

                                    CHAR szPathAnsi[260];
                                    WCHAR szPath[260];
                                    SHGetPathFromIDListA(pidlStart, szPathAnsi);
                                    SHGetPathFromIDListW(pidlStart, szPath);

                                    ici.fMask |= CMIC_MASK_UNICODE;
                                    ici.lpDirectory = szPathAnsi;
                                    ici.lpDirectoryW = szPath;
                                    pcm->InvokeCommand(reinterpret_cast<CMINVOKECOMMANDINFO*>(&ici));
                                }
                                else
                                {
                                    ITEMIDLIST* pidlSearchHome;
                                    if (SUCCEEDED(SHGetKnownFolderIDList(
                                        FOLDERID_SearchHome, KF_FLAG_DEFAULT, nullptr, &pidlSearchHome)))
                                    {
                                        SHELLEXECUTEINFOW shei = {};
                                        shei.cbSize = sizeof(shei);
                                        shei.fMask = SEE_MASK_IDLIST;
                                        shei.lpVerb = L"open";
                                        shei.nShow = SW_SHOWNORMAL;
                                        shei.lpIDList = pidlSearchHome;
                                        ShellExecuteExW(&shei);
                                        ILFree(pidlSearchHome);
                                    }
                                }
                                break;
                        }
                    }
                }
                pcm->Release();
            }
            DestroyMenu(hmenu);
        }
        psf->Release();
    }

    ILFree(pidlStart);
    field_40 = 0;
    return S_OK;
}

HRESULT CStartButton::CreateStartButtonBalloon(UINT idsTitle, UINT idsMessage)
{
    if (!_hwndStartBalloon)
    {
        _hwndStartBalloon = SHFusionCreateWindow(
            L"tooltips_class32", nullptr, TTS_ALWAYSTIP | TTS_NOPREFIX | TTS_BALLOON | WS_POPUP,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _hwndStart, nullptr, g_hinstCabinet, nullptr);
        if (_hwndStartBalloon)
        {
            SendMessageW(_hwndStartBalloon, CCM_SETVERSION, COMCTL32_VERSION, 0);
            SendMessageW(_hwndStartBalloon, TTM_SETMAXTIPWIDTH, 0, 300);
            SendMessageW(_hwndStartBalloon, TTM_SETWINDOWTHEME, 0, reinterpret_cast<LPARAM>(L"TaskBar"));
            SetPropW(_hwndStartBalloon, L"StartMenuBalloonTip", reinterpret_cast<HANDLE>(3));
        }
    }

    if (_hwndStartBalloon)
    {
        WCHAR szTip[260];
        if (LoadStringW(g_hinstCabinet, idsMessage, szTip, ARRAYSIZE(szTip)))
        {
            TTTOOLINFOW ti = {};
            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_IDISHWND | TTF_TRACK | TTF_TRANSPARENT;
            ti.hwnd = _hwndStart;
            ti.uId = reinterpret_cast<UINT_PTR>(ti.hwnd);
            SendMessageW(_hwndStartBalloon, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&ti));
            SendMessageW(_hwndStartBalloon, TTM_TRACKACTIVATE, 0, 0);

            ti.lpszText = szTip;
            SendMessageW(_hwndStartBalloon, TTM_UPDATETIPTEXTW, 0, reinterpret_cast<LPARAM>(&ti));
            if (LoadStringW(g_hinstCabinet, idsTitle, szTip, ARRAYSIZE(szTip)))
            {
                SendMessageW(_hwndStartBalloon, TTM_SETTITLEW, TTI_INFO, reinterpret_cast<LPARAM>(szTip));
            }

            RECT rc;
            GetWindowRect(_hwndStart, &rc);
            SendMessageW(_hwndStartBalloon, TTM_TRACKPOSITION, 0, MAKELPARAM((rc.left + rc.right) / 2, rc.top));
            SendMessageW(_hwndStartBalloon, TTM_TRACKACTIVATE, TRUE, reinterpret_cast<LPARAM>(&ti));
            SetTimer(_hwndStart, 1, 10000u, nullptr);
        }
    }

    return S_OK;
}

HWND CStartButton::CreateStartButton(HWND hwndParent)
{
    // dwStyle = 0x400 | 0x800 | 0x4000000 | 0x80000000 = 0x84000C00
    DWORD dwStyleEx = GetWindowLongPtr(hwndParent, GWL_EXSTYLE);
    _hwndStart = SHFusionCreateWindowEx(
        dwStyleEx & (WS_EX_LAYOUTRTL | WS_EX_RTLREADING | WS_EX_RIGHT) | (WS_EX_LAYERED | WS_EX_TOOLWINDOW),
        WC_BUTTON, nullptr, WS_POPUP | WS_CLIPSIBLINGS | BS_VCENTER, 0, 0, 1, 1, hwndParent, nullptr, g_hinstCabinet,
        nullptr);
    if (_hwndStart)
    {
        SetProp(_hwndStart, TEXT("StartButtonTag"), (HANDLE)0x130);
        SendMessage(_hwndStart, CCM_DPISCALE, TRUE, 0);

        SetWindowSubclass(_hwndStart, s_StartButtonSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));

        LoadString(g_hinstCabinet, IDS_STARTCLASSIC, _szWindowName, ARRAYSIZE(_szWindowName));
        SetWindowText(_hwndStart, _szWindowName);
    }
    return _hwndStart;
}

HRESULT CStartButton::SetStartPaneActive(BOOL fActive)
{
    if (fActive)
    {
        _uStartButtonState = 1;
    }
    else if (_uStartButtonState != 2)
    {
        _uStartButtonState = 0;

        UnlockStartPane();
    }
    return S_OK;
}

HRESULT CStartButton::OnStartMenuDismissed()
{
    _pStartButtonSite->OnStartMenuDismissed();
    return S_OK;
}

HRESULT CStartButton::UnlockStartPane()
{
    if (_fForegroundLocked)
    {
        _fForegroundLocked = FALSE;
        LockSetForegroundWindow(LSFW_UNLOCK);
    }
    return S_OK;
}

HRESULT CStartButton::LockStartPane()
{
    if (!_fForegroundLocked)
    {
        _fForegroundLocked = TRUE;
        LockSetForegroundWindow(LSFW_LOCK);
    }
    return S_OK;
}

HRESULT CStartButton::GetPopupPosition(DWORD* pdwPos)
{
    if (!_pStartButtonSite)
        return E_FAIL;

    UINT uStuckPlace = _pStartButtonSite->GetStartMenuStuckPlace();
    switch (uStuckPlace)
    {
        case STICK_LEFT:
            *pdwPos = MPPF_LEFT;
            break;
        case STICK_TOP:
            *pdwPos = MPPF_TOP;
            break;
        case STICK_RIGHT:
            *pdwPos = MPPF_RIGHT;
            break;
        default:
            *pdwPos = MPPF_BOTTOM;
            break;
    }

    return S_OK;
}

HRESULT CStartButton::GetWindow(HWND* phwndStart)
{
    *phwndStart = _hwndStart;
    return S_OK;
}

HRESULT CStartButton::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
{
    if (IsEqualGUID(guidService, __uuidof(IStartButton)))
    {
        return QueryInterface(riid, ppvObject);
    }
    return E_FAIL;
}

void CStartButton::BuildStartMenu()
{
    HRESULT hr;

    CloseStartMenu();
    _pStartButtonSite->PurgeRebuildRequests();
    DestroyStartMenu();

    if (Tray_StartPanelEnabled())
    {
        // SHTracePerfSQMSetValueImpl(&ShellTraceId_Explorer_StartMenu_Mode, 58, 0);
        void* pvStartPane;
        hr = DesktopV2_Create(&_pmpStartPane, &_pmbStartPane, &pvStartPane, &_punkSite, v_hwndTray);
        IUnknown_SetSite(_punkSite, static_cast<IServiceProvider*>(this));
        DesktopV2_Build(pvStartPane);
    }
    else
    {
        // SHTracePerfSQMSetValueImpl(&ShellTraceId_Explorer_StartMenu_Mode, 58, 1);
        hr = StartMenuHost_Create(&_pmpStartMenu, &_pmbStartMenu, &_punkSite);
        if (SUCCEEDED(hr))
        {
            IUnknown_SetSite(_punkSite, static_cast<IServiceProvider*>(this));

            HWND hwnd;
            if (SUCCEEDED(IUnknown_GetWindow(_pmpStartMenu, &hwnd)))
            {
                SetWindowSubclass(hwnd, s_StartMenuSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));
            }

            IBanneredBar* pbb;
            hr = _pmpStartMenu->QueryInterface(IID_PPV_ARGS(&pbb));
            if (SUCCEEDED(hr))
            {
                if (_hbmpStartBkg)
                {
                    pbb->SetBitmap(_hbmpStartBkg);
                }
                if (_pStartButtonSite->ShouldUseSmallIcons())
                {
                    pbb->SetIconSize(BMICON_SMALL);
                }
                else
                {
                    pbb->SetIconSize(BMICON_LARGE);
                }
                pbb->Release();
            }
        }
    }

    if (FAILED(hr))
    {
        // TraceMsg(TF_ERROR, "Could not create StartMenu");
    }
}

void CStartButton::CloseStartMenu()
{
    if (_pmpStartMenu)
    {
        _pmpStartMenu->OnSelect(MPOS_FULLCANCEL);
    }
    if (_pmpStartPane)
    {
        _pmpStartPane->OnSelect(MPOS_FULLCANCEL);
    }
}

void CStartButton::DestroyStartMenu()
{
    IUnknown_SetSite(_punkSite, NULL);
    ATOMICRELEASE(_punkSite);

    IUnknown_SetSite(_pmpStartMenu, NULL);
    ATOMICRELEASE(_pmpStartMenu);
    ATOMICRELEASE(_pmbStartMenu);

    IUnknown_SetSite(_pmpStartPane, NULL);
    ATOMICRELEASE(_pmpStartPane);
    ATOMICRELEASE(_pmbStartPane);
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

    IMenuPopup** ppmpToDisplay = &_pmpStartMenu;
    if (_pmpStartPane)
    {
        ppmpToDisplay = &_pmpStartPane;
    }

    if (!*ppmpToDisplay)
    {
        // TraceMsg(TF_WARNING, "e.tbm: Rebuilding Start Menu");
        BuildStartMenu();
    }

    // Recalculate the position of the taskbar and start button:
    bool fMoveTaskbar = false;

	DWORD_PTR exStyle = GetWindowLongPtr(v_hwndTray, GWL_EXSTYLE);
    if ((exStyle & WS_EX_TOPMOST) == 0
        && _pszThemeName == TEXT("StartBottom")
        || _pszThemeName == TEXT("StartTop"))
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

void CStartButton::DrawStartButton(int iStateId, bool bRepaint)
{
    POINT pptDst;
    HRGN hRgn;
    BOOL fCalculated = CStartButton::_CalcStartButtonPos(&pptDst, &hRgn);

    if (bRepaint)
    {
        HDC hdc = GetDC(_hwndStart);
        if (hdc)
        {
            HDC hdcMem = CreateCompatibleDC(hdc);
            if (hdcMem)
            {
                RECT rcStart = { 0, 0, _sizeStart.cx, _sizeStart.cy };

                BITMAPINFO bmi = { 0 };
                bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                bmi.bmiHeader.biWidth = rcStart.right;
                bmi.bmiHeader.biHeight = -rcStart.bottom;
                bmi.bmiHeader.biPlanes = 1;
                bmi.bmiHeader.biBitCount = 32;
                bmi.bmiHeader.biCompression = BI_RGB;

                HBITMAP hbmStartButton = CreateDIBSection(hdc, &bmi, 0, 0, 0, 0);
                if (hbmStartButton)
                {
                    HBITMAP hbmMem = (HBITMAP)SelectObject(hdcMem, hbmStartButton);
                    HDC hdcDst = GetDC(NULL);

                    POINT ptStart = { 0, 0 };
                    if (!_hTheme)
                    {
                        SendMessage(_hwndStart, WM_PRINTCLIENT, (WPARAM)hdcMem, PRF_CLIENT);
                        UpdateLayeredWindow(_hwndStart, hdcDst, 0, &_sizeStart, hdcMem, &ptStart, 0, 0, ULW_OPAQUE);
                    }
                    else
                    {
                        SHFillRectClr(hdcMem, &rcStart, 0);
                        DrawThemeBackground(_hTheme, hdcMem, 1, iStateId, &rcStart, 0);

                        BLENDFUNCTION bf;
                        bf.BlendOp = 0;
                        bf.BlendFlags = 0;
                        bf.SourceConstantAlpha = 255;
                        bf.AlphaFormat = 1;
                        UpdateLayeredWindow(_hwndStart, hdcDst, 0, &_sizeStart, hdcMem, &ptStart, 0, &bf, ULW_ALPHA);
                    }
                    ReleaseDC(NULL, hdcDst);
                    SelectObject(hdcMem, hbmMem);
                    DeleteObject(hbmStartButton);
                }
                DeleteDC(hdcMem);
            }
            ReleaseDC(_hwndStart, hdc);
        }
    }
    UpdateLayeredWindow(_hwndStart, nullptr, &pptDst, nullptr, nullptr, nullptr, 0, nullptr, 0);

    if (fCalculated)
        SetWindowRgn(_hwndStart, hRgn, TRUE);
}

void CStartButton::ExecRefresh()
{
    if (_pmbStartMenu)
    {
        IUnknown_Exec(_pmbStartMenu, &CLSID_MenuBand, MBANDCID_REFRESH, 0, nullptr, nullptr);
    }
    else if (_pmpStartPane)
    {
        IUnknown_Exec(_pmpStartPane, &CLSID_MenuBand, MBANDCID_REFRESH, 0, nullptr, nullptr);
    }
}

void CStartButton::ForceButtonUp()
{
    if (!_fAllowUp)
    {
        _fAllowUp = TRUE;

        MSG msg;
        PeekMessageW(&msg, _hwndStart, WM_LBUTTONDOWN, WM_LBUTTONDOWN, PM_REMOVE);
        PeekMessageW(&msg, _hwndStart, WM_LBUTTONDOWN, WM_LBUTTONDOWN, PM_REMOVE);
        SendMessageW(_hwndStart, BM_SETSTATE, FALSE, 0);
        PeekMessageW(&msg, _hwndStart, WM_LBUTTONDBLCLK, WM_LBUTTONDBLCLK, PM_REMOVE);
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
        GetThemePartSize(hTheme, hdc, BP_PUSHBUTTON, CBS_UNCHECKEDNORMAL, NULL, TS_TRUE, &_sizeStart);

        int iDpi = GetDeviceCaps(hdc, LOGPIXELSX);

        // XXX(isabella): Looks to be DPI resolution?
        if (iDpi < 120)
        {
            field_C = 8;
        }
        else if (iDpi < 144)
        {
            field_C = 9;
        }
        else if (iDpi < 192)
        {
            field_C = 11;
        }
        else
        {
            field_C = 14;
        }

        ReleaseDC(_hwndStart, hdc);
    }
    else // Classic theme:
    {
        int idbStart = SHGetCurColorRes() <= 8 ? IDB_START16 : IDB_STARTCLASSIC;

        // @MOD (isabella): Bitmap in Explorer instead of ShellBrd
        HBITMAP hbmFlag = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(idbStart));
        if (hbmFlag)
        {
            BITMAP bm;
            if (GetObject(hbmFlag, sizeof(BITMAP), &bm))
            {
                BUTTON_IMAGELIST biml = { 0 };
                if (_himlStartFlag)
                {
                    // Clean up any previously-existing image list:
                    ImageList_Destroy(_himlStartFlag);
                }


                DWORD dwFlags = ILC_COLOR32;
                HBITMAP hbmFlagMask = NULL;
                if (idbStart == IDB_START16)
                {
                    dwFlags = ILC_COLOR8 | ILC_MASK;

                    // @MOD (isabella): Bitmap in Explorer instead of ShellBrd
                    hbmFlagMask = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_START16MASK));
                }

                if ((GetWindowLongPtr(_hwndStart, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0)
                {
                    dwFlags |= ILC_MIRROR;
                }

                biml.himl = _himlStartFlag = ImageList_Create(bm.bmWidth, bm.bmHeight, dwFlags, 1, 1);
                ImageList_Add(_himlStartFlag, hbmFlag, hbmFlagMask);

                if (hbmFlagMask)
                {
                    DeleteObject(hbmFlagMask);
                }

                biml.uAlign = BUTTON_IMAGELIST_ALIGN_LEFT;

                Button_SetImageList(_hwndStart, &biml);
            }
            DeleteObject(hbmFlag);
        }

        if (_hStartFont)
        {
            DeleteObject(_hStartFont);
        }

        _hStartFont = _CreateStartFont();

        SendMessage(_hwndStart, WM_SETFONT, (WPARAM)_hStartFont, TRUE);

        // Recalculate the size:
        _sizeStart = { 0 };
        SendMessage(_hwndStart, BCM_GETIDEALSIZE, 0, (LPARAM)&_sizeStart);
    }
}

BOOL CStartButton::InitBackgroundBitmap()
{
    _fBackgroundBitmapInitialized = TRUE;

    // @MOD (isabella): Vista loads this bitmap from ShellBrd, but we store the bitmap in our own
    // module. Vista's original code is such (link against WinBrand.dll):
    //    _hbmpStartBkg = BrandingLoadBitmap(L"Shellbrd", 1001);
    _hbmpStartBkg = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_CLASSICSTARTBKG));

    return _hbmpStartBkg != NULL;
}

void CStartButton::InitTheme()
{
    _pszThemeName = _GetCurrentThemeName();
    SetWindowTheme(_hwndStart, _pszThemeName, nullptr);
}

BOOL CStartButton::IsButtonPushed()
{
    return SendMessageW(_hwndStart, BM_GETSTATE, 0, 0) & BST_PUSHED;
}

HRESULT CStartButton::IsMenuMessage(MSG *pmsg)
{
    return (!_pmbStartPane || _pmbStartPane->IsMenuMessage(pmsg))
        && (!_pmbStartMenu || _pmbStartMenu->IsMenuMessage(pmsg));
}

BOOL CStartButton::IsPopupMenuVisible()
{
    HWND hwnd;
    return SUCCEEDED(IUnknown_GetWindow(_pmpStartMenu, &hwnd)) && IsWindowVisible(hwnd)
        || SUCCEEDED(IUnknown_GetWindow(_pmpStartPane, &hwnd)) && IsWindowVisible(hwnd);
}

BOOL CStartButton::_CalcStartButtonPos(POINT *pPoint, HRGN *phRgn)
{
    RECT rcTrayWnd;
    GetWindowRect(v_hwndTray, &rcTrayWnd);

    LONG cyFrameHalf = g_cyFrame / 2;

    if (_pszThemeName == L"StartTop")
    {
        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _sizeStart.cx : rcTrayWnd.left;

        if (rcTrayWnd.bottom <= cyFrameHalf)
            pPoint->y = rcTrayWnd.top - _sizeStart.cy - cyFrameHalf;
        else
            pPoint->y = rcTrayWnd.bottom + field_C - _sizeStart.cy;
    }
    else if (_pszThemeName == L"StartBottom")
    {
        RECT rc;

        pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _sizeStart.cx : rcTrayWnd.left;

        HMONITOR hMon = MonitorFromRect(&rcTrayWnd, MONITOR_DEFAULTTONEAREST);
        GetMonitorRects(hMon, &rc, FALSE);

        if (rc.bottom - rcTrayWnd.top <= cyFrameHalf)
            pPoint->y = cyFrameHalf + rcTrayWnd.top;
        else
            pPoint->y = rcTrayWnd.top - field_C;
    }
    else if (_hTheme)
    {
        int height;

        if (STUCK_HORIZONTAL(c_tray._uStuckPlace))
        {
            pPoint->x = IsBiDiLocalizedSystem() ? rcTrayWnd.right - _sizeStart.cx : rcTrayWnd.left;
            height = rcTrayWnd.bottom - _sizeStart.cy - rcTrayWnd.top;
        }
        else
        {
            pPoint->x = rcTrayWnd.left + (rcTrayWnd.right - _sizeStart.cx - rcTrayWnd.left) / 2;
            height = g_cyTabSpace;
        }

        pPoint->y = rcTrayWnd.top +  height / 2;
    }
    else
    {
        int cyDlgFrame = GetSystemMetrics(SM_CYDLGFRAME);
        int cyBorder = GetSystemMetrics(SM_CYBORDER);
        if (IsBiDiLocalizedSystem() && (c_tray._uStuckPlace == 1 || c_tray._uStuckPlace == 3) != 0)
            pPoint->x = rcTrayWnd.right - _sizeStart.cx - cyBorder - cyDlgFrame;
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
                SetWindowRgn(_hwndStart, NULL, TRUE);
            }
        }
        else
        {
            c_tray.GetStuckMonitorRect(&rc);

            RECT rcSrc;
            rcSrc.left = pPoint->x;
            rcSrc.bottom = pPoint->y + _sizeStart.cy;
            rcSrc.top = pPoint->y;
            rcSrc.right = pPoint->x + _sizeStart.cx;
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
    return fRes;
}

void CStartButton::RecalcSize()
{
    if (!_hTheme)
    {
        RECT rc;
        GetClientRect(v_hwndTray, &rc);
        LPCWSTR pszWindowName = L"";
        if (rc.right >= _sizeStart.cx)
        {
            pszWindowName = _szWindowName;
        }

        SetWindowTextW(_hwndStart, pszWindowName);

        int height = _pStartButtonSite->GetStartButtonMinHeight();
        if (!height && _sizeStart.cy >= 0)
        {
            height = _sizeStart.cy;
        }
        _sizeStart.cy = height;
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
        SendMessage(_hwndStartBalloon, 0x412u, 0, MAKELONG(xCoordinate, yCoordinate));
    }
}

void CStartButton::StartButtonReset()
{
    GetSizeAndFont(_hTheme);
    RecalcSize();
    c_tray.UpdateStuckRect();
}

int CStartButton::TrackMenu(HMENU hMenu)
{
    TPMPARAMS tpm;
    int iRet;

    tpm.cbSize = sizeof(tpm);
    GetClientRect(_hwndStart, &tpm.rcExclude);

    RECT rcClient;
    GetClientRect(GetParent(_hwndStart), &rcClient);
    tpm.rcExclude.bottom = min(tpm.rcExclude.bottom, rcClient.bottom);

    MapWindowRect(_hwndStart, NULL, &tpm.rcExclude);

    UINT uFlag = TPM_RETURNCMD | TPM_RIGHTBUTTON;
    if (IS_BIDI_LOCALIZED_SYSTEM())
    {
        uFlag |= TPM_LAYOUTRTL;
    }

    _pStartButtonSite->EnableTooltips(FALSE);
    iRet = TrackPopupMenuEx(hMenu, uFlag, tpm.rcExclude.left, tpm.rcExclude.bottom, _hwndStart, &tpm);
    _pStartButtonSite->EnableTooltips(TRUE);

    return iRet;
}

HRESULT CStartButton::TranslateMenuMessage(MSG *pmsg, LRESULT *plRet)
{
    return (!_pmbStartMenu || _pmbStartMenu->TranslateMenuMessage(pmsg, plRet))
        && (!_pmbStartPane || _pmbStartPane->TranslateMenuMessage(pmsg, plRet));
}

void CStartButton::UpdateStartButton(bool a2)
{
    if (_hTheme && _GetCurrentThemeName() != _pszThemeName)
    {
        _pszThemeName = _GetCurrentThemeName();
        SetWindowTheme(_hwndStart, _GetCurrentThemeName(), nullptr);
    }
    else
    {
        DrawStartButton(PBS_NORMAL, a2);
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
    SHSetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED,
        TEXT("StartButtonBalloonTip"), REG_DWORD, (BYTE*)&dwData, sizeof(dwData));
}

LRESULT CStartButton::OnMouseClick(HWND hWndTo, LPARAM lParam)
{
    LRESULT lRes = S_OK;
    if (_hwndStartBalloon)
    {
        RECT rcBalloon;
        GetWindowRect(_hwndStartBalloon, &rcBalloon);
        MapWindowRect(NULL, hWndTo, &rcBalloon);
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
        rcStart.left = rcStuck.right - _sizeStart.cx;
    else
        rcStart.left = rcStuck.left;

    if (uStuckPlace == STICK_TOP || uStuckPlace == STICK_BOTTOM)
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

void CStartButton::_ExploreCommonStartMenu(BOOL bExplore)
{
    LPITEMIDLIST pidl;
    if (SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_COMMON_STARTMENU, NULL, KF_FLAG_DEFAULT, &pidl)))
    {
        SHELLEXECUTEINFO sei = { 0 };

		sei.cbSize       = sizeof(sei);
        sei.fMask        = SEE_MASK_IDLIST | SEE_MASK_ASYNCOK;
        sei.lpVerb       = bExplore ? TEXT("explore") : TEXT("open");
        sei.nShow        = SW_SHOWNORMAL;
        sei.lpIDList     = pidl;
        ShellExecuteEx(&sei);
        ILFree(pidl);
    }
}

const WCHAR* CStartButton::_GetCurrentThemeName()
{
    RECT rc;
    GetWindowRect(v_hwndTray, &rc);

    if (c_tray._uStuckPlace == STICK_BOTTOM && RECTHEIGHT(rc) < _sizeStart.cy)
    {
        return L"StartBottom";
    }
    else if (c_tray._uStuckPlace == STICK_TOP && RECTHEIGHT(rc) < _sizeStart.cy)
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
    if (_himlStartFlag)
    {
        ImageList_Destroy(_himlStartFlag);
    }

    RemovePropW(_hwndStart, L"StartButtonTag");
}

void CStartButton::_OnSettingChanged(UINT a2)
{
    if (!_hTheme && a2 != 0x2F)
    {
        bool v3 = !field_4C;
        if (field_4C)
        {
            PostMessage(_hwndStart, 0x31Au, 0, 0);
            v3 = !field_4C;
        }
        field_4C = v3;
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
        _pszThemeName = nullptr;
        if (!field_4C)
        {
            StartButtonReset();
            DrawStartButton(PBS_NORMAL, true);
            bThemeApplied = true;
        }
        else
        {
            PostMessage(_hwndStart, 0x31Au, 0, 0);
        }
        field_4C = !field_4C;
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
                field_14 = TRUE;
                LRESULT lRet = DefSubclassProc(hWnd, BM_SETSTATE, wParam, lParam);
                DrawStartButton(PBS_PRESSED, true);
                _pStartButtonSite->StartButtonClicked();
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
                return DefWindowProc(hWnd, uMsg, 0, lParam);
            }
            else
            {

                // Nope, Forward it on.
                field_14 = FALSE;
                LRESULT lr = DefSubclassProc(hWnd, BM_SETSTATE, 0, lParam);
                DrawStartButton(PBS_NORMAL, true);
                _pStartButtonSite->EnableTooltips(TRUE);
                _uDown = 0;
                return lr;
            }
        }
    }
    else
    {
        // all vista reimpl'd code

        if (_uStartButtonState == 2 && uMsg >= WM_MOUSEFIRST && uMsg <= WM_MOUSELAST)
        {
            _uStartButtonState = 0; // SBSM_NORMAL
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
                    _CalcStartButtonPos(&pt, NULL);

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
                if (!field_10 && _hTheme && !_uDown)
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
                if (_hTheme && !field_10 && !field_14)
                {
                    DrawStartButton(PBS_HOT, true);

                    TRACKMOUSEEVENT tme;
                    tme.cbSize = sizeof(TRACKMOUSEEVENT);
                    tme.dwFlags = TME_LEAVE;
                    tme.hwndTrack = _hwndStart;
                    tme.dwHoverTime = 0;

                    TrackMouseEvent(&tme);

                    field_10 = 1;
                }
                break;
            }

            case WM_MOUSELEAVE:
            {
                if (_hTheme && !field_14 && !c_tray.IsMouseOverStartButton())
                {
                    DrawStartButton(PBS_NORMAL, true);
                }
                field_10 = 0;
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
                    SendMessage(GetAncestor(hWnd, GA_ROOTOWNER), WM_UPDATEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
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
                return MAKELRESULT(_sizeStart.cx, _sizeStart.cy);
            }

            case WM_KEYDOWN:
            {
                SendMessage(GetAncestor(hWnd, GA_ROOTOWNER), WM_UPDATEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEACCEL), 0);

                if (wParam == VK_RETURN)
                {
                    PostMessage(GetAncestor(hWnd, GA_ROOTOWNER), WM_COMMAND, 0x131, 0);
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

LRESULT CStartButton::s_StartButtonSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    return reinterpret_cast<CStartButton*>(dwRefData)->_StartButtonSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CStartButton::s_StartMenuSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    if (uMsg == WM_WINDOWPOSCHANGING)
    {
        PWINDOWPOS pwp = (WINDOWPOS*)lParam;
        if ((pwp->flags & SWP_NOZORDER) != 0 || c_tray._uStuckPlace != 1 && c_tray._uStuckPlace != 3)
        {
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }
        pwp->hwndInsertAfter = v_hwndTray;

        return 0;
    }
    if (uMsg == WM_NCDESTROY)
    {
        RemoveWindowSubclass(hWnd, s_StartMenuSubclassProc, uIdSubclass);
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}