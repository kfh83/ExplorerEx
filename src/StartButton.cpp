#include "StartButton.h"

#include <shlwapi.h>
#include <windowsx.h>

#include "SHFusion.h"
#include "tray.h"
#include "Util.h"

CStartButton::CStartButton()
{
}

HRESULT CStartButton::QueryInterface(REFIID riid, void** ppvObject)  // taken from ep_taskbar 7-stuff (?)
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
    LPITEMIDLIST ppidlLast;
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

void CStartButton::GetRect(RECT* lpRect)
{
    GetWindowRect(_hwndStartBtn, lpRect);
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
