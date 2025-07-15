#include "pch.h"
#include "cabinet.h"
#include "trayclok.h"
#include "traynot.h"
#include "tray.h"
#include "util.h"
#include "shellapi.h"
#include "shundoc.h"
#include "cocreateinstancehook.h"

const WCHAR CTrayNotify::c_szTrayNotify[]                   = L"TrayNotifyWnd";

void ExplorerPlaySound(LPCTSTR pszSound)
{
    // note, we have access only to global system sounds here as we use "Apps\.Default"
    TCHAR szKey[256];
    StringCchPrintf(szKey, ARRAYSIZE(szKey), TEXT("AppEvents\\Schemes\\Apps\\.Default\\%s\\.current"), pszSound);

    TCHAR szFileName[MAX_PATH];
    szFileName[0] = 0;
    LONG cbSize = sizeof(szFileName);

    // test for an empty string, PlaySound will play the Default Sound if we
    // give it a sound it cannot find...

    if ((RegQueryValue(HKEY_CURRENT_USER, szKey, szFileName, &cbSize) == ERROR_SUCCESS)
        && szFileName[0])
    {
        // flags are relevant, we try to not stomp currently playing sounds
        PlaySound(szFileName, NULL, SND_FILENAME | SND_ASYNC | SND_NODEFAULT | SND_NOSTOP);
    }
}

STDMETHODIMP_(ULONG) CTrayNotify::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CTrayNotify::Release()
{
    ASSERT( 0 != m_cRef );
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if ( 0 == cRef )
    {
        //
        //  TODO:   gpease  27-FEB-2002
        //
        // delete this; Why is this statement missing? If on purpose, why even 
        //  bother with InterlockedXXX and even the refer counter?!?
        //
    }
    return cRef;
}

HRESULT CTrayNotify::SetPreference(NOTIFYITEM pNotifyItem)
{
    return E_NOTIMPL;
}

HRESULT CTrayNotify::RegisterCallback(INotificationCB* pNotifyCB)
{
    return E_NOTIMPL;
}

HRESULT CTrayNotify::EnableAutoTray(BOOL bTraySetting)
{
    return E_NOTIMPL;
}

#pragma optimize( "", off )
//
// CTrayNotifyStub functions...
//
HRESULT CTrayNotifyStub::SetPreference(NOTIFYITEM pNotifyItem)
{
    return c_tray._trayNotify.SetPreference(pNotifyItem);
}

HRESULT CTrayNotifyStub::RegisterCallback(INotificationCB* pNotifyCB, ULONG*)
{
    return c_tray._trayNotify.RegisterCallback(pNotifyCB);
}

HRESULT CTrayNotifyStub::UnregisterCallback(ULONG*)
{
    return S_OK;
}

HRESULT CTrayNotifyStub::EnableAutoTray(BOOL bTraySetting)
{
    return c_tray._trayNotify.EnableAutoTray(bTraySetting);
}

HRESULT CTrayNotifyStub::DoAction(BOOL bTraySetting)
{
    return S_OK;
}

STDMETHODIMP_(HRESULT __stdcall) CTrayNotifyStub::SetWindowingEnvironmentConfig(IUnknown* unk)
{
    return E_NOTIMPL;
}

HRESULT CTrayNotifyStub_CreateInstance(IUnknown* pUnkOuter, IUnknown** ppunk)
{
    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION;

    CComObject<CTrayNotifyStub> *pStub = new CComObject<CTrayNotifyStub>;
    if (pStub)
        return pStub->QueryInterface(IID_PPV_ARGS(ppunk));
    else
        return E_OUTOFMEMORY;
}
#pragma optimize( "", on )

HWND CTrayNotify::TrayNotifyCreate(HWND hwndParent, UINT uID, HINSTANCE hInst)
{
    WNDCLASSEX wc;
    DWORD dwExStyle = WS_EX_STATICEDGE;

    ZeroMemory(&wc, sizeof(wc));
    wc.cbSize = sizeof(WNDCLASSEX);

    if (!GetClassInfoEx(hInst, c_szTrayNotify, &wc))
    {
        wc.lpszClassName = c_szTrayNotify;
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = s_WndProc;
        wc.hInstance = hInst;
        //wc.hIcon = NULL;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_3DFACE+1);
        //wc.lpszMenuName  = NULL;
        //wc.cbClsExtra = 0;
        wc.cbWndExtra = sizeof(CTrayNotify *);
        //wc.hIconSm = NULL;

        if (!RegisterClassEx(&wc))
        {
            return(NULL);
        }

        auto err = GetLastError();

        if (!ClockCtl_Class(hInst))
        {
            return(NULL);
        }
    }

    return(CreateWindowEx(dwExStyle, c_szTrayNotify,
            NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | WS_CLIPCHILDREN, 0, 0, 0, 0,
            hwndParent, IntToPtr_(HMENU, uID), hInst, (void*)this));
}

LRESULT CTrayNotify::TrayNotify(HWND hwndNotify, HWND hwndFrom, PCOPYDATASTRUCT pcds, BOOL* pbRefresh)
{
    if (!hwndNotify || !pcds)
    {
        return (FALSE);
    }

    CTrayNotify* ptnd = (CTrayNotify*)GetWindowLongPtr(hwndNotify, 0);
    wprintf(TEXT("Find the goobber that prematurely removed the PRIVDATA pointer.\n"));
    if (ptnd)
    {
        if (pcds->cbData < sizeof(TRAYNOTIFYDATA))
        {
            return (FALSE);
        }

        // We'll add a signature just in case
        PTRAYNOTIFYDATA pnid = (PTRAYNOTIFYDATA)pcds->lpData;
        if (pnid->dwSignature != NI_SIGNATURE)
        {
            return (FALSE);
        }

        return ptnd->TrayNotifyIcon(pnid, pbRefresh);
    }

    return (TRUE);
}

WPARAM CTrayNotify::_CalculateAnchorPointWPARAMIfNecessary(DWORD inputType, HWND const hwnd, int itemIndex)
{
    if (inputType == TRAYITEM_ANCHORPOINT_INPUTTYPE_MOUSE)
    {
        POINT ptCursor;
        if (GetCursorPos(&ptCursor))
            return MAKEWPARAM(ptCursor.x, ptCursor.y);
        return 0;
    }
    if (inputType == TRAYITEM_ANCHORPOINT_INPUTTYPE_KEYBOARD)
    {
        RECT rcItem;
        if (SendMessageW(hwnd, TB_GETITEMRECT, itemIndex, (LPARAM)&rcItem))
        {
            MapWindowPoints(hwnd, nullptr, (LPPOINT)&rcItem, 2);
            return MAKEWPARAM((rcItem.left + rcItem.right) / 2, (rcItem.top + rcItem.bottom) / 2);
        }
        return 0;
    }
    return inputType;
}

LRESULT CTrayNotify::_SendNotify(PTNPRIVICON ptnpi, UINT uMsg, DWORD dwAnchorPoint /* = 0 */, HWND const hwnd /* = NULL */, int itemIndex /* = 0 */)
{
    ASSERT(!_fNoTrayItemsDisplayPolicyEnabled);

    WPARAM wParam;
    LPARAM lParam;
    if (ptnpi->uVersion >= NOTIFYICON_VERSION_4) // Version 4, used since Vista:
    {
        wParam = _CalculateAnchorPointWPARAMIfNecessary(dwAnchorPoint, hwnd, itemIndex);
        lParam = MAKELPARAM(uMsg, ptnpi->uID);
    }
    else // NOTIFYICON version 3 and prior:
    {
        wParam = ptnpi->uID;
        lParam = uMsg;
    }

    if (ptnpi->uCallbackMessage && ptnpi->hWnd)
        return SendNotifyMessage(ptnpi->hWnd, ptnpi->uCallbackMessage, wParam, lParam);
    return 0;
}

void CTrayNotify::_SetImage(INT_PTR iIndex, int iImage)
{
    TBBUTTONINFO tbbi;
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = TBIF_IMAGE | TBIF_BYINDEX;
    tbbi.iImage = iImage;

    SendMessage(_hwndToolbar, TB_SETBUTTONINFO, iIndex, (LPARAM)&tbbi);
}

void CTrayNotify::_SetText(INT_PTR iIndex, LPTSTR pszText)
{
    TBBUTTONINFO tbbi;
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = TBIF_TEXT | TBIF_BYINDEX;
    tbbi.pszText = pszText;
    tbbi.cchText = -1;
    SendMessage(_hwndToolbar, TB_SETBUTTONINFO, iIndex, (LPARAM)&tbbi);
}

int CTrayNotify::_GetImage(INT_PTR iIndex)
{
    TBBUTTONINFO tbbi;
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = TBIF_IMAGE | TBIF_BYINDEX;

    SendMessage(_hwndToolbar, TB_GETBUTTONINFO, iIndex, (LPARAM)&tbbi);
    return tbbi.iImage;
}

#define _GetDataByIndex(i) _GetData(i, TRUE)
PTNPRIVICON CTrayNotify::_GetData(INT_PTR i, BOOL byIndex)
{
    TBBUTTONINFO tbbi;
    tbbi.cbSize = sizeof(tbbi);
    tbbi.lParam = 0;
    tbbi.dwMask = TBIF_LPARAM;
    if (byIndex)
        tbbi.dwMask |= TBIF_BYINDEX;
    SendMessage(_hwndToolbar, TB_GETBUTTONINFO, i, (LPARAM)&tbbi);
    return (PTNPRIVICON)(LPVOID)tbbi.lParam;
}

INT_PTR CTrayNotify::_GetCount()
{
    return SendMessage(_hwndToolbar, TB_BUTTONCOUNT, 0, 0L);
}

INT_PTR CTrayNotify::_GetVisibleCount()
{
    if (_iVisCount != -1)
    {
        return _iVisCount;
    }
    else
    {
        INT_PTR i;
        INT_PTR iRet = 0;
        INT_PTR iCount = _GetCount();
        TBBUTTONINFO tbbi;
        tbbi.cbSize = sizeof(tbbi);
        tbbi.dwMask = TBIF_STATE | TBIF_BYINDEX;

        for (i = 0; i < iCount; i++)
        {
            SendMessage(_hwndToolbar, TB_GETBUTTONINFO, i, (LPARAM)&tbbi);
            if (!(tbbi.fsState & TBSTATE_HIDDEN))
            {
                iRet++;
            }
        }
        _iVisCount = iRet;
        return iRet;
    }
}

int CTrayNotify::_FindImageIndex(HICON hIcon, BOOL fSetAsSharedSource)
{
    INT_PTR iCount = _GetCount();

    for (INT_PTR i = 0; i < iCount; i++)
    {
        PTNPRIVICON ptnpi = _GetData(i, TRUE);
        if (ptnpi->hIcon == hIcon)
        {
            // if we're supposed to mark this as a shared icon source and it's not itself a shared icon
            // target, mark it now.  this is to allow us to recognize when the source icon changes and
            // that we can know that we need to find other indicies and update them too.
            if (fSetAsSharedSource && !(ptnpi->dwState & NIS_SHAREDICON))
                ptnpi->dwState |= NISP_SHAREDICONSOURCE;
            return _GetImage(i);
        }
    }
    return -1;
}

void CTrayNotify::_RemoveImage(UINT uIMLIndex)
{
    if (uIMLIndex != (UINT)-1)
    {
        ImageList_Remove(_himlIcons, uIMLIndex);

        INT_PTR nCount = _GetCount();
        for (INT_PTR i = nCount - 1; i >= 0; i--)
        {
            int iImage = _GetImage(i);
            if (iImage > (int)uIMLIndex)
                _SetImage(i, iImage - 1);
        }
    }
}

INT_PTR CTrayNotify::_FindNotify(PNOTIFYICONDATA32 pnid)
{
    INT_PTR i;
    INT_PTR nCount = _GetCount();

    for (i = nCount - 1; i >= 0; --i)
    {
        PTNPRIVICON ptnpi = _GetDataByIndex(i);
        ASSERT(ptnpi);
        if (ptnpi->hWnd == GetHWnd(pnid) && ptnpi->uID == pnid->uID)
        {
            break;
        }
    }

    return (i);
}

BOOL CTrayNotify::_CheckAndResizeImages()
{
    int cxSmIconOld, cySmIconOld;
    BOOL fOK = TRUE;

    HIMAGELIST himlOld = _himlIcons;

    // Do dimensions match current icons?
    int cxSmIconNew = GetSystemMetrics(SM_CXSMICON);
    int cySmIconNew = GetSystemMetrics(SM_CYSMICON);
    ImageList_GetIconSize(himlOld, &cxSmIconOld, &cySmIconOld);
    if (cxSmIconNew != cxSmIconOld || cySmIconNew != cySmIconOld)
    {
        HIMAGELIST himlNew = ImageList_Create(cxSmIconNew, cySmIconNew, SHGetImageListFlags(_hwndToolbar), 0, 1);
        if (himlNew)
        {
            // Copy the images over to the new image list.
            int cItems = ImageList_GetImageCount(himlOld);
            for (int i = 0; i < cItems; i++)
            {
                // REVIEW Lame - there's no way to copy images to an empty
                // imagelist, resizing it on the way.
                HICON hicon = ImageList_GetIcon(himlOld, i, ILD_NORMAL);
                if (hicon)
                {
                    if (ImageList_AddIcon(himlNew, hicon) == -1)
                    {
                        // Couldn't copy image so bail.
                        fOK = FALSE;
                    }
                    DestroyIcon(hicon);
                }
                else
                {
                    fOK = FALSE;
                }

                // FU - bail.
                if (!fOK)
                    break;
            }

            // Did everything copy over OK?
            if (fOK)
            {
                // Yep, Set things up to use the new one.
                _himlIcons = himlNew;
                // Destroy the old icon cache.
                ImageList_Destroy(himlOld);
                SendMessage(_hwndToolbar, TB_SETIMAGELIST, 0, (LPARAM)_himlIcons);
                SendMessage(_hwndToolbar, TB_AUTOSIZE, 0, 0);
            }
            else
            {
                // Nope, stick with what we have.
                ImageList_Destroy(himlNew);
            }
        }
    }

    return fOK;
}

void CTrayNotify::_ActivateTips(BOOL bActivate)
{
    if ((HWND)SendMessage(_hwndToolbar, TB_GETTOOLTIPS, 0, 0))
        SendMessage((HWND)SendMessage(_hwndToolbar, TB_GETTOOLTIPS, 0, 0), TTM_ACTIVATE, (WPARAM)bActivate, 0);
}

void CTrayNotify::_InfoTipMouseClick(int x, int y)
{
    if (_pinfo)
    {
        RECT rect;
        POINT pt;

        pt.x = x;
        pt.y = y;

        GetWindowRect(_hwndInfoTip, &rect);
        // x & y are mapped to our window so map the rect to our window as well
        MapWindowRect(HWND_DESKTOP, _hwndNotify, &rect);

        if (PtInRect(&rect, pt))
            _ShowInfoTip(_pinfo->nIcon, FALSE, FALSE);
    }
}

void CTrayNotify::_PositionInfoTip()
{
    int x = 0, y = 0;
    RECT rc;

    if (!_pinfo)
        return;

    if (SendMessage(_hwndToolbar, TB_GETITEMRECT, (WPARAM)_pinfo->nIcon, (LPARAM)&rc))
    {
        MapWindowPoints(_hwndToolbar, HWND_DESKTOP, (LPPOINT)&rc, 2);
        x = (rc.left + rc.right) / 2;
        y = (rc.top + rc.bottom) / 2;
    }

    SendMessage(_hwndInfoTip, TTM_TRACKPOSITION, 0, MAKELONG(x, y));
}

void CTrayNotify::_ShowInfoTip(INT_PTR nIcon, BOOL bShow, BOOL bAsync)
{
    TOOLINFO ti = { 0 };

    ti.cbSize = sizeof(ti);
    ti.hwnd = _hwndNotify;
    ti.uId = (INT_PTR)_hwndNotify;

    // make sure we only show/hide what we intended to show/hide
    if (_pinfo && _pinfo->nIcon == nIcon)
    {
        PTNPRIVICON ptnpi = _GetDataByIndex(nIcon);

        if (!ptnpi || ptnpi->dwState & NIS_HIDDEN)
        {
            // icon is hidden, cannot show it's balloon
            bShow = FALSE; //show the next balloon instead
        }

        if (bShow)
        {
            // If there is a fullscreen app or a screen saver running,
            // we don't show anything and we empty the queue
            if (c_tray._fStuckRudeApp || _IsScreenSaverRunning())
            {
                LocalFree(_pinfo);
                _pinfo = NULL;
                DPA_EnumCallback(_hdpaInfo, DeleteDPAPtrCB, NULL);
                DPA_DeleteAllPtrs(_hdpaInfo);
                return;
            }

            if (bAsync)
                PostMessage(_hwndNotify, TNM_ASYNCINFOTIP, (WPARAM)nIcon, 0);
            else
            {
                HICON hIcon = NULL;
                if (_pinfo->dwFlags & NIIF_USER)
                {
                    if (ptnpi->hBalloonIcon)
                        hIcon = ptnpi->hBalloonIcon;
                    else
                        hIcon = ptnpi->hIcon;
                }
                SendMessage(_hwndInfoTip,
                            TTM_SETTITLE,
                            hIcon ? (WPARAM)hIcon : _pinfo->dwFlags & INFO_ICON,
                            (LPARAM)_pinfo->szTitle);
                _PositionInfoTip();
                ti.lpszText = _pinfo->szInfo;
                // if tray is in auto hide mode unhide it
                c_tray.Unhide();
                c_tray._fBalloonUp = TRUE;
                SendMessage(_hwndInfoTip, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);
                // disable regular tooltips
                _ActivateTips(FALSE);
                // show the balloon
                SendMessage(_hwndInfoTip, TTM_TRACKACTIVATE, (WPARAM)TRUE, (LPARAM)&ti);
                _uInfoTipTimer = SetTimer(_hwndNotify, INFO_TIMER, MIN_INFO_TIME, NULL);
            }
        }
        else
        {
            LocalFree(_pinfo);
            _pinfo = NULL;
            if (_uInfoTipTimer)
                KillTimer(_hwndNotify, _uInfoTipTimer);
            // hide the previous tip if any
            SendMessage(_hwndInfoTip, TTM_TRACKACTIVATE, (WPARAM)FALSE, (LPARAM)0);
            ti.lpszText = NULL;
            SendMessage(_hwndInfoTip, TTM_UPDATETIPTEXT, 0, (LPARAM)&ti);

            // we are hiding the current balloon. are there any waiting? yes, then show the first one
            if (_hdpaInfo && DPA_GetPtrCount(_hdpaInfo) > 0)
            {
                TNINFOITEM* pii = (TNINFOITEM*)DPA_DeletePtr(_hdpaInfo, 0);

                ASSERT(pii);
                _pinfo = pii;
                _ShowInfoTip(pii->nIcon, TRUE, bAsync);
            }
            else
            {
                c_tray._fBalloonUp = FALSE; // this will take care of hiding the tray if necessary
                _ActivateTips(TRUE);
            }
        }
    }
    else if (_pinfo && !bShow)
    {
        // we wanted to hide something that wasn't showing up
        // maybe it's in the queue
        if (_hdpaInfo && DPA_GetPtrCount(_hdpaInfo) > 0)
        {
            int cItems = DPA_GetPtrCount(_hdpaInfo);

            for (int i = 0; i < cItems; i++)
            {
                TNINFOITEM* pii = (TNINFOITEM*)DPA_GetPtr(_hdpaInfo, i);

                ASSERT(pii);
                if (pii->nIcon == nIcon)
                {
                    DPA_DeletePtr(_hdpaInfo, i); // this just removes it from the dpa
                    LocalFree(pii);
                    return; // just remove the first one
                }
            }
        }
    }
}

void CTrayNotify::_SetInfoTip(INT_PTR nIcon, PNOTIFYICONDATA32 pnid, BOOL bAsync)
{
    if (!_hwndInfoTip)
    {
        _hwndInfoTip = CreateWindow(TOOLTIPS_CLASS, NULL,
                                    WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP | TTS_BALLOON,
                                    CW_USEDEFAULT, CW_USEDEFAULT,
                                    CW_USEDEFAULT, CW_USEDEFAULT,
                                    _hwndNotify, NULL, hinstCabinet,
                                    NULL);
        SetWindowPos(_hwndInfoTip, HWND_TOPMOST,
                     0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        if (_hwndInfoTip)
        {
            TOOLINFO ti = { 0 };
            RECT rc = { 0, -2, 0, 0 };

            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_IDISHWND | TTF_TRACK | TTF_TRANSPARENT;
            ti.hwnd = _hwndNotify;
            ti.uId = (UINT_PTR)_hwndNotify;
            //ti.lpszText = NULL;
            // set the version so we can have non buggy mouse event forwarding
            SendMessage(_hwndInfoTip, CCM_SETVERSION, COMCTL32_VERSION, 0);
            SendMessage(_hwndInfoTip, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
            SendMessage(_hwndInfoTip, TTM_SETMAXTIPWIDTH, 0, (LPARAM)MAX_TIP_WIDTH);
            SendMessage(_hwndInfoTip, TTM_SETMARGIN, 0, (LPARAM)&rc);
        }

        ASSERT(_hdpaInfo == NULL);
        _hdpaInfo = DPA_Create(10);
    }

    if (_hwndInfoTip)
    {
        // show the new one...
        if (pnid && pnid->szInfo[0] != TEXT('\0'))
        {
            TNINFOITEM* pii = (TNINFOITEM*)LocalAlloc(LPTR, sizeof(TNINFOITEM));
            if (pii)
            {
                pii->nIcon = nIcon;
                lstrcpyn(pii->szInfo, pnid->szInfo, ARRAYSIZE(pii->szInfo));
                lstrcpyn(pii->szTitle, pnid->szInfoTitle, ARRAYSIZE(pii->szTitle));
                pii->uTimeout = pnid->uTimeout;
                if (pii->uTimeout < MIN_INFO_TIME)
                    pii->uTimeout = MIN_INFO_TIME;
                else if (pii->uTimeout > MAX_INFO_TIME)
                    pii->uTimeout = MAX_INFO_TIME;
                pii->dwFlags = pnid->dwInfoFlags;
                // if pinfo != NULL then we have a balloon showing right now
                // if hdpaInfo != NULL then we just add the new balloon to the queue
                // else we show the new one right away -- it hides the old one but that's ok
                // this is just low mem case when we could not alloc hdpa
                // we also show the info tip right away if the current tip has been shown
                // for at least the min time and there are no tips waiting (queue size == 0)
                if (_pinfo && _hdpaInfo &&
                    (!_pinfo->bMinShown || DPA_GetPtrCount(_hdpaInfo) > 0))
                {
                    if (DPA_AppendPtr(_hdpaInfo, pii) == -1)
                    {
                        LocalFree(pii);
                    }
                    return;
                }

                if (_pinfo)
                {
                    LocalFree(_pinfo);
                }

                _pinfo = pii;

                _ShowInfoTip(nIcon, TRUE, bAsync);
            }
        }
        else
        {
            _ShowInfoTip(nIcon, FALSE, FALSE);
        }
    }
}

BOOL CTrayNotify::_ModifyNotify(PNOTIFYICONDATA32 pnid, INT_PTR nIcon, BOOL* pbRefresh)
{
    BOOL fResize = FALSE;
    BOOL bHideButton = FALSE;
    PTNPRIVICON ptnpi = _GetDataByIndex(nIcon);
    if (!ptnpi)
    {
        return (FALSE);
    }

    _CheckAndResizeImages();

    if (pnid->uFlags & NIF_STATE)
    {
#define NIS_VALIDMASK (NIS_HIDDEN | NIS_SHAREDICON)
        DWORD dwOldState = ptnpi->dwState;
        // validate mask
        if (pnid->dwStateMask & ~NIS_VALIDMASK)
            return FALSE;

        ptnpi->dwState = (pnid->dwState & pnid->dwStateMask) | (ptnpi->dwState & ~pnid->dwStateMask);

        if (pnid->dwStateMask & NIS_HIDDEN)
        {
            TBBUTTONINFO tbbi;
            tbbi.cbSize = sizeof(tbbi);
            tbbi.dwMask = TBIF_STATE | TBIF_BYINDEX;

            SendMessage(_hwndToolbar, TB_GETBUTTONINFO, nIcon, (LPARAM)&tbbi);
            if (ptnpi->dwState & NIS_HIDDEN)
            {
                tbbi.fsState |= TBSTATE_HIDDEN;
                bHideButton = TRUE;
            }
            else
                tbbi.fsState &= ~TBSTATE_HIDDEN;
            SendMessage(_hwndToolbar, TB_SETBUTTONINFO, nIcon, (LPARAM)&tbbi);
            _iVisCount = -1;
        }

        if ((pnid->dwState ^ dwOldState) & NIS_SHAREDICON)
        {
            if (dwOldState & NIS_SHAREDICON)
            {
                // if we're going from shared to not shared,
                // clear the icon
                _SetImage(nIcon, -1);
                ptnpi->hIcon = NULL;
            }
        }
        fResize |= ((pnid->dwState ^ dwOldState) & NIS_HIDDEN);
    }

    // The icon is the only thing that can fail, so I will do it first
    if (pnid->uFlags & NIF_ICON)
    {
        int iImageNew;
        if (ptnpi->dwState & NIS_SHAREDICON)
        {
            iImageNew = _FindImageIndex(GetHIcon(pnid), TRUE);
            if (iImageNew == -1)
                return FALSE;
        }
        else
        {
            int iImageOld = _GetImage(nIcon);
            if (GetHIcon(pnid))
            {
                // Replace icon knows how to handle -1 for add
                iImageNew = ImageList_ReplaceIcon(_himlIcons, iImageOld,
                                                  GetHIcon(pnid));
                if (iImageNew < 0)
                {
                    return (FALSE);
                }
            }
            else
            {
                _RemoveImage(iImageOld);
                iImageNew = -1;
            }

            if (ptnpi->dwState & NISP_SHAREDICONSOURCE)
            {
                INT_PTR i;
                INT_PTR iCount = _GetCount();
                // if we're the source of shared icons, we need to go update all the other icons that
                // are using our icon
                for (i = 0; i < iCount; i++)
                {
                    if (_GetImage(i) == iImageOld)
                    {
                        PTNPRIVICON ptnpiTemp = _GetDataByIndex(i);
                        ptnpiTemp->hIcon = GetHIcon(pnid);
                        _SetImage(i, iImageNew);
                    }
                }
            }
            if (iImageOld == -1 || iImageNew == -1)
                fResize = TRUE;
        }
        ptnpi->hIcon = GetHIcon(pnid);
        _SetImage(nIcon, iImageNew);
    }

    if (pnid->uFlags & NIF_MESSAGE)
    {
        ptnpi->uCallbackMessage = pnid->uCallbackMessage;
    }

    if (pnid->uFlags & NIF_TIP)
    {
        _SetText(nIcon, pnid->szTip);
        if (!bHideButton && pbRefresh && pnid->uFlags == NIF_TIP)
        {
            *pbRefresh = FALSE;
        }

        if (pnid->dwInfoFlags & NIIF_USER)
        {
            if (ptnpi->hBalloonIcon)
                DestroyIcon(ptnpi->hBalloonIcon);
            if (pnid->dwBalloonIcon)
                ptnpi->hBalloonIcon = GetHBalloonIcon(pnid);
        }
    }

    if (fResize)
        _Size();

    // infotip is up and we are hiding the button it corresponds to...
    if (bHideButton && _pinfo && _pinfo->nIcon == nIcon)
        _ShowInfoTip(_pinfo->nIcon, FALSE, TRUE);

    // need to have info stuff done after resize because we need to
    // position the infotip relative to the hwndToolbar
    if (pnid->uFlags & NIF_INFO)
    {
        // if button is hidden we don't show infotip
        if (!(ptnpi->dwState & NIS_HIDDEN))
            _SetInfoTip(nIcon, pnid, fResize);
    }

#if XXX_RENIM
    CacheNID(COP_ADD, nIcon, pnid);
#endif

    return (TRUE);
}

BOOL CTrayNotify::_SetVersionNotify(PNOTIFYICONDATA32 pnid, INT_PTR nIcon)
{
    PTNPRIVICON ptnpi = _GetDataByIndex(nIcon);
    if (!ptnpi)
        return FALSE;

    if (pnid->uVersion < NOTIFYICON_VERSION)
    {
        ptnpi->uVersion = 0;
        return TRUE;
    }
    else if (pnid->uVersion == NOTIFYICON_VERSION) // Version 3, used by XP.
    {
        ptnpi->uVersion = NOTIFYICON_VERSION;
        return TRUE;
    }
    else if (pnid->uVersion == NOTIFYICON_VERSION_4) // Version 4, used since Vista.
    {
        ptnpi->uVersion = NOTIFYICON_VERSION_4;
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

void CTrayNotify::_FreeNotify(PTNPRIVICON ptnpi, int iImage)
{
    //if it wasn't sharing an icon with another guy, go ahead and delete it
    if (!(ptnpi->dwState & NIS_SHAREDICON))
    {
        _RemoveImage(iImage);
    }
    LocalFree(ptnpi);
}

INT_PTR CTrayNotify::_DeleteNotify(INT_PTR nIcon)
{
    // delete info tip if showing
    if (_pinfo && _pinfo->nIcon == nIcon)
    {
        KillTimer(_hwndNotify, _uInfoTipTimer);
        SendMessage(_hwndInfoTip, TTM_TRACKACTIVATE, (WPARAM)FALSE, (LPARAM)0);
        _uInfoTipTimer = 0;
    }
    else
    {
        PostMessage(_hwndNotify, TNM_ASYNCINFOTIP, 0, 0);
    }

    if (_hdpaInfo && DPA_GetPtrCount(_hdpaInfo) > 0)
    {
        int cItems = DPA_GetPtrCount(_hdpaInfo);
        for (int i = cItems - 1; i >= 0; i--)
        {
            TNINFOITEM* pii = (TNINFOITEM*)DPA_GetPtr(_hdpaInfo, i);

            ASSERT(pii);
            if (pii->nIcon == nIcon)
            {
                DPA_DeletePtr(_hdpaInfo, i); // this just removes it from the dpa
                LocalFree(pii);
            }
            else if (nIcon < pii->nIcon)
            {
                pii->nIcon--;
            }
        }
    }

    if (_pinfo)
    {
        if (nIcon == _pinfo->nIcon)
        {
            // frees pinfo and shows the next balloon if any
            _ShowInfoTip(_pinfo->nIcon, FALSE, TRUE);
        }
        else if (nIcon < _pinfo->nIcon)
        {
            _pinfo->nIcon--;
        }
        else
        {
            PostMessage(_hwndNotify, TNM_ASYNCINFOTIP, (WPARAM)_pinfo->nIcon, 0);
        }
    }

#if XXX_RENIM
    CacheNID(COP_DEL, nIcon, NULL);
#endif

    _iVisCount = -1;
    return SendMessage(_hwndToolbar, TB_DELETEBUTTON, nIcon, 0);
}

BOOL CTrayNotify::_InsertNotify(PNOTIFYICONDATA32 pnid)
{
    TBBUTTON tbb;
    static int s_iNextID = 0;

    // First insert a totally "default" icon
    PTNPRIVICON ptnpi = (PTNPRIVICON)LocalAlloc(LPTR, sizeof(TNPRIVICON));
    if (!ptnpi)
    {
        return FALSE;
    }

    ptnpi->hWnd = GetHWnd(pnid);
    ptnpi->uID = pnid->uID;
    
    tbb.dwData = (DWORD_PTR)ptnpi;
    tbb.iBitmap = -1;
    tbb.idCommand = s_iNextID++;
    tbb.fsStyle = TBSTYLE_BUTTON;
    tbb.fsState = TBSTATE_ENABLED;
    tbb.iString = -1;
    if (SendMessage(_hwndToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb))
    {
        INT_PTR insert = _GetCount() - 1;

        // Then modify this icon with the specified info
        if (!_ModifyNotify(pnid, insert, NULL))
        {
            _DeleteNotify(insert);
            // Note that we do not go to the LocalFree
            return FALSE;
        }
    }
    return TRUE;
}

void CTrayNotify::_SetCursorPos(INT_PTR i)
{
    RECT rc;
    if (SendMessage(_hwndToolbar, TB_GETITEMRECT, i, (LPARAM)&rc))
    {
        MapWindowPoints(_hwndToolbar, HWND_DESKTOP, (LPPOINT)&rc, 2);
        SetCursorPos((rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2);
    }
}

LRESULT CTrayNotify::_ToolbarWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    BOOL fClickDown = FALSE;

    switch (uMsg)
    {
        case WM_LBUTTONDOWN:
        case WM_MBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
        case WM_MBUTTONDBLCLK:
        case WM_RBUTTONDBLCLK:
            fClickDown = TRUE;
        // Fall through

        case WM_MOUSEMOVE:
        case WM_LBUTTONUP:
        case WM_MBUTTONUP:
        case WM_RBUTTONUP:
            _MouseEvent(uMsg, wParam, lParam, fClickDown);
            break;

        case WM_KEYDOWN:
            switch (wParam)
            {
                case VK_RETURN:
                    _fReturn = TRUE;
                    break;

                case VK_SPACE:
                    _fReturn = FALSE;
                    break;
            }
            break;

        case WM_CONTEXTMENU:
        {
            HWND hwnd;
            INT_PTR i = SendMessage(_hwndToolbar, TB_GETHOTITEM, 0, 0);
            if (i != -1)
            {
                PTNPRIVICON ptnpi = _GetDataByIndex(i);
                if (lParam == (LPARAM)-1)
                    _fKey = TRUE;

                if (_fKey)
                {
                    _SetCursorPos(i);
                }

                hwnd = (HWND)SendMessage(_hwndToolbar, TB_GETTOOLTIPS, 0, 0);
                if (hwnd)
                    SendMessage(hwnd, TTM_POP, 0, 0);

                if (ptnpi)
                {
                    // Determine the anchor point command for V4:
                    DWORD dwAnchorPointCmd = ptnpi->uVersion >= NOTIFYICON_VERSION_4
                        ? (lParam != TRAYITEM_ANCHORPOINT_INPUTTYPE_MOUSE
                            ? (DWORD)lParam
                            : TRAYITEM_ANCHORPOINT_INPUTTYPE_KEYBOARD)
                        : 0;

                    SHAllowSetForegroundWindow(ptnpi->hWnd);
                    if (ptnpi->uVersion >= KEYBOARD_VERSION)
                    {
                        _SendNotify(ptnpi, WM_CONTEXTMENU, dwAnchorPointCmd, hwnd, i);
                    }
                    else
                    {
                        if (_fKey)
                        {
                            _SendNotify(ptnpi, WM_RBUTTONDOWN);
                            _SendNotify(ptnpi, WM_RBUTTONUP);
                        }
                    }
                }
                return 0;
            }
        }
    }
    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CTrayNotify::_Create(HWND hWnd)
{
    UINT flags = ILC_MASK;

    HWND hwndClock = ClockCtl_Create(hWnd, IDC_CLOCK, hinstCabinet);
    if (!hwndClock)
    {
        delete this;
        SetWindowLongPtr(hWnd,0,NULL);
        //LocalFree(this);
        return (-1);
    }


    //SetWindowLongPtr(hWnd, 0, (LONG_PTR)this);

    _iVisCount = -1;
    _hwndNotify = hWnd;
    _hwndClock = hwndClock;
    _hwndToolbar = CreateWindowEx(WS_EX_TOOLWINDOW, TOOLBARCLASSNAME, NULL,
                                  WS_VISIBLE | WS_CHILD | TBSTYLE_FLAT |
                                  TBSTYLE_TOOLTIPS |
                                  WS_CLIPCHILDREN |
                                  WS_CLIPSIBLINGS | CCS_NODIVIDER | CCS_NOPARENTALIGN |
                                  CCS_NORESIZE | TBSTYLE_WRAPABLE,
                                  0, 0, 0, 0, hWnd, 0, hinstCabinet, NULL);

    SendMessage(_hwndToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
    HWND hwndTooltips = (HWND)SendMessage(_hwndToolbar, TB_GETTOOLTIPS, 0, 0);
    SHSetWindowBits(hwndTooltips, GWL_STYLE, TTS_ALWAYSTIP, TTS_ALWAYSTIP);
    SendMessage(_hwndToolbar, TB_SETPADDING, 0, MAKELONG(2, 2));
    SendMessage(_hwndToolbar, TB_SETMAXTEXTROWS, 0, 0);
    SendMessage(_hwndToolbar, CCM_SETVERSION, COMCTL32_VERSION, 0);
    SendMessage(_hwndToolbar, TB_SETEXTENDEDSTYLE,
                TBSTYLE_EX_INVERTIBLEIMAGELIST, TBSTYLE_EX_INVERTIBLEIMAGELIST);
    _himlIcons = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON),
        SHGetImageListFlags(_hwndToolbar), 0, 1);
    if (!_himlIcons)
    {
        return (-1);
    }


    SendMessage(_hwndToolbar, TB_SETIMAGELIST, 0, (LPARAM)_himlIcons);

    // if this fails, not that big a deal... we'll still show, but won't handle clicks
    SetWindowSubclass(_hwndToolbar, s_ToolbarWndProc, 0, (DWORD_PTR)this);

#if XXX_RENIM
    RestoreNIDList(ptnd);
#endif

    return (0);
}

LRESULT CTrayNotify::_Destroy()
{
    RemoveWindowSubclass(_hwndToolbar, s_ToolbarWndProc, 0);
    while (_DeleteNotify(0))
    {
        // Continue while there are icondata's to delete
    }

    if (_himlIcons)
    {
        ImageList_Destroy(_himlIcons);
    }

    if (_hwndInfoTip)
    {
        DestroyWindow(_hwndInfoTip);
        _hwndInfoTip = NULL;
    }
    if (_pinfo)
    {
        LocalFree(_pinfo);
        _pinfo = NULL;
    }
    if (_hdpaInfo)
    {
        DPA_DestroyCallback(_hdpaInfo, DeleteDPAPtrCB, NULL);
    }

    SetWindowLongPtr(_hwndNotify, 0, 0);
    
    return(0);
}

LRESULT CTrayNotify::_Paint()
{
    PAINTSTRUCT ps;
    HWND hWnd = _hwndNotify;

    HDC hdc = BeginPaint(hWnd, &ps);

    // Setting to the current color returns immediately, so don't bother checking
    ImageList_SetBkColor(_himlIcons, GetSysColor(COLOR_3DFACE));

    EndPaint(hWnd, &ps);
    return (0);
}

int CTrayNotify::_MatchIconsHorz(int nMatchHorz, INT_PTR nIcons, POINT* ppt)
{
    if (!nIcons)
    {
        ppt->x = ppt->y = 0;
        return (0);
    }

    int xIcon = GetSystemMetrics(SM_CXSMICON) + 2 * g_cxBorder;
    int yIcon = GetSystemMetrics(SM_CYSMICON) + 2 * g_cyBorder;

    // We can put the icons below or to the left of the clock, which
    // one gives a smaller total height?
    int nCols = (nMatchHorz - 3 * g_cxBorder) / xIcon;
    if (nCols < 1)
    {
        nCols = 1;
    }
    int nRows = ROWSCOLS((int)nIcons, nCols);

    ppt->x = g_cxEdge + (nCols * xIcon) + 2 * g_cxBorder;
    ppt->y = (nRows * yIcon) + 2 * g_cyBorder;

    return (nCols);
}

int CTrayNotify::_MatchIconsVert(int nMatchVert, INT_PTR nIcons, POINT* ppt)
{
    if (!nIcons)
    {
        ppt->x = ppt->y = 0;
        return (0);
    }

    int xIcon = GetSystemMetrics(SM_CXSMICON) + 2 * g_cxBorder;
    int yIcon = GetSystemMetrics(SM_CYSMICON) + 2 * g_cyBorder;

    // We can put the icons below or to the left of the clock, which
    // one gives a smaller total height?
    int nRows = (nMatchVert - 3 * g_cyBorder) / yIcon;
    if (nRows < 1)
    {
        nRows = 1;
    }
    int nCols = ROWSCOLS((int)nIcons, nRows);

    ppt->x = g_cxEdge + (nCols * xIcon) + 2 * g_cxBorder;
    ppt->y = (nRows * yIcon) + 2 * g_cyBorder;

    return (nCols);
}

UINT CTrayNotify::_CalcRects(int nMaxHorz, int nMaxVert, LPRECT prClock, LPRECT prNotifies)
{
    UINT nCols;
    POINT ptBelow, ptLeft;
    int nHorzLeft, nVertLeft;
    int nHorzBelow, nVertBelow;
    int nColsBelow, nColsLeft;

    // Check whether we should try to match the horizontal or vertical
    // size (the smaller is the one to match)
    enum { MATCH_HORZ, MATCH_VERT } eMatch =
        nMaxHorz <= nMaxVert ? MATCH_HORZ : MATCH_VERT;
    enum { SIDE_LEFT, SIDE_BELOW } eSide;

    LRESULT lRes = SendMessage(_hwndClock, WM_CALCMINSIZE, 0, 0);
    prClock->left = prClock->top = 0;
    prClock->right = LOWORD(lRes);
    prClock->bottom = HIWORD(lRes);

    INT_PTR nIcons = _GetVisibleCount();
    if (nIcons == 0)
    {
        SetRectEmpty(prNotifies);
        goto CalcDone;
    }

    if (eMatch == MATCH_HORZ)
    {
        nColsBelow = _MatchIconsHorz(nMaxHorz, nIcons, &ptBelow);
        // Add cxBorder because the clock goes right next to the icons
        nColsLeft = _MatchIconsHorz(nMaxHorz - prClock->right + g_cxBorder,
                                    nIcons, &ptLeft);
    }
    else
    {
        nColsBelow = _MatchIconsVert(nMaxVert - prClock->bottom,
                                     nIcons, &ptBelow);
        nColsLeft = _MatchIconsVert(nMaxVert, nIcons, &ptLeft);
    }

    nVertBelow = ptBelow.y + prClock->bottom;
    nHorzBelow = ptBelow.x;
    nVertLeft = ptLeft.y;
    nHorzLeft = ptLeft.x + prClock->right;

    eSide = SIDE_LEFT;
    if (eMatch == MATCH_HORZ)
    {
        // If there is no room on the left, or putting it below makes a
        // smaller rectangle
        if (nMaxHorz < nHorzLeft || nVertBelow < nVertLeft)
        {
            eSide = SIDE_BELOW;
        }
    }
    else
    {
        // If there is room below and putting it below makes a
        // smaller rectangle
        if (nMaxVert > nVertBelow && nHorzBelow < nHorzLeft)
        {
            eSide = SIDE_BELOW;
        }
    }

    prNotifies->left = 0;
    if (eSide == SIDE_LEFT)
    {
        prNotifies->right = ptLeft.x + g_cxBorder;
        prNotifies->top = 0;
        prNotifies->bottom = ptLeft.y;

        OffsetRect(prClock, prNotifies->right, 0);

        nCols = nColsLeft;
    }
    else
    {
        if (ptBelow.x < prClock->right && eMatch == MATCH_VERT)
        {
            // Let's recalc using the whole clock width
            nColsBelow = _MatchIconsHorz(prClock->right, nIcons, &ptBelow);
        }
        ptBelow.y += 2;
        prNotifies->right = ptBelow.x;
        prNotifies->top = prClock->bottom + g_cyBorder;
        prNotifies->bottom = prNotifies->top + ptBelow.y;

        if (prClock->right && (prClock->right < prNotifies->right))
        {
            // Use the larger value to center properly
            prClock->right = prNotifies->right;
        }

        nCols = nColsBelow;
    }

CalcDone:
    // At least as tall as a gripper
    if (prClock->bottom < g_cySize + g_cyEdge)
    {
        prClock->bottom = g_cySize + g_cyEdge;
    }

    // Never wider than the space we allotted
    if (prClock->right > nMaxHorz - 4 * g_cxBorder)
    {
        prClock->right = nMaxHorz - 4 * g_cxBorder;
    }

    // Add back the border around the whole window
    OffsetRect(prClock, g_cxBorder, g_cyBorder);

    return (nCols);
}

LRESULT CTrayNotify::_CalcMinSize(int nMaxHorz, int nMaxVert)
{
    RECT rTotal, rClock, rNotifies;

    if (!(GetWindowLong(_hwndClock, GWL_STYLE) & WS_VISIBLE) &&
        !_GetCount())
    {
        ShowWindow(_hwndNotify, SW_HIDE);
        return 0L;
    }
    else
    {
        if (!IsWindowVisible(_hwndNotify))
        {
            ShowWindow(_hwndNotify, SW_SHOW);
        }
    }

    _CalcRects(nMaxHorz, nMaxVert, &rClock, &rNotifies);

    UnionRect(&rTotal, &rClock, &rNotifies);

    // this can happen if rClock's hidden width is 0;
    // make sure the rTotal height is at least the clock's height.
    // it can be smaller if the clock is hidden and thus has a 0 width
    if ((rTotal.bottom - rTotal.top) < (rClock.bottom - rClock.top))
        rTotal.bottom = rTotal.top + (rClock.bottom - rClock.top);

    // Add on room for borders
    return (MAKELRESULT(rTotal.right+g_cxBorder, rTotal.bottom+g_cyBorder));
}

LRESULT CTrayNotify::_Size()
{
    RECT rTotal, rClock;
    RECT rNotifies;
    HWND hWnd = _hwndNotify;

    // use GetWindowRect because _TNCalcRects includes the borders
    GetWindowRect(hWnd, &rTotal);
    rTotal.right -= rTotal.left;
    rTotal.bottom -= rTotal.top;
    rTotal.left = rTotal.top = 0;

    // Account for borders on the left and right
    _nCols = _CalcRects(rTotal.right,
                        rTotal.bottom, &rClock, &rNotifies);


    SetWindowPos(_hwndClock, NULL, rClock.left, rClock.top,
                 rClock.right - rClock.left, rClock.bottom - rClock.top, SWP_NOZORDER);

    SetWindowPos(_hwndToolbar, NULL, rNotifies.left + g_cxEdge, rNotifies.top,
                 rNotifies.right - rNotifies.left - g_cxEdge, rNotifies.bottom - rNotifies.top, SWP_NOZORDER);
    return (0);
}

LRESULT CTrayNotify::_Timer(UINT uTimerID)
{
    if (uTimerID == _uInfoTipTimer)
    {
        KillTimer(_hwndNotify, _uInfoTipTimer);
        _uInfoTipTimer = 0;
        if (_pinfo)
        {
            if (_pinfo->bMinShown
                || (_hdpaInfo && DPA_GetPtrCount(_hdpaInfo) > 0)
                || _pinfo->uTimeout == MIN_INFO_TIME)
                _ShowInfoTip(_pinfo->nIcon, FALSE, FALSE); // hide this balloon and show new one
            else
            {
                _pinfo->bMinShown = TRUE;
                if (_pinfo->uTimeout > MIN_INFO_TIME)
                    _uInfoTipTimer = SetTimer(_hwndNotify, INFO_TIMER, _pinfo->uTimeout - MIN_INFO_TIME, NULL);
            }
        }
    }

    return 0;
}

LRESULT CTrayNotify::_MouseEvent(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL fClickDown)
{
    POINT pt;
    pt.x = GET_X_LPARAM(lParam);
    pt.y = GET_Y_LPARAM(lParam);

    INT_PTR i = SendMessage(_hwndToolbar, TB_HITTEST, 0, (LPARAM)&pt);
    PTNPRIVICON ptnpi = _GetDataByIndex(i);
    if (ptnpi)
    {
        if (IsWindow(ptnpi->hWnd))
        {
            if (fClickDown)
            {
                SHAllowSetForegroundWindow(ptnpi->hWnd);
                if (_pinfo && i == _pinfo->nIcon)
                    _ShowInfoTip(_pinfo->nIcon, FALSE, FALSE);
            }
            _SendNotify(ptnpi, uMsg);
        }
        else
        {
            _DeleteNotify(i);
            c_tray.SizeWindows();
        }
        return 1;
    }
    return (0);
}

LRESULT CTrayNotify::_OnCDNotify(LPNMTBCUSTOMDRAW pnm)
{
    switch (pnm->nmcd.dwDrawStage)
    {
        case CDDS_PREPAINT:
            return CDRF_NOTIFYITEMDRAW;

        case CDDS_ITEMPREPAINT:
        {
            LRESULT lRet = TBCDRF_NOOFFSET;

            // notify us for the hot tracked item please
            if (pnm->nmcd.uItemState & CDIS_HOT)
                lRet |= CDRF_NOTIFYPOSTPAINT;

            // we want the buttons to look totally flat all the time
            pnm->nmcd.uItemState = 0;

            return lRet;
        }

        case CDDS_ITEMPOSTPAINT:
        {
            // draw the hot tracked item as a focus rect, since
            // the tray notify area doesn't behave like a button:
            //   you can SINGLE click or DOUBLE click or RCLICK
            //   (kybd equiv: SPACE, ENTER, SHIFT F10)

            LRESULT lRes = SendMessage(_hwndNotify, WM_QUERYUISTATE, 0, 0);
            if (!(LOWORD(lRes) & UISF_HIDEFOCUS) && _hwndToolbar == GetFocus())
            {
                DrawFocusRect(pnm->nmcd.hdc, &pnm->nmcd.rc);
            }
            break;
        }
    }

    return CDRF_DODEFAULT;
}

LRESULT CTrayNotify::_Notify(LPNMHDR pNmhdr)
{
    switch (pNmhdr->code)
    {
        case NM_KEYDOWN:
            _fKey = TRUE;
            break;

        case TBN_ENDDRAG:
            _fKey = FALSE;
            break;

        case TBN_DELETINGBUTTON:
        {
            TBNOTIFY* ptbn = (TBNOTIFY*)pNmhdr;
            PTNPRIVICON ptnpi = (PTNPRIVICON)(LPVOID)ptbn->tbButton.dwData;
            _FreeNotify(ptnpi, ptbn->tbButton.iBitmap);
            break;
        }

        case NM_CUSTOMDRAW:
            return _OnCDNotify((LPNMTBCUSTOMDRAW)pNmhdr);

        default:
            break;
    }

    return(0);
}

void CTrayNotify::_SysChange(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (uMsg == WM_WININICHANGE)
    {
        _CheckAndResizeImages();
    }
    
    if (_hwndClock)
    {
        SendMessage(_hwndClock, uMsg, wParam, lParam);
    }
}

void CTrayNotify::_Command(UINT id, UINT uCmd)
{
    switch (uCmd)
    {
        case BN_CLICKED:
        {
            PTNPRIVICON ptnpi = _GetData(id, FALSE);
            if (ptnpi)
            {
                if (_fKey)
                    _SetCursorPos(SendMessage(_hwndToolbar, TB_COMMANDTOINDEX, id, 0));

                SHAllowSetForegroundWindow(ptnpi->hWnd);
                if (ptnpi->uVersion >= KEYBOARD_VERSION)
                {
                    // if they are a new version that understands the keyboard messages,
                    // send the real message to them.
                    _SendNotify(ptnpi, _fKey ? NIN_KEYSELECT : NIN_SELECT);
                    // Hitting RETURN is like double-clicking (which in the new
                    // style means keyselecting twice)
                    if (_fKey && _fReturn)
                        _SendNotify(ptnpi, NIN_KEYSELECT);
                }
                else
                {
                    // otherwise mock up a mouse event if it was a keyboard select
                    // (if it wasn't a keyboard select, we assume they handled it already on
                    // the WM_MOUSE message
                    if (_fKey)
                    {
                        _SendNotify(ptnpi, WM_LBUTTONDOWN);
                        _SendNotify(ptnpi, WM_LBUTTONUP);
                        if (_fReturn)
                        {
                            _SendNotify(ptnpi, WM_LBUTTONDBLCLK);
                            _SendNotify(ptnpi, WM_LBUTTONUP);
                        }
                    }
                }
            }
            break;
        }
    }
}

LRESULT CTrayNotify::v_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (WM_CREATE == uMsg)
    {
        return _Create(hWnd);
    }
    else
    {
        {
            switch (uMsg)
            {
                case WM_DESTROY:
                    return _Destroy();

                case WM_COMMAND:
                    _Command(GET_WM_COMMAND_ID(wParam, lParam), GET_WM_COMMAND_CMD(wParam, lParam));
                    break;

                case WM_SETFOCUS:
                    SetFocus(_hwndToolbar);
                    break;

                case WM_PAINT:
                    return _Paint();

                case WM_CALCMINSIZE:
                    return _CalcMinSize((int)wParam, (int)lParam);

                case WM_TIMECHANGE:
                case WM_WININICHANGE:
                case WM_POWERBROADCAST:
                case WM_POWER:
                    _SysChange(uMsg, wParam, lParam);
                    goto DoDefault;

                case WM_NCHITTEST:
                    return (_IsOverClock(lParam) ? HTTRANSPARENT : HTCLIENT);

                case WM_NOTIFY:
                    return (_Notify((LPNMHDR)lParam));

                case TNM_GETCLOCK:
                    return (LRESULT)_hwndClock;

                case TNM_TRAYHIDE:
                    if (lParam && IsWindowVisible(_hwndClock))
                        SendMessage(_hwndClock, TCM_RESET, 0, 0);
                    break;

                case TNM_HIDECLOCK:
                    ShowWindow(_hwndClock, lParam ? SW_HIDE : SW_SHOW);
                    break;

                case TNM_TRAYPOSCHANGED:
                    if (_pinfo)
                        PostMessage(_hwndNotify, TNM_ASYNCINFOTIPPOS, (WPARAM)_pinfo->nIcon, 0);
                    break;

                case TNM_ASYNCINFOTIPPOS:
                    _PositionInfoTip();
                    break;

                case TNM_ASYNCINFOTIP:
                    _ShowInfoTip((INT_PTR)wParam, TRUE, FALSE);
                    break;

                case WM_SIZE:
                    _Size();
                    break;

                case WM_TIMER:
                    _Timer((UINT)wParam);
                    break;

                case TNM_RUDEAPP:
                    // rude app is getting displayed, hide the balloon
                    if (wParam && _pinfo)
                        _ShowInfoTip(_pinfo->nIcon, FALSE, FALSE);
                    break;

                // only button down, mouse move msgs are forwarded down to us from info tip
                //case WM_LBUTTONUP:
                //case WM_MBUTTONUP:
                //case WM_RBUTTONUP:
                case WM_LBUTTONDOWN:
                case WM_MBUTTONDOWN:
                case WM_RBUTTONDOWN:
                    _InfoTipMouseClick(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
                    break;

                default:
                    goto DoDefault;
            }
        }
    DoDefault:
        return (DefWindowProc(hWnd, uMsg, wParam, lParam));
    }

    return 0;
}

LRESULT CTrayNotify::s_ToolbarWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    return ((CTrayNotify *)dwRefData)->_ToolbarWndProc(hwnd, uMsg, wParam, lParam, uIdSubclass, dwRefData);
}

BOOL CTrayNotify::TrayNotifyIcon(PTRAYNOTIFYDATA pnid, BOOL* pbRefresh)
{
    PNOTIFYICONDATA32 pNID = &pnid->nid;
    if (pNID->cbSize<sizeof(NOTIFYICONDATA32))
    {
        return(FALSE);
    }
    INT_PTR nIcon = _FindNotify(pNID);

    switch (pnid->dwMessage)
    {
        case NIM_SETFOCUS:
            // the notify client is trying to return focus to us
            if (nIcon >= 0) {
                SetForegroundWindow(v_hwndTray);
                SetFocus(_hwndToolbar);
                SendMessage(_hwndToolbar, TB_SETHOTITEM, nIcon, 0);
                return TRUE;
            }
            return FALSE;

        case NIM_ADD:
            _iVisCount = -1;
            if (nIcon >= 0)
            {
                // already there
                return(FALSE);
            }

            if (!_InsertNotify(pNID))
            {
                return(FALSE);
            }
            // if balloon is up we have to move it, but we cannot straight up
            // position it because traynotify will be moved around by tray
            // so we do it async
            if (_pinfo)
                PostMessage(_hwndNotify, TNM_ASYNCINFOTIPPOS, (WPARAM)_pinfo->nIcon, 0);
            break;

        case NIM_MODIFY:
            if (nIcon < 0)
            {
                return(FALSE);
            }

            if (!_ModifyNotify(pNID, nIcon, pbRefresh))
            {
                return(FALSE);
            }
            // see comment above
            if (_pinfo)
                PostMessage(_hwndNotify, TNM_ASYNCINFOTIPPOS, (WPARAM)_pinfo->nIcon, 0);
            break;

        case NIM_DELETE:
            if (nIcon < 0)
            {
                return(FALSE);
            }

            _DeleteNotify(nIcon);
            _Size();
            break;

        case NIM_SETVERSION:
            if (nIcon < 0)
            {
                return(FALSE);
            }

            return _SetVersionNotify(pNID, nIcon);

        default:
            return(FALSE);
    }

    return(TRUE);
}

int CTrayNotify::DeleteDPAPtrCB(void* pItem, void* pData)
{
    LocalFree(pItem);
    return TRUE;
}

BOOL CTrayNotify::_IsScreenSaverRunning()
{
    BOOL fRunning;
    if (SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0, &fRunning, 0))
    {
        return fRunning;
    }

    return FALSE;
}
