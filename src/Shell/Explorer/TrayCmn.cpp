#include "pch.h"
#include "traycmn.h"
#include "trayreg.h"

//
// CNotificationItem - encapsulate the data needed to communicate between the tray
// and the tray properties dialog
//

CNotificationItem::CNotificationItem()
{
    _Init();
}

CNotificationItem::CNotificationItem(const NOTIFYITEM& no)
{
    _Init();
    CopyNotifyItem(no);
}

CNotificationItem::CNotificationItem(const CNotificationItem& no)
{
    _Init();
    CopyNotifyItem(no);
}

CNotificationItem::CNotificationItem(const TNPersistStreamData* ptnpd)
{
    _Init();
    CopyPTNPD(ptnpd);
}

inline void CNotificationItem::_Init()
{
    hIcon               = nullptr;
    pszExeName          = nullptr;
    pszIconText         = nullptr;
    pszAppId            = nullptr;
    fIsExplicitAppId    = FALSE;
    guidItem            = GUID_NULL;
    nDisplayIndex       = -1;
    dwFlags             = 0;
    uVersion            = 0;
    fUseSystemTip       = FALSE;
    uCallbackMsg        = 0;
}

void CNotificationItem::CopyNotifyItem(const NOTIFYITEM& no, BOOL bInsert /* = TRUE */)
{
    hWnd = no.hWnd;
    uID = no.uID;
    if (bInsert)
        dwUserPref = no.dwUserPref;
    hIcon = CopyIcon(no.hIcon);
    SetExeName(no.pszExeName);
    SetIconText(no.pszIconText);
    SetAppId(no.pszAppId);
    fIsExplicitAppId = no.fIsExplicitAppId;
    memcpy(&guidItem, &(no.guidItem), sizeof(no.guidItem));
    nDisplayIndex = no.nDisplayIndex;
    uVersion = no.uVersion;
    fUseSystemTip = no.fUseSystemTip;
    uCallbackMsg = no.uCallbackMsg;
}

const CNotificationItem& CNotificationItem::operator=(const TNPersistStreamData* ptnpd)
{
    if (ptnpd)
    {
        _Free();
        _Init();
        CopyPTNPD(ptnpd);
    }
    return *this;
}

const CNotificationItem& CNotificationItem::operator=(const CNotificationItem& ni)
{
    _Free();
    _Init();
    CopyNotifyItem(ni, FALSE);
    return *this;
}

void CNotificationItem::CopyPTNPD(const TNPersistStreamData* ptnpd)
{
    if (ptnpd)
    {
        hWnd = nullptr;
        uID = ptnpd->uID;
        dwUserPref = ptnpd->dwUserPref;
        hIcon = nullptr;

        SetExeName(ptnpd->szExeName);
        SetIconText(ptnpd->szIconText);

        memcpy(&guidItem, &ptnpd->guidItem, sizeof(ptnpd->guidItem));

        nDisplayIndex = -1;
        dwFlags = 0;
        uVersion = 0;
        fUseSystemTip = FALSE;
        uCallbackMsg = 0;
        pszAppId = nullptr;
        fIsExplicitAppId = FALSE;
    }
}

inline void CNotificationItem::CopyBuffer(LPCTSTR lpszSrc, LPTSTR* plpszDest)
{
    if (*plpszDest)
    {
        delete[] *plpszDest;
        *plpszDest = NULL;
    }

    int nStringLen = (lpszSrc == NULL) ? 0 : lstrlen(lpszSrc);
    if (nStringLen)
    {
        *plpszDest = new TCHAR[(nStringLen + 1)];
        if (*plpszDest)
        {
            if (SUCCEEDED(StringCchCopy(*plpszDest, nStringLen+1, lpszSrc)))
                return;
            else
                delete [] *plpszDest;
        }
    }
    *plpszDest = NULL;
}

inline void CNotificationItem::SetExeName(LPCTSTR lpszExeName)
{
    // CopyBuffer(lpszExeName, &pszExeName);
    CoTaskMemFree(pszExeName);
    CoAllocStringOpt(lpszExeName, &pszExeName);
}

inline void CNotificationItem::SetIconText(LPCTSTR lpszIconText)
{
    // CopyBuffer(lpszIconText, &pszIconText);
    CoTaskMemFree(pszIconText);
    CoAllocStringOpt(lpszIconText, &pszIconText);
}

inline void CNotificationItem::SetAppId(LPCTSTR lpszAppId)
{
    CoTaskMemFree(pszAppId);
    CoAllocStringOpt(lpszAppId, &pszAppId);
}

CNotificationItem::~CNotificationItem()
{
    _Free();
}

void CNotificationItem::_Free()
{
    if (hIcon)
        DestroyIcon(hIcon);

    CoTaskMemFree(pszExeName);
    pszExeName = nullptr;

    CoTaskMemFree(pszIconText);
    pszIconText = nullptr;

    CoTaskMemFree(pszAppId);
    pszAppId = nullptr;
}

BOOL CNotificationItem::operator==(CNotificationItem& ni) const
{
    if (uID != ni.uID)
        return FALSE;

    if (!hWnd)
    {
        return pszExeName && ni.pszExeName && !lstrcmpiW(pszExeName, ni.pszExeName);
    }
    return hWnd == ni.hWnd;
}