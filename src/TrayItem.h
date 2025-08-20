#ifndef _TRAYITEM_H
#define _TRAYITEM_H

#include "pch.h"
#include "traycmn.h"

#define TNUP_AUTOMATIC  0
#define TNUP_DEMOTED    1
#define TNUP_PROMOTED   2

// IMPORTANT : Should be kept in sync with any flags defined in shellapi.w
typedef enum ICONSTATEFLAG
{
    TIF_HIDDEN = 0,
    TIF_DEMOTED,
    TIF_STARTUPICON,
    TIF_SHARED,
    TIF_SHAREDICONSOURCE,
    TIF_ONCEVISIBLE,
    TIF_ITEMCLICKED,
    TIF_ITEMSAMEICONMODIFY,
} ICONSTATEFLAG;

#if 0
// Stolen from ep_taskbar.
class CTrayItemIdentity
{
public:
    CTrayItemIdentity() = default;

    CTrayItemIdentity(const CTrayItemIdentity& other) // TODO Pass-by-reference type assumed
    {
        _guidItem = other._guidItem;
        _hWnd = other._hWnd;
        _uID = other._uID;
        _uVersion = other._uVersion;
        _uCallbackMsg = other._uCallbackMsg;
        StringCchCopyW(_szExeName, MAX_PATH, other._szExeName);
    }

    CTrayItemIdentity(const NOTIFYITEM& notifyItem)
    {
        _guidItem = notifyItem.guidItem;
        _hWnd = notifyItem.hWnd;
        _uID = notifyItem.uID;
        _uVersion = notifyItem.uVersion;
        _uCallbackMsg = notifyItem.uCallbackMsg;
        StringCchCopyW(_szExeName, MAX_PATH, notifyItem.pszExeName);
    }

    CTrayItemIdentity(const CTrayItem* trayItem)
    {
        _guidItem = trayItem->m_guidItem;
        _hWnd = trayItem->m_hWnd;
        _uID = trayItem->m_uID;
        _uVersion = trayItem->m_uVersion;
        _uCallbackMsg = trayItem->m_uCallbackMsg;
        StringCchCopyW(_szExeName, MAX_PATH, trayItem->m_szExeName);
    }

    bool operator==(const CTrayItemIdentity& other) const
    {
        if (!IsEqualGUID(GUID_NULL, _guidItem)
            && IsEqualGUID(_guidItem, other._guidItem))
        {
            return true;
        }
        return _hWnd == other._hWnd && _uID == other._uID;
    }

    bool operator!=(const CTrayItemIdentity& other) const
    {
        if (!IsEqualGUID(_guidItem, other._guidItem)
            && !IsEqualGUID(_guidItem, GUID_NULL)
            && !IsEqualGUID(other._guidItem, GUID_NULL))
        {
            return true;
        }
        return _hWnd != other._hWnd && _uID != other._uID;
    }

    HRESULT GetApplicationIdentity(WCHAR *szAppIdOut, size_t cch, bool *pfOutIsExplicitAppId);
    HRESULT GetExeName(WCHAR *pszExeName, size_t cch) const;
    GUID GetGuid() const;

private:
    GUID _guidItem = GUID_NULL;
    HWND _hWnd = nullptr;
    WCHAR _szExeName[MAX_PATH];
    UINT _uID = 0;
    UINT _uVersion = 3;
    UINT _uCallbackMsg = 0; // TODO Check values set by constructors
};
#endif

class CTrayItem
{
    public:
        CTrayItem() : dwUserPref(TNUP_AUTOMATIC), guidItem(GUID_NULL) { }
        ~CTrayItem() { }

        // The icon should have an associated exe name, and it should either be
        // demoted or its preference should be set to something other than automatic
        BOOL ShouldSaveIcon()
        {
            return ( (szExeName[0] != 0) && (WasOnceVisible()) && (szIconText[0] != 0) ); 
        }

        BOOL IsDemoted()            { return (_CheckIconState(TIF_DEMOTED)); }
        BOOL IsHidden()             { return (_CheckIconState(TIF_HIDDEN)); }
        BOOL IsIconShared()         { return (_CheckIconState(TIF_SHARED)); }
        BOOL IsSharedIconSource()   { return (_CheckIconState(TIF_SHAREDICONSOURCE)); }
        BOOL IsStartupIcon()        { return (_CheckIconState(TIF_STARTUPICON)); }
        BOOL WasOnceVisible()       { return (_CheckIconState(TIF_ONCEVISIBLE)); }
        BOOL IsItemClicked()        { return (_CheckIconState(TIF_ITEMCLICKED)); }
        BOOL IsItemSameIconModify() { return (_CheckIconState(TIF_ITEMSAMEICONMODIFY)); }

        BOOL IsGuidItemValid()      { return (guidItem != GUID_NULL); }

        BOOL IsIconTimerCurrent()   { return (uIconDemoteTimerID != 0); }

        void SetDemoted(BOOL bDemoted) { _SetIconState(TIF_DEMOTED, bDemoted); }
        void SetStartupIcon(BOOL bIsStartupIcon) { _SetIconState(TIF_STARTUPICON, bIsStartupIcon); }
        void SetSharedIconSource(BOOL bSharedIconSource) 
        {
            _SetIconState(TIF_SHAREDICONSOURCE, bSharedIconSource); 
        }
        void SetOnceVisible(BOOL bOnceVisible) { _SetIconState(TIF_ONCEVISIBLE, bOnceVisible); }
        void SetItemClicked(BOOL bItemClicked) { _SetIconState(TIF_ITEMCLICKED, bItemClicked); }
        void SetItemSameIconModify(BOOL bItemSameIconModify)
        {
            _SetIconState(TIF_ITEMSAMEICONMODIFY, bItemSameIconModify);
        }
   
    public:
        HWND        hWnd;
        UINT        uID;
        UINT        uCallbackMessage;
        DWORD       dwState;
        UINT        uVersion;
        HICON       hIcon;        // *** may be stale, not guaranteed to be a valid hicon ***
        HICON       hBalloonIcon; // *** may be stale, not guaranteed to be a valid hicon ***

        ULONG       uIconDemoteTimerID;
        DWORD       dwUserPref;  // user preference (hidden? visible? automatic?)
        DWORD       dwLastSoundTime;
        TCHAR       szExeName[MAX_PATH];
        TCHAR       szIconText[MAX_PATH];
        UINT        uNumSeconds;
        GUID        guidItem;

        DWORD _GetStateFlag(ICONSTATEFLAG sf); // XXX(isabella): Made public so CTrayItemManager can access.

    private:
        void _SetIconState(ICONSTATEFLAG sf, BOOL bSet);
        BOOL _CheckIconState(ICONSTATEFLAG sf);
};

// Item Count helper flags..
#define     GIC_PROMOTED       0
#define     GIC_DEMOTED        1
#define     GIC_ALL            2

class CTrayItemManager
{
    public:
        CTrayItemManager()  { }
        ~CTrayItemManager() { }

        INT_PTR FindItemAssociatedWithGuid(GUID guidItemToCheck);
        INT_PTR FindItemAssociatedWithTimer(UINT_PTR uIconDemoteTimerID);
        INT_PTR FindItemAssociatedWithHwndUid(HWND hwnd, UINT uID);

        INT_PTR GetItemCount(int nItemCountThreshold = -1)
        {
            return _GetItemCountHelper(GIC_ALL, nItemCountThreshold);
        }

        INT_PTR GetDemotedItemCount(int nItemCountThreshold = -1)
        {
            return _GetItemCountHelper(GIC_DEMOTED, nItemCountThreshold);
        }

        INT_PTR GetPromotedItemCount(int nItemCountThreshold = -1)
        {
            return _GetItemCountHelper(GIC_PROMOTED, nItemCountThreshold);
        }

        void SetTBBtnImage(INT_PTR iIndex, int iImage);
        int GetTBBtnImage(INT_PTR iIndex, BOOL fByIndex = TRUE);
        
        void SetTBBtnText(INT_PTR iIndex, LPTSTR pszText);

        BOOL SetTBBtnStateHelper(INT_PTR iIndex, BYTE fsState, BOOL_PTR bSet);

        BOOL GetTrayItem(INT_PTR nIndex, CNotificationItem * pni, BOOL * pbStat);

        int FindImageIndex(HICON hIcon, BOOL fSetAsSharedSource);

        BOOL DemotedItemsPresent(int nMinDemotedItemsThreshold);

        CTrayItem * GetItemData(INT_PTR i, BOOL byIndex, HWND hwndToolbar);
        CTrayItem * GetItemDataByIndex(INT_PTR i) 
        {
            return GetItemData(i, TRUE, m_hwndToolbar);
        }

        void SetIconList(HIMAGELIST hImageList) { m_himlIcons = hImageList; }
        void SetTrayToolbar(HWND hwndToolbar)   { m_hwndToolbar = hwndToolbar; }

    private:
        INT_PTR _GetItemCountHelper(int nItemFlag, int nItemCountThreshold);

        HIMAGELIST      m_himlIcons;
        HWND            m_hwndToolbar;
};

#endif // _TRAYITEM_H
