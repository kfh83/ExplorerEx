#include "pch.h"
#include "cocreateinstancehook.h"
#include "shundoc.h"
#include "stdafx.h"
#include "specfldr.h"
#include "hostutil.h"
#include "cowsite.h"
#include "startmnu.h"

//
//  This definition is stolen from shell32\unicpp\dcomp.h
//
#define REGSTR_PATH_HIDDEN_DESKTOP_ICONS_STARTPANEL \
     REGSTR_PATH_EXPLORER TEXT("\\HideDesktopIcons\\NewStartPanel")

HRESULT CRecentShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
HRESULT CNoSubdirShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
HRESULT CMyComputerShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
HRESULT CNoFontsShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
HRESULT CConnectToShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
LPTSTR _Static_LoadString(const struct SpecialFolderDesc *pdesc);
BOOL ShouldShowWindowsSecurity();
BOOL ShouldShowOEMLink();
BOOL ShouldShowSPADLink();

typedef HRESULT (CALLBACK *CREATESHELLMENUCALLBACK)(IShellMenuCallback **ppsmc);
typedef BOOL (CALLBACK *SHOULDSHOWFOLDERCALLBACK)();

EXTERN_C HINSTANCE g_hinstCabinet;

void ShowFolder(UINT csidl);
BOOL IsNetConPidlRAS(IShellFolder2 *psfNetCon, LPCITEMIDLIST pidlNetConItem);

//****************************************************************************
//
//  SpecialFolderDesc
//
//  Describes a special folder.
//

#define SFD_SEPARATOR       ((LPTSTR)-1)

// Flags for SpecialFolderDesc._uFlags

enum {
    // These values are collectively known as the
    // "display mode" and match the values set by regtreeop.

    SFD_HIDE     = 0x0000,
    SFD_SHOW     = 0x0001,
    SFD_CASCADE  = 0x0002,
    SFD_MODEMASK = 0x0003,

    SFD_DROPTARGET      = 0x0004,
    SFD_CANCASCADE      = 0x0008,
    SFD_FORCECASCADE    = 0x0010,
    SFD_BOLD            = 0x0020,
    SFD_WASSHOWN        = 0x0040,
    SFD_PREFIX          = 0x0080,
    SFD_USEBGTHREAD     = 0x0100,
};

#pragma region ResourceStringHelpers

template <typename T>
HRESULT TCoTaskMemAllocCb(SIZE_T cb, T **out)
{
    T *p = (T *)CoTaskMemAlloc(cb);
    HRESULT hr = p ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        *out = p;
    }
    return S_OK;
}

HRESULT CALLBACK _ResourceStringAllocCopyExCoAlloc(HANDLE hHeap, SIZE_T cb, WCHAR **ppsz)
{
    return TCoTaskMemAllocCb(cb, ppsz);
}

template <typename T>
HRESULT TResourceStringAllocCopyEx(HMODULE hInstance,
    UINT uId,
    WORD wLanguage,
    HRESULT(CALLBACK *pfnAlloc)(HANDLE, SIZE_T, T *),
    HANDLE hHeap,
    T *pt);

HRESULT ResourceStringCoAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR **ppsz)
{
    return TResourceStringAllocCopyEx(hModule, uId, wLanguage, _ResourceStringAllocCopyExCoAlloc, nullptr, ppsz);
}

HRESULT ResourceStringCoAllocCopy(HINSTANCE hModule, UINT uId, WCHAR **ppsz)
{
    return ResourceStringCoAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

#pragma endregion

signed int ResultFromLastError()
{
    signed int result = GetLastError();
    if (result > 0)
    {
        return (unsigned __int16)result | 0x80070000;
    }
    return result;
}

HRESULT SHFormatMessageArg(DWORD dwFlags, const void* lpSource, DWORD dwMessageId, DWORD dwLangID, WCHAR* pszBuffer, DWORD cchSize, ...)
{
    va_list va;
    va_start(va, cchSize);

    va_list vaParamList;
    va_copy(vaParamList, va);
    if (FormatMessage(dwFlags, lpSource, dwMessageId, dwLangID, pszBuffer, cchSize, &vaParamList))
    {
        return S_OK;
    }
    return ResultFromLastError();
}

DEFINE_GUID(POLID_NoUserFolderInStartMenu, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoSMMyDocs, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoSMMyPictures, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuMyMusic, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuHomegroup, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuVideos, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuDownloads, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuRecordedTV, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuMyGames, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoFavoritesMenu, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoRecentDocsMenu, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoMyComputerIcon, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoStartMenuNetworkPlaces, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoNetworkConnections, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoSetFolders, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoControlPanel, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoSMConfigurePrograms, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoSMHelp, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoRun, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_ForceRunOnStartMenu, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
DEFINE_GUID(POLID_NoNTSecurity, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

#include <evntprov.h>

class SpecialFolderListItem;

class CMenuDescriptor
{
public:
    typedef int(CMenuDescriptor::* CUSTOMTOOLTIPCALLBACK)(SpecialFolderList*, const SpecialFolderListItem*, LPWSTR, DWORD) const;

    LPCWSTR _pszTarget;
    const KNOWNFOLDERID _kfId;
    LPCWSTR _pszPath;
    const KNOWNFOLDERID _kfMenuId;
    LPCWSTR _pszIconPath;
    LPCGUID _pguidPolicyHide;
    LPCGUID _pguidPolicyRestrict;
    LPCGUID _pguidPolicyForceShow;
    LPCWSTR _pszShow;
    UINT _uFlags;
    CREATESHELLMENUCALLBACK _CreateShellMenuCallback;
    LPCWSTR _pszCustomizeKey;
    DWORD _dwShellFolderFlags;
    DWORD _dwShellMenuSetFlags;
    UINT _idsCustomName;
    UINT _iToolTip;
    CUSTOMTOOLTIPCALLBACK _CustomTooltipCallback;
    SHOULDSHOWFOLDERCALLBACK _ShowFolder;
    LPCWSTR _pszCanHideOnDesktop;
    PCEVENT_DESCRIPTOR _pEventDescriptor;
    int _iEventId;
    DWORD _dwEventFlags;

    REFKNOWNFOLDERID GetFolderID() const
    {
        return _kfId;
    }

    REFKNOWNFOLDERID GetMenuFolderID() const
    {
        return _kfMenuId;
    }


    BOOL IsDropTarget() const { return _uFlags & SFD_DROPTARGET; }
    BOOL IsCacheable() const { return _uFlags & SFD_USEBGTHREAD; }

    BOOL IsBold() const { return _uFlags & SFD_BOLD; }
    int IsSeparator() const { return _pszTarget == SFD_SEPARATOR; }

    HRESULT CreateShellMenuCallback(IShellMenuCallback **ppsmc) const
    {
        return _CreateShellMenuCallback ? _CreateShellMenuCallback(ppsmc) : S_OK;
    }

    BOOL GetCustomName(LPTSTR *ppsz) const
    {
        return _idsCustomName && SUCCEEDED(ResourceStringCoAllocCopy(g_hinstCabinet, _idsCustomName, ppsz));
    }

	// Taken from ep_taskbar by @amrsatrio
    LPCWSTR GetItemName() const
    { 
        return _pszPath == SFD_SEPARATOR ? NULL : _pszPath;
    }

    // Taken from ep_taskbar by @amrsatrio
    LPWSTR GetShowCacheRegName() const
    {
        WCHAR szCached[] = TEXT("_ShouldShow");
        SIZE_T cch = lstrlen(_pszShow) + ARRAYSIZE(szCached) + 1;
        WCHAR *pszShowCache = (WCHAR *)LocalAlloc(LPTR, sizeof(WCHAR) * cch);
        if (pszShowCache)
        {
            StringCchCopy(pszShowCache, cch, _pszShow);
            StringCchCat(pszShowCache, cch, szCached);
        }
        return pszShowCache;
    }

    // Taken from ep_taskbar by @amrsatrio
    DWORD GetDisplayMode(int *bOutIgnoreRule)
    {
       
        *bOutIgnoreRule = FALSE;

        if (_pguidPolicyHide && SHWindowsPolicy(*_pguidPolicyHide)
            || _pguidPolicyRestrict && SHWindowsPolicy(*_pguidPolicyRestrict))
        {
            return SFD_HIDE;
        }

        DWORD dwMode = _uFlags & SFD_MODEMASK;

        if (_pszShow)
        {
            DWORD dwNewMode;
            DWORD cb = sizeof(dwNewMode);
            if (_SHRegGetValueFromHKCUHKLM(L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _pszShow, SRRF_RT_DWORD, NULL, &dwNewMode, &cb) == ERROR_SUCCESS)
            {
                dwMode = dwNewMode;
                *bOutIgnoreRule = TRUE;
            }
            else
            {
                WCHAR *pszShowCache = GetShowCacheRegName();
                if (pszShowCache)
                {
                    cb = sizeof(dwNewMode);
                    if (SHGetValue(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", pszShowCache, NULL, &dwNewMode, &cb) == ERROR_SUCCESS)
                    {
                        dwMode = dwNewMode;
                    }
                    LocalFree(pszShowCache);
                }
            }
        }

        if (_pguidPolicyForceShow && SHWindowsPolicy(*_pguidPolicyForceShow))
        {
            dwMode = SFD_SHOW;
        }

        dwMode &= SFD_MODEMASK;

        if (dwMode == SFD_CASCADE && (_uFlags & SFD_CANCASCADE) == 0)
        {
            dwMode = SFD_SHOW;
        }
        else if (dwMode == SFD_SHOW && (_uFlags & SFD_FORCECASCADE) != 0)
        {
            dwMode = SFD_CASCADE;
        }

        return dwMode;
    }

    void AdjustForSKU(DWORD dwProductType)
    {
        LSTATUS Value; // eax
        DWORD v5; // [esp+10h] [ebp-28h] SPLIT BYREF
        int v6; // [esp+14h] [ebp-24h]
        BOOL bIsServer; // [esp+18h] [ebp-20h]
        int v8; // [esp+1Ch] [ebp-1Ch]
        //CPPEH_RECORD ms_exc; // [esp+20h] [ebp-18h]

        v6 = 1;
        bIsServer = IsOS(OS_ANYSERVER);
        v8 = 0;
        if (dwProductType == 4 || dwProductType == 6 || dwProductType == 16 || dwProductType == 27)
            v6 = 0;
        if (bIsServer)
            v6 = 0;
        if (IsEqualGUID(_kfId, GUID_NULL))
        {
            if (_pszTarget == (LPCWSTR)-1 || !_pszTarget)// -1 = SFD_SEPARATOR
            {
                if (_pszPath && !StrCmpICW(L"Microsoft.AdministrativeTools", _pszPath) && bIsServer)
                {
                    _uFlags = _uFlags & 0xFFFFFFFC | 2;
                    goto LABEL_33;
                }
                goto LABEL_34;
            }
            //if (memcmp(&_kfId, &GUID_NULL, 0x10u)
            //    && CcshellAssertFailedW(
            //        L"d:\\longhorn\\shell\\explorer\\desktop2\\specfldr.cpp",
            //        673,
            //        L"!IsSeparator() && !HasFolderID()",
            //        0))
            //{
            //    AttachUserModeDebugger();
            //    do
            //    {
            //        __debugbreak();
            //        ms_exc.registration.TryLevel = -2;
            //    } while (dword_108BAD4);
            //}
            if (!StrCmpIC(TEXT("::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}"), _pszTarget == SFD_SEPARATOR ? NULL : _pszTarget))
            {
                if (bIsServer)
                {
                    _uFlags = _uFlags & 0xFFFFFFFC | 1;
                    v8 = 1;
                }
                goto LABEL_34;
            }
            if (StrCmp(
                TEXT("::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}"),
                _pszTarget == SFD_SEPARATOR ? NULL : _pszTarget))
            {
                goto LABEL_34;
            }
        }
        else if (!IsEqualGUID(GetFolderID(), FOLDERID_Profile)
            && !IsEqualGUID(GetFolderID(), FOLDERID_Pictures)
            && !IsEqualGUID(GetFolderID(), FOLDERID_Music)
            && !IsEqualGUID(GetFolderID(), FOLDERID_Recent))
        {
            if (IsEqualGUID(GetFolderID(), FOLDERID_Games) && !v6)
                goto LABEL_15;
            goto LABEL_34;
        }
        if (bIsServer)
        {
        LABEL_15:
            _uFlags &= 0xFFFFFFFC;
        LABEL_33:
            v8 = 1;
        }
    LABEL_34:
        if (v8)
        {
            v5 = 4;
            Value = SHRegGetValue(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _pszShow, 24, nullptr, &dwProductType, &v5);
            if (Value == 2 || Value == 3)
            {
                dwProductType = _uFlags & 3;
                _SHRegSetValue(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", _pszShow, 24, 4u, &dwProductType, 4u);
            }
        }
    }

    BOOL UserTooltip(SpecialFolderList* pList, const SpecialFolderListItem* pitem, WCHAR* pszText, DWORD cch) const
    {
        BOOL bRet = FALSE;

        WCHAR szText[256];
        if (pList->GetLVText((PaneItem*)pitem, szText, ARRAYSIZE(szText)) >= 0)
        {
            WCHAR szBuf[256];
            LoadString(g_hinstCabinet, 317, szBuf, ARRAYSIZE(szBuf));
            bRet = SUCCEEDED(SHFormatMessageArg(0x400, szBuf, 0, 0, pszText, cch, szText)) ? TRUE : FALSE;
        }
        return bRet;
    }
};

static CMenuDescriptor s_rgsfd[] =
{
    { // User (00)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_UsersFiles,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoUserFolderInStartMenu,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowUser"),
        /* _uFlags */                       0x609,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        &CMenuDescriptor::UserTooltip,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          TEXT("{59031a47-3f72-44a7-89c5-5595fe6b30ee}"),
        /* _pEventDescriptor */             nullptr, /*ExplorerFrame_OpenProfile*/
        /* _iEventId */                     582,
        /* _dwEventFlags */                 0
    },
    { // Documents (01)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Documents,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoSMMyDocs,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowMyDocs"),
        /* _uFlags */                       0x1E09,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     IDS_CUSTOMTIP_MYDOCS,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*ExplorerFrame_OpenDocuments*/
        /* _iEventId */                     129,
        /* _dwEventFlags */                 0x10
    },
    { // Pictures (02)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Pictures,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoSMMyPictures,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowMyPics"),
        /* _uFlags */                       0x1E09,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     IDS_CUSTOMTIP_MYPICS,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*ExplorerFrame_OpenPictures*/
        /* _iEventId */                     525,
        /* _dwEventFlags */                 0x40
    },
    { // Music (03)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Music,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoStartMenuMyMusic,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowMyMusic"),
        /* _uFlags */                       0x1E09,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     IDS_CUSTOMTIP_MYMUSIC,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*ExplorerFrame_OpenMusic*/
        /* _iEventId */                     524,
        /* _dwEventFlags */                 0x20
    },
    { // Games (04)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Games,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoStartMenuMyGames,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowMyGames"),
        /* _uFlags */                       0x0E09,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     312,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Games_Launch*/
        /* _iEventId */                     856,
        /* _dwEventFlags */                 0x8000
    },
    { // Favorites (05)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Favorites,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoFavoritesMenu,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("StartMenuFavorites"),
        /* _uFlags */                       0x1A98,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Favorites"),
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                IDS_STARTPANE_FAVORITES,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Favorites_Launch*/
        /* _iEventId */                     143,
        /* _dwEventFlags */                 0x4
    },
	{ // Separator (06)
        /* _pszTarget */                    SFD_SEPARATOR,
        /* _kfId */                         {},
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              nullptr,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      nullptr,
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr,
        /* _iEventId */                     0,
        /* _dwEventFlags */                 0
    },
    { // Recent (07)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_Recent,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  TEXT("::{0c39a5cf-1a7a-40c8-ba74-8900e6df5fcd}"),
        /* _pguidPolicyHide */              &POLID_NoRecentDocsMenu,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_TrackDocs"),
        /* _uFlags */                       0x9A,
        /* _CreateShellMenuCallback */      CRecentShellMenuCallback_CreateInstance,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0x2,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                IDS_STARTPANE_RECENT,
        /* _iToolTip */                     IDS_CUSTOMTIP_RECENT,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_RecentItems_Launch*/
        /* _iEventId */                     523,
        /* _dwEventFlags */                 0x4000
    },
    { // Computer (08)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_ComputerFolder,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoMyComputerIcon,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowMyComputer"),
        /* _uFlags */                       0x409,
        /* _CreateShellMenuCallback */      CMyComputerShellMenuCallback_CreateInstance,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     IDS_CUSTOMTIP_MYCOMP,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          TEXT("{20D04FE0-3AEA-1069-A2D8-08002B30309D}"),
        /* _pEventDescriptor */             nullptr, /*ExplorerFrame_OpenComputer*/
        /* _iEventId */                     522,
        /* _dwEventFlags */                 0x1
    },
    { // Network (09)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         FOLDERID_NetworkFolder,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoStartMenuNetworkPlaces,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowNetPlaces"),
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                IDS_STARTPANE_CONNECTTO,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Network_Launch*/
        /* _iEventId */                     563,
        /* _dwEventFlags */                 0x100
    },
    { // Network Connections (10)
        /* _pszTarget */                    TEXT("::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoNetworkConnections,
        /* _pguidPolicyRestrict */          &POLID_NoSetFolders,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowNetConn"),
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     318,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_ConnectTo_Launch*/
        /* _iEventId */                     564,
        /* _dwEventFlags */                 0x80
    },
    { // Separator (11)
        /* _pszTarget */                    SFD_SEPARATOR,
        /* _kfId */                         {},
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              nullptr,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      nullptr,
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr,
        /* _iEventId */                     0,
        /* _dwEventFlags */                 0
    },
    { // Control Panel (12)
        /* _pszTarget */                    TEXT("::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     FOLDERID_ControlPanelFolder,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoControlPanel,
        /* _pguidPolicyRestrict */          &POLID_NoSetFolders,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowControlPanel"),
        /* _uFlags */                       0x89,
        /* _CreateShellMenuCallback */      nullptr, //CNoSubMenuShellMenuCallback_CreateInstance,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                IDS_STARTPANE_CONTROLPANEL,
        /* _iToolTip */                     IDS_CUSTOMTIP_CTRLPANEL,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          TEXT("{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"),
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_ControlPanel_Launch*/
        /* _iEventId */                     128,
        /* _dwEventFlags */                 0x2
    },
    { // Set Program Access and Defaults (13)
        /* _pszTarget */                    TEXT("::{E44E5D18-0652-4508-A4E2-8A090067BCB0}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoSMConfigurePrograms,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowSetProgramAccessAndDefaults"),
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     316,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   ShouldShowSPADLink,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_SPAD_Launch*/
        /* _iEventId */                     133,
        /* _dwEventFlags */                 0x1000
    },
    { // Administrative Tools (14)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      TEXT("Microsoft.AdministrativeTools"),
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              nullptr,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_AdminToolsRoot"),
        /* _uFlags */                       0x18,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_AdminTools_Launch*/
        /* _iEventId */                     140,
        /* _dwEventFlags */                 0x2000
    },
    { // Printers (15)
        /* _pszTarget */                    nullptr,
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      TEXT("Microsoft.Printers"),
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              nullptr,
        /* _pguidPolicyRestrict */          &POLID_NoSetFolders,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowPrinters"),
        /* _uFlags */                       0,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     319,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Printers_Launch*/
        /* _iEventId */                     138,
        /* _dwEventFlags */                 0x400
    },
    { // Help (16)
        /* _pszTarget */                    TEXT("::{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoSMHelp,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      TEXT("Start_ShowHelp"),
        /* _uFlags */                       0x81,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                /*0*/IDS_STARTPANE_TITLE_HELP, // @MOD: Use string directly in binary instead of reading from registry which points to normal explorer.exe
        /* _iToolTip */                     /*0*/IDS_STARTPANE_INFOTIP_HELP, // @MOD: Use string directly in binary instead of reading from registry which points to normal explorer.exe
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Help_Launch*/
        /* _iEventId */                     131,
        /* _dwEventFlags */                 0x8
    },
    { // Run (17)
        /* _pszTarget */                    TEXT("::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoRun,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         &POLID_ForceRunOnStartMenu,
        /* _pszShow */                      TEXT("Start_ShowRun"),
        /* _uFlags */                       0x80,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr, /*Explorer_StartMenu_Run_Launch*/
        /* _iEventId */                     139,
        /* _dwEventFlags */                 0
    },
    { // Separator (18)
        /* _pszTarget */                    SFD_SEPARATOR,
        /* _kfId */                         {},
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     0,
        /* _pszIconPath */                  0,
        /* _pguidPolicyHide */              0,
        /* _pguidPolicyRestrict */          0,
        /* _pguidPolicyForceShow */         0,
        /* _pszShow */                      0,
        /* _uFlags */                       SFD_SHOW,
        /* _CreateShellMenuCallback */      0,
        /* _pszCustomizeKey */              0,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                nullptr,
        /* _iToolTip */                     nullptr,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   nullptr,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             0,
        /* _iEventId */                     nullptr,
        /* _dwEventFlags */                 nullptr
    },
    { // Windows Security (19)
        /* _pszTarget */                    TEXT("::{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}"),
        /* _kfId */                         GUID_NULL,
        /* _pszPath */                      nullptr,
        /* _kfMenuId */                     GUID_NULL,
        /* _pszIconPath */                  nullptr,
        /* _pguidPolicyHide */              &POLID_NoNTSecurity,
        /* _pguidPolicyRestrict */          nullptr,
        /* _pguidPolicyForceShow */         nullptr,
        /* _pszShow */                      nullptr,
        /* _uFlags */                       0x81,
        /* _CreateShellMenuCallback */      nullptr,
        /* _pszCustomizeKey */              nullptr,
        /* _dwShellFolderFlags */           0,
        /* _dwShellMenuSetFlags */          0,
        /* _idsCustomName */                0,
        /* _iToolTip */                     0,
        /* _CustomTooltipCallback */        nullptr,
        /* _ShowFolder */                   ShouldShowWindowsSecurity,
        /* _pszCanHideOnDesktop */          nullptr,
        /* _pEventDescriptor */             nullptr,
        /* _iEventId */                     0,
        /* _dwEventFlags */                 0,
    }
};

// d:\\longhorn\\Shell\\inc\\idllib.h
PCIDLIST_ABSOLUTE _SHILMakeFull(const void *pv)
{
    PCIDLIST_ABSOLUTE pidl = reinterpret_cast<PCIDLIST_ABSOLUTE>(pv);
    //RIP(ILIsAligned(reinterpret_cast<PCUIDLIST_RELATIVE>(pidl))); // 183
    return pidl;
}

STDAPI BindCtx_SetMode(IBindCtx *pbcIn, DWORD grfMode, IBindCtx **ppbcOut)
{
    HRESULT hr = S_OK;

    *ppbcOut = pbcIn;
    if (pbcIn)
    {
        pbcIn->AddRef();
    }
    else
    {
        hr = CreateBindCtx(0, ppbcOut);
    }

    if (SUCCEEDED(hr))
    {
        BIND_OPTS bo;
        bo.cbStruct = sizeof(bo);
        bo.grfFlags = 0;
        bo.grfMode = grfMode;
        bo.dwTickCountDeadline = 0;
        if (pbcIn)
        {
            hr = pbcIn->GetBindOptions(&bo);
            bo.grfMode = grfMode;
        }

        if (SUCCEEDED(hr)) // @Note: This check is new in Windows 10
            hr = (*ppbcOut)->SetBindOptions(&bo);

        if (FAILED(hr))
			IUnknown_SafeReleaseAndNullPtr(ppbcOut);
    }

    return hr;
}

HRESULT BindCtx_CreateWithMode(DWORD grfMode, IBindCtx **ppbc)
{
    return BindCtx_SetMode(nullptr, grfMode, ppbc);
}

//****************************************************************************
//
//  SpecialFolderListItem
//
//  A PaneItem for the benefit of SFTBarHost.
//

class SpecialFolderListItem : public PaneItem
{

public:
	ITEMIDLIST_ABSOLUTE *_pidl;         // Full Pidl to each item
	ITEMIDLIST_ABSOLUTE *_pidlCascade;  // Pidl for cascade menu
	ITEMIDLIST_ABSOLUTE *_pidlSimple;   // Pidl for "simple" items with a KNOWNFOLDERID
	const CMenuDescriptor *_psfd;       // Describes this item
	int _iImageIndex;                   // Index of private icon
	WCHAR _chMnem;                      // Keyboard accelerator
	LPWSTR _pszDispName;                // Display name

    SpecialFolderListItem(const CMenuDescriptor *psfd)
        : _pidl(nullptr)
        , _psfd(psfd)
    {
        if (_psfd->IsSeparator())
        {
            _iPinPos = PINPOS_SEPARATOR;
        }
        else if (IsEqualGUID(_psfd->GetFolderID(), GUID_NULL))
        {
            if (_psfd->_pszTarget)
            {
                IBindCtx *pbc;
                if (SUCCEEDED(BindCtx_CreateWithMode(STGM_CREATE, &pbc)))
                {
                    SHParseDisplayName(_psfd->GetItemName(), pbc, &_pidl, 0, nullptr);
                    pbc->Release();
                }
            }

            IOpenControlPanel *pocp = nullptr;
            if (SUCCEEDED(CoCreateInstance(CLSID_OpenControlPanel, NULL,
                CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, IID_PPV_ARGS(&pocp))))
            {
                WCHAR szPath[MAX_PATH];
                if (SUCCEEDED(pocp->GetPath(_psfd->_pszPath, szPath, ARRAYSIZE(szPath))))
                {
                    SHParseDisplayName(szPath, nullptr, &_pidl, 0, nullptr);
                }
            }

            if (pocp)
            {
                pocp->Release();
            }
        }
        else if ((_psfd->_uFlags & 0x1000) != 0) // Probably something like "_psfd->IsSimple()"
        {
            SHGetKnownFolderIDList(_psfd->GetFolderID(), KF_FLAG_SIMPLE_IDLIST | KF_FLAG_DONT_VERIFY, nullptr, &_pidlSimple);
        }
        else
        {
            SHGetKnownFolderIDList(_psfd->GetFolderID(), 0, nullptr, &_pidl);
        }

        if (!IsEqualGUID(_psfd->GetMenuFolderID(), GUID_NULL))
        {
            SHGetKnownFolderIDList(_psfd->GetMenuFolderID(), 0, nullptr, &_pidlCascade);
        }

        _iImageIndex = -1;
    }

    ~SpecialFolderListItem() override
    {
        ILFree(_pidl);
        ILFree(_pidlCascade);
        ILFree(_pidlSimple);
        CoTaskMemFree(_pszDispName);
    }

    int GetPrivateIcon() override
    {
        return _iImageIndex;
    }

    void ReplaceLastPidlElement(LPCITEMIDLIST pidlNew)
    {
        ASSERT(ILIsChild(pidlNew)); // 873
        ILRemoveLastID(_pidl);
        PIDLIST_ABSOLUTE pidlCombined = (PIDLIST_ABSOLUTE)_SHILMakeFull(ILAppendID(_pidl, &pidlNew->mkid, TRUE));
        if (pidlCombined)
        {
            _pidl = pidlCombined;
        }
    }
};

HRESULT DisplayNameOfAsString(IShellFolder *psf, const ITEMIDLIST_RELATIVE *pidl, SHGDNF flags, WCHAR **ppsz);

HRESULT SpecialFolderList::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut)
{
    HRESULT hr = E_INVALIDARG;

    if (pguidCmdGroup)
    {
        if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup) && nCmdID == 329)
        {
            CKnownFolderInformation *pkfi = new CKnownFolderInformation();
            if (pkfi)
            {
                hr = SHGetKnownFolderIDList(FOLDERID_UsersFiles, 0, 0, &pkfi->_pidl);
                if (SUCCEEDED(hr))
                {
                    IShellFolder *psf;
                    LPCITEMIDLIST pidl;
                    hr = SHBindToParent(pkfi->_pidl, IID_PPV_ARGS(&psf), &pidl);
                    if (SUCCEEDED(hr))
                    {
                        hr = DisplayNameOfAsString(psf, pidl, 0, &pkfi->_pszDispName);
                        if (SUCCEEDED(hr))
                        {
                            PostMessage(_hwnd, SBM_REBUILDMENU, (WPARAM)pkfi, 0);
                            pkfi = nullptr;
                        }
                        psf->Release();
                    }
                }

                if (pkfi)
                {
                    delete pkfi;
                }
            }
        }
    }
    return hr;
}

typedef DWORD(*SHExpandEnvironmentStringsW_t)(const WCHAR *pszIn, WCHAR *pszOut, DWORD cchOut);
DWORD SHExpandEnvironmentStringsW(const WCHAR *pszIn, WCHAR *pszOut, DWORD cchOut)
{
    static SHExpandEnvironmentStringsW_t fn = nullptr;
    if (!fn)
    {
        HMODULE h = GetModuleHandleW(L"shlwapi.dll");
        if (h)
            fn = (SHExpandEnvironmentStringsW_t)GetProcAddress(h, MAKEINTRESOURCEA(460));
    }
    return fn ? fn(pszIn, pszOut, cchOut) : 0;
}

HRESULT SpecialFolderList::CLoadFullPidlTask::InternalResumeRT()
{
    CKnownFolderInformation* pkfi = new CKnownFolderInformation();
    if (pkfi)
    {
        const CMenuDescriptor* pdesc = &s_rgsfd[this->_dwIndex];
        if (SHGetKnownFolderIDList(pdesc->GetFolderID(), 0, nullptr, &pkfi->_pidl) >= 0)
        {
            IShellFolder* psf;
            LPCITEMIDLIST ppidlLast;
            if (SUCCEEDED(SHBindToParent(pkfi->_pidl, IID_PPV_ARGS(&psf), &ppidlLast)))
            {
                HRESULT hr = DisplayNameOfAsString(psf, ppidlLast, 0, &pkfi->_pszDispName);
                if (SUCCEEDED(hr))
                {
                    IKnownFolderManager* pkfm = nullptr;
                    hr = CoCreateInstance(CLSID_KnownFolderManager, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pkfm));
                    if (SUCCEEDED(hr))
                    {
                        IKnownFolder* pkf = nullptr;
                        hr = pkfm->GetFolder(pdesc->GetFolderID(), &pkf);
                        if (SUCCEEDED(hr))
                        {
                            KNOWNFOLDER_DEFINITION kfd;
                            hr = pkf->GetFolderDefinition(&kfd);
                            if (SUCCEEDED(hr))
                            {
                                if (kfd.pszLocalizedName)
                                {
                                    WCHAR szOutBuf[260];
                                    hr = SHLoadIndirectString(kfd.pszLocalizedName, szOutBuf, ARRAYSIZE(szOutBuf), nullptr);
                                    if (SUCCEEDED(hr))
                                    {
                                        if (!StrCmp(pkfi->_pszDispName, szOutBuf))
                                        {
                                            CoTaskMemFree(pkfi->_pszDispName);
                                            pkfi->_pszDispName = nullptr;
                                        }
                                    }
                                }

                                WCHAR szIconPath[260];
                                int iIndex = PathParseIconLocation(kfd.pszIcon);
                                if (SHExpandEnvironmentStringsW(kfd.pszIcon, szIconPath, ARRAYSIZE(szIconPath)))
                                {
                                    pkfi->_hIcon = _IconOf(psf, ppidlLast, _iIconSize, szIconPath, iIndex);
                                }
                                FreeKnownFolderDefinitionFields(&kfd);
                            }
                        }
                        if (pkf)
                            pkf->Release();
                    }
                    if (pkfm)
                        pkfm->Release();
                }

                psf->Release();

                if (SUCCEEDED(hr) && IsWindow(_hwnd))
                {
                    PostMessage(_hwnd, SBM_REBUILDMENU, (WPARAM)pkfi, _dwIndex);
                    pkfi = nullptr;
                }
            }
        }

        if (pkfi)
            delete pkfi;
    }
    return S_OK;
}

SpecialFolderList::~SpecialFolderList()
{
}

HRESULT SpecialFolderList::Initialize()
{
    DWORD dwProductType;
    if (RtlGetProductInfo(6, 0, 0, 0, &dwProductType) && dwProductType != 0xABCDABCD)
    {
        for (int i = 0; i < ARRAYSIZE(s_rgsfd); i++)
        {
            s_rgsfd[i].AdjustForSKU(dwProductType);
        }
    }
    return S_OK;
}

BOOL ShouldShowWindowsSecurity()
{
    return SHGetMachineInfo(GMI_TSCLIENT);
}

BOOL ShouldShowSPADLink()
{
    return !IsOS(OS_ANYSERVER);
}

DWORD WINAPI SpecialFolderList::_HasEnoughChildrenThreadProc(void* pvData)
{
    SpecialFolderList* pThis = static_cast<SpecialFolderList*>(pvData);

    HRESULT hr = SHCoInitialize();
    if (SUCCEEDED(hr))
    {
        for (DWORD dwIndex = 0; dwIndex < ARRAYSIZE(s_rgsfd); dwIndex++)
        {
            CMenuDescriptor* pdesc = &s_rgsfd[dwIndex];

            BOOL bIgnoreRule;
            DWORD dwMode = pdesc->GetDisplayMode(&bIgnoreRule);
            if (pdesc->IsCacheable() && pdesc->_ShowFolder)
            {
                ASSERT(pdesc->_pszShow); // 943

                if (!bIgnoreRule && pdesc->_ShowFolder())
                {
                    if (!(dwMode & SFD_WASSHOWN))
                    {
                        WCHAR* pszShowCache = pdesc->GetShowCacheRegName();
                        if (pszShowCache)
                        {
                            dwMode |= SFD_WASSHOWN;
                            SHSetValue(HKEY_CURRENT_USER,REGSTR_EXPLORER_ADVANCED, pszShowCache, REG_DWORD, &dwMode, sizeof(dwMode));
                            pThis->Invalidate();
                            LocalFree(pszShowCache);
                        }
                    }
                }
                else
                {
                    SpecialFolderListItem* pitem = new SpecialFolderListItem(pdesc);
                    if (pitem)
                    {
                        if (pitem->_pidl)
                        {
                            ASSERT(pThis->_cNotify < SFTHOST_MAXNOTIFY); // 973
                            if (pThis->RegisterNotify(pThis->_cNotify, 4106, pitem->_pidl, FALSE))
                            {
                                pThis->_cNotify++;
                            }
                        }
                        pitem->Release();
                    }

                    if (dwMode & SFD_WASSHOWN)
                    {
                        WCHAR* pszShowCache = pdesc->GetShowCacheRegName();
                        if (pszShowCache)
                        {
                            SHDeleteValue(HKEY_CURRENT_USER,REGSTR_EXPLORER_ADVANCED, pszShowCache);
                            pThis->Invalidate();
                            LocalFree(pszShowCache);
                        }
                    }
                }
            }
        }
    }

    SHCoUninitialize(hr);
    pThis->field_188 = 0;
    pThis->Release();
    return 0;
}

BOOL ShouldShowItem(const CMenuDescriptor* pdesc, BOOL bIgnoreRule, DWORD dwMode)
{
    if (bIgnoreRule)
        return TRUE;

    if (pdesc->_ShowFolder)
    {
        if (pdesc->IsCacheable())
        {
            return (dwMode & SFD_WASSHOWN) != 0;
        }
        return pdesc->_ShowFolder();
    }

    return TRUE;
}

#define SPECIAL_FOLDER_LIST_LOGGING

void SpecialFolderList::EnumItems()
{
    if (!field_188)
    {
        field_188 = 1;

        for (UINT id = 0; id < _cNotify; id++)
        {
            UnregisterNotify(id);
        }
        _cNotify = 0;

        AddRef();
        if (!QueueUserWorkItem(_HasEnoughChildrenThreadProc, this, 0))
        {
            Release();
        }
    }
    field_190 = 0;

    BOOL fIgnoreSeparators = TRUE;
    int iItems = 0;

    DWORD dwVisibleEventFlags = 0;
    DWORD dwCascadingEventFlags = 0;

    if (_IsPrivateImageList())
    {
        ImageList_Remove(_himl, -1);
    }

    IKnownFolderManager* pkfm = nullptr;

    for (DWORD dwIndex = 0; dwIndex < ARRAYSIZE(s_rgsfd); ++dwIndex)
    {
        CMenuDescriptor* pdesc = &s_rgsfd[dwIndex];

        BOOL bIgnoreRule;
        DWORD dwMode = pdesc->GetDisplayMode(&bIgnoreRule);

        if (dwMode != SFD_HIDE)
        {
            SpecialFolderListItem* pitem = new SpecialFolderListItem(pdesc);
            if (pitem)
            {
                if (pitem->IsSeparator() && !fIgnoreSeparators
                    || (pitem->_pidl || pitem->_pidlSimple)
                    && ShouldShowItem(pdesc, bIgnoreRule, dwMode))
                {
#ifdef SPECIAL_FOLDER_LIST_LOGGING
                    WCHAR szBuffer[512];
                    const WCHAR* pszTarget = pdesc->_pszTarget;
                    if (!pszTarget || pszTarget == SFD_SEPARATOR)
                        pszTarget = L"(null)";

                    // Format as a neat table row
                    swprintf_s(szBuffer, ARRAYSIZE(szBuffer),
                               L"[FolderList] | Index: %2d | Target: %-60ls | CustomName: %5u | Flags: 0x%04X\r\n",
                               dwIndex, pszTarget, pdesc->_idsCustomName, pdesc->_uFlags);
                    OutputDebugStringW(szBuffer);
#endif
                    if ((dwMode & SFD_MODEMASK) == SFD_CASCADE)
                    {
                        pitem->EnableCascade();
                    }

                    if (pdesc->IsDropTarget())
                    {
                        pitem->EnableDropTarget();
                    }

                    if (!pitem->IsSeparator())
                    {
                        if (pitem->_pidl)
                        {
#ifdef SPECIAL_FOLDER_LIST_LOGGING
                            wprintf(L"Creating full list item for %ls with a dwIndex of %d\n", pdesc->_pszTarget, dwIndex);
                            if (pdesc->IsSeparator())
                            {
                                wprintf(L"Creating SFD_SEPARATOR at dwIndex: %d\n", dwIndex);
                            }
#endif
                            _CreateFullListItem(pitem);
                        }
                        else
                        {
                            ASSERT(pitem->_pidlSimple != NULL); // 1422
                            ASSERT(pitem->_psfd->HasFolderID()); // 1423

                            if (!pkfm)
                            {
                                CoCreateInstance(CLSID_KnownFolderManager, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pkfm));
                            }

                            if (pkfm)
                            {
#ifdef SPECIAL_FOLDER_LIST_LOGGING
                                wprintf(L"Creating simple list item for %ls with a dwIndex of %d\n", pdesc->_pszTarget, dwIndex);
                                if (pdesc->IsSeparator())
                                {
                                    wprintf(L"Creating SFD_SEPARATOR at dwIndex: %d\n", dwIndex);
                                }
#endif
                                _CreateSimpleListItem(pkfm, pitem, dwIndex);
                            }
                        }
                    }

                    fIgnoreSeparators = pitem->IsSeparator();
                    if (AddItem(pitem) && !fIgnoreSeparators)
                    {
                        ++iItems;
                        dwVisibleEventFlags |= pitem->_psfd->_dwEventFlags;
                        if ((dwMode & SFD_MODEMASK) == SFD_CASCADE)
                        {
                            dwCascadingEventFlags |= pitem->_psfd->_dwEventFlags;
                        }
                    }
                }

                pitem->Release();
            }
        }
    }
    if (!field_18C)
    {
        field_18C = 1;
        //SHTracePerfSQMSetValueImpl(&ShellTraceId_Explorer_StartMenu_Visible_Menu_Items, 53, dwVisibleEventFlags);
        //SHTracePerfSQMSetValueImpl(&ShellTraceId_Explorer_StartMenu_Cascading_Menu_Items, 162, dwCascadingEventFlags);
    }
    SetDesiredSize(0, iItems);

    if (pkfm)
    {
        pkfm->Release();
    }
}

/*----------------------------------------------------------
  Purpose: Removes '&'s from a string, returning the character after
           the last '&'.  Double-ampersands are collapsed into a single
           ampersand.  (This is important so "&Help && Support" works.)

           If a string has multiple mnemonics ("&t&wo") USER is inconsistent.
           DrawText uses the last one, but the dialog manager uses the first
           one.  So we use whichever one is most convenient.

           EXEX USE: For Japanese language use, suffix mnemonics are common.
           They look like "�R���g���[�� �p�l��(&C)". XP intended to remove the
           full suffix, however, it neglects to do so and only the ampersand
           is removed. This behaviour was fixed to remove the full "(&C)"
           sequence in later versions of Windows. Ordinarily, this wouldn't
           be a problem, however the special folder display names in the start
           menu rely on this function's faulty behaviour in XP to display
           correctly. Otherwise, the full mnenomic sequence is dropped, but the
           underline is still drawn in the correct position as if the mnenomic
           were drawn. Overall, this is bad for accuracy to XP, so we'll just
           use a copy of the XP function.
*/
STDAPI_(WCHAR) SHStripMneumonicXP(LPWSTR pszMenu)
{
    ASSERT(pszMenu);
    WCHAR cMneumonic = pszMenu[0]; // Default is first char

    // Early-out:  Many strings don't have ampersands at all
    LPWSTR pszAmp = StrChrW(pszMenu, L'&');
    if (pszAmp)
    {
        LPWSTR pszCopy = pszAmp - 1;

        //  FAREAST some localized builds have an mnemonic that looks like
        //    "Localized Text (&L)"  we should remove that, too
        if (pszAmp > pszMenu && *pszCopy == L'(')
        {
            if (pszAmp[2] == L')')
            {
                cMneumonic = *pszAmp;
                // move amp so that we arent past the potential terminator
                pszAmp += 3;
                pszAmp = pszCopy;
            }
        }
        else
        {
            //  move it up so that we copy on top of amp
            pszCopy++;
        }

        while (*pszAmp)
        {
            // Protect against string that ends in '&' - don't read past the end!
            if (*pszAmp == L'&' && pszAmp[1])
            {
                pszAmp++;                   // Don't copy the ampersand itself
                if (*pszAmp != L'&')        // && is not a mnemonic
                {
                    cMneumonic = *pszAmp;
                }
            }
            *pszCopy++ = *pszAmp++;
        }
        *pszCopy = 0;
    }

    return cMneumonic;
}

LPTSTR SpecialFolderList::DisplayNameOfItem(PaneItem* p, IShellFolder* psf, LPCITEMIDLIST pidlItem, SHGDNF shgno)
{
    LPWSTR psz = nullptr;
    SpecialFolderListItem* pitem = static_cast<SpecialFolderListItem*>(p);
    if (shgno == SHGDN_NORMAL)
    {
        WCHAR* pszDispName = pitem->_pszDispName;
        if (pszDispName)
        {
            pitem->_pszDispName = nullptr;
            psz = pszDispName;
        }
    }
    else
    {
        if (!pitem->_psfd->GetCustomName(&psz))
        {
            psz = SFTBarHost::DisplayNameOfItem(p, psf, pidlItem, shgno);
        }
    }

    if ((pitem->_psfd->_uFlags & SFD_PREFIX) && psz)
    {
        CoTaskMemFree(pitem->_pszAccelerator);
        pitem->_pszAccelerator = nullptr;
        SHStrDupW(psz, &pitem->_pszAccelerator);
        pitem->_chMnem = (unsigned __int16)CharUpperW((LPWSTR)SHStripMneumonicW(psz));
    }
    return psz;
}

int SpecialFolderList::CompareItems(PaneItem *p1, PaneItem *p2)
{
//    SpecialFolderListItem *pitem1 = static_cast<SpecialFolderListItem *>(p1);
//    SpecialFolderListItem *pitem2 = static_cast<SpecialFolderListItem *>(p2);

    return 0; // we added them in the right order the first time
}

HRESULT SpecialFolderList::GetFolderAndPidl(PaneItem *p,
        IShellFolder **ppsfOut, PCITEMID_CHILD *ppidlOut)
{
#ifdef DEAD_CODE
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    return SHBindToIDListParent(pitem->_pidl, IID_PPV_ARGS(ppsfOut), ppidlOut);
#else
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    if (pitem->_pidl)
    {
        return SHBindToParent(pitem->_pidl, IID_PPV_ARGS(ppsfOut), ppidlOut);
    }

    if (!pitem->_pidlSimple)
    {
        return SHBindToParent(pitem->_pidl, IID_PPV_ARGS(ppsfOut), ppidlOut);
    }
    else
    {
        return SHBindToParent(pitem->_pidlSimple, IID_PPV_ARGS(ppsfOut), ppidlOut);
    }
#endif
}

void SpecialFolderList::GetItemInfoTip(PaneItem *p, LPTSTR pszText, DWORD cch)
{
#ifdef DEAD_CODE
    SpecialFolderListItem *pitem = (SpecialFolderListItem*)p;
    if (pitem->_psfd->_iToolTip)
        LoadString(_AtlBaseModule.GetResourceInstance(), pitem->_psfd->_iToolTip, pszText, cch);
    else
        SFTBarHost::GetItemInfoTip(p, pszText, cch);    // call the base class
#else
    SpecialFolderListItem* pitem = (SpecialFolderListItem*)p;

    int v6;
    if (pitem->_psfd->_CustomTooltipCallback)
        //v6 = pitem->_psfd->_CustomTooltipCallback(this, pitem, pszText, cch);
		v6 = (pitem->_psfd->*(pitem->_psfd->_CustomTooltipCallback))(this, pitem, pszText, cch);
    else
        v6 = 0;

    if (!v6)
    {
        UINT _iToolTip = pitem->_psfd->_iToolTip;
        if (_iToolTip)
            LoadString(g_hinstCabinet, _iToolTip, pszText, cch);
        else
            SFTBarHost::GetItemInfoTip(p, pszText, cch);
    }
#endif
}

HRESULT SpecialFolderList::ContextMenuRenameItem(PaneItem *p, LPCTSTR ptszNewName)
{
#ifdef DEAD_CODE
    SpecialFolderListItem *pitem = (SpecialFolderListItem*)p;
    IShellFolder *psf;
    LPCITEMIDLIST pidlItem;
    HRESULT hr = GetFolderAndPidl(pitem, &psf, &pidlItem);
    if (SUCCEEDED(hr))
    {
        LPITEMIDLIST pidlNew;
        hr = psf->SetNameOf(_hwnd, pidlItem, ptszNewName, SHGDN_INFOLDER, &pidlNew);
        if (SUCCEEDED(hr))
        {
            pitem->ReplaceLastPidlElement(pidlNew);
        }
        psf->Release();
    }

    return hr;
#else
    SpecialFolderListItem *pitem = (SpecialFolderListItem *)p;
    IShellFolder *psf;
    LPCITEMIDLIST pidlItem;
    HRESULT hr = GetFolderAndPidl(pitem, &psf, &pidlItem);
    if (SUCCEEDED(hr))
    {
        LPITEMIDLIST pidlNew;
        hr = psf->SetNameOf(_hwnd, pidlItem, ptszNewName, SHGDN_INFOLDER, &pidlNew);
        if (SUCCEEDED(hr) && pidlNew)
        {
            pitem->ReplaceLastPidlElement(pidlNew);
            ILFree(pidlNew);
        }
        psf->Release();
    }
    return hr;
#endif
}

//
//  If we get any changenotify, it means that somebody added (or thought about
//  adding) an item to one of our minkids folders, so we'll have to look to see
//  if it crossed the minkids threshold.
//
void SpecialFolderList::OnChangeNotify(UINT id, LONG lEvent, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
{
#ifdef DEAD_CODE
    Invalidate();
    for (id = 0; id < _cNotify; id++)
    {
        UnregisterNotify(id);
    }
    _cNotify = 0;
    PostMessage(_hwnd, SFTBM_REFRESH, TRUE, 0);
#else
    UINT v6 = 0;
    UINT* p_cNotify = &this->_cNotify;
    bool v8 = this->_cNotify == 0;
    Invalidate();
    if (!v8)
    {
        do
            SFTBarHost::UnregisterNotify(v6++);
        while (v6 < *p_cNotify);
    }
    *p_cNotify = 0;
    PostMessageW(this->_hwnd, 0x40Au, 1u, 0);
#endif
}


#ifdef DEAD_CODE
BOOL SpecialFolderList::IsItemStillValid(PaneItem *p)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    return pitem->IsStillValid();
}
#endif

BOOL SpecialFolderList::IsBold(PaneItem *p)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    return pitem->_psfd->IsBold();
}

HRESULT SpecialFolderList::GetCascadeMenu(PaneItem *p, IShellMenu **ppsm)
{
#ifdef DEAD_CODE
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    IShellFolder *psf;
    HRESULT hr = SHBindToObjectEx(NULL, pitem->_pidl, NULL, IID_PPV_ARGS(&psf));
    if (SUCCEEDED(hr))
    {
        IShellMenu *psm;
        hr = CoCreateInstanceHook(CLSID_MenuBand, NULL, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&psm));
        if (SUCCEEDED(hr))
        {

            //
            //  Recent Documents requires special treatment.
            //
            IShellMenuCallback *psmc = NULL;
            hr = pitem->_psfd->CreateShellMenuCallback(&psmc);

            if (SUCCEEDED(hr))
            {
                DWORD dwFlags = SMINIT_TOPLEVEL | SMINIT_VERTICAL | pitem->_psfd->_dwShellFolderFlags;
                if (IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOCHANGESTARMENU,
                    TEXT("Advanced"), TEXT("Start_EnableDragDrop"),
                    ROUS_DEFAULTALLOW | ROUS_KEYALLOWS))
                {
                    dwFlags |= SMINIT_RESTRICT_DRAGDROP | SMINIT_RESTRICT_CONTEXTMENU;
                }
                psm->Initialize(psmc, 0, 0, dwFlags);

                HKEY hkCustom = NULL;
                if (pitem->_psfd->_pszCustomizeKey)
                {
                    RegCreateKeyEx(HKEY_CURRENT_USER, pitem->_psfd->_pszCustomizeKey,
                        NULL, NULL, REG_OPTION_NON_VOLATILE,
                        KEY_READ | KEY_WRITE, NULL, &hkCustom, NULL);
                }

                dwFlags = SMSET_USEBKICONEXTRACTION;
                hr = psm->SetShellFolder(psf, pitem->_pidl, hkCustom, dwFlags);
                if (SUCCEEDED(hr))
                {
                    // SetShellFolder takes ownership of hkCustom
                    *ppsm = psm;
                    psm->AddRef();
                }
                else
                {
                    // Clean up the registry key since SetShellFolder
                    // did not take ownership
                    if (hkCustom)
                    {
                        RegCloseKey(hkCustom);
                    }
                }

                ATOMICRELEASE(psmc); // psmc can be NULL
            }
            psm->Release();
        }
        psf->Release();
    }

    return hr;
#else
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    PIDLIST_RELATIVE pidl = pitem->_pidlCascade;
    if (!pidl)
    {
        pidl = pitem->_pidl;
    }

    IShellFolder *psf = 0;
    HRESULT hr = pidl ? SHBindToObject(0, pidl, 0, IID_IShellFolder, (void **)&psf) : E_FAIL;
    if (hr >= 0)
    {
        IShellMenu *psm;
        hr = CoCreateInstance(CLSID_MenuBand, NULL, CLSCTX_INPROC_SERVER, IID_IShellMenu, (LPVOID *)&psm);
        if (hr >= 0)
        {
            IShellMenuCallback *psmc = NULL;
            hr = pitem->_psfd->CreateShellMenuCallback(&psmc);
            if (hr >= 0)
            {
                DWORD dwFlags = pitem->_psfd->_dwShellFolderFlags | 0x10000004;
                if (IsRestrictedOrUserSetting(
                    HKEY_CURRENT_USER,
                    REST_NOCHANGESTARMENU,
                    TEXT("Advanced"),
                    TEXT("Start_EnableDragDrop"),
                    0))
                {
                    dwFlags |= 3;
                }
                psm->Initialize(psmc, 0, 0, dwFlags);

                HKEY hkCustom = 0;
                if (pitem->_psfd->_pszCustomizeKey)
                {
                    RegCreateKeyEx(HKEY_CURRENT_USER, pitem->_psfd->_pszCustomizeKey, 0, 0, 0, 0x2001Fu, 0, &hkCustom, 0);
                }

                hr = psm->SetShellFolder(psf, (PIDLIST_ABSOLUTE)pidl, hkCustom, pitem->_psfd->_dwShellMenuSetFlags | 8);
                if (hr < 0)
                {
                    if (hkCustom)
                    {
                        RegCloseKey(hkCustom);
                    }
                }
                else
                {
                    *ppsm = psm;
                    psm->AddRef();
                }

                IUnknown_SafeReleaseAndNullPtr(&psmc);
            }
            psm->Release();
        }
        psf->Release();
    }

    return hr;
#endif
}

TCHAR SpecialFolderList::GetItemAccelerator(PaneItem *p, int iItemStart)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);

    if (pitem->_chMnem)
    {
        return pitem->_chMnem;
    }
    else
    {
        // Default: First letter is accelerator.
        return SFTBarHost::GetItemAccelerator(p, iItemStart);
    }
}

LRESULT SpecialFolderList::OnWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_NOTIFY:
        switch (((NMHDR*)(lParam))->code)
        {
        // When the user connects/disconnects via TS, we need to recalc
        // the "Windows Security" item
        case SMN_REFRESHLOGOFF:
            Invalidate();
            break;
        }
    }

    // Else fall back to parent implementation
    return SFTBarHost::OnWndProc(hwnd, uMsg, wParam, lParam);
}

BOOL _IsItemHiddenOnDesktop(LPCTSTR pszGuid)
{
    return _SHRegGetBoolValueFromHKCUHKLM(REGSTR_PATH_HIDDEN_DESKTOP_ICONS_STARTPANEL, pszGuid, FALSE);
    //return SHRegGetBoolUSValue(REGSTR_PATH_HIDDEN_DESKTOP_ICONS_STARTPANEL,
    //                           pszGuid, FALSE, FALSE);
}

UINT SpecialFolderList::AdjustDeleteMenuItem(PaneItem *p, UINT *puiFlags)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    if (pitem->_psfd->_pszCanHideOnDesktop)
    {
        // Set MF_CHECKED if the item is visible on the desktop
        if (!_IsItemHiddenOnDesktop(pitem->_psfd->_pszCanHideOnDesktop))
        {
            // Item is visible - show the checkbox
            *puiFlags |= MF_CHECKED;
        }

        return IDS_SFTHOST_SHOWONDESKTOP;
    }
    else
    {
        return 0; // not deletable
    }
}

HRESULT SpecialFolderList::ContextMenuInvokeItem(PaneItem *p, IContextMenu *pcm, CMINVOKECOMMANDINFOEX *pici, LPCTSTR pszVerb)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    HRESULT hr;

    if (StrCmpIC(pszVerb, TEXT("delete")) == 0)
    {
        ASSERT(pitem->_psfd->_pszCanHideOnDesktop);

        // Toggle the hide/unhide state
        DWORD dwHide = !_IsItemHiddenOnDesktop(pitem->_psfd->_pszCanHideOnDesktop);
        LONG lErr = SHRegSetUSValue(REGSTR_PATH_HIDDEN_DESKTOP_ICONS_STARTPANEL,
                                    pitem->_psfd->_pszCanHideOnDesktop,
                                    REG_DWORD, &dwHide, sizeof(dwHide),
                                    SHREGSET_FORCE_HKCU);
        hr = HRESULT_FROM_WIN32(lErr);
        if (SUCCEEDED(hr))
        {
            // explorer\rcids.h and shell32\unicpp\resource.h have DIFFERENT
            // VALUES FOR FCIDM_REFRESH!  We want the one in unicpp\resource.h
            // because that's the correct one...
#define FCIDM_REFRESH_REAL 0x0a220
            PostMessage(GetShellWindow(), WM_COMMAND, FCIDM_REFRESH_REAL, 0); // refresh desktop
        }
    }
    else
    {
        hr = SFTBarHost::ContextMenuInvokeItem(pitem, pcm, pici, pszVerb);
    }

    return hr;
}

#ifdef DEAD_CODE

HRESULT SpecialFolderList::_GetUIObjectOfItem(PaneItem *p, REFIID riid, LPVOID *ppv)
{
    SpecialFolderListItem *pitem = static_cast<SpecialFolderListItem *>(p);
    if (pitem->_psfd->IsCSIDL() && (CSIDL_RECENT == pitem->_psfd->GetCSIDL()))
    {
        *ppv = NULL;
        return E_NOTIMPL;
    }
    return SFTBarHost::_GetUIObjectOfItem(p, riid, ppv);
}

#endif

HRESULT SpecialFolderList::OnItemUpdate(PaneItem *p, WPARAM wParam, LPARAM lParam)
{
    SpecialFolderListItem *pitem = reinterpret_cast<SpecialFolderListItem *>(p);
    CKnownFolderInformation *pkfi = reinterpret_cast<CKnownFolderInformation *>(wParam);

    HRESULT hr = S_FALSE;

    if ((DWORD)lParam == (pitem->_psfd - s_rgsfd))
    {
        ILFree(pitem->_pidl);
        pitem->_pidl = pkfi->_pidl;
        pkfi->_pidl = NULL;

        if (pkfi->_pszDispName)
        {
            CoTaskMemFree(pitem->_pszDispName);
            pitem->_pszDispName = pkfi->_pszDispName;
            pkfi->_pszDispName = 0;
            hr = S_OK;
        }

        if (pkfi->_hIcon)
        {
            int iImageCount = ImageList_GetImageCount(_himl);
            printf("SpecialFolderList::OnItemUpdate: ImageList_GetImageCount returned %d images\n", iImageCount);

            pitem->_iImageIndex = ImageList_ReplaceIcon(_himl, -1, pkfi->_hIcon);

            int iNewImageCount = ImageList_GetImageCount(_himl);
            printf("SpecialFolderList::OnItemUpdate: ImageList_GetImageCount after ReplaceIcon returned %d images\n", iNewImageCount);

            DestroyIcon(pkfi->_hIcon);
            pkfi->_hIcon = NULL;
            hr = S_OK;
        }
    }
    return hr;
}

int SpecialFolderList::GetMinTextWidth()
{
    if (!this->field_190)
    {
        field_190 = _CalcMaxTextWith();
    }
    return field_190;
}

HRESULT DisplayNameOfAsString(IShellFolder *psf, const ITEMIDLIST_RELATIVE *pidl, SHGDNF flags, WCHAR **ppsz)
{
    *ppsz = nullptr;

    IShellFolder *psfParent;
    const ITEMID_CHILD *pidlChild;
    HRESULT hr = SHBindToFolderIDListParent(psf, pidl, IID_PPV_ARGS(&psfParent), &pidlChild);
    if (SUCCEEDED(hr))
    {
        STRRET sr;
        hr = psfParent->GetDisplayNameOf(pidlChild, flags, &sr);
        if (SUCCEEDED(hr))
            hr = StrRetToStrW(&sr, pidlChild, ppsz);

        psfParent->Release();
    }

    return hr;
}

HRESULT SpecialFolderList::_CreateFullListItem(SpecialFolderListItem *pitem)
{
    HRESULT hr = E_FAIL;

    if (pitem->_pidl)
    {
        IShellFolder *psf;
        LPCITEMIDLIST pidlLast;
        hr = SHBindToParent(pitem->_pidl, IID_PPV_ARGS(&psf), &pidlLast);
        if (SUCCEEDED(hr))
        {
            if (!pitem->_psfd->GetCustomName(&pitem->_pszDispName))
            {
                hr = DisplayNameOfAsString(psf, pidlLast, 0, &pitem->_pszDispName);
            }

            if (_IsPrivateImageList())
            {
                if (pitem->_psfd->_pszIconPath)
                {
					printf("SpecialFolderList::_CreateFullListItem (Private): Loading icon from path: %ls\n", pitem->_psfd->_pszIconPath);
                    IBindCtx *pbc;
                    if (SUCCEEDED(BindCtx_CreateWithMode(0x1000u, &pbc)))
                    {
                        LPITEMIDLIST pidl;
                        if (SHParseDisplayName(pitem->_psfd->_pszIconPath, pbc, &pidl, 0, 0) >= 0)
                        {
                            IShellFolder *psf;
                            LPCITEMIDLIST v8;
                            if (SUCCEEDED(SHBindToParent(pidl, IID_PPV_ARGS(&psf), &v8)))
                            {
                                HICON hIcon = _IconOf(psf, v8, _cxIcon, 0, -1);
								int iImageCount = ImageList_GetImageCount(_himl);
								printf("SpecialFolderList::_CreateFullListItem (Private): ImageList_GetImageCount returned %d images\n", iImageCount);
                                pitem->_iImageIndex = ImageList_ReplaceIcon(_himl, -1, hIcon);
								int iNewImageCount = ImageList_GetImageCount(_himl);
								printf("SpecialFolderList::_CreateFullListItem (Private): ImageList_GetImageCount after ReplaceIcon returned %d images\n", iNewImageCount);
                                psf->Release();
                                DestroyIcon(hIcon);
                            }
                            ILFree(pidl);
                        }
                        pbc->Release();
                    }
                }
                else
                {
                    HICON hIcon = _IconOf(psf, pidlLast, _cxIcon, 0, -1);
					int iImageCount = ImageList_GetImageCount(_himl);
					printf("SpecialFolderList::_CreateFullListItem: ImageList_GetImageCount returned %d images\n", iImageCount);
                    pitem->_iImageIndex = ImageList_ReplaceIcon(_himl, -1, hIcon);
					int iNewImageCount = ImageList_GetImageCount(_himl);
					printf("SpecialFolderList::_CreateFullListItem: ImageList_GetImageCount after ReplaceIcon returned %d images\n", iNewImageCount);
                    DestroyIcon(hIcon);
                }
            }
            psf->Release();
        }
    }
    return hr;
}

TASKOWNERID stru_1017900 = { 156425175u, 57748u, 17760u, { 176u, 158u, 222u, 109u, 45u, 227u, 102u, 62u } };

HRESULT SpecialFolderList::_CreateSimpleListItem(IKnownFolderManager *pkfm, SpecialFolderListItem *pitem, DWORD a4)
{
    // eax
    // esi
    int v7; // eax
    IRunnableTask *pclfpt; // ebx
    IKnownFolder *pkf; // [esp+Ch] [ebp-260h] BYREF
    int v14; // [esp+10h] [ebp-25Ch] SPLIT
    HICON hIcon; // [esp+10h] [ebp-25Ch] MAPDST SPLIT BYREF
    KNOWNFOLDER_DEFINITION kfd; // [esp+14h] [ebp-258h] BYREF
    WCHAR szOutBuf[260]; // [esp+60h] [ebp-20Ch] BYREF

    pkf = 0;

    HRESULT hr = pkfm->GetFolder(pitem->_psfd->GetFolderID(), &pkf);
    if (hr >= 0)
    {
        hr = pkf->GetFolderDefinition(&kfd);

        WCHAR guidStr[64];
        StringFromGUID2(pitem->_psfd->GetFolderID(), guidStr, ARRAYSIZE(guidStr));
        if (kfd.pszIcon == NULL)
        {
            wprintf(L"[DEBUG] s_rgsfdNEW[%d] (KNOWNFOLDERID: %ls) has kfd.pszIcon == NULL\n", a4, guidStr);
        }

        if (hr >= 0)
        {
            if (!kfd.pszLocalizedName)
            {
                goto LABEL_6;
            }

            hr = SHLoadIndirectString(kfd.pszLocalizedName, szOutBuf, 260u, 0);
            if (hr >= 0)
            {
                hr = SHStrDup(szOutBuf, &pitem->_pszDispName);
            LABEL_6:
                if (hr >= 0)
                {
                    v14 = PathParseIconLocationW(kfd.pszIcon);
                    if (_IsPrivateImageList())
                    {
						printf("SpecialFolderList::_CreateSimpleListItem (Private): Loading icon from path: %ls\n", kfd.pszIcon);
                        hr = SHDefExtractIcon(kfd.pszIcon, v14, 2u, &hIcon, 0, _cxIcon);
                        if (hr >= 0)
                        {
							printf("SpecialFolderList::_CreateSimpleListItem: hIcon: 0x%p\n", hIcon);
                            int iImageCount = ImageList_GetImageCount(this->_himl);
                            printf("SpecialFolderList::_CreateSimpleListItem: ImageList_GetImageCount returned %d images\n", iImageCount);
                            v7 = ImageList_ReplaceIcon(this->_himl, -1, hIcon);
                            int iNewImageCount = ImageList_GetImageCount(this->_himl);
                            printf("SpecialFolderList::_CreateSimpleListItem: ImageList_GetImageCount after ReplaceIcon returned %d images\n", iNewImageCount);
                            pitem->_iImageIndex = v7;
                            DestroyIcon(hIcon);
                        }
                    }
                    else
                    {
                        pitem->_iImageIndex = Shell_GetCachedImageIndexW(kfd.pszIcon, v14, 2u);
                    }
                }
            }

            FreeKnownFolderDefinitionFields(&kfd);
            if (hr >= 0 && _psched)
            {
                pclfpt = new CLoadFullPidlTask(_hwnd, a4, _cxIcon);
                if (pclfpt)
                {
                    hr = _psched->AddTask(pclfpt, stru_1017900, (DWORD_PTR)this, 0x10000000);
                    pclfpt->Release();
                }
            }
        }
    }

    if (pkf)
    {
        pkf->Release();
    }
    return hr;
}

//****************************************************************************
//
//  IShellMenuCallback helper for Recent Documents
//
//  We want to restrict to the first MAXRECDOCS items.
//

class CRecentShellMenuCallback
    : public CUnknown
    , public IShellMenuCallback
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    friend HRESULT CRecentShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
    HRESULT _FilterRecentPidl(IShellFolder *psf, LPCITEMIDLIST pidlItem);

    int     _nShown;
    int     _iMaxRecentDocs;
};

HRESULT CRecentShellMenuCallback::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CRecentShellMenuCallback, IShellMenuCallback),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

HRESULT CRecentShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_BEGINENUM:
        _nShown = 0;
        _iMaxRecentDocs = SHRestricted(REST_MaxRecentDocs);
        if (_iMaxRecentDocs < 1)
            _iMaxRecentDocs = 15;       // default from shell32\recdocs.h
        return S_OK;

    case SMC_FILTERPIDL:
        ASSERT(psmd->dwMask & SMDM_SHELLFOLDER);
        return _FilterRecentPidl(psmd->psf, psmd->pidlItem);

    }
    return S_FALSE;
}

//
//  Return S_FALSE to allow the item to show, S_OK to hide it
//

HRESULT CRecentShellMenuCallback::_FilterRecentPidl(IShellFolder *psf, LPCITEMIDLIST pidlItem)
{
    HRESULT hrRc = S_OK;      // Assume hidden

    if (_nShown < _iMaxRecentDocs)
    {
        IShellLink *psl;
        if (SUCCEEDED(psf->GetUIObjectOf(NULL, 1, &pidlItem, IID_IShellLink, NULL, reinterpret_cast<void**>(static_cast<IShellLinkW**>(&psl)))))
        {
            LPITEMIDLIST pidlTarget;
            if (SUCCEEDED(psl->GetIDList(&pidlTarget)) && pidlTarget)
            {
                DWORD dwAttr = SFGAO_FOLDER;
                if (SUCCEEDED(SHGetAttributesOf(pidlTarget, &dwAttr)) &&
                    !(dwAttr & SFGAO_FOLDER))
                {
                    // We found a shortcut to a nonfolder - keep it!
                    _nShown++;
                    hrRc = S_FALSE;
                }
                ILFree(pidlTarget);
            }

            psl->Release();
        }
    }

    return hrRc;
}

HRESULT CRecentShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc)
{
    *ppsmc = new CRecentShellMenuCallback;
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}

//****************************************************************************
//
//  IShellMenuCallback helper that disallows cascading into subfolders
//

class CNoSubdirShellMenuCallback
    : public CUnknown
    , public IShellMenuCallback
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    friend HRESULT CNoSubdirShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
};

HRESULT CNoSubdirShellMenuCallback::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CNoSubdirShellMenuCallback, IShellMenuCallback),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

HRESULT CNoSubdirShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_GETSFINFO:
        {
            // Turn off the SMIF_SUBMENU flag on everybody.  This
            // prevents us from cascading more than one level deel.
            SMINFO *psminfo = reinterpret_cast<SMINFO *>(lParam);
            psminfo->dwFlags &= ~SMIF_SUBMENU;
            return S_OK;
        }
    }
    return S_FALSE;
}

HRESULT CNoSubdirShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc)
{
    *ppsmc = new CNoSubdirShellMenuCallback;
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}

//****************************************************************************
//
//  IShellMenuCallback helper for My Computer
//
//  Disallow cascading into subfolders and also force the default
//  drag/drop effect to DROPEFFECT_LINK.
//

class CMyComputerShellMenuCallback
    : public CNoSubdirShellMenuCallback
{
public:
    typedef CNoSubdirShellMenuCallback super;

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    friend HRESULT CMyComputerShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
};

HRESULT CMyComputerShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_BEGINDRAG:
        *(DWORD*)wParam = DROPEFFECT_LINK;
        return S_OK;

    }
    return super::CallbackSM(psmd, uMsg, wParam, lParam);
}

HRESULT CMyComputerShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc)
{
    *ppsmc = new CMyComputerShellMenuCallback;
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}

//****************************************************************************
//
//  IShellMenuCallback helper that prevents Fonts from cascading
//  Used by Control Panel.
//

class CNoFontsShellMenuCallback
    : public CUnknown
    , public IShellMenuCallback
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    friend HRESULT CNoFontsShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
};

HRESULT CNoFontsShellMenuCallback::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CNoFontsShellMenuCallback, IShellMenuCallback),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

BOOL _IsFontsFolderShortcut(IShellFolder *psf, LPCITEMIDLIST pidl)
{
    TCHAR sz[MAX_PATH];
    return SUCCEEDED(DisplayNameOf(psf, pidl, SHGDN_FORPARSING | SHGDN_INFOLDER, sz, ARRAYSIZE(sz))) &&
           lstrcmpi(sz, TEXT("::{D20EA4E1-3957-11d2-A40B-0C5020524152}")) == 0;
}

HRESULT CNoFontsShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_GETSFINFO:
        {
            // If this is the Fonts item, then remove the SUBMENU attribute.
            SMINFO *psminfo = reinterpret_cast<SMINFO *>(lParam);
            if ((psminfo->dwMask & SMIM_FLAGS) &&
                (psminfo->dwFlags & SMIF_SUBMENU) &&
                _IsFontsFolderShortcut(psmd->psf, psmd->pidlItem))
            {
                psminfo->dwFlags &= ~SMIF_SUBMENU;
            }
            return S_OK;
        }
    }
    return S_FALSE;
}

HRESULT CNoFontsShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc)
{
    *ppsmc = new CNoFontsShellMenuCallback;
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}

//****************************************************************************
//
//  IShellMenuCallback helper that filters the "connect to" menu
//

class CConnectToShellMenuCallback
    : public CUnknown
    , public IShellMenuCallback
    , public CObjectWithSite
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    // *** IObjectWithSite ***
    // inherited from CObjectWithSite

private:
    HRESULT _OnGetSFInfo(SMDATA *psmd, SMINFO *psminfo);
    HRESULT _OnGetInfo(SMDATA *psmd, SMINFO *psminfo);
    HRESULT _OnEndEnum(SMDATA *psmd);

    friend HRESULT CConnectToShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc);
    BOOL _bAnyRAS;
};

#define ICOL_NETCONMEDIATYPE       0x101 // from netshell
#define ICOL_NETCONSUBMEDIATYPE    0x102 // from netshell
#define ICOL_NETCONSTATUS          0x103 // from netshell
#define ICOL_NETCONCHARACTERISTICS 0x104 // from netshell

BOOL IsMediaRASType(NETCON_MEDIATYPE ncm)
{
    return (ncm == NCM_DIRECT || ncm == NCM_ISDN || ncm == NCM_PHONE || ncm == NCM_TUNNEL || ncm == NCM_PPPOE);  // REVIEW DIRECT correct?
}

BOOL IsNetConPidlRAS(IShellFolder2 *psfNetCon, LPCITEMIDLIST pidlNetConItem)
{
    BOOL bRet = FALSE;
    SHCOLUMNID scidMediaType, scidSubMediaType, scidCharacteristics;
    VARIANT v;
    
    scidMediaType.fmtid       = GUID_NETSHELL_PROPS;
    scidMediaType.pid         = ICOL_NETCONMEDIATYPE;

    scidSubMediaType.fmtid    = GUID_NETSHELL_PROPS;
    scidSubMediaType.pid      = ICOL_NETCONSUBMEDIATYPE;
    
    scidCharacteristics.fmtid = GUID_NETSHELL_PROPS;
    scidCharacteristics.pid   = ICOL_NETCONCHARACTERISTICS;

    if (SUCCEEDED(psfNetCon->GetDetailsEx(pidlNetConItem, &scidMediaType, &v)))
    {
        // Is this a RAS connection
        if (IsMediaRASType((NETCON_MEDIATYPE)v.lVal))
        {
            VariantClear(&v);
         
            // Make sure it's not incoming
            if (SUCCEEDED(psfNetCon->GetDetailsEx(pidlNetConItem, &scidCharacteristics, &v)))
            {
                if (!(NCCF_INCOMING_ONLY & v.lVal))
                    bRet = TRUE;
            }
        }

        // Is this a Wireless LAN connection?
        if (NCM_LAN == (NETCON_MEDIATYPE)v.lVal)
        {
            VariantClear(&v);
            
            if (SUCCEEDED(psfNetCon->GetDetailsEx(pidlNetConItem, &scidSubMediaType, &v)))
            {
                
                if (NCSM_WIRELESS == (NETCON_SUBMEDIATYPE)v.lVal)
                    bRet = TRUE;
            }
        }

        VariantClear(&v);
    }
    return bRet;
}

HRESULT CConnectToShellMenuCallback::_OnGetInfo(SMDATA *psmd, SMINFO *psminfo)
{
    HRESULT hr = S_FALSE;
    if (psminfo->dwMask & SMIM_ICON)
    {
        if (psmd->uId == IDM_OPENCONFOLDER)
        {
            LPITEMIDLIST pidl = SHCloneSpecialIDList(NULL, CSIDL_CONNECTIONS, FALSE);
            if (pidl)
            {
                LPCITEMIDLIST pidlObject;
                IShellFolder *psf;
                hr = SHBindToParent(pidl, IID_PPV_ARGS(&psf), &pidlObject);
                if (SUCCEEDED(hr))
                {
                    SHMapPIDLToSystemImageListIndex(psf, pidlObject, &psminfo->iIcon);
                    psminfo->dwFlags |= SMIF_ICON;
                    psf->Release();
                }
                ILFree(pidl);
            }
        }
    }
    return hr;
}

HRESULT CConnectToShellMenuCallback::_OnGetSFInfo(SMDATA *psmd, SMINFO *psminfo)
{
    IShellFolder2 *psf2;
    ASSERT(psminfo->dwMask & SMIM_FLAGS);                       // ??
    psminfo->dwFlags &= ~SMIF_SUBMENU;

    if (SUCCEEDED(psmd->psf->QueryInterface(IID_PPV_ARGS(&psf2))))
    {
        if (!IsNetConPidlRAS(psf2, psmd->pidlItem))
            psminfo->dwFlags |= SMIF_HIDDEN;
        else
            _bAnyRAS = TRUE;

        psf2->Release();
    }

    return S_OK;
}

HRESULT CConnectToShellMenuCallback::_OnEndEnum(SMDATA *psmd)
{
    HRESULT hr = S_FALSE;
    IShellMenu* psm;

    if (psmd->punk && SUCCEEDED(hr = psmd->punk->QueryInterface(IID_PPV_ARGS(&psm))))
    {
        // load the static portion of the connect to menu, and add it to the bottom
        HMENU hmStatic = LoadMenu(_AtlBaseModule.GetResourceInstance(), MAKEINTRESOURCE(MENU_CONNECTTO));

        if (hmStatic)
        {
            // if there aren't any dynamic items (RAS connections), then delete the separator
            if (!_bAnyRAS)
                DeleteMenu(hmStatic, 0, MF_BYPOSITION);

            HWND hwnd = NULL;
            IUnknown *punk;
            if (SUCCEEDED(IUnknown_QueryService(_punkSite, SID_SMenuPopup, IID_PPV_ARGS(&punk))))
            {
                IUnknown_GetWindow(punk, &hwnd);
                punk->Release();
            }
            psm->SetMenu(hmStatic, hwnd, SMSET_NOEMPTY | SMSET_BOTTOM);
        }
        psm->Release();
    }
    return hr;
}


HRESULT CConnectToShellMenuCallback::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CConnectToShellMenuCallback, IShellMenuCallback),
        QITABENT(CConnectToShellMenuCallback, IObjectWithSite),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

HRESULT CConnectToShellMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case SMC_GETINFO:
        return _OnGetInfo(psmd, (SMINFO *)lParam);

    case SMC_GETSFINFO:
        return _OnGetSFInfo(psmd, (SMINFO *)lParam);

    case SMC_BEGINENUM:
        _bAnyRAS = FALSE;

    case SMC_ENDENUM:
        return _OnEndEnum(psmd);

    case SMC_EXEC:
        switch (psmd->uId)
        {
            case IDM_OPENCONFOLDER:
                ShowFolder(CSIDL_CONNECTIONS);
                return S_OK;
        }
        break;
    }

    return S_FALSE;
}

HRESULT CConnectToShellMenuCallback_CreateInstance(IShellMenuCallback **ppsmc)
{
    *ppsmc = new CConnectToShellMenuCallback;
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}

void ShowFolder(UINT csidl)
{
    LPITEMIDLIST pidl;
    if (SUCCEEDED(SHGetFolderLocation(NULL, csidl, NULL, 0, &pidl)))
    {
        SHELLEXECUTEINFO shei = { 0 };

        shei.cbSize     = sizeof(shei);
        shei.fMask      = SEE_MASK_IDLIST;
        shei.nShow      = SW_SHOWNORMAL;
        shei.lpVerb     = TEXT("open");
        shei.lpIDList   = pidl;
        ShellExecuteEx(&shei);
        ILFree(pidl);
    }
}