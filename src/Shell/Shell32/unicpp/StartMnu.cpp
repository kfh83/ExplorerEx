#include "pch.h"

#include "ShGuidP.h"
#include "cocreateinstancehook.h"
#include "shundoc.h"
#include "stdafx.h"
#pragma hdrstop
#include "resource.h"
#include "runtask.h"
#include "msi.h"
#include "cabinet.h"
#include "util.h"
#include "cowsite.h"
#include "uemapp.h"

#define REGSTR_EXPLORER_WINUPDATE REGSTR_PATH_EXPLORER TEXT("\\WindowsUpdate")

#define IDM_TOPLEVELSTARTMENU  0

// StartMenuInit Flags
#define STARTMENU_DISPLAYEDBEFORE       0x00000001
#define STARTMENU_CHEVRONCLICKED        0x00000002

// New item counts for UEM stuff
#define UEM_NEWITEMCOUNT 2
typedef DWORD   BITBOOL;

// Menuband per pane user data
typedef struct
{
    BITBOOL _fInitialized;
} SMUSERDATA;

// for g_hdpaDarwinAds
EXTERN_C CRITICAL_SECTION g_csDarwinAds = { 0 };

//#define ENTERCRITICAL_DARWINADS EnterCriticalSection(&g_csDarwinAds)
//#define LEAVECRITICAL_DARWINADS LeaveCriticalSection(&g_csDarwinAds)
#define ENTERCRITICAL_DARWINADS 
#define LEAVECRITICAL_DARWINADS 

// The threading concern with this variable is create/delete/add/remove. We will only remove an item 
// and delete the hdpa on the main thread. We will however add and create on both threads.
// We need to serialize access to the dpa, so we're going to grab the shell crisec.
HDPA g_hdpaDarwinAds = NULL;

class CDarwinAd
{
public:
    LPITEMIDLIST    _pidl;
    LPTSTR          _pszDescriptor;
    LPTSTR          _pszLocalPath;
    INSTALLSTATE    _state;

    CDarwinAd(LPITEMIDLIST pidl, LPTSTR psz)
    {
        // I take ownership of this pidl
        _pidl = pidl;
        Str_SetPtrW(&_pszDescriptor, psz);
    }

    void CheckInstalled()
    {
        TCHAR szProduct[GUIDSTR_MAX];
        TCHAR szFeature[MAX_FEATURE_CHARS];
        TCHAR szComponent[GUIDSTR_MAX];

        if (MsiDecomposeDescriptorW(_pszDescriptor, szProduct, szFeature, szComponent, NULL) == ERROR_SUCCESS)
        {
            _state = MsiQueryFeatureState(szProduct, szFeature);
        }
        else
        {
            _state = INSTALLSTATE_INVALIDARG;
        }

        // Note: Cannot use ParseDarwinID since that bumps the usage count
        // for the app and we're not running the app, just looking at it.
        // Also because ParseDarwinID tries to install the app (eek!)
        //
        // Must ignore INSTALLSTATE_SOURCE because MsiGetComponentPath will
        // try to install the app even though we're just querying...
        TCHAR szCommand[MAX_PATH];
        DWORD cch = ARRAYSIZE(szCommand);

        if (_state == INSTALLSTATE_LOCAL &&
            MsiGetComponentPath(szProduct, szComponent, szCommand, &cch) == _state)
        {
            PathUnquoteSpaces(szCommand);
            Str_SetPtrW(&_pszLocalPath, szCommand);
        }
        else
        {
            Str_SetPtrW(&_pszLocalPath, NULL);
        }
    }

    BOOL IsAd()
    {
        return _state == INSTALLSTATE_ADVERTISED;
    }

    ~CDarwinAd()
    {
        ILFree(_pidl);
        Str_SetPtrW(&_pszDescriptor, NULL);
        Str_SetPtrW(&_pszLocalPath, NULL);
    }
};

int GetDarwinIndex(LPCITEMIDLIST pidlFull, CDarwinAd** ppda);

HRESULT GetMyPicsDisplayName(LPTSTR pszBuffer, UINT cchBuffer)
{
    LPITEMIDLIST pidlMyPics = SHCloneSpecialIDList(NULL, CSIDL_MYPICTURES, FALSE);
    if (pidlMyPics)
    {
        HRESULT hRet = SHGetNameAndFlags(pidlMyPics, SHGDN_NORMAL, pszBuffer, cchBuffer, NULL);
        ILFree(pidlMyPics);
        return hRet;
    }
    return E_FAIL;
}

#define RESTOPT_INTELLIMENUS_USER       0
#define RESTOPT_INTELLIMENUS_DISABLED   1       // Match Restriction assumption: 1 == Off
#define RESTOPT_INTELLIMENUS_ENABLED    2


BOOL AreIntelliMenusEnabled()
{
    DWORD dwRest = SHRestricted(REST_INTELLIMENUS);
    if (dwRest != RESTOPT_INTELLIMENUS_USER)
        return (dwRest == RESTOPT_INTELLIMENUS_ENABLED);

    return SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("IntelliMenus"),
        FALSE, TRUE); // Don't ignore HKCU, Enable Menus by default
}

BOOL FeatureEnabled(LPCTSTR pszFeature)
{
    return SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, pszFeature,
        FALSE, // Don't ignore HKCU
        FALSE); // Disable this cool feature.
}


// Since we can be presented with an Augmented shellfolder and we need a Full pidl,
// we have been given the responsibility to unwrap it for perf reasons.
LPITEMIDLIST FullPidlFromSMData(LPSMDATA psmd)
{
    LPITEMIDLIST pidlItem;
    LPITEMIDLIST pidlFolder = NULL;
    LPITEMIDLIST pidlFull = NULL;
    IAugmentedShellFolder2* pasf2;
    if (SUCCEEDED(psmd->psf->QueryInterface(IID_IAugmentedFolder, (LPVOID*)&pasf2)))
    {
        if (SUCCEEDED(pasf2->UnWrapIDList(psmd->pidlItem, 1, NULL, &pidlFolder, &pidlItem, NULL)))
        {
            pidlFull = ILCombine(pidlFolder, pidlItem);
            ILFree(pidlFolder);
            ILFree(pidlItem);
        }
        pasf2->Release();
    }

    if (!pidlFolder)
    {
        pidlFull = ILCombine(psmd->pidlFolder, psmd->pidlItem);
    }

    return pidlFull;
}

//
//  Determine whether a namespace pidl in a merged shellfolder came
//  from the specified object GUID.
//
BOOL IsMergedFolderGUID(IShellFolder* psf, LPCITEMIDLIST pidl, REFGUID rguid)
{
    IAugmentedShellFolder2* pasf;
    BOOL fMatch = FALSE;
    if (SUCCEEDED(psf->QueryInterface(IID_IAugmentedFolder, (LPVOID*)&pasf)))
    {
        GUID guid;
        if (SUCCEEDED(pasf->GetNameSpaceID(pidl, &guid)))
        {
            fMatch = IsEqualGUID(guid, rguid);
        }
        pasf->Release();
    }

    return fMatch;
}

STDMETHODIMP_(int) s_DarwinAdsDestroyCallback(LPVOID pData1, LPVOID pData2)
{
    CDarwinAd* pda = (CDarwinAd*)pData1;
    if (pda)
        delete pda;
    return TRUE;
}


// SHRegisterDarwinLink takes ownership of the pidl
BOOL SHRegisterDarwinLink(LPITEMIDLIST pidlFull, LPWSTR pszDarwinID, BOOL fUpdate)
{
    BOOL fRetVal = FALSE;

    ENTERCRITICAL_DARWINADS;

    if (pidlFull)
    {
        CDarwinAd* pda = NULL;

        if (GetDarwinIndex(pidlFull, &pda) != -1 && pda)
        {
            // We already know about this link; don't need to add it
            fRetVal = TRUE;
        }
        else
        {
            pda = new CDarwinAd(pidlFull, pszDarwinID);
            if (pda)
            {
                pidlFull = NULL;    // take ownership

                // Do we have a global cache?
                if (g_hdpaDarwinAds == NULL)
                {
                    // No; This is either the first time this is called, or we
                    // failed the last time.
                    g_hdpaDarwinAds = DPA_Create(5);
                }

                if (g_hdpaDarwinAds)
                {
                    // DPA_AppendPtr returns the zero based index it inserted it at.
                    if (DPA_AppendPtr(g_hdpaDarwinAds, (void*)pda) >= 0)
                    {
                        fRetVal = TRUE;
                    }

                }
            }
        }

        if (!fRetVal)
        {
            // if we failed to create a dpa, delete this.
            delete pda;
        }
        else if (fUpdate)
        {
            // update the entry if requested
            pda->CheckInstalled();
        }
        ILFree(pidlFull);

    }
    else if (!pszDarwinID)
    {
        // NULL, NULL means "destroy darwin info, we're shutting down"
        HDPA hdpa = g_hdpaDarwinAds;
        g_hdpaDarwinAds = NULL;
        if (hdpa)
            DPA_DestroyCallback(hdpa, s_DarwinAdsDestroyCallback, NULL);
    }

    LEAVECRITICAL_DARWINADS;

    return fRetVal;
}

BOOL ProcessDarwinAd(IShellLinkDataList* psldl, LPCITEMIDLIST pidlFull)
{
    // This function does not check for the existance of a member before adding it,
    // so it is entirely possible for there to be duplicates in the list....
    BOOL fIsLoaded = FALSE;
    BOOL fFreesldl = FALSE;
    BOOL fRetVal = FALSE;

    if (!psldl)
    {
        // We will detect failure of this at use time.
        if (FAILED(CoCreateInstanceHook(CLSID_ShellLink, NULL, CLSCTX_INPROC, IID_PPV_ARG(IShellLinkDataList, &psldl))))
        {
            return FALSE;
        }

        fFreesldl = TRUE;

        IPersistFile* ppf;
        OLECHAR sz[MAX_PATH];
        if (SHGetPathFromIDListW(pidlFull, sz))
        {
            if (SUCCEEDED(psldl->QueryInterface(IID_PPV_ARG(IPersistFile, &ppf))))
            {
                if (SUCCEEDED(ppf->Load(sz, 0)))
                {
                    fIsLoaded = TRUE;
                }
                ppf->Release();
            }
        }
    }
    else
        fIsLoaded = TRUE;

    CDarwinAd* pda = NULL;
    if (fIsLoaded)
    {
        EXP_DARWIN_LINK* pexpDarwin;

        if (SUCCEEDED(psldl->CopyDataBlock(EXP_DARWIN_ID_SIG, (void**)&pexpDarwin)))
        {
            fRetVal = SHRegisterDarwinLink(ILClone(pidlFull), pexpDarwin->szwDarwinID, TRUE);
            LocalFree(pexpDarwin);
        }
    }

    if (fFreesldl)
        psldl->Release();

    return fRetVal;
}

// This routine creates the IShellFolder and pidl for one of the many
// merged folders on the Start Menu / Start Panel.

typedef struct
{
    UINT csidl;
    UINT uANSFlags;
    const GUID* pguidObj;
    int field_C;
} MERGEDFOLDERINFO, * LPMERGEDFOLDERINFO;

typedef const MERGEDFOLDERINFO* LPCMERGEDFOLDERINFO;

#define MAKEINTIDLIST(csidl)    (LPCITEMIDLIST)MAKEINTRESOURCE(csidl)

MIDL_INTERFACE("76347b91-9846-4ce7-9a57-69b910d16123")
ISetFolderEnumRestriction : IUnknown
{
    virtual HRESULT SetEnumRestriction(DWORD dwRequired, DWORD dwForbidden) PURE;
};

DEFINE_GUID(POLID_NoStartMenuSubFolders, 0x78627E11, 0xEC80, 0x49A3, 0xB3, 0xB7, 0x6A, 0xBE, 0x73, 0x96, 0x93, 0xFC);

EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** psfOut, ITEMIDLIST_ABSOLUTE** pidlOut);

DEFINE_GUID(CLSID_StartMenuCommon, 0x2981C306, 0x09EA, 0x405D, 0x9D, 0x0F, 0x83, 0x33, 0x78, 0x27, 0xBD, 0x35);
DEFINE_GUID(CLSID_ProgramsFolderCommon, 0xFC1EE10B, 0x7EF6, 0x41B5, 0xBB, 0x60, 0x98, 0xD2, 0x6D, 0xD9, 0xFC, 0xD1);

HRESULT CreateMergedFolderHelper(REFCLSID rclsid, const MERGEDFOLDERINFO* rgmfi, UINT cmfi, REFIID riid, void** ppv)
{
    *ppv = nullptr;

    IAugmentedShellFolder* pasf;
    HRESULT hr = CoCreateInstance(CLSID_MergedFolder, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pasf));
    if (SUCCEEDED(hr))
    {
        ITEMIDLIST* pidl;
        hr = BindToGetFolderAndPidl(rclsid, nullptr, &pidl);
        if (SUCCEEDED(hr))
        {
            for (UINT i = 0; SUCCEEDED(hr) && i < cmfi; ++i)
            {
                const MERGEDFOLDERINFO* prgmfi = &rgmfi[i];
                if ((prgmfi->uANSFlags & 4) != 0 && SHRestricted(REST_NOCOMMONGROUPS))
                {
                    continue;
                }

                PERSIST_FOLDER_TARGET_INFO pfti = {};
                pfti.dwAttributes = FILE_ATTRIBUTE_SYSTEM | FILE_ATTRIBUTE_DIRECTORY;
                pfti.csidl = prgmfi->csidl | CSIDL_FLAG_PFTI_TRACKTARGET;

                IShellFolder2* psf;
                hr = CFSFolder_CreateFolder(nullptr, nullptr, pidl, &pfti, IID_PPV_ARGS(&psf));
                if (SUCCEEDED(hr))
                {
                    if (prgmfi->pguidObj == &CLSID_StartMenu || prgmfi->pguidObj == &CLSID_StartMenuCommon)
                    {
                        ISetFolderEnumRestriction* prest;
                        if (SHWindowsPolicy(POLID_NoStartMenuSubFolders) && SUCCEEDED(psf->QueryInterface(IID_PPV_ARGS(&prest))))
                        {
                            prest->SetEnumRestriction(0, SHCONTF_FOLDERS);
                            prest->Release();
                        }
                    }
                    else
                    {
                        ASSERT(rgmfi[i].pguidObj == nullptr || !IsEqualGUID(*rgmfi[i].pguidObj, CLSID_StartMenu)
                            && !IsEqualGUID(*rgmfi[i].pguidObj, CLSID_StartMenuCommon)); // 370
                    }

                    hr = pasf->AddNameSpace(rgmfi[i].pguidObj, psf, nullptr, rgmfi[i].uANSFlags, rgmfi[i].field_C);
                    psf->Release();
                }
            }
            ILFree(pidl);

            if (SUCCEEDED(hr))
            {
                hr = pasf->QueryInterface(riid, ppv);
            }
        }
        pasf->Release();
    }

    return hr;
}

const MERGEDFOLDERINFO c_rgmfiStartMenu[] =
{
    { 0x800B, 0xFF02, &CLSID_StartMenu, 1 },
    { 0x0016, 0x0006, &CLSID_StartMenuCommon, 1 }
};

const MERGEDFOLDERINFO c_rgmfiProgramsFolder[] =
{
    { 0x8002u, 0xFF02u, &CLSID_ProgramsFolder, 2 },
    { 0x0017u, 0x0006u, &CLSID_ProgramsFolderCommon, 2 }
};

const MERGEDFOLDERINFO c_rgmfiProgramsFolderAndFastItems[] =
{
    { 0x8002u, 0xFF0Au, &CLSID_ProgramsFolder, 2 },
    { 0x0017u, 0x000Eu, &CLSID_ProgramsFolderCommon, 2 },
    { 0x800Bu, 0x000Au, &CLSID_StartMenu, 1 },
    { 0x0016u, 0x000Eu, &CLSID_StartMenu, 1 }
};

EXTERN_C HRESULT CStartMenuFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
{
    return CreateMergedFolderHelper(CLSID_StartMenuFolder, c_rgmfiStartMenu, ARRAYSIZE(c_rgmfiStartMenu), riid, ppv);
}

EXTERN_C HRESULT CProgramsFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
{
    return CreateMergedFolderHelper(
        CLSID_ProgramsFolder, c_rgmfiProgramsFolder, ARRAYSIZE(c_rgmfiProgramsFolder), riid, ppv);
}

EXTERN_C HRESULT CProgramsFolderAndFastItems_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv)
{
    return CreateMergedFolderHelper(
        CLSID_ProgramsFolderAndFastItems, c_rgmfiProgramsFolderAndFastItems,
        ARRAYSIZE(c_rgmfiProgramsFolderAndFastItems), riid, ppv);
}

HRESULT GetFilesystemInfo(IShellFolder* psf, ITEMIDLIST_ABSOLUTE** ppidlRoot, int* pcsidl)
{
    ASSERT(psf);

    *pcsidl = 0;
    *ppidlRoot = nullptr;

    HRESULT hr = E_FAIL;

    IPersistFolder3* ppf;
    if (SUCCEEDED(psf->QueryInterface(IID_PPV_ARGS(&ppf))))
    {
        PERSIST_FOLDER_TARGET_INFO pfti = {};
        if (SUCCEEDED(ppf->GetFolderTargetInfo(&pfti)))
        {
            *pcsidl = pfti.csidl;
            if (pfti.csidl != -1)
            {
                hr = S_OK;
            }

            ILFree(pfti.pidlTargetFolder);
        }

        if (SUCCEEDED(hr))
        {
            hr = ppf->GetCurFolder(ppidlRoot);
        }
        ppf->Release();
    }

    return hr;
}

#define IDM_MYPICTURES          518

HRESULT ExecStaticStartMenuItem(int idCmd, BOOL fAllUsers, BOOL fOpen)
{
    int csidl = -1;

    HRESULT hr = E_OUTOFMEMORY;
    SHELLEXECUTEINFO shei = { 0 };
    switch (idCmd)
    {
        case IDM_PROGRAMS:          csidl = fAllUsers ? CSIDL_COMMON_PROGRAMS : CSIDL_PROGRAMS; break;
        case IDM_FAVORITES:         csidl = CSIDL_FAVORITES; break;
        case IDM_MYDOCUMENTS:       csidl = CSIDL_PERSONAL; break;
        case IDM_MYPICTURES:        csidl = CSIDL_MYPICTURES; break;
        case IDM_CONTROLS:          csidl = CSIDL_CONTROLS;  break;
        case IDM_PRINTERS:          csidl = CSIDL_PRINTERS;  break;
        case IDM_NETCONNECT:        csidl = CSIDL_CONNECTIONS; break;
        default:
            return E_FAIL;
    }

    if (csidl != -1)
    {
        SHGetFolderLocation(NULL, csidl, NULL, 0, (LPITEMIDLIST*)&shei.lpIDList);
    }

    if (shei.lpIDList)
    {
        shei.cbSize = sizeof(shei);
        shei.fMask = SEE_MASK_IDLIST;
        shei.nShow = SW_SHOWNORMAL;
        shei.lpVerb = fOpen ? TEXT("open") : TEXT("explore");
        hr = ShellExecuteEx(&shei) ? S_OK : E_FAIL;
        ILFree((LPITEMIDLIST)shei.lpIDList);
    }

    return hr;
}

DEFINE_GUID(CLSID_StartMenuGroupPolicyFilter, 0x23613363, 0x0028, 0x431D, 0xA4, 0x9E, 0xA3, 0xCD, 0x48, 0x2D, 0x39, 0x26);

MIDL_INTERFACE("8cc5baf5-11e9-4dd2-9810-45214134c800")
IStartMenuGroupPolicyFilter : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE IsFileRestrictedByPolicyOrSystemMetrics(const WCHAR*) = 0;
};

//
//  Base class for Classic and Personal start menus.
//

class CStartMenuCallbackBase : public CObjectWithSite
{
public:
    //~ Begin IUnknown Interface
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

protected:
    CStartMenuCallbackBase(BOOL fIsStartPanel = FALSE);
    ~CStartMenuCallbackBase() override;

    void _InitializePrograms();
    HRESULT _FilterPidl(UINT uParent, IShellFolder* psf, const ITEMID_CHILD* pidl);
    HRESULT _Promote(LPSMDATA psmd, DWORD dwFlags);
    HRESULT _HandleNew(LPSMDATA psmd);
    HRESULT _GetSFInfo(SMDATA* psmd, SMINFO* psminfo);
    HRESULT _ProcessChangeNotify(SMDATA* psmd, LONG lEvent, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2);

    virtual DWORD _GetDemote(SMDATA* psmd) { return 0; }
    BOOL _IsDarwinAdvertisement(LPCITEMIDLIST pidlFull);

    void _RefreshSettings();

protected:
    LONG _cRef;

#ifdef DEBUG
    DWORD _dwThreadID;
#endif

    WCHAR* _pszPrograms;
    WCHAR* _pszAdminTools;
    WCHAR* _pszCommonPrograms;

    ITrayPriv2* _ptp2;

    BOOL            _fExpandoMenus;
    BOOL            _fShowAdminTools;
    BOOL            _fIsStartPanel;
    BOOL            _fInitPrograms;

    IStartMenuGroupPolicyFilter* _pgpf;
};

MIDL_INTERFACE("fe787bcb-0ee8-44fb-8c89-12f508913c40")
IMruDataList : IUnknown
{

    typedef int (*MRUDATALISTCOMPARE)(const BYTE*, const BYTE*, int);

    enum
    {
        MRULISTF_USE_MEMCMP = 0x0000,   // default, find uses memcmp()
        MRULISTF_USE_STRCMPIW = 0x0001,   // find uses StrCmpIW()
        MRULISTF_USE_STRCMPW = 0x0002,   // find uses StrCmpW()
        MRULISTF_USE_ILISEQUAL = 0x0003,   // find uses ILIsEqual()
    };
    typedef DWORD MRULISTF;

    virtual HRESULT InitData(
        UINT uMax,
        MRULISTF flags,
        HKEY hKey,
         LPCWSTR pszSubKey,
        MRUDATALISTCOMPARE pfnCompare) PURE;

    virtual HRESULT AddData(
        const BYTE* pData,
        DWORD cbData,
        DWORD* pdwSlot) PURE;

    virtual HRESULT FindData(
        const BYTE* pData,
        DWORD cbData,
        int* piIndex) PURE;

    virtual HRESULT GetData(
        int iIndex,
        BYTE* pData,
        DWORD cbData) PURE;

    virtual HRESULT QueryInfo(
        int iIndex,
        DWORD* pdwSlot,
        DWORD* pcbData) PURE;

    virtual HRESULT Delete(int iIndex) PURE;
};

class CStartMenuCallback : public IShellMenuCallback, public CStartMenuCallbackBase
{
public:
    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

    //~ Begin IShellMenuCallback Interface
    STDMETHODIMP CallbackSM(SMDATA* psmd, UINT uMsg, WPARAM wParam, LPARAM lParam) override;
    //~ End IShellMenuCallback Interface

    //~ Begin CObjectWithSite::IObjectWithSite Interface
    STDMETHODIMP SetSite(IUnknown* punk) override;
    STDMETHODIMP GetSite(REFIID riid, void** ppvOut) override;
    //~ End CObjectWithSite::IObjectWithSite Interface

    CStartMenuCallback();

    HRESULT InitializeFastItemsShellMenu(IShellMenu* psm);
    HRESULT InitializeCSIDLShellMenu(int uId, int csidl, LPCTSTR pszRoot, LPCTSTR pszValue,
        DWORD dwPassInitFlags, DWORD dwSetFlags, BOOL fAddOpen,
        IShellMenu* psm);
    HRESULT InitializeDocumentsShellMenu(IShellMenu* psm);
    HRESULT InitializeProgramsShellMenu(IShellMenu* psm);
    HRESULT InitializeSubShellMenu(int idCmd, IShellMenu* psm);

private:
    ~CStartMenuCallback() override;

    IContextMenu* _pcmFind;
    ITrayPriv* _ptp;
    IUnknown* _punkSite;
    IOleCommandTarget* _poct;
    BITBOOL         _fAddOpenFolder : 1;
    BITBOOL         _fCascadeMyDocuments : 1;
    BITBOOL         _fCascadePrinters : 1;
    BITBOOL         _fCascadeControlPanel : 1;
    BITBOOL         _fFindMenuInvalid : 1;
    BITBOOL         _fCascadeNetConnections : 1;
    BITBOOL         _fShowInfoTip : 1;
    BITBOOL         _fInitedShowTopLevelStartMenu : 1;
    BITBOOL         _fCascadeMyPictures : 1;

    BITBOOL         _fHasMyDocuments : 1;
    BITBOOL         _fHasMyPictures : 1;

    TCHAR           _szFindMnemonic[2];

    HWND            _hwnd;

    IMruDataList* _pmruRecent;
    DWORD           _cRecentDocs;

    DWORD           _dwFlags;
    DWORD           _dwChevronCount;

    HRESULT _ExecHmenuItem(SMDATA* psmdata);
    HRESULT _Init(SMDATA* psmdata);
    HRESULT _Create(SMDATA* psmdata, void** pvUserData);
    HRESULT _Destroy(SMDATA* psmdata);
    HRESULT _GetHmenuInfo(SMDATA* psmd, SMINFO* sminfo);
    HRESULT _GetObject(SMDATA* psmd, REFIID riid, void** ppvObj);
    HRESULT _FilterRecentPidl(IShellFolder* psf, PCUITEMID_CHILD pidl);
    HRESULT _Demote(SMDATA* psmd);
    HRESULT _GetTip(WCHAR* pstrTitle, WCHAR* pstrTip);
    DWORD _GetDemote(SMDATA* psmd);
    HRESULT _HandleAccelerator(WCHAR ch, SMDATA* psmdata);
    HRESULT _GetDefaultIcon(WCHAR* psz, int* piIndex);
    void _GetStaticStartMenu(HMENU* phmenu, HWND* phwnd);
    HRESULT _GetStaticInfoTip(SMDATA* psmd, WCHAR* pszTip, int cch);

    // helper functions
    DWORD GetInitFlags();
    void SetInitFlags(DWORD dwFlags);
    HRESULT _InitializeFindMenu(IShellMenu* psm);
    HRESULT _ExecItem(LPSMDATA, UINT);
    HRESULT VerifyCSIDL(int idCmd, int csidl, IShellMenu* psm);
    HRESULT VerifyMergedGuy(BOOL fPrograms, IShellMenu* psm);
    void _UpdateDocsMenuItemNames(IShellMenu* psm);
    void _UpdateDocumentsShellMenu(IShellMenu* psm);
};


class CStartContextMenu : IContextMenu
{
public:
    // IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void);

    // IContextMenu
    STDMETHOD(QueryContextMenu)(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    STDMETHOD(InvokeCommand)(LPCMINVOKECOMMANDINFO lpici);
    STDMETHOD(GetCommandString)(UINT_PTR idCmd, UINT uType, UINT* pRes, LPSTR pszName, UINT cchMax);

    CStartContextMenu(int idCmd) : _idCmd(idCmd), _cRef(1) {};
private:
    int _cRef;
    virtual ~CStartContextMenu() {};

    int _idCmd;
};

void CStartMenuCallbackBase::_RefreshSettings()
{
    _fShowAdminTools = FeatureEnabled(L"StartMenuAdminTools");
}

CStartMenuCallbackBase::CStartMenuCallbackBase(BOOL fIsStartPanel)
    : _cRef(1)
    , _fIsStartPanel(fIsStartPanel)
{
#ifdef DEBUG
    _dwThreadID = GetCurrentThreadId();
#endif

    WCHAR szBuf[260];
    if (SUCCEEDED(SHGetFolderPathEx(FOLDERID_CommonAdminTools, KF_FLAG_CREATE, nullptr, szBuf, ARRAYSIZE(szBuf))))
    {
        Str_SetPtrW(&_pszAdminTools, PathFindFileNameW(szBuf));
    }

    _RefreshSettings();
    CoCreateInstance(CLSID_StartMenuGroupPolicyFilter, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&_pgpf));
    SHReValidateDarwinCache();
}

#define IDS_FIND_MNEMONIC       0x7674

CStartMenuCallback::CStartMenuCallback()
    : _cRecentDocs(-1)
{
    LoadString(LoadLibraryW(L"shell32.dll"), IDS_FIND_MNEMONIC, _szFindMnemonic, ARRAYSIZE(_szFindMnemonic));
}

CStartMenuCallbackBase::~CStartMenuCallbackBase()
{
    ASSERT(_dwThreadID == GetCurrentThreadId()); // 701

    Str_SetPtrW(&_pszAdminTools, nullptr);
    Str_SetPtrW(&_pszPrograms, nullptr);
    Str_SetPtrW(&_pszCommonPrograms, nullptr);

    IUnknown_SafeReleaseAndNullPtr(&_pgpf);
    IUnknown_SafeReleaseAndNullPtr(&_ptp2);
}

CStartMenuCallback::~CStartMenuCallback()
{
    ATOMICRELEASE(_pcmFind);
    ATOMICRELEASE(_ptp);
    ATOMICRELEASE(_pmruRecent);
}

ULONG CStartMenuCallbackBase::AddRef()
{
    return ++_cRef;
}

ULONG CStartMenuCallbackBase::Release()
{
    ASSERT(0 != _cRef); // 739

    LONG cRef = InterlockedDecrement(&_cRef);
    if (cRef == 0 && this)
    {
        delete this;
    }
    return cRef;
}

STDMETHODIMP CStartMenuCallback::SetSite(IUnknown* punk)
{
    ATOMICRELEASE(_punkSite);
    _punkSite = punk;
    if (punk)
    {
        _punkSite->AddRef();
    }

    return S_OK;
}

STDMETHODIMP CStartMenuCallback::GetSite(REFIID riid, void** ppvOut)
{
    if (_ptp)
        return _ptp->QueryInterface(riid, ppvOut);
    else
        return E_NOINTERFACE;
}

#ifdef DEBUG
void DBUEMQueryEvent(const IID* pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam)
{
#if 1
    return;
#else
    UAINFO uei;

    uei.cbSize = sizeof(uei);
    uei.dwMask = ~0;    // UEIM_HIT etc.
    UEMQueryEvent(pguidGrp, eCmd, wParam, lParam, &uei);

    TCHAR szBuf[20];
    wsprintf(szBuf, TEXT("hit=%d"), uei.cHit);
    MessageBox(NULL, szBuf, TEXT("UEM"), MB_OKCANCEL);

    return;
#endif
}
#endif

DWORD CStartMenuCallback::GetInitFlags()
{
    DWORD dwType;
    DWORD cbSize = sizeof(DWORD);
    DWORD dwFlags = 0;
    SHGetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("StartMenuInit"),
        &dwType, (BYTE*)&dwFlags, &cbSize);
    return dwFlags;
}

void CStartMenuCallback::SetInitFlags(DWORD dwFlags)
{
    SHSetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("StartMenuInit"), REG_DWORD, &dwFlags, sizeof(DWORD));
}

DWORD GetClickCount()
{

    //This function retrieves the number of times the user has clicked on the chevron item.

    DWORD dwType;
    DWORD cbSize = sizeof(DWORD);
    DWORD dwCount = 1;      // Default to three clicks before we give up.
    // PMs what it to 1 now. Leaving back end in case they change their mind.
    SHGetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("StartMenuChevron"),
        &dwType, (BYTE*)&dwCount, &cbSize);

    return dwCount;

}

void SetClickCount(DWORD dwClickCount)
{
    SHSetValue(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED, TEXT("StartMenuChevron"), REG_DWORD, &dwClickCount, sizeof(DWORD));
}

void _ValidateShellNoRoam(HKEY hk)
{
    WCHAR szOld[MAX_COMPUTERNAME_LENGTH + 1] = L"";
    WCHAR szNew[MAX_COMPUTERNAME_LENGTH + 1] = L"";
    DWORD cb = sizeof(szOld);
    SHGetValueW(hk, NULL, NULL, NULL, szOld, &cb);
    cb = ARRAYSIZE(szNew);
    GetComputerNameW(szNew, &cb);
    if (StrCmpICW(szNew, szOld))
    {
        //  need to delete this key's kids
        SHDeleteKey(hk, NULL);
        SHSetValueW(hk, NULL, NULL, REG_SZ, szNew, CbFromCchW(lstrlenW(szNew) + 1));
    }
}

LANGID MLGetUILanguage(void)
{
    static LANGID LangID = 0;
    CHAR szLangID[8] = "";

    if (0 == LangID)  // no cached LANGID
    {
        if (true)
        {
            static LANGID(CALLBACK * pfnGetUserDefaultUILanguage)(void) = NULL;

            if (pfnGetUserDefaultUILanguage == NULL)
            {
                HMODULE hmod = GetModuleHandle(TEXT("KERNEL32"));

                if (hmod)
                    pfnGetUserDefaultUILanguage = (LANGID(CALLBACK*)(void))GetProcAddress(hmod, "GetUserDefaultUILanguage");
            }
            if (pfnGetUserDefaultUILanguage)
                return pfnGetUserDefaultUILanguage();
        }

    }
    return LangID;
}

void _ValidateMUICache(HKEY hk)
{
    LANGID lidOld = 0;
    //  if we are running on legacy platforms, we aggressively invalidate
    LANGID lidNew = MLGetUILanguage();
    DWORD cb = sizeof(lidOld);
    SHGetValueW(hk, NULL, L"LangID", NULL, &lidOld, &cb);

    if (lidOld != lidNew)
    {
        SHDeleteKey(hk, NULL);
        SHSetValueW(hk, NULL, L"LangID", REG_BINARY, &lidNew, sizeof(lidNew));
    }
}

typedef void (*PFNVALIDATE)(HKEY);

typedef struct
{
    LPCWSTR psz;
    DWORD dwOption;
    PFNVALIDATE pfnValidate;
    HKEY hkCU;
    HKEY hkLM;
} SKCACHE;

#define SKENTRY(s)  {s, REG_OPTION_NON_VOLATILE, NULL, NULL, NULL}
#define SKENTRYOPT(s, o)  {s, o, NULL, NULL, NULL}
#define SKENTRYVAL(s, pfnV) {s, REG_OPTION_NON_VOLATILE, pfnV, NULL, NULL}

static SKCACHE s_skPath[] =
{
    SKENTRY(L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer"),
    SKENTRY(L"Software\\Microsoft\\Windows\\Shell"),
    SKENTRYVAL(L"Software\\Microsoft\\Windows\\ShellNoRoam", _ValidateShellNoRoam),
    SKENTRY(L"Software\\Classes"),
};

static SKCACHE s_skSub[] =
{
    SKENTRY(L"LocalizedResourceName"),
    SKENTRY(L"Handlers"),
    SKENTRY(L"Associations"),
    SKENTRYOPT(L"Volatile", REG_OPTION_VOLATILE),
    SKENTRYVAL(L"MUICache", _ValidateMUICache),
    SKENTRY(L"FileExts"),
};

LONG RegCreateKeyExWrapW(HKEY hKey, LPCTSTR lpSubKey, DWORD Reserved, LPTSTR lpClass, DWORD dwOptions, REGSAM samDesired, LPSECURITY_ATTRIBUTES lpSecurityAttributes, PHKEY phkResult, LPDWORD lpdwDisposition)
{


    {
        return RegCreateKeyExW(hKey, lpSubKey, Reserved, lpClass, dwOptions, samDesired, lpSecurityAttributes, phkResult, lpdwDisposition);
    }

}

LONG APIENTRY
RegOpenKeyExWrapW(
    HKEY    hKey,
    LPCWSTR lpSubKey,
    DWORD   ulOptions,
    REGSAM  samDesired,
    PHKEY   phkResult)
{

    {
        return RegOpenKeyExW(hKey, lpSubKey, ulOptions, samDesired, phkResult);
    }

}

HKEY _OpenKey(HKEY hk, LPCWSTR psz, BOOL fCreate, DWORD dwOption)
{
    HKEY hkRet = NULL;
    DWORD err;
    if (fCreate && psz)
    {
        DWORD dwDisp;
        err = RegCreateKeyExWrapW(hk, psz, 0, NULL, dwOption, MAXIMUM_ALLOWED, NULL, &hkRet, &dwDisp);
    }
    else
    {
        err = RegOpenKeyExWrapW(hk, psz, 0, MAXIMUM_ALLOWED, &hkRet);
    }

    if (!hkRet)
    {
        //  if ERROR_KEY_DELETED
        //  should we invalidate our cache??
        //  cause we will fail forever...
        SetLastError(err);
    }

    return hkRet;
}

HKEY _OpenSKCache(HKEY hk, BOOL fHKLM, BOOL fNoCaching, BOOL fCreateSub, SKCACHE* psk, DWORD* pdwOption)
{
    HKEY hkSub = fHKLM ? psk->hkLM : psk->hkCU;
    *pdwOption = psk->dwOption;

    if (!hkSub || fNoCaching)
    {
        hkSub = _OpenKey(hk, psk->psz, fCreateSub, psk->dwOption);
        if (hkSub)
        {
            if (psk->pfnValidate)
                psk->pfnValidate(hkSub);

            if (!fNoCaching)
            {
                ENTERCRITICAL;
                HKEY* phk = fHKLM ? &psk->hkLM : &psk->hkCU;
                if (!*phk)
                {
                    *phk = hkSub;
                }
                else
                {
                    RegCloseKey(hkSub);
                    hkSub = *phk;
                }
                LEAVECRITICAL;
            }
        }
    }
    return hkSub;
}

#define HKEY_FROM_SKROOT(sk)    ((sk & SKROOT_MASK) == SKROOT_HKLM ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER)

typedef enum _SHELLKEY
{
    SKROOT_HKCU = 0x00000001,       //  internal to the function
    SKROOT_HKLM = 0x00000002,       //  internal to the function
    SKROOT_MASK = 0x0000000F,       //  internal to the function
    SKPATH_EXPLORER = 0x00000000,       //  internal to the function
    SKPATH_SHELL = 0x00000010,       //  internal to the function
    SKPATH_SHELLNOROAM = 0x00000020,       //  internal to the function
    SKPATH_CLASSES = 0x00000030,       //  internal to the function
    SKPATH_MASK = 0x00000FF0,       //  internal to the function
    SKSUB_NONE = 0x00000000,       //  internal to the function
    SKSUB_LOCALIZEDNAMES = 0x00001000,       //  internal to the function
    SKSUB_HANDLERS = 0x00002000,       //  internal to the function
    SKSUB_ASSOCIATIONS = 0x00003000,       //  internal to the function
    SKSUB_VOLATILE = 0x00004000,       //  internal to the function
    SKSUB_MUICACHE = 0x00005000,       //  internal to the function
    SKSUB_FILEEXTS = 0x00006000,       //  internal to the function
    SKSUB_MASK = 0x000FF000,       //  internal to the function

    SHELLKEY_HKCU_EXPLORER = SKROOT_HKCU | SKPATH_EXPLORER | SKSUB_NONE,
    SHELLKEY_HKLM_EXPLORER = SKROOT_HKLM | SKPATH_EXPLORER | SKSUB_NONE,
    SHELLKEY_HKCU_SHELL = SKROOT_HKCU | SKPATH_SHELL | SKSUB_NONE,
    SHELLKEY_HKLM_SHELL = SKROOT_HKLM | SKPATH_SHELL | SKSUB_NONE,
    SHELLKEY_HKCU_SHELLNOROAM = SKROOT_HKCU | SKPATH_SHELLNOROAM | SKSUB_NONE,
    SHELLKEY_HKCULM_SHELL = SHELLKEY_HKCU_SHELLNOROAM,
    SHELLKEY_HKCULM_CLASSES = SKROOT_HKCU | SKPATH_CLASSES | SKSUB_NONE,
    SHELLKEY_HKCU_LOCALIZEDNAMES = SKROOT_HKCU | SKPATH_SHELL | SKSUB_LOCALIZEDNAMES,
    SHELLKEY_HKCULM_HANDLERS = SKROOT_HKCU | SKPATH_SHELLNOROAM | SKSUB_HANDLERS,
    SHELLKEY_HKCULM_ASSOCIATIONS = SKROOT_HKCU | SKPATH_SHELLNOROAM | SKSUB_ASSOCIATIONS,
    SHELLKEY_HKCULM_VOLATILE = SKROOT_HKCU | SKPATH_SHELLNOROAM | SKSUB_VOLATILE,
    SHELLKEY_HKCULM_MUICACHE = SKROOT_HKCU | SKPATH_SHELLNOROAM | SKSUB_MUICACHE,
    SHELLKEY_HKCU_FILEEXTS = SKROOT_HKCU | SKPATH_EXPLORER | SKSUB_FILEEXTS,

    SHELLKEY_HKCULM_HANDLERS_RO = SHELLKEY_HKCULM_HANDLERS,      //  deprecated
    SHELLKEY_HKCULM_HANDLERS_RW = SHELLKEY_HKCULM_HANDLERS,      //  deprecated
    SHELLKEY_HKCULM_ASSOCIATIONS_RO = SHELLKEY_HKCULM_ASSOCIATIONS,    //  deprecated
    SHELLKEY_HKCULM_ASSOCIATIONS_RW = SHELLKEY_HKCULM_ASSOCIATIONS,    //  deprecated
    SHELLKEY_HKCULM_RO = SHELLKEY_HKCU_SHELLNOROAM,     //  deprecated
    SHELLKEY_HKCULM_RW = SHELLKEY_HKCU_SHELLNOROAM,     //  deprecated
} SHELLKEY;

HKEY _OpenShellKey(SHELLKEY sk, HKEY hkRoot, BOOL fNoCaching, BOOL fCreateSub, DWORD* pdwOption)
{
    BOOL fHKLM = (sk & SKROOT_MASK) == SKROOT_HKLM;
    ULONG uPath = (sk & SKPATH_MASK) >> 4;
    ULONG uSub = (sk & SKSUB_MASK) >> 12;

    ASSERT(uPath < ARRAYSIZE(s_skPath));
    HKEY hkPath = NULL;
    if (uPath < ARRAYSIZE(s_skPath))
    {
        hkPath = _OpenSKCache(hkRoot, fHKLM, fNoCaching, fCreateSub, &s_skPath[uPath], pdwOption);
    }
    else
        SetLastError(E_INVALIDARG);

    //  see if there is a sub value to add
    if (hkPath && uSub != SKSUB_NONE && --uSub < ARRAYSIZE(s_skSub))
    {
        HKEY hkSub = _OpenSKCache(hkPath, fHKLM, fNoCaching, fCreateSub, &s_skSub[uSub], pdwOption);
        if (fNoCaching)
            RegCloseKey(hkPath);
        hkPath = hkSub;
    }

    return hkPath;
}
typedef LONG(*PFNREGOPENCURRENTUSER) (REGSAM, HKEY*);
LONG NT5_RegOpenCurrentUser(REGSAM sam, HKEY* phk)
{
    static PFNREGOPENCURRENTUSER s_pfn = (PFNREGOPENCURRENTUSER)-1;

    if (s_pfn == (PFNREGOPENCURRENTUSER)-1)
    {

        s_pfn = (PFNREGOPENCURRENTUSER)GetProcAddress(GetModuleHandle(TEXT("advapi32")), "RegOpenCurrentUser");

    }

    if (s_pfn)
    {
        return s_pfn(sam, phk);
    }
    else
    {
        *phk = NULL;
        return ERROR_CAN_NOT_COMPLETE;
    }
}

HKEY _GetRootKey(SHELLKEY sk, BOOL* pfNoCaching)
{
    HKEY hkRoot = HKEY_FROM_SKROOT(sk);
    HANDLE hToken;
    if (hkRoot == HKEY_CURRENT_USER && OpenThreadToken(GetCurrentThread(), TOKEN_QUERY | TOKEN_IMPERSONATE, TRUE, &hToken))
    {
        //  we dont support ARBITRARY tokens
        //  but RegOpenCurrentUser() opens the current thread token
        NT5_RegOpenCurrentUser(MAXIMUM_ALLOWED, &hkRoot);
        //  if we wanted to we would have to do something
        //  like shell32!GetUserProfileKey(hToken, &hkRoot);

        CloseHandle(hToken);
    }

    *pfNoCaching = HKEY_FROM_SKROOT(sk) != hkRoot;
    return hkRoot;
}

HKEY SHGetShellKey(SHELLKEY sk, LPCWSTR pszSubKey, BOOL fCreateSub)
{
    BOOL fNoCaching;
    HKEY hkRoot = _GetRootKey(sk, &fNoCaching);
    HKEY hkRet = NULL;
    if (hkRoot)
    {
        DWORD dwOption;
        HKEY hkPath = _OpenShellKey(sk, hkRoot, fNoCaching, fCreateSub, &dwOption);

        //  this duplicates when there is no subkey
        if (hkPath)
        {
            hkRet = _OpenKey(hkPath, pszSubKey, fCreateSub, dwOption);

            if (fNoCaching)
                RegCloseKey(hkPath);
        }

        if (fNoCaching)
            RegCloseKey(hkRoot);
    }
    else
        SetLastError(ERROR_ACCESS_DENIED);

    return hkRet;
}
#define REGSTR_KEY_RECENTDOCS TEXT("RecentDocs")
#define MAXRECENT_DEFAULTDOC      10
#define MAXRECENT_MAJORDOC        20
#define MAXRECENTDOCS 15

enum
{
    MRULISTF_USE_MEMCMP = 0x0000,   // default, find uses memcmp()
    MRULISTF_USE_STRCMPIW = 0x0001,   // find uses StrCmpIW()
    MRULISTF_USE_STRCMPW = 0x0002,   // find uses StrCmpW()
    MRULISTF_USE_ILISEQUAL = 0x0003,   // find uses ILIsEqual()
};
#define SRMLF_COMPPIDL  0x00000001   // use the pidl in the recent folder

typedef struct
{
    WORD        cb;                     // pidl size
    BYTE        bFlags;                 // SHID_FS_* bits
    DWORD       dwSize;                 // -1 implies > 4GB, hit the disk to get the real size
    WORD        dateModified;
    WORD        timeModified;
    WORD        wAttrs;                 // FILE_ATTRIBUTES_* cliped to 16bits
    CHAR        cFileName[MAX_PATH];    // this is WCHAR for names that don't round trip
    CHAR        cAltFileName[8 + 1 + 3 + 1];  // ANSI version of cFileName (some chars not converted)
} IDFOLDER;
typedef UNALIGNED IDFOLDER* LPIDFOLDER;
typedef const UNALIGNED IDFOLDER* LPCIDFOLDER;
#define SHID_GROUPMASK          0x70
#define SHID_FS                   0x30  // base simple IDList, we don't generate these anymore

#define GETRECNAME(p) ((LPCTSTR)(p))
#define GETRECPIDL(p) ((LPCITEMIDLIST) (((LPBYTE) (p)) + CbFromCchW(lstrlen(GETRECNAME(p)) +1)))
LPCIDFOLDER CFSFolder_IsValidID(LPCITEMIDLIST pidl)
{
    if (pidl && pidl->mkid.cb && (((LPCIDFOLDER)pidl)->bFlags & SHID_GROUPMASK) == SHID_FS)
        return (LPCIDFOLDER)pidl;

    return NULL;
}


HRESULT CreateRecentMRUList(IMruDataList** ppmru)
{
    *ppmru = 0;
    return *ppmru ? S_OK : E_OUTOFMEMORY;
}

#define SMC_DESTROY             0x0000002B  // Called when a pane is being destroyed.
#define SMC_EXEC                0x00000004  // The callback is called to execute an item
#define SMC_GETINFOTIP          0x0000000D  // The callback is called to get some object
#define SMC_BEGINENUM           0x00000013  // tell callback that we are beginning to ENUM the indicated parent
#define SMC_ENDENUM             0x00000014  // tell callback that we are ending the ENUM of the indicated paren
#define SMC_DUMPONUPDATE        0x00000035  // S_OK if host wants old trash-everything-on-update behavior (recent docs)
#define SMC_INSERTINDEX         0x0000000E  // New item insert index
#define SMSET_MERGE                 0x00000002
#define SMC_MAPACCELERATOR      0x00000015  // Called when processing an accelerator.
#define SMC_GETMINPROMOTED      0x00000018  // Returns the minimum number of promoted items
#define STARTMENU_CHEVRONCLICKED        0x00000002

HRESULT CStartMenuCallback::QueryInterface(const IID& riid, void** ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CStartMenuCallback, IShellMenuCallback), // IID_IShellMenuCallback
        QITABENT(CStartMenuCallback, IObjectWithSite), // IID_IObjectWithSite
        {},
    };
    return QISearch(this, qit, riid, ppvObj);
}

ULONG CStartMenuCallback::AddRef()
{
    return CStartMenuCallbackBase::AddRef();
}

ULONG CStartMenuCallback::Release()
{
    return CStartMenuCallbackBase::Release();
}

STDMETHODIMP CStartMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    HRESULT hr = S_FALSE;
    switch (uMsg)
    {

        case SMC_CREATE:
            hr = _Create(psmd, (void**)lParam);
            break;

        case SMC_DESTROY:
            hr = _Destroy(psmd);
            break;

        case SMC_INITMENU:
            hr = _Init(psmd);
            break;

        case SMC_SFEXEC:
            hr = _ExecItem(psmd, uMsg);
            break;

        case SMC_EXEC:
            hr = _ExecHmenuItem(psmd);
            break;

        case SMC_GETOBJECT:
            hr = _GetObject(psmd, (GUID) * ((GUID*)wParam), (void**)lParam);
            break;

        case SMC_GETINFO:
            hr = _GetHmenuInfo(psmd, (SMINFO*)lParam);
            break;

        case SMC_GETSFINFOTIP:
            if (!_fShowInfoTip)
                hr = E_FAIL;  // E_FAIL means don't show. S_FALSE means show default
            break;

        case SMC_GETINFOTIP:
            hr = _GetStaticInfoTip(psmd, (LPWSTR)wParam, (int)lParam);
            break;

        case SMC_GETSFINFO:
            hr = _GetSFInfo(psmd, (SMINFO*)lParam);
            break;

        case SMC_BEGINENUM:
            if (psmd->uIdParent == IDM_RECENT)
            {
                ASSERT(_cRecentDocs == -1);
                ASSERT(!_pmruRecent);
                CreateRecentMRUList(&_pmruRecent);

                _cRecentDocs = 0;
                hr = S_OK;
            }
            break;

        case SMC_ENDENUM:
            if (psmd->uIdParent == IDM_RECENT)
            {
                ASSERT(_cRecentDocs != -1);
                ATOMICRELEASE(_pmruRecent);

                _cRecentDocs = -1;
                hr = S_OK;
            }
            break;

        case SMC_DUMPONUPDATE:
            if (psmd->uIdParent == IDM_RECENT)
            {
                hr = S_OK;
            }
            break;

        case SMC_FILTERPIDL:
            ASSERT(psmd->dwMask & SMDM_SHELLFOLDER);

            if (psmd->uIdParent == IDM_RECENT)
            {
                //  we need to filter out all but the first MAXRECENTITEMS
                //  and no folders allowed!
                hr = _FilterRecentPidl(psmd->psf, psmd->pidlItem);
            }
            else
            {
                hr = _FilterPidl(psmd->uIdParent, psmd->psf, psmd->pidlItem);
            }
            break;

        case SMC_INSERTINDEX:
            ASSERT(lParam && IS_VALID_WRITE_PTR(lParam, int));
            *((int*)lParam) = 0;
            hr = S_OK;
            break;

        case SMC_SHCHANGENOTIFY:
        {
            PSMCSHCHANGENOTIFYSTRUCT pshf = (PSMCSHCHANGENOTIFYSTRUCT)lParam;
            hr = _ProcessChangeNotify(psmd, pshf->lEvent, pshf->pidl1, pshf->pidl2);
        }
        break;

        case SMC_REFRESH:
            if (psmd->uIdParent == IDM_TOPLEVELSTARTMENU)
            {
                hr = S_OK;

                // Refresh is only called on the top level.
                HMENU hmenu;
                IShellMenu* psm;
                _GetStaticStartMenu(&hmenu, &_hwnd);
                if (hmenu && psmd->punk && SUCCEEDED(psmd->punk->QueryInterface(IID_PPV_ARG(IShellMenu, &psm))))
                {
                    hr = psm->SetMenu(hmenu, _hwnd, SMSET_BOTTOM | SMSET_MERGE);
                    psm->Release();
                }

                _RefreshSettings();
                _fExpandoMenus = !_fIsStartPanel && AreIntelliMenusEnabled();
                _fCascadeMyDocuments = FeatureEnabled(TEXT("CascadeMyDocuments"));
                _fCascadePrinters = FeatureEnabled(TEXT("CascadePrinters"));
                _fCascadeControlPanel = FeatureEnabled(TEXT("CascadeControlPanel"));
                _fCascadeNetConnections = FeatureEnabled(TEXT("CascadeNetworkConnections"));
                _fAddOpenFolder = FeatureEnabled(TEXT("StartMenuOpen"));
                _fShowInfoTip = FeatureEnabled(TEXT("ShowInfoTip"));
                _fCascadeMyPictures = FeatureEnabled(TEXT("CascadeMyPictures"));
                _fFindMenuInvalid = TRUE;
                _dwFlags = GetInitFlags();
            }
            break;

        case SMC_DEMOTE:
            hr = _Demote(psmd);
            break;

        case SMC_PROMOTE:
            hr = _Promote(psmd, (DWORD)wParam);
            break;

        case SMC_NEWITEM:
            hr = _HandleNew(psmd);
            break;

        case SMC_MAPACCELERATOR:
            hr = _HandleAccelerator((TCHAR)wParam, (SMDATA*)lParam);
            break;

        case SMC_DEFAULTICON:
            ASSERT(psmd->uIdAncestor == IDM_FAVORITES); // This is only valid for the Favorites menu
            hr = _GetDefaultIcon((LPWSTR)wParam, (int*)lParam);
            break;

        case SMC_GETMINPROMOTED:
            // Only do this for the programs menu
            if (psmd->uIdParent == IDM_PROGRAMS)
                *((int*)lParam) = 4;        // 4 was choosen by RichSt 9.15.98
            break;

        case SMC_CHEVRONEXPAND:

            // Has the user already seen the chevron tip enough times? (We set the bit when the count goes to zero.
            if (!(_dwFlags & STARTMENU_CHEVRONCLICKED))
            {
                // No; Then get the current count from the registry. We set a default of 3, but an admin can set this
                // to -1, that would make it so that they user sees it all the time.
                DWORD dwClickCount = GetClickCount();
                if (dwClickCount > 0)
                {
                    // Since they clicked, take one off.
                    dwClickCount--;

                    // Set it back in.
                    SetClickCount(dwClickCount);
                }

                if (dwClickCount == 0)
                {
                    // Ah, the user has seen the chevron tip enought times... Stop being annoying.
                    _dwFlags |= STARTMENU_CHEVRONCLICKED;
                    SetInitFlags(_dwFlags);
                }
            }
            hr = S_OK;
            break;

        case SMC_DISPLAYCHEVRONTIP:

            // We only want to see the tip on the top level programs case, no where else. We also don't
            // want to see it if they've had enough.
            if (psmd->uIdParent == IDM_PROGRAMS &&
                !(_dwFlags & STARTMENU_CHEVRONCLICKED) &&
                !SHRestricted(REST_NOSMBALLOONTIP))
            {
                hr = S_OK;
            }
            break;

        case SMC_CHEVRONGETTIP:
            if (!SHRestricted(REST_NOSMBALLOONTIP))
                hr = _GetTip((LPWSTR)wParam, (LPWSTR)lParam);
            break;
    }

    return hr;
}

// For the Favorites menu, since their icon handler is SO slow, we're going to fake the icon
// and have it get the real ones on the background thread...
HRESULT CStartMenuCallback::_GetDefaultIcon(LPWSTR psz, int* piIndex)
{
    DWORD cbSize = MAX_PATH;
    HRESULT hr = AssocQueryString(0, ASSOCSTR_DEFAULTICON, TEXT("InternetShortcut"), NULL, psz, &cbSize);
    if (SUCCEEDED(hr))
    {
        *piIndex = PathParseIconLocation(psz);
    }

    return hr;
}

HRESULT CStartMenuCallback::_ExecItem(LPSMDATA psmd, UINT uMsg)
{
    ASSERT(_dwThreadID == GetCurrentThreadId());
    return _ptp->ExecItem(psmd->psf, psmd->pidlItem);
}

HRESULT CStartMenuCallback::_Demote(LPSMDATA psmd)
{
    //We want to for the UEM to demote pidlFolder, 
    // then tell the Parent menuband (If there is one)
    // to invalidate this pidl.
    HRESULT hr = S_FALSE;

    if (_fExpandoMenus &&
        (psmd->uIdAncestor == IDM_PROGRAMS ||
            psmd->uIdAncestor == IDM_FAVORITES))
    {
        UAINFO uei;
        uei.cbSize = sizeof(uei);
        uei.dwMask = UEIM_HIT;
        uei.cLaunches = 0;
        hr = UEMSetEvent(psmd->uIdAncestor == IDM_PROGRAMS ? &UEMIID_SHELL : &UEMIID_BROWSER,
            UEME_RUNPIDL, (WPARAM)psmd->psf, (LPARAM)psmd->pidlItem, &uei);
    }
    return hr;
}

#define SMINV_FORCE          0x00000080
#define UEMIID_SHELL    CLSID_ActiveDesktop     // FEATURE need better one
#define UEMIID_BROWSER  CLSID_InternetToolbar   // FEATURE need better one
#define UEME_RUNPIDL    18
#define UEMF_EVENTMON   0x00000001 
#define UEMF_INSTRUMENT 0x00000002
#define UEMF_XEVENT     (UEMF_EVENTMON | UEMF_INSTRUMENT)

// Even if intellimenus are off, fire a UEM event if it was an Exec from
// the More Programs menu of the Start Panel [SMINV_FORCE will be set]
// so we can detect which are the user's most popular apps.

HRESULT CStartMenuCallbackBase::_Promote(LPSMDATA psmd, DWORD dwFlags)
{
    if ((_fExpandoMenus || (_fIsStartPanel && (dwFlags & SMINV_FORCE))) &&
        (psmd->uIdAncestor == IDM_PROGRAMS ||
            psmd->uIdAncestor == IDM_FAVORITES))
    {
        UEMFireEvent(psmd->uIdAncestor == IDM_PROGRAMS ? &UEMIID_SHELL : &UEMIID_BROWSER,
            UEME_RUNPIDL, UEMF_XEVENT, (WPARAM)psmd->psf, (LPARAM)psmd->pidlItem);
    }
    return S_OK;
}

HRESULT CStartMenuCallbackBase::_HandleNew(LPSMDATA psmd)
{
    HRESULT hr = S_FALSE;
    if (_fExpandoMenus &&
        (psmd->uIdAncestor == IDM_PROGRAMS ||
            psmd->uIdAncestor == IDM_FAVORITES))
    {
        UAINFO uei;
        uei.cbSize = sizeof(uei);
        uei.dwMask = UEIM_HIT;
        uei.cLaunches = UEM_NEWITEMCOUNT;
        hr = UEMSetEvent(psmd->uIdAncestor == IDM_PROGRAMS ? &UEMIID_SHELL : &UEMIID_BROWSER,
            UEME_RUNPIDL, (WPARAM)psmd->psf, (LPARAM)psmd->pidlItem, &uei);
    }

    if (psmd->uIdAncestor == IDM_PROGRAMS)
    {
        LPITEMIDLIST pidlFull = FullPidlFromSMData(psmd);
        if (pidlFull)
        {
            ProcessDarwinAd(nullptr, pidlFull);
            ILFree(pidlFull);
        }
    }
    return hr;
}

HRESULT ShowFolder(UINT csidl)
{
    LPITEMIDLIST pidl;
    if (SUCCEEDED(SHGetFolderLocation(NULL, csidl, NULL, 0, &pidl)))
    {
        SHELLEXECUTEINFO shei = { 0 };

        shei.cbSize = sizeof(shei);
        shei.fMask = SEE_MASK_IDLIST;
        shei.nShow = SW_SHOWNORMAL;
        shei.lpVerb = TEXT("open");
        shei.lpIDList = pidl;
        ShellExecuteEx(&shei);
        ILFree(pidl);
    }
    return S_OK;
}

void _ExecRegValue(LPCTSTR pszValue)
{
    TCHAR szPath[MAX_PATH];
    DWORD cbSize = ARRAYSIZE(szPath);

    if (ERROR_SUCCESS == SHGetValue(HKEY_LOCAL_MACHINE, REGSTR_EXPLORER_ADVANCED, pszValue,
        NULL, szPath, &cbSize))
    {
        SHELLEXECUTEINFO shei = { 0 };
        shei.cbSize = sizeof(shei);
        shei.nShow = SW_SHOWNORMAL;
        shei.lpParameters = PathGetArgs(szPath);
        PathRemoveArgs(szPath);
        shei.lpFile = szPath;
        ShellExecuteEx(&shei);
    }
}

//#define IsInRange(item,min_val,max_val) \
//            (((item) >= min_val) && ((item) <= max_val))

extern "C" void WINAPI SetICIKeyModifiers(DWORD* pfMask)
{
    //assert(pfMask);

    if (GetKeyState(VK_SHIFT) < 0)
    {
        *pfMask |= CMIC_MASK_SHIFT_DOWN;
    }

    if (GetKeyState(VK_CONTROL) < 0)
    {
        *pfMask |= CMIC_MASK_CONTROL_DOWN;
    }
}
#define IDM_BLANKITEM           515
#define IDM_MYDOCUMENTS         516
#define IDM_OPEN_FOLDER         517
#define IDM_MYPICTURES          518
#define IDM_CSC                 553
HRESULT CStartMenuCallback::_ExecHmenuItem(LPSMDATA psmd)
{
    HRESULT hr = S_FALSE;
    if (IsInRange(psmd->uId, TRAY_IDM_FINDFIRST, TRAY_IDM_FINDLAST) && _pcmFind)
    {
        CMINVOKECOMMANDINFOEX ici = { 0 };
        ici.cbSize = sizeof(CMINVOKECOMMANDINFOEX);
        ici.lpVerb = (LPSTR)MAKEINTRESOURCE(psmd->uId - TRAY_IDM_FINDFIRST);
        ici.nShow = SW_NORMAL;

        // record if shift or control was being held down
        SetICIKeyModifiers(&ici.fMask);

        _pcmFind->InvokeCommand((LPCMINVOKECOMMANDINFO)&ici);
        hr = S_OK;
    }
    else
    {
        switch (psmd->uId)
        {
            case IDM_OPEN_FOLDER:
                switch (psmd->uIdParent)
                {
                    case IDM_CONTROLS:
                        hr = ShowFolder(CSIDL_CONTROLS);
                        break;

                    case IDM_PRINTERS:
                        hr = ShowFolder(CSIDL_PRINTERS);
                        break;

                    case IDM_NETCONNECT:
                        hr = ShowFolder(CSIDL_CONNECTIONS);
                        break;

                    case IDM_MYPICTURES:
                        hr = ShowFolder(CSIDL_MYPICTURES);
                        break;

                    case IDM_MYDOCUMENTS:
                        hr = ShowFolder(CSIDL_PERSONAL);
                        break;
                }
                break;

            case IDM_NETCONNECT:
                hr = ShowFolder(CSIDL_CONNECTIONS);
                break;

            case IDM_MYDOCUMENTS:
                hr = ShowFolder(CSIDL_PERSONAL);
                break;

            case IDM_MYPICTURES:
                hr = ShowFolder(CSIDL_MYPICTURES);
                break;

            case IDM_CSC:
                _ExecRegValue(TEXT("StartMenuSyncAll"));
                break;

            default:
                hr = ExecStaticStartMenuItem(psmd->uId, FALSE, TRUE);
                break;
        }
    }
    return hr;
}

void CStartMenuCallback::_GetStaticStartMenu(HMENU* phmenu, HWND* phwnd)
{
    *phmenu = nullptr;
    *phwnd = nullptr;

    IMenuPopup* pmp;
    if (SUCCEEDED(IUnknown_QueryService(_punkSite, SID_SMenuPopup, IID_PPV_ARGS(&pmp))))
    {
        if (SUCCEEDED(IUnknown_GetSite(pmp, IID_PPV_ARGS(&_ptp))))
        {
            _ptp->QueryInterface(IID_PPV_ARGS(&_ptp2));
            _ptp->GetStaticStartMenu(phmenu);
            IUnknown_GetWindow(_ptp, phwnd);

            if (_poct == nullptr)
            {
                _ptp->QueryInterface(IID_PPV_ARGS(&_poct));
            }
        }
        else
        {
            //CcshellDebugMsgW(0x2000000, "CStartMenuCallback::_SetSite : Failed to aquire CStartMenuHost");
        }

        pmp->Release();
    }
}

HRESULT CStartMenuCallback::_Create(SMDATA* psmdata, void** ppvUserData)
{
    *ppvUserData = new SMUSERDATA;

    return S_OK;
}

HRESULT CStartMenuCallback::_Destroy(SMDATA* psmdata)
{
    if (psmdata->pvUserData)
    {
        delete (SMUSERDATA*)psmdata->pvUserData;
        psmdata->pvUserData = NULL;
    }

    return S_OK;
}

void CStartMenuCallbackBase::_InitializePrograms()
{
    if (!_fInitPrograms)
    {
        WCHAR szTemp[260];
        SHGetFolderPathEx(FOLDERID_Programs, KF_FLAG_DEFAULT, nullptr, szTemp, ARRAYSIZE(szTemp));
        Str_SetPtrW(&_pszPrograms, PathFindFileNameW(szTemp));

        SHGetFolderPathEx(FOLDERID_CommonPrograms, KF_FLAG_DEFAULT, nullptr, szTemp, ARRAYSIZE(szTemp));
        Str_SetPtrW(&_pszCommonPrograms, PathFindFileNameW(szTemp));

        _fInitPrograms = TRUE;
    }
}



// Given a CSIDL and a Shell menu, this will verify if the IShellMenu
// is pointing at the same place as the CSIDL is. If not, then it will
// update the shell menu to the new location.
HRESULT CStartMenuCallback::VerifyCSIDL(int idCmd, int csidl, IShellMenu* psm)
{
    DWORD dwFlags;
    LPITEMIDLIST pidl;
    IShellFolder* psf;
    HRESULT hr = S_OK;
    if (SUCCEEDED(psm->GetShellFolder(&dwFlags, &pidl, IID_PPV_ARG(IShellFolder, &psf))))
    {
        psf->Release();

        LPITEMIDLIST pidlCSIDL;
        if (SUCCEEDED(SHGetFolderLocation(NULL, csidl, NULL, 0, &pidlCSIDL)))
        {
            // If the pidl of the IShellMenu is not equal to the
            // SpecialFolder Location, then we need to update it so they are...
            if (!ILIsEqual(pidlCSIDL, pidl))
            {
                hr = InitializeSubShellMenu(idCmd, psm);
            }
            ILFree(pidlCSIDL);
        }
        ILFree(pidl);
    }

    return hr;
}

HRESULT CStartMenuCallback::VerifyMergedGuy(BOOL fPrograms, IShellMenu* psm)
{
    HRESULT hr = S_OK;

    DWORD dwFlags;
    ITEMIDLIST* pidl;
    IAugmentedShellFolder* pasf;
    if (SUCCEEDED(psm->GetShellFolder(&dwFlags, &pidl, IID_PPV_ARGS(&pasf))))
    {
        for (int i = 0; i < 2; ++i)
        {
            IShellFolder* psf;
            if (SUCCEEDED(pasf->QueryNameSpace(i, nullptr, &psf)))
            {
                ITEMIDLIST_ABSOLUTE* pidlFolder;
                int csidl;
                if (SUCCEEDED(GetFilesystemInfo(psf, &pidlFolder, &csidl)))
                {
                    ITEMIDLIST_ABSOLUTE* pidlCSIDL;
                    if (SUCCEEDED(SHGetFolderLocation(nullptr, csidl, nullptr, 0, &pidlCSIDL)))
                    {
                        if (!ILIsEqual(pidlCSIDL, pidlFolder))
                        {
                            _fInitPrograms = FALSE;
                            if (fPrograms)
                            {
                                hr = InitializeProgramsShellMenu(psm);
                            }
                            else
                            {
                                hr = InitializeFastItemsShellMenu(psm);
                            }
                            // i = 100;
                            break; // @MOD break directly rather than setting i to 100 for some reason.
                        }
                        ILFree(pidlCSIDL);
                    }
                    ILFree(pidlFolder);
                }
                psf->Release();
            }
        }

        ILFree(pidl);
        pasf->Release();
    }

    return hr;
}

void _FixMenuItemName(IShellMenu* psm, UINT uID, LPTSTR pszNewMenuName)
{
    HMENU hMenu;
    ASSERT(NULL != psm);
    if (SUCCEEDED(psm->GetMenu(&hMenu, NULL, NULL)))
    {
        MENUITEMINFO mii = { 0 };
        TCHAR szMenuName[256];
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_TYPE;
        mii.dwTypeData = szMenuName;
        mii.cch = ARRAYSIZE(szMenuName);
        szMenuName[0] = TEXT('\0');
        if (::GetMenuItemInfo(hMenu, uID, FALSE, &mii))
        {
            if (0 != StrCmp(szMenuName, pszNewMenuName))
            {
                // The mydocs name has changed, update the menu item:
                mii.dwTypeData = pszNewMenuName;
                if (::SetMenuItemInfo(hMenu, uID, FALSE, &mii))
                {
                    SMDATA smd;
                    smd.dwMask = SMDM_HMENU;
                    smd.uId = uID;
                    psm->InvalidateItem(&smd, SMINV_ID | SMINV_REFRESH);
                }
            }
        }
    }
}

#define MENU_STARTMENU_MYDOCS           401

void CStartMenuCallback::_UpdateDocumentsShellMenu(IShellMenu* psm)
{
    // Add/Remove My Documents and My Pictures items of menu

    BOOL fMyDocs = !SHRestricted(REST_NOSMMYDOCS);
    if (fMyDocs)
    {
        LPITEMIDLIST pidl;
        fMyDocs = SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_PERSONAL, NULL, 0, &pidl));
        if (fMyDocs)
            ILFree(pidl);
    }

    BOOL fMyPics = !SHRestricted(REST_NOSMMYPICS);
    if (fMyPics)
    {
        LPITEMIDLIST pidl;
        fMyPics = SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_MYPICTURES, NULL, 0, &pidl));
        if (fMyPics)
            ILFree(pidl);
    }

    HMENU hMenu = SHLoadMenuPopup(LoadLibraryW(L"shell32.dll"), MENU_STARTMENU_MYDOCS);
    if (hMenu)
    {
        // Modern Windows has this blank item that XP doesn't use.
        DeleteMenu(hMenu, IDM_BLANKITEM, MF_BYCOMMAND);
        if (!fMyDocs)
            DeleteMenu(hMenu, IDM_MYDOCUMENTS, MF_BYCOMMAND);
        if (!fMyPics)
            DeleteMenu(hMenu, IDM_MYPICTURES, MF_BYCOMMAND);
        // Reset section of menu
        psm->SetMenu(hMenu, _hwnd, SMSET_TOP);
    }

    // Cache what folders are available
    _fHasMyDocuments = fMyDocs;
    _fHasMyPictures = fMyPics;
}

STDAPI GetCSIDLDisplayName(int csidl, WCHAR* pszPath, UINT cch)
{
    *pszPath = 0;

    ITEMIDLIST_ABSOLUTE* pidl;
    if (SUCCEEDED(SHGetFolderLocation(nullptr, csidl, nullptr, 0, &pidl)))
    {
        SHGetNameAndFlags(pidl, SHGDN_NORMAL, pszPath, cch, nullptr);
        ILFree(pidl);
    }
    return *pszPath ? S_OK : E_FAIL;
}

void CStartMenuCallback::_UpdateDocsMenuItemNames(IShellMenu* psm)
{
    TCHAR szBuffer[MAX_PATH];

    if (_fHasMyDocuments)
    {
        if (SUCCEEDED(GetCSIDLDisplayName(CSIDL_MYDOCUMENTS, szBuffer, ARRAYSIZE(szBuffer))))
            _FixMenuItemName(psm, IDM_MYDOCUMENTS, szBuffer);
    }

    if (_fHasMyPictures)
    {
        if (SUCCEEDED(GetCSIDLDisplayName(CSIDL_MYPICTURES, szBuffer, ARRAYSIZE(szBuffer))))
            _FixMenuItemName(psm, IDM_MYPICTURES, szBuffer);
    }
}

HRESULT CStartMenuCallback::_Init(SMDATA* psmdata)
{
    HRESULT hr = S_FALSE;
    IShellMenu* psm;
    if (psmdata->punk && SUCCEEDED(hr = psmdata->punk->QueryInterface(IID_PPV_ARG(IShellMenu, &psm))))
    {
        switch (psmdata->uIdParent)
        {
            case IDM_TOPLEVELSTARTMENU:
            {
                if (psmdata->pvUserData && !((SMUSERDATA*)psmdata->pvUserData)->_fInitialized)
                {
                    //TraceMsg(TF_MENUBAND, "CStartMenuCallback::_Init : Initializing Toplevel Start Menu");
                    ((SMUSERDATA*)psmdata->pvUserData)->_fInitialized = TRUE;

                    HMENU hmenu;

                    //TraceMsg(TF_MENUBAND, "CStartMenuCallback::_Init : First Time, and correct parameters");

                    _GetStaticStartMenu(&hmenu, &_hwnd);
                    if (hmenu)
                    {
                        HMENU   hmenuOld = NULL;
                        HWND    hwnd;
                        DWORD   dwFlags;

                        psm->GetMenu(&hmenuOld, &hwnd, &dwFlags);
                        if (hmenuOld != NULL)
                        {
                            TBOOL(DestroyMenu(hmenuOld));
                        }
                        hr = psm->SetMenu(hmenu, _hwnd, SMSET_BOTTOM);
                        //TraceMsg(TF_MENUBAND, "CStartMenuCallback::_Init : SetMenu(HMENU 0x%x, HWND 0x%x", hmenu, _hwnd);
                    }

                    _fExpandoMenus = !_fIsStartPanel && AreIntelliMenusEnabled();
                    _fCascadeMyDocuments = FeatureEnabled(TEXT("CascadeMyDocuments"));
                    _fCascadePrinters = FeatureEnabled(TEXT("CascadePrinters"));
                    _fCascadeControlPanel = FeatureEnabled(TEXT("CascadeControlPanel"));
                    _fCascadeNetConnections = FeatureEnabled(TEXT("CascadeNetworkConnections"));
                    _fAddOpenFolder = FeatureEnabled(TEXT("StartMenuOpen"));
                    _fShowInfoTip = FeatureEnabled(TEXT("ShowInfoTip"));
                    _fCascadeMyPictures = FeatureEnabled(TEXT("CascadeMyPictures"));
                    _dwFlags = GetInitFlags();
                }
                else if (!_fInitedShowTopLevelStartMenu)
                {
                    _fInitedShowTopLevelStartMenu = TRUE;
                    psm->InvalidateItem(NULL, SMINV_REFRESH);
                }

                // Verify that the Fast items is still pointing to the right location
                if (SUCCEEDED(hr))
                {
                    hr = VerifyMergedGuy(FALSE, psm);
                }
                break;
            }
            case IDM_MENU_FIND:
                if (_fFindMenuInvalid)
                {
                    hr = _InitializeFindMenu(psm);
                    _fFindMenuInvalid = FALSE;
                }
                break;

            case IDM_PROGRAMS:
                // Verify the programs menu is still pointing to the right location
                hr = VerifyMergedGuy(TRUE, psm);
                break;

            case IDM_FAVORITES:
                hr = VerifyCSIDL(IDM_FAVORITES, CSIDL_FAVORITES, psm);
                break;

            case IDM_MYDOCUMENTS:
                hr = VerifyCSIDL(IDM_MYDOCUMENTS, CSIDL_PERSONAL, psm);
                break;

            case IDM_MYPICTURES:
                hr = VerifyCSIDL(IDM_MYPICTURES, CSIDL_MYPICTURES, psm);
                break;

            case IDM_RECENT:
                _UpdateDocumentsShellMenu(psm);
                _UpdateDocsMenuItemNames(psm);
                hr = VerifyCSIDL(IDM_RECENT, CSIDL_RECENT, psm);
                break;
            case IDM_CONTROLS:
                hr = VerifyCSIDL(IDM_CONTROLS, CSIDL_CONTROLS, psm);
                break;
            case IDM_PRINTERS:
                hr = VerifyCSIDL(IDM_PRINTERS, CSIDL_PRINTERS, psm);
                break;
        }

        psm->Release();
    }

    return hr;
}
#define IDS_CONTROL_TIP         0x768A
#define IDS_PRINTERS_TIP        0x768B
#define IDS_TRAYPROP_TIP        0x768C
#define IDS_MYDOCS_TIP          0x768D
#define IDS_NETCONNECT_TIP      0x768E
#define IDS_MYPICS_TIP          0x76A5

HRESULT CStartMenuCallback::_GetStaticInfoTip(SMDATA* psmd, LPWSTR pszTip, int cch)
{
    if (!_fShowInfoTip)
        return E_FAIL;

    HRESULT hr = E_FAIL;

    const static struct
    {
        UINT idCmd;
        UINT idInfoTip;
    } s_mpcmdTip[] =
    {
#if 0   // No tips for the Toplevel. Keep this here because I bet that someone will want them...
       { IDM_PROGRAMS,       IDS_PROGRAMS_TIP },
       { IDM_FAVORITES,      IDS_FAVORITES_TIP },
       { IDM_RECENT,         IDS_RECENT_TIP },
       { IDM_SETTINGS,       IDS_SETTINGS_TIP },
       { IDM_MENU_FIND,      IDS_FIND_TIP },
       { IDM_HELPSEARCH,     IDS_HELP_TIP },        // Redundant?
       { IDM_FILERUN,        IDS_RUN_TIP },
       { IDM_LOGOFF,         IDS_LOGOFF_TIP },
       { IDM_EJECTPC,        IDS_EJECT_TIP },
       { IDM_EXITWIN,        IDS_SHUTDOWN_TIP },
#endif
       // Settings Submenu
       { IDM_CONTROLS,       IDS_CONTROL_TIP },
       { IDM_PRINTERS,       IDS_PRINTERS_TIP },
       { IDM_TRAYPROPERTIES, IDS_TRAYPROP_TIP },
       { IDM_NETCONNECT,     IDS_NETCONNECT_TIP },

       // Recent Folder
       { IDM_MYDOCUMENTS,    IDS_MYDOCS_TIP },
       { IDM_MYPICTURES,     IDS_MYPICS_TIP },
    };


    for (int i = 0; i < ARRAYSIZE(s_mpcmdTip); i++)
    {
        if (s_mpcmdTip[i].idCmd == psmd->uId)
        {
            TCHAR szTip[MAX_PATH];
            if (LoadString(LoadLibraryW(L"shell32.dll"), s_mpcmdTip[i].idInfoTip, szTip, ARRAYSIZE(szTip)))
            {
                SHTCharToUnicode(szTip, pszTip, cch);
                hr = S_OK;
            }
            break;
        }
    }

    return hr;
}

typedef struct
{
    WCHAR wszMenuText[MAX_PATH];
    WCHAR wszHelpText[MAX_PATH];
    int   iIcon;
} SEARCHEXTDATA, * LPSEARCHEXTDATA;

#define IDM_FILERUN                 401
#define IDM_LOGOFF                  402
#define IDM_EJECTPC                 410
#define IDM_SETTINGSASSIST          411
#define IDM_TRAYPROPERTIES          413
#define IDM_UPDATEWIZARD            414
#define IDM_UPDATE_SEP              415
#define IDM_MU_DISCONNECT           5000
#define IDM_MU_SECURITY             5001

#define IDI_DVDDRIVE                291
#define IDI_MEDIACDAUDIOPLUS        292
#define IDI_MEDIACDEMPTY            293
#define IDI_MEDIACDROM              294
#define IDI_MEDIACDR                295
#define IDI_MEDIACDRW               296
#define IDI_MEDIADVDRAM             297
#define IDI_MEDIADVDR               298
#define IDI_AUDIOPLAYER             299
#define IDI_DEVICETAPEDRIVE         300
#define IDI_DEVICEOPTICALDRIVE      301
#define IDI_MEDIABLANKCD            302
#define IDI_MEDIACOMPFLASH          303
#define IDI_MEDIADVDROM             304
#define IDI_MEDIAMEMSTICK           305
#define IDI_MEDIAPCMCIA             306
#define IDI_MEDIASECUREDIGITALMEDIA 307
#define IDI_MEDIASMARTMEDIA         308
#define IDI_DEVICECAMERA            309
#define IDI_DEVICECELLPHONE         310
#define IDI_DEVICEHTTPPRINT         311
#define IDI_DEVICEJAZDRIVE          312
#define IDI_DEVICEZIPDRIVE          313
#define IDI_DEVICEPOCKETPC          314
#define IDI_DEVICESCANNER           315
#define IDI_DEVICESTI               316
#define IDI_DEVICEVIDEOCAM          317
#define IDI_MEDIADVDRW              318
#define IDI_TASK_NEWFOLDER          319
#define IDI_TASK_SENDTOCD           320
#define IDI_CPTASK_32CPLS           321
#define IDI_CLASSICSM_FAVORITES     322
#define IDI_CLASSICSM_FIND          323
#define IDI_CLASSICSM_HELP          324
#define IDI_CLASSICSM_LOGOFF        325
#define IDI_CLASSICSM_PROGS         326
#define IDI_CLASSICSM_RECENTDOCS    327
#define IDI_CLASSICSM_RUN           328
#define IDI_CLASSICSM_SHUTDOWN      329
#define IDI_CLASSICSM_SETTINGS      330
#define IDI_CLASSICSM_UNDOCK        331
#define IDI_TASK_SEARCHDS           337
#define IDI_NONE                    338
#define II_MU_STSECURITY     47
#define II_MU_STDISCONN      48

#define II_STCPANEL          35
#define II_STSPROGS          36
#define II_STPRNTRS          37
#define II_STFONTS           38
#define II_STTASKBR          39

#define II_CDAUDIO           40
#define II_TREE              41
#define II_STCPROGS          42
#define II_STFAVORITES       43
#define II_STLOGOFF          44
#define II_STFLDRPROP        45
#define II_WINUPDATE         46

#define IDI_MYDOCS                      100
#define IDI_MYPICS                      101
#define IDI_CSC                 179     // ClientSideCaching
#define IDI_NETCONNECT          175

HRESULT CStartMenuCallback::_GetHmenuInfo(SMDATA* psmd, SMINFO* psminfo)
{
    const static struct
    {
        UINT idCmd;
        int  iImage;
    } s_mpcmdimg[] = { // Top level menu
                       { IDM_PROGRAMS,       -IDI_CLASSICSM_PROGS },
                       { IDM_FAVORITES,      -IDI_CLASSICSM_FAVORITES },
                       { IDM_RECENT,         -IDI_CLASSICSM_RECENTDOCS },
                       { IDM_SETTINGS,       -IDI_CLASSICSM_SETTINGS },
                       { IDM_MENU_FIND,      -IDI_CLASSICSM_FIND },
                       { IDM_HELPSEARCH,     -IDI_CLASSICSM_HELP },
                       { IDM_FILERUN,        -IDI_CLASSICSM_RUN },
                       { IDM_LOGOFF,         -IDI_CLASSICSM_LOGOFF },
                       { IDM_EJECTPC,        -IDI_CLASSICSM_UNDOCK },
                       { IDM_EXITWIN,        -IDI_CLASSICSM_SHUTDOWN },
                       { IDM_MU_SECURITY,    II_MU_STSECURITY },
                       { IDM_MU_DISCONNECT,  II_MU_STDISCONN  },
                       { IDM_SETTINGSASSIST, -IDI_CLASSICSM_SETTINGS },
                       { IDM_CONTROLS,       II_STCPANEL },
                       { IDM_PRINTERS,       II_STPRNTRS },
                       { IDM_TRAYPROPERTIES, II_STTASKBR },
                       { IDM_MYDOCUMENTS,    -IDI_MYDOCS},
                       { IDM_CSC,            -IDI_CSC},
                       { IDM_NETCONNECT,     -IDI_NETCONNECT},
    };


    ASSERT(IS_VALID_WRITE_PTR(psminfo, SMINFO));

    int iIcon = -1;
    DWORD dwFlags = psminfo->dwFlags;
    MENUITEMINFO mii = { 0 };
    HRESULT hr = S_FALSE;

    if (psminfo->dwMask & SMIM_ICON)
    {
        if (IsInRange(psmd->uId, TRAY_IDM_FINDFIRST, TRAY_IDM_FINDLAST))
        {
            // The find menu extensions pack their icon into their data member of
            // Menuiteminfo....
            mii.cbSize = sizeof(mii);
            mii.fMask = MIIM_DATA;
            if (GetMenuItemInfo(psmd->hmenu, psmd->uId, MF_BYCOMMAND, &mii))
            {
                LPSEARCHEXTDATA psed = (LPSEARCHEXTDATA)mii.dwItemData;

                if (psed)
                    psminfo->iIcon = mii.dwItemData;
                else
                    psminfo->iIcon = -1;

                dwFlags |= SMIF_ICON;
                hr = S_OK;
            }
        }
        else
        {
            if (psmd->uId == IDM_MYPICTURES)
            {
                LPITEMIDLIST pidlMyPics = SHCloneSpecialIDList(NULL, CSIDL_MYPICTURES, FALSE);
                if (pidlMyPics)
                {
                    LPCITEMIDLIST pidlObject;
                    IShellFolder* psf;
                    hr = SHBindToParent(pidlMyPics, IID_PPV_ARG(IShellFolder, &psf), &pidlObject);
                    if (SUCCEEDED(hr))
                    {
                        SHMapPIDLToSystemImageListIndex(psf, pidlObject, &psminfo->iIcon);
                        dwFlags |= SMIF_ICON;
                        psf->Release();
                    }
                    ILFree(pidlMyPics);
                }
            }
            else
            {
                UINT uIdLocal = psmd->uId;
                if (uIdLocal == IDM_OPEN_FOLDER)
                    uIdLocal = psmd->uIdAncestor;

                for (int i = 0; i < ARRAYSIZE(s_mpcmdimg); i++)
                {
                    if (s_mpcmdimg[i].idCmd == uIdLocal)
                    {
                        iIcon = s_mpcmdimg[i].iImage;
                        break;
                    }
                }

                if (iIcon != -1)
                {
                    dwFlags |= SMIF_ICON;
                    psminfo->iIcon = Shell_GetCachedImageIndex(TEXT("shell32.dll"), iIcon, 0);
                    hr = S_OK;
                }
            }
        }
    }

    if (psminfo->dwMask & SMIM_FLAGS)
    {
        psminfo->dwFlags = dwFlags;

        if ((psmd->uId == IDM_CONTROLS && _fCascadeControlPanel) ||
            (psmd->uId == IDM_PRINTERS && _fCascadePrinters) ||
            (psmd->uId == IDM_MYDOCUMENTS && _fCascadeMyDocuments) ||
            (psmd->uId == IDM_NETCONNECT && _fCascadeNetConnections) ||
            (psmd->uId == IDM_MYPICTURES && _fCascadeMyPictures))
        {
            psminfo->dwFlags |= SMIF_SUBMENU;
            hr = S_OK;
        }
        else switch (psmd->uId)
        {
            case IDM_FAVORITES:
            case IDM_PROGRAMS:
                psminfo->dwFlags |= SMIF_DROPCASCADE;
                hr = S_OK;
                break;
        }
    }

    return hr;
}

DWORD CStartMenuCallback::_GetDemote(SMDATA* psmd)
{
    UAINFO uei;
    DWORD dwFlags = 0;

    uei.cbSize = sizeof(uei);
    uei.dwMask = UEIM_HIT;
    if (SUCCEEDED(UEMQueryEvent(psmd->uIdAncestor == IDM_PROGRAMS ? &UEMIID_SHELL : &UEMIID_BROWSER,
        UEME_RUNPIDL, (WPARAM)psmd->psf, (LPARAM)psmd->pidlItem, &uei)))
    {
        if (uei.cLaunches == 0)
        {
            dwFlags |= SMIF_DEMOTED;
        }
    }

    return dwFlags;
}

//
//  WARNING!  Since this function returns a pointer from our Darwin cache,
//  it must be called while the Darwin critical section is held.
//
int GetDarwinIndex(LPCITEMIDLIST pidlFull, CDarwinAd** ppda)
{
    int iRet = -1;
    if (g_hdpaDarwinAds)
    {
        int chdpa = DPA_GetPtrCount(g_hdpaDarwinAds);
        for (int ihdpa = 0; ihdpa < chdpa; ihdpa++)
        {
            *ppda = (CDarwinAd*)DPA_FastGetPtr(g_hdpaDarwinAds, ihdpa);
            if (*ppda)
            {
                if (ILIsEqual((*ppda)->_pidl, pidlFull))
                {
                    iRet = ihdpa;
                    break;
                }
            }
        }
    }
    return iRet;
}

BOOL CStartMenuCallbackBase::_IsDarwinAdvertisement(LPCITEMIDLIST pidlFull)
{
    // What this is doing is comparing the passed in pidl with the
    // list of pidls in g_hdpaDarwinAds. That hdpa contains a list of
    // pidls that are darwin ads.

    // If the background thread is not done, then we must assume that
    // it has not processed the shortcut that we are on. That is why we process it
    // in line.


    ENTERCRITICAL_DARWINADS;

    // NOTE: There can be two items in the hdpa. This is ok.
    BOOL fAd = FALSE;
    CDarwinAd* pda = NULL;
    int iIndex = GetDarwinIndex(pidlFull, &pda);
    // Are there any ads?
    if (iIndex != -1 && pda != NULL)
    {
        //This is a Darwin pidl. Is it installed?
        fAd = pda->IsAd();
    }

    LEAVECRITICAL_DARWINADS;

    return fAd;
}

HRESULT SHParseDarwinIDFromCacheW(LPWSTR pszDarwinDescriptor, LPWSTR* ppwszOut)
{
    HRESULT hr = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);

    if (g_hdpaDarwinAds)
    {
        ENTERCRITICAL_DARWINADS;
        int chdpa = DPA_GetPtrCount(g_hdpaDarwinAds);
        for (int ihdpa = 0; ihdpa < chdpa; ihdpa++)
        {
            CDarwinAd* pda = (CDarwinAd*)DPA_FastGetPtr(g_hdpaDarwinAds, ihdpa);
            if (pda && pda->_pszLocalPath && pda->_pszDescriptor &&
                StrCmpCW(pszDarwinDescriptor, pda->_pszDescriptor) == 0)
            {
                hr = SHStrDupW(pda->_pszLocalPath, ppwszOut);
                break;
            }
        }
        LEAVECRITICAL_DARWINADS;
    }

    return hr;
}


// REARCHITECT (lamadio): There is a duplicate of this helper in browseui\browmenu.cpp
//                   When modifying this, rev that one as well.
void UEMRenamePidl(const GUID* pguidGrp1, IShellFolder* psf1, LPCITEMIDLIST pidl1,
    const GUID* pguidGrp2, IShellFolder* psf2, LPCITEMIDLIST pidl2)
{
    UAINFO uei;
    uei.cbSize = sizeof(uei);
    uei.dwMask = UEIM_HIT | UEIM_FILETIME;
    if (SUCCEEDED(UEMQueryEvent(pguidGrp1,
        UEME_RUNPIDL, (WPARAM)psf1,
        (LPARAM)pidl1, &uei)) &&
        uei.cLaunches > 0)
    {
        UEMSetEvent(pguidGrp2,
            UEME_RUNPIDL, (WPARAM)psf2, (LPARAM)pidl2, &uei);

        uei.cLaunches = 0;
        UEMSetEvent(pguidGrp1,
            UEME_RUNPIDL, (WPARAM)psf1, (LPARAM)pidl1, &uei);
    }
}

// REARCHITECT (lamadio): There is a duplicate of this helper in browseui\browmenu.cpp
//                   When modifying this, rev that one as well.
void UEMDeletePidl(const GUID* pguidGrp, IShellFolder* psf, LPCITEMIDLIST pidl)
{
    UAINFO uei;
    uei.cbSize = sizeof(uei);
    uei.dwMask = UEIM_HIT;
    uei.cLaunches = 0;
    UEMSetEvent(pguidGrp, UEME_RUNPIDL, (WPARAM)psf, (LPARAM)pidl, &uei);
}

//
//  Sortof safe version of ILIsParent which catches when pidlParent or
//  pidlBelow is NULL.  pidlParent can be NULL on systems that don't
//  have a Common Program Files folder.  pidlBelow should never be NULL
//  but it doesn't hurt to check.
//
STDAPI_(BOOL) SMILIsAncestor(LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidlBelow)
{
    if (pidlParent && pidlBelow)
        return ILIsParent(pidlParent, pidlBelow, FALSE);
    else
        return FALSE;
}

HRESULT CStartMenuCallbackBase::_ProcessChangeNotify(SMDATA* psmd, LONG lEvent,
    LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
{
    switch (lEvent)
    {
        case SHCNE_ASSOCCHANGED:
            SHReValidateDarwinCache();
            return S_OK;

        case SHCNE_RENAMEFOLDER:
            // NTRAID89654-2000/03/13 (lamadio): We should move the MenuOrder stream as well. 5.5.99
        case SHCNE_RENAMEITEM:
        {
            LPITEMIDLIST pidlPrograms;
            LPITEMIDLIST pidlProgramsCommon;
            LPITEMIDLIST pidlFavorites;
            SHGetFolderLocation(NULL, CSIDL_PROGRAMS, NULL, 0, &pidlPrograms);
            SHGetFolderLocation(NULL, CSIDL_COMMON_PROGRAMS, NULL, 0, &pidlProgramsCommon);
            SHGetFolderLocation(NULL, CSIDL_FAVORITES, NULL, 0, &pidlFavorites);

            BOOL fPidl1InStartMenu = SMILIsAncestor(pidlPrograms, pidl1) ||
                SMILIsAncestor(pidlProgramsCommon, pidl1);
            BOOL fPidl1InFavorites = SMILIsAncestor(pidlFavorites, pidl1);


            // If we're renaming something from the Start Menu
            if (fPidl1InStartMenu || fPidl1InFavorites)
            {
                IShellFolder* psfFrom;
                LPCITEMIDLIST pidlFrom;
                if (SUCCEEDED(SHBindToParent(pidl1, IID_PPV_ARG(IShellFolder, &psfFrom), &pidlFrom)))
                {
                    // Into the Start Menu
                    BOOL fPidl2InStartMenu = SMILIsAncestor(pidlPrograms, pidl2) ||
                        SMILIsAncestor(pidlProgramsCommon, pidl2);
                    BOOL fPidl2InFavorites = SMILIsAncestor(pidlFavorites, pidl2);
                    if (fPidl2InStartMenu || fPidl2InFavorites)
                    {
                        IShellFolder* psfTo;
                        LPCITEMIDLIST pidlTo;

                        if (SUCCEEDED(SHBindToParent(pidl2, IID_PPV_ARG(IShellFolder, &psfTo), &pidlTo)))
                        {
                            // Then we need to rename it
                            UEMRenamePidl(fPidl1InStartMenu ? &UEMIID_SHELL : &UEMIID_BROWSER,
                                psfFrom, pidlFrom,
                                fPidl2InStartMenu ? &UEMIID_SHELL : &UEMIID_BROWSER,
                                psfTo, pidlTo);
                            psfTo->Release();
                        }
                    }
                    else
                    {
                        // Otherwise, we delete it.
                        UEMDeletePidl(fPidl1InStartMenu ? &UEMIID_SHELL : &UEMIID_BROWSER,
                            psfFrom, pidlFrom);
                    }

                    psfFrom->Release();
                }
            }

            ILFree(pidlPrograms);
            ILFree(pidlProgramsCommon);
            ILFree(pidlFavorites);
        }
        break;

        case SHCNE_DELETE:
            // NTRAID89654-2000/03/13 (lamadio): We should nuke the MenuOrder stream as well. 5.5.99
        case SHCNE_RMDIR:
        {
            IShellFolder* psf;
            LPCITEMIDLIST pidl;

            if (SUCCEEDED(SHBindToParent(pidl1, IID_PPV_ARG(IShellFolder, &psf), &pidl)))
            {
                // NOTE favorites is the only that will be initialized
                UEMDeletePidl(psmd->uIdAncestor == IDM_FAVORITES ? &UEMIID_BROWSER : &UEMIID_SHELL,
                    psf, pidl);
                psf->Release();
            }

        }
        break;

        case SHCNE_CREATE:
        case SHCNE_MKDIR:
        {
            IShellFolder* psf;
            LPCITEMIDLIST pidl;

            if (SUCCEEDED(SHBindToParent(pidl1, IID_PPV_ARG(IShellFolder, &psf), &pidl)))
            {
                UAINFO uei;
                uei.cbSize = sizeof(uei);
                uei.dwMask = UEIM_HIT;
                uei.cLaunches = UEM_NEWITEMCOUNT;
                UEMSetEvent(psmd->uIdAncestor == IDM_FAVORITES ? &UEMIID_BROWSER : &UEMIID_SHELL,
                    UEME_RUNPIDL, (WPARAM)psf, (LPARAM)pidl, &uei);
                psf->Release();
            }

        }
        break;
    }

    return S_FALSE;
}

HRESULT CStartMenuCallbackBase::_GetSFInfo(SMDATA* psmd, SMINFO* psminfo)
{
    if (psminfo->dwMask & SMIM_FLAGS &&
        (psmd->uIdAncestor == IDM_PROGRAMS ||
            psmd->uIdAncestor == IDM_FAVORITES))
    {
        if (_fExpandoMenus)
        {
            psminfo->dwFlags |= _GetDemote(psmd);
        }

        // This is a little backwards. If the Restriction is On, Then we allow the feature.
        if (SHRestricted(REST_GREYMSIADS) &&
            psmd->uIdAncestor == IDM_PROGRAMS)
        {
            LPITEMIDLIST pidlFull = FullPidlFromSMData(psmd);
            if (pidlFull)
            {
                if (_IsDarwinAdvertisement(pidlFull))
                {
                    psminfo->dwFlags |= SMIF_ALTSTATE;
                }
                ILFree(pidlFull);
            }
        }

        if (_ptp2)
        {
            _ptp2->ModifySMInfo(psmd, psminfo);
        }
    }
    return S_OK;
}

void SHReValidateDarwinCache()
{
    if (g_hdpaDarwinAds)
    {
        ENTERCRITICAL_DARWINADS;
        int chdpa = DPA_GetPtrCount(g_hdpaDarwinAds);
        for (int ihdpa = 0; ihdpa < chdpa; ihdpa++)
        {
            CDarwinAd* pda = (CDarwinAd*)DPA_FastGetPtr(g_hdpaDarwinAds, ihdpa);
            if (pda)
            {
                pda->CheckInstalled();
            }
        }
        LEAVECRITICAL_DARWINADS;
    }
}

// Determines if a CSIDL is a child of another CSIDL
// e.g.
//  CSIDL_STARTMENU = c:\foo\bar\Start Menu
//  CSIDL_PROGRAMS  = c:\foo\bar\Start Menu\Programs
//  Return true
BOOL IsCSIDLChild(int csidlParent, int csidlChild)
{
    BOOL fChild = FALSE;
    TCHAR sz1[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(NULL, csidlParent, NULL, 0, sz1)))
    {
        TCHAR sz2[MAX_PATH];
        if (SUCCEEDED(SHGetFolderPath(NULL, csidlChild, NULL, 0, sz2)))
        {
            TCHAR szCommonRoot[MAX_PATH];
            if (PathCommonPrefix(sz1, sz2, szCommonRoot) ==
                lstrlen(sz1))
            {
                fChild = TRUE;
            }
        }
    }

    return fChild;
}

//
// Now StartMenuChange value is stored seperately for classic startmenu and new startmenu.
// So, we need to check which one is currently on before we read the value.
//
BOOL IsStartMenuChangeNotAllowed(BOOL fStartPanel)
{
    return(IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOCHANGESTARMENU,
        TEXT("Advanced"),
        (fStartPanel ? TEXT("Start_EnableDragDrop") : TEXT("StartMenuChange")),
        ROUS_DEFAULTALLOW | ROUS_KEYALLOWS));
}

#define SMINIT_DEFAULT              0x00000000  // No Options
#define SMINIT_RESTRICT_CONTEXTMENU 0x00000001  // Don't allow Context Menus
#define SMINIT_RESTRICT_DRAGDROP    0x00000002  // Don't allow Drag and Drop
#define SMINIT_TOPLEVEL             0x00000004  // This is the top band.
#define SMINIT_DEFAULTTOTRACKPOPUP  0x00000008  // When no callback is specified, 
#define SMINIT_CACHED               0x00000010
#define SMINIT_USEMESSAGEFILTER     0x00000020
#define SMINIT_LEGACYMENU           0x00000040  // Old Menu behaviour.
#define SMINIT_CUSTOMDRAW           0x00000080   // Send SMC_CUSTOMDRAW
#define SMINIT_NOSETSITE            0x00010000  // Internal setting
#define SMINIT_VERTICAL             0x10000000  // This is a vertical menu
#define SMINIT_HORIZONTAL           0x20000000  // This is a horizontal menu    (does not inherit)
#define SMINIT_MULTICOLUMN          0x40000000  // this is a multi column menu

#define SMSET_SEPARATEMERGEFOLDER   0x00000200    //Insert separator when MergedFolder host changes

#define SMSET_DONTREGISTERCHANGENOTIFY 0x00000020 // ShellFolder is a discontiguous child of a parent shell folder

HRESULT CStartMenuCallback::InitializeProgramsShellMenu(IShellMenu* psm)
{
    HKEY hkeyPrograms = nullptr;
    ITEMIDLIST* pidl = nullptr;

    DWORD dwInitFlags = 0x10000000;
    if (!FeatureEnabled(L"StartMenuScrollPrograms") && !_fIsStartPanel)
        dwInitFlags = 0x50000000;
    if (IsStartMenuChangeNotAllowed(_fIsStartPanel))
        dwInitFlags |= 0x3;
    if (_fIsStartPanel)
        dwInitFlags |= 0x20004;

    HRESULT hr = psm->Initialize(this, 504, 504, dwInitFlags);
    if (SUCCEEDED(hr))
    {
        _InitializePrograms();
        const WCHAR* pszOrderKey =
            _fIsStartPanel
            ? L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Start Menu2\\Programs"
            : L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Start Menu\\Programs";
        RegCreateKeyExW(
            HKEY_CURRENT_USER, pszOrderKey, 0, nullptr, REG_OPTION_NON_VOLATILE, (KEY_READ | KEY_WRITE),
            nullptr, &hkeyPrograms, nullptr);

        IShellFolder* psf;
        DWORD dwSmset = SMSET_TOP;
        if (_fIsStartPanel)
        {
            dwSmset |= 0x200;
            hr = BindToGetFolderAndPidl(CLSID_ProgramsFolderAndFastItems, &psf, &pidl);
        }
        else
        {
            hr = BindToGetFolderAndPidl(CLSID_ProgramsFolder, &psf, &pidl);
        }
        if (SUCCEEDED(hr))
        {
            ASSERT(pidl); // 2189
            ASSERT(psf); // 2192
            hr = psm->SetShellFolder(psf, pidl, hkeyPrograms, dwSmset);
            psf->Release();
            ILFree(pidl);
        }

        if (FAILED(hr))
        {
            RegCloseKey(hkeyPrograms);
        }
    }

    return hr;
}

HRESULT GetFolderAndPidl(UINT csidl, IShellFolder** ppsf, LPITEMIDLIST* ppidl)
{
    *ppsf = NULL;
    HRESULT hr = SHGetFolderLocation(NULL, csidl, NULL, 0, ppidl);
    if (SUCCEEDED(hr))
    {
        hr = SHBindToObject(NULL, IID_X_PPV_ARG(IShellFolder, *ppidl, ppsf));
        if (FAILED(hr))
        {
            ILFree(*ppidl);
            *ppidl = NULL;
        }
    }
    return hr;
}

#define MENU_STARTMENU_OPENFOLDER       402

// Creates the "Start Menu\\<csidl>" section of the start menu by
// looking up the csidl, generating the Hkey from HKCU\pszRoot\pszValue,
//  Initializing the IShellMenu with dwPassInitFlags, then setting the locations 
// into the passed IShellMenu passing the flags dwSetFlags.
HRESULT CStartMenuCallback::InitializeCSIDLShellMenu(int uId, int csidl, LPCTSTR pszRoot, LPCTSTR pszValue,
    DWORD dwPassInitFlags, DWORD dwSetFlags, BOOL fAddOpen,
    IShellMenu* psm)
{
    DWORD dwInitFlags = SMINIT_VERTICAL | dwPassInitFlags;

    if (IsStartMenuChangeNotAllowed(_fIsStartPanel))
        dwInitFlags |= SMINIT_RESTRICT_DRAGDROP | SMINIT_RESTRICT_CONTEXTMENU;

    psm->Initialize(this, uId, uId, dwInitFlags);

    LPITEMIDLIST pidl;
    IShellFolder* psfFolder;
    HRESULT hr = GetFolderAndPidl(csidl, &psfFolder, &pidl);
    if (SUCCEEDED(hr))
    {
        HKEY hKey = NULL;

        if (pszRoot)
        {
            TCHAR szPath[MAX_PATH];
            StrCpyN(szPath, pszRoot, ARRAYSIZE(szPath));
            if (pszValue)
            {
                PathAppend(szPath, pszValue);
            }

            RegCreateKeyEx(HKEY_CURRENT_USER, szPath, NULL, NULL,
                REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE,
                NULL, &hKey, NULL);
        }

        // Point the menu to the shellfolder
        hr = psm->SetShellFolder(psfFolder, pidl, hKey, dwSetFlags);
        if (SUCCEEDED(hr))
        {
            if (fAddOpen && _fAddOpenFolder)
            {
                HMENU hMenu = SHLoadMenuPopup(LoadLibraryW(L"shell32.dll"), MENU_STARTMENU_OPENFOLDER);
                if (hMenu)
                {
                    psm->SetMenu(hMenu, _hwnd, SMSET_BOTTOM);
                }
            }
        }
        else
            RegCloseKey(hKey);

        psfFolder->Release();
        ILFree(pidl);
    }

    return hr;
}

// This generates the Recent | My Documents, My Pictures sub menu.
HRESULT CStartMenuCallback::InitializeDocumentsShellMenu(IShellMenu* psm)
{
    HRESULT hr = InitializeCSIDLShellMenu(IDM_RECENT, CSIDL_RECENT, NULL, NULL,
        SMINIT_RESTRICT_DRAGDROP, SMSET_BOTTOM, FALSE,
        psm);

    // Initializing, reset cache bits for top part of menu
    _fHasMyDocuments = FALSE;
    _fHasMyPictures = FALSE;

    return hr;
}

HRESULT CStartMenuCallback::InitializeFastItemsShellMenu(IShellMenu* psm)
{
    DWORD dwFlags = 0x4 | 0x20000 | 0x10000000;
    if (IsStartMenuChangeNotAllowed(_fIsStartPanel))
        dwFlags |= 0x1 | 0x2;

    HRESULT hr = psm->Initialize(this, 0, -1, dwFlags);
    if (SUCCEEDED(hr))
    {
        _InitializePrograms();

        IShellFolder* psf;
        ITEMIDLIST_ABSOLUTE* pidl;
        hr = BindToGetFolderAndPidl(CLSID_StartMenuFolder, &psf, &pidl);
        if (SUCCEEDED(hr))
        {
            HKEY hMenuKey = nullptr;
            RegCreateKeyExW(
                HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Start Menu",
                0, nullptr, REG_OPTION_NON_VOLATILE, (KEY_READ | KEY_WRITE), nullptr, &hMenuKey, nullptr);

            //CcshellDebugMsgW(0x2000000, "Root Start Menu Key Is %d", hMenuKey);
            hr = psm->SetShellFolder(psf, pidl, hMenuKey, 0x10000004);
            psf->Release();
            ILFree(pidl);
        }
    }

    return hr;
}

HRESULT CStartMenuCallback::_InitializeFindMenu(IShellMenu* psm)
{
    HRESULT hr = E_FAIL;

    psm->Initialize(this, IDM_MENU_FIND, IDM_MENU_FIND, SMINIT_VERTICAL);

    HMENU hmenu = CreatePopupMenu();
    if (hmenu)
    {
        ATOMICRELEASE(_pcmFind);

        if (_ptp)
        {
            if (SUCCEEDED(_ptp->GetFindCM(hmenu, TRAY_IDM_FINDFIRST, TRAY_IDM_FINDLAST, &_pcmFind)))
            {
                IContextMenu2* pcm2;
                _pcmFind->QueryInterface(IID_PPV_ARG(IContextMenu2, &pcm2));
                if (pcm2)
                {
                    pcm2->HandleMenuMsg(WM_INITMENUPOPUP, (WPARAM)hmenu, 0);
                    pcm2->Release();
                }
            }

            if (_pcmFind)
            {
                hr = psm->SetMenu(hmenu, NULL, SMSET_TOP);
                // Don't Release _pcmFind
            }
        }

        // Since we failed to create the ShellMenu
        // we need to dispose of this HMENU
        if (FAILED(hr))
            DestroyMenu(hmenu);
    }

    return hr;
}

#define SMSET_USEBKICONEXTRACTION   0x00000008   // Use the background icon extractor
#define SMSET_HASEXPANDABLEFOLDERS  0x00000010   // Need to call SHIsExpandableFolder

HRESULT CStartMenuCallback::InitializeSubShellMenu(int idCmd, IShellMenu* psm)
{
    HRESULT hr = E_FAIL;

    switch (idCmd)
    {
        case IDM_PROGRAMS:
            hr = InitializeProgramsShellMenu(psm);
            break;

        case IDM_RECENT:
            hr = InitializeDocumentsShellMenu(psm);
            break;

        case IDM_MENU_FIND:
            hr = _InitializeFindMenu(psm);
            break;

        case IDM_FAVORITES:
            hr = InitializeCSIDLShellMenu(
                IDM_FAVORITES, CSIDL_FAVORITES, STRREG_FAVORITES, nullptr, 0,
                SMSET_HASEXPANDABLEFOLDERS | SMSET_USEBKICONEXTRACTION, FALSE, psm);
            break;

        case IDM_CONTROLS:
            hr = InitializeCSIDLShellMenu(
                IDM_CONTROLS, CSIDL_CONTROLS, STRREG_STARTMENU, L"ControlPanel", 0, 0, TRUE, psm);
            break;

        case IDM_PRINTERS:
            hr = InitializeCSIDLShellMenu(
                IDM_PRINTERS, CSIDL_PRINTERS, STRREG_STARTMENU, L"Printers", 0, 0, TRUE, psm);
            break;

        case IDM_MYDOCUMENTS:
            hr = InitializeCSIDLShellMenu(
                IDM_MYDOCUMENTS, CSIDL_PERSONAL, STRREG_STARTMENU, L"MyDocuments", 0, 0, TRUE, psm);
            break;

        case IDM_MYPICTURES:
            hr = InitializeCSIDLShellMenu(
                IDM_MYPICTURES, CSIDL_MYPICTURES, STRREG_STARTMENU, L"MyPictures", 0, 0, TRUE, psm);
            break;

        case IDM_NETCONNECT:
            hr = InitializeCSIDLShellMenu(
                IDM_NETCONNECT, CSIDL_CONNECTIONS, STRREG_STARTMENU, L"NetConnections", 0, 0, TRUE, psm);
            break;
    }

    return hr;
}

HRESULT CStartMenuCallback::_GetObject(LPSMDATA psmd, REFIID riid, void** ppvOut)
{
    HRESULT hr = E_FAIL;
    UINT    uId = psmd->uId;

    ASSERT(ppvOut);
    ASSERT(IS_VALID_READ_PTR(psmd, SMDATA));

    *ppvOut = NULL;

    if (IsEqualGUID(riid, IID_IShellMenu))
    {
        IShellMenu* psm = NULL;
        hr = CoCreateInstanceHook(CLSID_MenuBand, NULL, CLSCTX_INPROC, IID_PPV_ARG(IShellMenu, &psm));
        if (SUCCEEDED(hr))
        {
            hr = InitializeSubShellMenu(uId, psm);

            if (FAILED(hr))
            {
                psm->Release();
                psm = NULL;
            }
        }

        *ppvOut = psm;
    }
    else if (IsEqualGUID(riid, IID_IContextMenu))
    {
        //
        //  NOTE - we dont allow users to open the recent folder this way - ZekeL - 1-JUN-99
        //  because this is really an internal folder and not a user folder.
        //

        switch (uId)
        {
            case IDM_PROGRAMS:
            case IDM_FAVORITES:
            case IDM_MYDOCUMENTS:
            case IDM_MYPICTURES:
            case IDM_CONTROLS:
            case IDM_PRINTERS:
            case IDM_NETCONNECT:
            {
                CStartContextMenu* pcm = new CStartContextMenu(uId);
                if (pcm)
                {
                    hr = pcm->QueryInterface(riid, ppvOut);
                    pcm->Release();
                }
                else
                    hr = E_OUTOFMEMORY;
            }
        }
    }
    return hr;
}

//
//  Return S_OK to remove the pidl from enumeration
//
HRESULT CStartMenuCallbackBase::_FilterPidl(UINT uParent, IShellFolder* psf, const ITEMID_CHILD* pidl)
{
    ASSERT(IS_VALID_PIDL(pidl)); // 2495

    HRESULT hr = S_FALSE;

    WCHAR szChild[260];
    if (SUCCEEDED(DisplayNameOfW(psf, pidl, SHGDN_FORPARSING, szChild, ARRAYSIZE(szChild))))
    {
        if (_pgpf && _pgpf->IsFileRestrictedByPolicyOrSystemMetrics(szChild) == S_OK)
        {
            return S_OK;
        }
        if (uParent == IDM_PROGRAMS || uParent == IDM_TOPLEVELSTARTMENU)
        {
            const WCHAR* pszChild = PathFindFileNameW(szChild);

            if (_pszPrograms && lstrcmpiW(pszChild, _pszPrograms) == 0)
                return S_OK;
            if (_pszCommonPrograms && lstrcmpiW(pszChild, _pszCommonPrograms) == 0)
                return S_OK;
            if (!_fShowAdminTools && _pszAdminTools && lstrcmpiW(pszChild, _pszAdminTools) == 0)
                return S_OK;
        }
    }

    return hr;
}

BOOL LinkGetInnerPidl(IShellFolder* psf, LPCITEMIDLIST pidl, LPITEMIDLIST* ppidlOut, DWORD* pdwAttr)
{
    *ppidlOut = NULL;

    IShellLink* psl;
    HRESULT hr = psf->GetUIObjectOf(NULL, 1, &pidl, __uuidof(**(&psl)), 0, IID_PPV_ARGS_Helper(&psl));
    if (SUCCEEDED(hr))
    {
        psl->GetIDList(ppidlOut);

        if (*ppidlOut)
        {
            if (FAILED(SHGetAttributesOf(*ppidlOut, pdwAttr)))
            {
                ILFree(*ppidlOut);
                *ppidlOut = NULL;
            }
        }
        psl->Release();
    }
    return (*ppidlOut != NULL);
}


//
//  _FilterRecentPidl() 
//  the Recent Documents folder can now (NT5) have more than 15 or 
//  so documents, but we only want to show the 15 most recent that we always have on
//  the start menu.  this means that we need to filter out all folders and 
//  anything more than MAXRECENTDOCS
//
HRESULT CStartMenuCallback::_FilterRecentPidl(IShellFolder* psf, PCUITEMID_CHILD pidl)
{
    HRESULT hr = S_OK;

    ASSERT(IS_VALID_PIDL(pidl));
    ASSERT(IS_VALID_CODE_PTR(psf, IShellFolder));
    ASSERT(_cRecentDocs != -1);

    ASSERT(_cRecentDocs <= MAXRECENTDOCS);

    //  if we already reached our limit, dont go over...
    if (_pmruRecent && (_cRecentDocs < MAXRECENTDOCS))
    {
        //  we now must take a looksee for it...
        int iItem;
        DWORD dwAttr = SFGAO_FOLDER | SFGAO_BROWSABLE;
        LPITEMIDLIST pidlTrue;

        //  need to find out if the link points to a folder...
        //  because we dont want
        if (SUCCEEDED(_pmruRecent->FindData((BYTE*)pidl, ILGetSize(pidl), &iItem))
            && LinkGetInnerPidl(psf, pidl, &pidlTrue, &dwAttr))
        {
            if (!(dwAttr & SFGAO_FOLDER))
            {
                //  we have a link to something that isnt a folder 
                hr = S_FALSE;
                _cRecentDocs++;
            }

            ILFree(pidlTrue);
        }
    }

    //  return S_OK if you dont want to show this item...

    return hr;
}


HRESULT CStartMenuCallback::_HandleAccelerator(TCHAR ch, SMDATA* psmdata)
{
    // Since we renamed the 'Find' menu to 'Search' the PMs wanted to have
    // an upgrade path for users (So they can continue to use the old accelerator
    // on the new menu item.)
    // To enable this, when toolbar detects that there is not an item in the menu
    // that contains the key that has been pressed, then it sends a TBN_ACCL.
    // This is intercepted by mnbase, and translated into SMC_ACCEL. 
    if (CharUpper((LPTSTR)ch) == CharUpper((LPTSTR)_szFindMnemonic[0]))
    {
        psmdata->uId = IDM_MENU_FIND;
        return S_OK;
    }

    return S_FALSE;
}

#define IDS_CHEVRONTIPTITLE     0x768F
#define IDS_CHEVRONTIP          0x7690

HRESULT CStartMenuCallback::_GetTip(LPWSTR pstrTitle, LPWSTR pstrTip)
{
    if (pstrTitle == NULL ||
        pstrTip == NULL)
    {
        return S_FALSE;
    }

    LoadString(LoadLibraryW(L"shell32.dll"), IDS_CHEVRONTIPTITLE, pstrTitle, MAX_PATH);
    LoadString(LoadLibraryW(L"shell32.dll"), IDS_CHEVRONTIP, pstrTip, MAX_PATH);

    // Why would this fail?
    ASSERT(pstrTitle[0] != L'\0' && pstrTip[0] != L'\0');
    return S_OK;
}

EXTERN_C HRESULT CStartMenu_CreateInstance(LPUNKNOWN punkOuter, REFIID riid, void** ppvOut)
{
    HRESULT hr = E_FAIL;
    IMenuPopup* pmp = NULL;

    *ppvOut = NULL;

    CStartMenuCallback* psmc = new CStartMenuCallback();
    if (psmc)
    {
        IShellMenu* psm;

        hr = CoCreateInstanceHook(CLSID_MenuBand, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARG(IShellMenu, &psm));
        if (SUCCEEDED(hr))
        {
            hr = CoCreateInstanceHook(CLSID_MenuDeskBar, punkOuter, CLSCTX_INPROC_SERVER, IID_PPV_ARG(IMenuPopup, &pmp));
            if (SUCCEEDED(hr))
            {
                IBandSite* pbs;
                hr = CoCreateInstanceHook(CLSID_MenuBandSite, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARG(IBandSite, &pbs));
                if (SUCCEEDED(hr))
                {
                    hr = pmp->SetClient(pbs);
                    if (SUCCEEDED(hr))
                    {
                        IDeskBand* pdb;
                        hr = psm->QueryInterface(IID_PPV_ARG(IDeskBand, &pdb));
                        if (SUCCEEDED(hr))
                        {
                            hr = pbs->AddBand(pdb);
                            pdb->Release();
                        }
                    }
                    pbs->Release();
                }
                // Don't free pmp. We're using it below.
            }

            if (SUCCEEDED(hr))
            {
                // This is so the ref counting happens correctly.
                hr = psm->Initialize(psmc, 0, 0, SMINIT_VERTICAL | SMINIT_TOPLEVEL);
                if (SUCCEEDED(hr))
                {
                    // if this fails, we don't get that part of the menu
                    // this is okay since it can happen if the start menu is redirected
                    // to where we dont have access.
                    psmc->InitializeFastItemsShellMenu(psm);
                }
            }

            psm->Release();
        }
        psmc->Release();
    }

    if (SUCCEEDED(hr))
    {
        hr = pmp->QueryInterface(riid, ppvOut);
    }
    else
    {
        // We need to do this so that it does a cascading delete
        IUnknown_SetSite(pmp, NULL);
    }

    if (pmp)
        pmp->Release();

    return hr;
}

// IUnknown
STDMETHODIMP CStartContextMenu::QueryInterface(REFIID riid, void** ppvObj)
{

    static const QITAB qit[] =
    {
        QITABENT(CStartContextMenu, IContextMenu),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObj);
}

STDMETHODIMP_(ULONG) CStartContextMenu::AddRef(void)
{
    return ++_cRef;
}

STDMETHODIMP_(ULONG) CStartContextMenu::Release(void)
{
    ASSERT(_cRef > 0);
    _cRef--;

    if (_cRef > 0)
        return _cRef;

    delete this;
    return 0;
}

#define SMCM_STARTMENU_FIRST        0x5000
#define SMCM_OPEN                   (SMCM_STARTMENU_FIRST + 0)
#define SMCM_EXPLORE                (SMCM_STARTMENU_FIRST + 1)
#define SMCM_OPEN_ALLUSERS          (SMCM_STARTMENU_FIRST + 2)
#define SMCM_EXPLORE_ALLUSERS       (SMCM_STARTMENU_FIRST + 3)
#define MENU_STARTMENUSTATICITEMS   359

// IContextMenu
STDMETHODIMP CStartContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    HRESULT hr = E_FAIL;
    HMENU hmenuStartMenu = SHLoadMenuPopup(LoadLibraryW(L"shell32.dll"), MENU_STARTMENUSTATICITEMS);

    if (hmenuStartMenu)
    {
        TCHAR szCommon[MAX_PATH];
        BOOL fAddCommon = (S_OK == SHGetFolderPath(NULL, CSIDL_COMMON_STARTMENU, NULL, 0, szCommon));

        if (fAddCommon)
            fAddCommon = IsUserAnAdmin();

        // Since we don't show this on the start button when the user is not an admin, don't show it here... I guess...
        if (_idCmd != IDM_PROGRAMS || !fAddCommon)
        {
            DeleteMenu(hmenuStartMenu, SMCM_OPEN_ALLUSERS, MF_BYCOMMAND);
            DeleteMenu(hmenuStartMenu, SMCM_EXPLORE_ALLUSERS, MF_BYCOMMAND);
        }

        if (Shell_MergeMenus(hmenu, hmenuStartMenu, 0, indexMenu, idCmdLast, uFlags))
        {
            SetMenuDefaultItem(hmenu, 0, MF_BYPOSITION);
            _SHPrettyMenu(hmenu);
            hr = ResultFromShort(GetMenuItemCount(hmenuStartMenu));
        }

        DestroyMenu(hmenuStartMenu);
    }

    return hr;
}

STDMETHODIMP CStartContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO lpici)
{
    HRESULT hr = E_FAIL;
    if (HIWORD64(lpici->lpVerb) == 0)
    {
        BOOL fAllUsers = FALSE;
        BOOL fOpen = TRUE;
        switch (LOWORD(lpici->lpVerb))
        {
            case SMCM_OPEN_ALLUSERS:
                fAllUsers = TRUE;
            case SMCM_OPEN:
                // fOpen = TRUE;
                break;

            case SMCM_EXPLORE_ALLUSERS:
                fAllUsers = TRUE;
            case SMCM_EXPLORE:
                fOpen = FALSE;
                break;

            default:
                return S_FALSE;
        }

        hr = ExecStaticStartMenuItem(_idCmd, fAllUsers, fOpen);
    }

    // Ahhh Don't handle verbs!!!
    return hr;

}

STDMETHODIMP CStartContextMenu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pRes, LPSTR pszName, UINT cchMax)
{
    return E_NOTIMPL;
}

MIDL_INTERFACE("f1763f2a-6e44-426d-ac5b-641c866dcd63")
IStartMenuMSIAds : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE IsMSIAds(ITEMIDLIST_ABSOLUTE*) = 0;
};

//****************************************************************************
//
//  CPersonalStartMenuCallback

class CPersonalProgramsMenuCallback
    : public CStartMenuCallbackBase
    , public IShellItemFilter
    , public IStartMenuMSIAds
{
public:
    CPersonalProgramsMenuCallback()
        : CStartMenuCallbackBase(TRUE)
    {
    }

    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override { return CStartMenuCallbackBase::AddRef(); }
    STDMETHODIMP_(ULONG) Release() override { return CStartMenuCallbackBase::Release(); }
    //~ End IUnknown Interface

    //~ Begin IShellItemFilter Interface
    STDMETHODIMP IncludeItem(IShellItem* psi) override;
    STDMETHODIMP GetEnumFlagsForItem(IShellItem* psi, SHCONTF* pgrfFlags) override;
    //~ End IShellItemFilter Interface

    //~ Begin IStartMenuMSIAds Interface
    STDMETHODIMP IsMSIAds(ITEMIDLIST_ABSOLUTE* pidl) override;
    //~ End IStartMenuMSIAds Interface
};

HRESULT CPersonalProgramsMenuCallback::QueryInterface(REFIID riid, void** ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CPersonalProgramsMenuCallback, IShellItemFilter),
        QITABENT(CPersonalProgramsMenuCallback, IObjectWithSite),
        QITABENT(CPersonalProgramsMenuCallback, IStartMenuMSIAds),
        {}
    };
    return QISearch(this, qit, riid, ppvObj);
}

DEFINE_PROPERTYKEY(PKEY_AppUserModel_HostEnvironment, 0x9F4C2855, 0x9F79, 0x4B39, 0xA8, 0xD0, 0xE1, 0xD4, 0x2D, 0xE1, 0xD5, 0xF3, 14);

bool IsImmersiveShortcut(IShellItem* psi)
{
    bool fImmersive = false;

    CComPtr<IShellItem2> spsi2;
    HRESULT hr = psi->QueryInterface(IID_PPV_ARGS(&spsi2));
    if (SUCCEEDED(hr))
    {
        ULONG uHostEnv;
        hr = spsi2->GetUInt32(PKEY_AppUserModel_HostEnvironment, &uHostEnv);
        if (SUCCEEDED(hr))
        {
            fImmersive = (uHostEnv == 1 || uHostEnv == 2);
        }
    }

    return fImmersive;
}

HRESULT CPersonalProgramsMenuCallback::IncludeItem(IShellItem* psi)
{
    VARIANT vt = {};
    vt.vt = VT_BOOL;

    _InitializePrograms();
    IUnknown_QueryServiceExec(_punkSite, SID_SM_NSCHOST, &SID_SM_DV2ControlHost, 318, 0, nullptr, &vt);

    IParentAndItem* ppai;
    HRESULT hr = psi->QueryInterface(IID_PPV_ARGS(&ppai));
    if (SUCCEEDED(hr))
    {
        IShellFolder* psf;
        ITEMIDLIST* pidl;
        hr = ppai->GetParentAndItem(nullptr, &psf, &pidl);
        if (SUCCEEDED(hr))
        {
            if (_FilterPidl(vt.boolVal == VARIANT_TRUE ? 0 : -1, psf, pidl) == S_OK || IsImmersiveShortcut(psi))
            {
                hr = S_FALSE;
            }
            psf->Release();
            ILFree(pidl);
        }
        ppai->Release();
    }

    return hr;
}

HRESULT CPersonalProgramsMenuCallback::GetEnumFlagsForItem(IShellItem* psi, SHCONTF* pgrfFlags)
{
    return E_NOTIMPL;
}

HRESULT CPersonalProgramsMenuCallback::IsMSIAds(ITEMIDLIST_ABSOLUTE* pidl)
{
    return _IsDarwinAdvertisement(pidl) ? S_OK : S_FALSE;
}

EXTERN_C HRESULT CPersonalStartMenu_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppvOut)
{
    *ppvOut = nullptr;

    CPersonalProgramsMenuCallback* psmc = new CPersonalProgramsMenuCallback();
    if (!psmc)
        return E_OUTOFMEMORY;
    HRESULT hr = psmc->QueryInterface(riid, ppvOut);
    psmc->Release();
    return hr;
}
