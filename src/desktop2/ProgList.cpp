#include "pch.h"
#include "cocreateinstancehook.h"
#include "shundoc.h"
#include "stdafx.h"
#include "proglist.h"
#include "uemapp.h"
#include "shdguid.h"
#include "shguidp.h"        // IID_IInitializeObject
#include "tray.h"
#include "path.h"

#define PRINT_UEMINFO(ueminfo) \
    printf( \
        "[%s] UEMINFO: cbSize=%u, dwMask=0x%08x, R=%.2f, cLaunches=%u, cSwitches=%u, dwTime=%u, ftExecute={%u,%u}, fExcludeFromMFU=%d\n", \
        __FUNCTION__, \
        (ueminfo).cbSize, \
        (ueminfo).dwMask, \
        (ueminfo).R, \
        (ueminfo).cLaunches, \
        (ueminfo).cSwitches, \
        (ueminfo).dwTime, \
        (ueminfo).ftExecute.dwLowDateTime, \
        (ueminfo).ftExecute.dwHighDateTime, \
        (ueminfo).fExcludeFromMFU \
    )

typedef UNALIGNED const WCHAR* LPNCWSTR;
typedef UNALIGNED WCHAR* LPNWSTR;

// Global cache item being passed from the background task to the start panel ByUsage.
CMenuItemsCache *g_pMenuCache;


// From startmnu.cpp
HRESULT Tray_RegisterHotKey(WORD wHotKey, LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidl);

#define TF_PROGLIST 0x00000020

#define CH_DARWINMARKER TEXT('\1')  // Illegal filename character

#define IsDarwinPath(psz) ((psz)[0] == CH_DARWINMARKER)

// Largest date representable in FILETIME - rolls over in the year 30828
static const FILETIME c_ftNever = { 0xFFFFFFFF, 0x7FFFFFFF };

void GetStartTime(FILETIME *pft)
{
    //
    //  If policy says "Don't offer new apps", then set the New App
    //  threshold to some impossible time in the future.
    //
    if (!SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, REGSTR_VAL_DV2_NOTIFYNEW, FALSE, TRUE))
    {
        *pft = c_ftNever;
        return;
    }

    FILETIME ftNow;
    GetSystemTimeAsFileTime(&ftNow);

    DWORD dwSize = sizeof(*pft);
    //Check to see if we have the StartMenu start Time for this user saved in the registry.
    if(SHRegGetUSValue(DV2_REGPATH, DV2_SYSTEM_START_TIME, NULL,
                       pft, &dwSize, FALSE, NULL, 0) != ERROR_SUCCESS)
    {
        // Get the current system time as the start time. If any app is launched after
        // this time, that will result in the OOBE message going away!
        *pft = ftNow;

        dwSize = sizeof(*pft);

        //Save this time as the StartMenu start time for this user.
        SHRegSetUSValue(DV2_REGPATH, DV2_SYSTEM_START_TIME, REG_BINARY,
                        pft, dwSize, SHREGSET_FORCE_HKCU);
    }

    //
    //  Thanks to roaming and reinstalls, the user may have installed a new
    //  OS since they first turned on the Start Menu, so bump forwards to
    //  the time the OS was installed (plus some fudge to account for the
    //  time it takes to run Setup) so all the Accessories don't get marked
    //  as "New".
    //
    DWORD dwTime;
    dwSize = sizeof(dwTime);
    if (SHGetValue(HKEY_LOCAL_MACHINE, TEXT("Software\\Microsoft\\Windows NT\\CurrentVersion"),
                   TEXT("InstallDate"), NULL, &dwTime, &dwSize) == ERROR_SUCCESS)
    {
        // Sigh, InstallDate is stored in UNIX time format, not FILETIME,
        // so convert it to FILETIME.  Q167296 shows how to do this.
        LONGLONG ll = Int32x32To64(dwTime, 10000000) + 116444736000000000;

        // Add some fudge to account for how long it takes to run Setup
        ll += FT_ONEHOUR * 5;       // five hours should be plenty

        FILETIME ft;
        SetFILETIMEfromInt64(&ft, ll);

        //
        //  But don't jump forwards into the future
        //
        if (::CompareFileTime(&ft, &ftNow) > 0)
        {
            ft = ftNow;
        }

        if (::CompareFileTime(pft, &ft) < 0)
        {
            *pft = ft;
        }

    }

    //
    //  If this is a roaming profile, then don't count anything that
    //  happened before the user logged on because we don't want to mark
    //  apps as new just because it roamed in with the user at logon.
    //  We actually key off of the start time of Explorer, because
    //  Explorer starts after the profiles are sync'd so we don't need
    //  a fudge factor.
    //
    DWORD dwType;
    if (GetProfileType(&dwType) && (dwType & (PT_TEMPORARY | PT_ROAMING)))
    {
        FILETIME ft, ftIgnore;
        if (GetProcessTimes(GetCurrentProcess(), &ft, &ftIgnore, &ftIgnore, &ftIgnore))
        {
            if (::CompareFileTime(pft, &ft) < 0)
            {
                *pft = ft;
            }
        }
    }

}

//****************************************************************************
//
//  How the pieces fit together...
//
//
//  Each ByUsageRoot consists of a ByUsageShortcutList.
//
//  A ByUsageShortcutList is a list of shortcuts.
//
//  Each shortcut references a ByUsageAppInfo.  Multiple shortcuts can
//  reference the same ByUsageAppInfo if they are all shortcuts to the same
//  app.
//
//  The ByUsageAppInfo::_cRef is the number of CByUsageShortcut's that
//  reference it.
//
//  A master list of all ByUsageAppInfo's is kept in _dpaAppInfo.
//

//****************************************************************************

//
//  Helper template for destroying DPA's.
//
//  DPADELETEANDDESTROY(dpa);
//
//  invokes the "delete" method on each pointer in the DPA,
//  then destroys the DPA.
//

template<class T>
BOOL CALLBACK _DeleteCB(T *self, LPVOID)
{
    delete self;
    return TRUE;
}

template<class T>
void DPADELETEANDDESTROY(CDPA<T, CTContainer_PolicyUnOwned<T>> &dpa)
{
    if (dpa)
    {
        dpa.DestroyCallback(_DeleteCB<T>, NULL);
        ASSERT(dpa == NULL);
    }
}

template <class T>
int DPA_DeleteCB(T *self, LPVOID)
{
    if (self)
    {
        delete self;
    }
    return 1;   // continue
}

//****************************************************************************

void CByUsageRoot::Reset()
{
    _sl.DestroyCallback(DPA_DeleteCB);
	_slOld.DestroyCallback(DPA_DeleteCB);

    ILFree(_pidl);
    _pidl = NULL;
};

//****************************************************************************

class CByUsageDir
{
    IShellFolder *_psf;         // the folder interface
    LPITEMIDLIST _pidl;         // the absolute pidl for this folder
    LONG         _cRef;         // Reference count

    CByUsageDir() : _cRef(1)
    {
    }

    ~CByUsageDir()
    {
        ATOMICRELEASE(_psf);
        ILFree(_pidl);
    }

public:
    // Creates a CByUsageDir for CSIDL_DESKTOP.  All other
    // CByUsageDir's come from this.
    static CByUsageDir *CreateDesktop()
    {
        CByUsageDir *self = new CByUsageDir();
        if (self)
        {
            ASSERT(self->_pidl == NULL);
            if (SUCCEEDED(SHGetDesktopFolder(&self->_psf)))
            {
                // all is well
                return self;
            }
            
            delete self;
        }

        return NULL;
    }

    // pdir = parent folder
    // pidl = new folder location relative to parent
    static CByUsageDir *Create(CByUsageDir *pdir, LPCITEMIDLIST pidl)
    {
        CByUsageDir *self = new CByUsageDir();
        if (self)
        {
            LPCITEMIDLIST pidlRoot = pdir->_pidl;
            self->_pidl = ILCombine(pidlRoot, pidl);
            if (self->_pidl)
            {
                IShellFolder *psfRoot = pdir->_psf;
                if (SUCCEEDED(SHBindToObject(psfRoot,
                              IID_X_PPV_ARG(IShellFolder, pidl, &self->_psf))))
                {
                    // all is well
                    return self;
                }
            }

            delete self;
        }
        
        return NULL;
    }

    void AddRef();
    void Release();
    IShellFolder *Folder() const { return _psf; }
    LPCITEMIDLIST Pidl() const { return _pidl; }
};

void CByUsageDir::AddRef()
{
    InterlockedIncrement(&_cRef);
}

void CByUsageDir::Release()
{
    ASSERT( 0 != _cRef );
    if (InterlockedDecrement(&_cRef) == 0)
    {
        delete this;
    }
}


//****************************************************************************

class CByUsageItem : public PaneItem
{
public:
    LPITEMIDLIST _pidl;     // relative pidl
    CByUsageDir *_pdir;      // Parent directory
    UEMINFO _uei;           /* Usage info (for sorting) */
    LPITEMIDLIST _pidl1;    // Vista - NEW

    static CByUsageItem *Create(CByUsageShortcut *pscut);
    static CByUsageItem *CreateSeparator();

    CByUsageItem() { EnableDropTarget(); }
    ~CByUsageItem() override;

    CByUsageDir *Dir() const { return _pdir; }
    IShellFolder *ParentFolder() const { return _pdir->Folder(); }
    LPCITEMIDLIST RelativePidl() const { return _pidl; }
    LPITEMIDLIST CreateFullPidl() const { return ILCombine(_pdir->Pidl(), _pidl); }

    void SetRelativePidl(LPITEMIDLIST pidlNew)
    {
        ILFree(_pidl);
        _pidl = pidlNew;

        ILFree(_pidl1);
        _pidl1 = ILCombine(_pdir->Pidl(), _pidl);
    }

    virtual BOOL IsEqual(PaneItem *pItem) const 
    {
        CByUsageItem *pbuItem = reinterpret_cast<CByUsageItem *>(pItem);
        BOOL fIsEqual = FALSE;
        if (_pdir == pbuItem->_pdir)
        {
            // Do we have identical pidls?
            // Note: this test needs to be fast, and does not need to be exact, so we don't use pidl binding here.
            UINT usize1 = ILGetSize(_pidl);
            UINT usize2 = ILGetSize(pbuItem->_pidl);
            if (usize1 == usize2)
            {
                fIsEqual = (memcmp(_pidl, pbuItem->_pidl, usize1) == 0);
            }
        }

        return fIsEqual; 
    }
};

inline BOOL ByUsage::_IsPinned(CByUsageItem *pitem)
{
    return pitem->Dir() == _pdirDesktop;
}

//****************************************************************************

// Each app referenced by a command line is saved in one of these
// structures.

class CByUsageAppInfo // "papp"
{          
public:
    CByUsageShortcut *_pscutBest;// best candidate so far
    CByUsageShortcut *_pscutBestSM;// best candidate so far on Start Menu (excludes Desktop)
    UEMINFO _ueiBest;           // info about best candidate so far
    UEMINFO _ueiTotal;          // cumulative information
    LPTSTR  _pszAppPath;        // path to app in question
    FILETIME _ftCreated;        // when was file created?
    bool    _fNew;              // is app new?
    bool    _fPinned;           // is app referenced by a pinned item?
    bool    _fIgnoreTimestamp;  // Darwin apps have unreliable timestamps
private:
    LONG    _cRef;              // reference count

public:

    // WARNING!  Initialize() is called multiple times, so make sure
    // that you don't leak anything
    BOOL Initialize(LPCTSTR pszAppPath, CMenuItemsCache *pmenuCache, BOOL fCheckNew, bool fIgnoreTimestamp)
    {
        ASSERT(IsBlank());

        _fIgnoreTimestamp = fIgnoreTimestamp || IsDarwinPath(pszAppPath);

        // Note! Str_SetPtr is last so there's nothing to free on failure

        if (_fIgnoreTimestamp)
        {
            // just save the path; ignore the timestamp
            if (Str_SetPtrW(&_pszAppPath, pszAppPath))
            {
                _fNew = TRUE;
                return TRUE;
            }
        }
        else
        if (Str_SetPtrW(&_pszAppPath, pszAppPath))
        {
            if (fCheckNew && GetAppCreationTime(pszAppPath, &_ftCreated))
            {
                _fNew = pmenuCache->IsNewlyCreated(&_ftCreated);
            }
            return TRUE;
        }

        return FALSE;
    }

    static int GetAppCreationTime(LPCWSTR pszApp, FILETIME *pftCreate)
    {
        int result; // eax
        FILETIME FileTime2; // [esp+8h] [ebp-214h] BYREF
        WCHAR FileName[260]; // [esp+10h] [ebp-20Ch] BYREF

        if (PathIsNetworkPath(pszApp))
        {
            pftCreate->dwHighDateTime = 0;
            pftCreate->dwLowDateTime = 0;
        }
        else
        {
            result = GetFileCreationTime(pszApp, pftCreate);
            if (!result)
                return result;
            StringCchCopy(FileName, ARRAYSIZE(FileName), pszApp);
            if (PathRemoveFileSpec(FileName)
                && !PathIsRoot(FileName)
                && GetFileCreationTime(FileName, &FileTime2)
                && CompareFileTime(pftCreate, &FileTime2) < 0)
            {
                *pftCreate = FileTime2;
            }
        }
        return 1;
    }

    static CByUsageAppInfo *Create()
    {
        CByUsageAppInfo *papp = new CByUsageAppInfo;
        if (papp)
        {
            ASSERT(papp->IsBlank());        // will be AddRef()d by caller
            ASSERT(papp->_pszAppPath == NULL);
        }
        return papp;
    }

    ~CByUsageAppInfo()
    {
        ASSERT(IsBlank());
        Str_SetPtrW(&_pszAppPath, NULL);
    }

    // Notice! When the reference count goes to zero, we do not delete
    // the item.  That's because there is still a reference to it in
    // the _dpaAppInfo DPA.  Instead, we'll recycle items whose refcount
    // is zero.

    // NTRAID:135696 potential race condition against background enumeration
    void AddRef() { InterlockedIncrement(&_cRef); }
    void Release() { ASSERT( 0 != _cRef ); InterlockedDecrement(&_cRef); }
    inline BOOL IsBlank() { return _cRef == 0; }
    inline BOOL IsNew() { return _fNew; }
    inline BOOL IsAppPath(LPCTSTR pszAppPath)
        { return lstrcmpi(_pszAppPath, pszAppPath) == 0; }
    const FILETIME& GetCreatedTime() const { return _ftCreated; }
    inline const FILETIME *GetFileTime() const { return &_ueiTotal.ftExecute; }

    inline LPTSTR GetAppPath() const { return _pszAppPath; }

    void GetUAInfo(OUT UEMINFO *puei)
    {
#ifdef DEBUG
        PRINT_UEMINFO(*puei);
#endif
		_GetUAInfo(&UAIID_APPLICATIONS, _pszAppPath, puei);
    }

    void CombineUAInfo(IN const UEMINFO *pueiNew, BOOL fNew = TRUE, BOOL fIsDesktop = FALSE, BOOL a5 = TRUE /*Guess*/)
    {
        if (a5)
        {
            _ueiTotal.R = pueiNew->R;
        }

        if (CompareFileTime(&pueiNew->ftExecute, &_ueiTotal.ftExecute) > 0)
        {
            _ueiTotal.ftExecute = pueiNew->ftExecute;
        }

        if (!fIsDesktop && !fNew)
            _fNew = FALSE;
    }

    //
    //  The app UEM info is "old" if the execution time is more
    //  than an hour after the install time.
    //
    inline BOOL _IsUAINFONew(const UEMINFO *puei)
    {
        return FILETIMEtoInt64(puei->ftExecute) <
               FILETIMEtoInt64(_ftCreated) + ByUsage::FT_NEWAPPGRACEPERIOD();
    }

    //
    //  Prepare for a new enumeration.
    //
    static BOOL CALLBACK EnumResetCB(CByUsageAppInfo *self, CMenuItemsCache *pmenuCache)
    {
        self->_pscutBest = NULL;
        self->_pscutBestSM = NULL;
        self->_fPinned = FALSE;
        ZeroMemory(&self->_ueiBest, sizeof(self->_ueiBest));
        ZeroMemory(&self->_ueiTotal, sizeof(self->_ueiTotal));
        if (self->_fNew && !self->_fIgnoreTimestamp)
        {
            self->_fNew = pmenuCache->IsNewlyCreated(&self->_ftCreated);
        }
        return TRUE;
    }

    static BOOL CALLBACK EnumGetFileCreationTime(CByUsageAppInfo *self, CMenuItemsCache *pmenuCache)
    {
        if (!self->IsBlank() &&
            !self->_fIgnoreTimestamp &&
            GetAppCreationTime(self->_pszAppPath, &self->_ftCreated))
        {
            self->_fNew = pmenuCache->IsNewlyCreated(&self->_ftCreated);
        }
        return TRUE;
    }

    CByUsageItem *CreateCByUsageItem();
};

//****************************************************************************

class CByUsageShortcut
{
#ifdef DEBUG
    CByUsageShortcut() { }           // make constructor private
#endif
    CByUsageDir *        _pdir;      // folder that contains this shortcut
    LPITEMIDLIST        _pidl;      // pidl relative to parent
    CByUsageAppInfo *    _papp;      // associated EXE
    FILETIME            _ftCreated; // time shortcut was created
    bool                _fNew;      // new app?
    bool                _fInteresting; // Is this a candidate for the MFU list?
    bool                _fDarwin;   // Is this a Darwin shortcut?

public: // @TEMP
	bool                field_17;    // Vista - NEW
public:

    // Accessors
    LPCITEMIDLIST RelativePidl() const { return _pidl; }
    CByUsageDir *Dir() const { return _pdir; }
    LPCITEMIDLIST ParentPidl() const { return _pdir->Pidl(); }
    IShellFolder *ParentFolder() const { return _pdir->Folder(); }
    CByUsageAppInfo *App() const { return _papp; }
    bool IsNew() const { return _fNew; }
    bool SetNew(bool fNew) { return _fNew = fNew; }
    const FILETIME& GetCreatedTime() const { return _ftCreated; }
    bool IsInteresting() const { return _fInteresting; }
    bool SetInteresting(bool fInteresting) { return _fInteresting = fInteresting; }
    bool IsDarwin() { return _fDarwin; }

    LPITEMIDLIST CreateFullPidl() const
        { return ILCombine(ParentPidl(), RelativePidl()); }

    LPCITEMIDLIST UpdateRelativePidl(CByUsageHiddenData *phd);

    void SetApp(CByUsageAppInfo *papp)
    {
        if (_papp) _papp->Release();
        _papp = papp;
        if (papp) papp->AddRef();
    }

    static CByUsageShortcut *Create(CByUsageDir *pdir, LPCITEMIDLIST pidl, CByUsageAppInfo *papp, bool fDarwin, BOOL fForce = FALSE)
    {
        //ASSERT(pdir);
        //ASSERT(pidl);

        CByUsageShortcut *pscut = new CByUsageShortcut;
        if (pscut)
        {

            pscut->_fNew = TRUE;        // will be set to FALSE later
            pscut->_pdir = pdir; pdir->AddRef();
            pscut->SetApp(papp);
            pscut->_fDarwin = fDarwin;
            pscut->_pidl = ILClone(pidl);
            if (pscut->_pidl)
            {
                LPTSTR pszShortcutName = _DisplayNameOf(pscut->ParentFolder(), pscut->RelativePidl(), SHGDN_FORPARSING);
                if (pszShortcutName &&
                    GetFileCreationTime(pszShortcutName, &pscut->_ftCreated))
                {
                    // Woo-hoo, all is well
                }
                else if (fForce)
                {
                    // The item was forced -- create it even though
                    // we don't know what it is.
                }
                else
                {
                    delete pscut;
                    pscut = NULL;
                }
                SHFree(pszShortcutName);
            }
            else
            {
                delete pscut;
                pscut = NULL;
            }
        }
        return pscut;
    }

    ~CByUsageShortcut()
    {
        if (_pdir) _pdir->Release();
        if (_papp) _papp->Release();
        ILFree(_pidl);          // ILFree ignores NULL
    }

    CByUsageItem *CreatePinnedItem(int iPinPos);

    HRESULT GetUAInfo(OUT UEMINFO *puei)
    {
        HRESULT hr = E_OUTOFMEMORY;

        LPITEMIDLIST pidlFull = ILCombine(_pdir->Pidl(), _pidl);
        if (pidlFull)
        {
            WCHAR szPath[MAX_PATH];
            hr = SHGetNameAndFlags(pidlFull, SHGDN_FORPARSING, szPath, ARRAYSIZE(szPath), 0);
            if (SUCCEEDED(hr))
            {
                _GetUAInfo(&UAIID_SHORTCUTS, szPath, puei);
#ifdef DEBUG
                PRINT_UEMINFO(*puei);
#endif
            }
            ILFree(pidlFull);
        }
        return hr;
    }
};

__inline HRESULT SHILClone(PCUIDLIST_RELATIVE pidl, PIDLIST_RELATIVE *ppidl)
{
    *ppidl = ILClone(pidl);
    return *ppidl ? S_OK : E_OUTOFMEMORY;
}

//****************************************************************************

CByUsageItem *CByUsageItem::Create(CByUsageShortcut *pscut)
{
#ifdef DEAD_CODE
    //ASSERT(pscut);
    CByUsageItem *pitem = new CByUsageItem;
    if (pitem)
    {
        pitem->_pdir = pscut->Dir();
        pitem->_pdir->AddRef();

        pitem->_pidl = ILClone(pscut->RelativePidl());
        if (pitem->_pidl)
        {
            return pitem;
        }
    }
    delete pitem;           // "delete" can handle NULL
    return NULL;
#else
    ASSERT(pscut); // 623
    CByUsageItem *pitem = new CByUsageItem;
    if (pitem)
    {
		LPITEMIDLIST v4;
        pitem->_pdir = pscut->Dir();
		pitem->_pdir->AddRef();
        if (SHILClone(pscut->RelativePidl(), &pitem->_pidl) < 0
            || (v4 = ILCombine(pitem->_pdir->Pidl(), pitem->_pidl), (pitem->_pidl1 = v4) == 0))
        {
            pitem->Release();
            return NULL;
        }
    }
    return pitem;
#endif
}

CByUsageItem *CByUsageItem::CreateSeparator()
{
    CByUsageItem *pitem = new CByUsageItem;
    if (pitem)
    {
        pitem->_iPinPos = PINPOS_SEPARATOR;
    }
    return pitem;
}

CByUsageItem::~CByUsageItem()
{
    ILFree(_pidl);
	ILFree(_pidl1);
    if (_pdir)
        _pdir->Release();
}

CByUsageItem *CByUsageAppInfo::CreateCByUsageItem()
{
    ASSERT(_pscutBest);
    CByUsageItem *pitem = CByUsageItem::Create(_pscutBest);
    if (pitem)
    {
        pitem->_uei = _ueiTotal;
    }
    return pitem;
}

CByUsageItem *CByUsageShortcut::CreatePinnedItem(int iPinPos)
{
    ASSERT(iPinPos >= 0);

    CByUsageItem *pitem = CByUsageItem::Create(this);
    if (pitem)
    {
        pitem->_iPinPos = iPinPos;
    }
    return pitem;
}

//****************************************************************************
//
//  The hidden data for the Programs list is of the following form:
//
//      WORD    cb;             // hidden item size
//      WORD    wPadding;       // padding
//      IDLHID  idl;            // IDLHID_STARTPANEDATA
//      int     iUnused;        // (was icon index)
//      WCHAR   LocalMSIPath[]; // variable-length string
//      WCHAR   TargetPath[];   // variable-length string
//      WCHAR   AltName[];      // variable-length string
//
//  AltName is the alternate display name for the EXE.
//
//  The hidden data is never accessed directly.  We always operate on the
//  ByUsageHiddenData structure, and there are special methods for
//  transferring data between this structure and the pidl.
//
//  The hidden data is appended to the pidl, with a "wOffset" backpointer
//  at the very end.
//
//  The TargetPath is stored as an unexpanded environment string.
//  (I.e., you have to ExpandEnvironmentStrings before using them.)
//
//  As an added bonus, the TargetPath might be a GUID (product code)
//  if it is really a Darwin shortcut.
//
class CByUsageHiddenData
{
public:
    WORD    _wHotKey;            // Hot Key
    LPWSTR  _pwszMSIPath;        // SHAlloc
    LPWSTR  _pwszTargetPath;     // SHAlloc
    LPWSTR  _pwszAltName;        // SHAlloc

public:

    void Init()
    {
        _wHotKey = 0;
        _pwszMSIPath = NULL;
        _pwszTargetPath = NULL;
        _pwszAltName = NULL;
    }

    BOOL IsClear()              // are all fields at initial values?
    {
        return _wHotKey == 0 &&
               _pwszMSIPath == NULL &&
               _pwszTargetPath == NULL &&
               _pwszAltName == NULL;
    }

    CByUsageHiddenData() { Init(); }

    enum {
        BUHD_HOTKEY       = 0x0000,             // so cheap we always fetch it
        BUHD_MSIPATH      = 0x0001,
        BUHD_TARGETPATH   = 0x0002,
        BUHD_ALTNAME      = 0x0004,
        BUHD_ALL          = -1,
    };

    BOOL Get(LPCITEMIDLIST pidl, UINT buhd);    // Load from pidl
    LPITEMIDLIST Set(LPITEMIDLIST pidl);        // Save to pidl

    void Clear()
    {
        SHFree(_pwszMSIPath);
        SHFree(_pwszTargetPath);
        SHFree(_pwszAltName);
        Init();
    }

    static LPTSTR GetAltName(LPCITEMIDLIST pidl);
    static LPITEMIDLIST SetAltName(LPITEMIDLIST pidl, LPCTSTR ptszNewName);
    void LoadFromShellLink(IShellLink *psl);
    HRESULT UpdateMSIPath();

private:
    static LPBYTE _ParseString(LPBYTE pbHidden, LPWSTR *ppwszOut, LPITEMIDLIST pidlMax, BOOL fSave);
    static LPBYTE _AppendString(LPBYTE pbHidden, LPWSTR pwsz);
};

//
//  We are parsing a string out of a pidl, so we have to be extra careful
//  to watch out for data corruption.
//
//  pbHidden = next byte to parse (or NULL to stop parsing)
//  ppwszOut receives parsed string if fSave = TRUE
//  pidlMax = start of next pidl; do not parse past this point
//  fSave = should we save the string in ppwszOut?
//
LPBYTE CByUsageHiddenData::_ParseString(LPBYTE pbHidden, LPWSTR *ppwszOut, LPITEMIDLIST pidlMax, BOOL fSave)
{
    if (!pbHidden)
        return NULL;

    LPNWSTR pwszSrc = (LPNWSTR)pbHidden;
    LPNWSTR pwsz = pwszSrc;
    LPNWSTR pwszLast = (LPNWSTR)pidlMax - 1;

    //
    //  We cannot use ualstrlenW because that might scan past pwszLast
    //  and fault.
    //
    while (pwsz < pwszLast && *pwsz)
    {
        pwsz++;
    }

    //  Corrupted data -- no null terminator found.
    if (pwsz >= pwszLast)
        return NULL;

    pwsz++;     // Step pwsz over the terminating NULL

    UINT cb = (UINT)((LPBYTE)pwsz - (LPBYTE)pwszSrc);
    if (fSave)
    {
        *ppwszOut = (LPWSTR)SHAlloc(cb);
        if (*ppwszOut)
        {
            CopyMemory(*ppwszOut, pbHidden, cb);
        }
        else
        {
            return NULL;
        }
    }
    pbHidden += cb;
    ASSERT(pbHidden == (LPBYTE)pwsz);
    return pbHidden;
}

BOOL CByUsageHiddenData::Get(LPCITEMIDLIST pidl, UINT buhd)
{
    ASSERT(IsClear());

    PCIDHIDDEN pidhid = ILFindHiddenID(pidl, IDLHID_STARTPANEDATA);
    if (!pidhid)
    {
        return FALSE;
    }

    // Do not access bytes after pidlMax
    LPITEMIDLIST pidlMax = _ILNext((LPITEMIDLIST)pidhid);

    LPBYTE pbHidden = ((LPBYTE)pidhid) + sizeof(HIDDENITEMID);

    // Skip the iUnused value
    // Note: if you someday choose to use it, you must read it as
    //      _iWhatever = *(UNALIGNED int *)pbHidden;
    pbHidden += sizeof(int);

    // HotKey
    _wHotKey = *(UNALIGNED WORD *)pbHidden;
    pbHidden += sizeof(_wHotKey);

    pbHidden = _ParseString(pbHidden, &_pwszMSIPath,    pidlMax, buhd & BUHD_MSIPATH);
    pbHidden = _ParseString(pbHidden, &_pwszTargetPath, pidlMax, buhd & BUHD_TARGETPATH);
    pbHidden = _ParseString(pbHidden, &_pwszAltName,    pidlMax, buhd & BUHD_ALTNAME);

    if (pbHidden)
    {
        return TRUE;
    }
    else
    {
        Clear();
        return FALSE;
    }
}


LPBYTE CByUsageHiddenData::_AppendString(LPBYTE pbHidden, LPWSTR pwsz)
{
    LPWSTR pwszOut = (LPWSTR)pbHidden;

    // The pointer had better already be aligned for WCHARs
    ASSERT(((ULONG_PTR)pwszOut & 1) == 0);

    if (pwsz)
    {
        lstrcpyW(pwszOut, pwsz);
    }
    else
    {
        pwszOut[0] = L'\0';
    }
    return (LPBYTE)(pwszOut + 1 + lstrlenW(pwszOut));
}

//
//  Note!  On failure, the source pidl is freed!
//  (This behavior is inherited from ILAppendHiddenID.)
//
LPITEMIDLIST CByUsageHiddenData::Set(LPITEMIDLIST pidl)
{
    UINT cb = sizeof(HIDDENITEMID);
    cb += sizeof(int);
    cb += sizeof(_wHotKey);
    cb += (UINT)(CbFromCchW(1 + (_pwszMSIPath ? lstrlenW(_pwszMSIPath) : 0)));
    cb += (UINT)(CbFromCchW(1 + (_pwszTargetPath ? lstrlenW(_pwszTargetPath) : 0)));
    cb += (UINT)(CbFromCchW(1 + (_pwszAltName ? lstrlenW(_pwszAltName) : 0)));

    // We can use the aligned version here since we allocated it ourselves
    // and didn't suck it out of a pidl.
    HIDDENITEMID *pidhid = (HIDDENITEMID*)alloca(cb);

    pidhid->cb = (WORD)cb;
    pidhid->wVersion = 0;
    pidhid->id = IDLHID_STARTPANEDATA;

    LPBYTE pbHidden = ((LPBYTE)pidhid) + sizeof(HIDDENITEMID);

    // The pointer had better already be aligned for ints
    ASSERT(((ULONG_PTR)pbHidden & 3) == 0);
    *(int *)pbHidden = 0;   // iUnused
    pbHidden += sizeof(int);

    *(DWORD *)pbHidden = _wHotKey;
    pbHidden += sizeof(_wHotKey);

    pbHidden = _AppendString(pbHidden, _pwszMSIPath);
    pbHidden = _AppendString(pbHidden, _pwszTargetPath);
    pbHidden = _AppendString(pbHidden, _pwszAltName);

    // Make sure our math was correct
    ASSERT(cb == (UINT)((LPBYTE)pbHidden - (LPBYTE)pidhid));

    // Remove and expunge the old data
    ILRemoveHiddenID(pidl, IDLHID_STARTPANEDATA);
    ILExpungeRemovedHiddenIDs(pidl);

    return ILAppendHiddenID(pidl, pidhid);
}

LPWSTR CByUsageHiddenData::GetAltName(LPCITEMIDLIST pidl)
{
    LPWSTR pszRet = NULL;
    CByUsageHiddenData hd;
    if (hd.Get(pidl, BUHD_ALTNAME))
    {
        pszRet = hd._pwszAltName;   // Keep this string
        hd._pwszAltName = NULL;     // Keep the upcoming assert happy
    }
    ASSERT(hd.IsClear());           // make sure we aren't leaking
    return pszRet;
}

//
//  Note!  On failure, the source pidl is freed!
//  (Propagating weird behavior of ByUsageHiddenData::Set)
//
LPITEMIDLIST CByUsageHiddenData::SetAltName(LPITEMIDLIST pidl, LPCTSTR ptszNewName)
{
    CByUsageHiddenData hd;

    // Attempt to overlay the existing values, but if they aren't available,
    // don't freak out.
    hd.Get(pidl, BUHD_ALL & ~BUHD_ALTNAME);

    ASSERT(hd._pwszAltName == NULL); // we excluded it from the hd.Get()
    hd._pwszAltName = const_cast<LPTSTR>(ptszNewName);

    pidl = hd.Set(pidl);
    hd._pwszAltName = NULL;     // so hd.Clear() won't SHFree() it
    hd.Clear();
    return pidl;

}

//
//  Returns S_OK if the item changed; S_FALSE if it remained the same
//
HRESULT CByUsageHiddenData::UpdateMSIPath()
{
    HRESULT hr = S_FALSE;

    if (_pwszTargetPath && IsDarwinPath(_pwszTargetPath))
    {
        LPWSTR pwszMSIPath = NULL;
        //
        //  If we can't resolve the Darwin ID to a filename, then leave
        //  the filename in the HiddenData alone - it's better than
        //  nothing.
        //
        if (SUCCEEDED(SHParseDarwinIDFromCacheW(_pwszTargetPath+1, &pwszMSIPath)) && pwszMSIPath)
        {
            //
            //  See if the MSI path has changed...
            //
            if (_pwszMSIPath == NULL ||
                StrCmpCW(pwszMSIPath, _pwszMSIPath) != 0)
            {
                hr = S_OK;
                SHFree(_pwszMSIPath);
                _pwszMSIPath = pwszMSIPath; // take ownership
            }
            else
            {
                // Unchanged; happy, free the path we aren't going to use
                SHFree(pwszMSIPath);
            }
        }
    }
    return hr;
}

LPCITEMIDLIST CByUsageShortcut::UpdateRelativePidl(CByUsageHiddenData *phd)
{
    return _pidl = phd->Set(_pidl);     // frees old _pidl even on failure
}

//
//  We must key off the Darwin ID and not the product code.
//
//  The Darwin ID is unique for each app in an application suite.
//  For example, PowerPoint and Outlook have different Darwin IDs.
//
//  The product code is the same for all apps in an application suite.
//  For example, PowerPoint and Outlook have the same product code.
//
//  Since we want to treat PowerPoint and Outlook as two independent
//  applications, we want to use the Darwin ID and not the product code.
//
HRESULT _GetDarwinID(IShellLinkDataList *pdl, DWORD dwSig, LPWSTR pszPath, UINT cchPath)
{
    LPEXP_DARWIN_LINK pedl;
    HRESULT hr;
    ASSERT(cchPath > 0);

    hr = pdl->CopyDataBlock(dwSig, (LPVOID*)&pedl);

    if (SUCCEEDED(hr))
    {
        pszPath[0] = CH_DARWINMARKER;
        hr = StringCchCopy(pszPath+1, cchPath - 1, pedl->szwDarwinID);
        LocalFree(pedl);
    }

    return hr;
}

HRESULT _GetPathOrDarwinID(IShellLink *psl, LPTSTR pszPath, UINT cchPath, DWORD dwFlags)
{
    HRESULT hr;

    ASSERT(cchPath);
    pszPath[0] = TEXT('\0');

    //
    //  See if it's a Darwin thingie.
    //
    IShellLinkDataList *pdl;
    hr = psl->QueryInterface(IID_PPV_ARGS(&pdl));
    if (SUCCEEDED(hr))
    {
        //
        //  Maybe this is a Darwin shortcut...  If so, then
        //  use the Darwin ID.
        //
        DWORD dwSLFlags;
        hr = pdl->GetFlags(&dwSLFlags);
        if (SUCCEEDED(hr))
        {
            if (dwSLFlags & SLDF_HAS_DARWINID)
            {
                hr = _GetDarwinID(pdl, EXP_DARWIN_ID_SIG, pszPath, cchPath);
            }
            else
            {
                hr = E_FAIL;            // No Darwin ID found
            }

            pdl->Release();
        }
    }

    if (FAILED(hr))
    {
        hr = psl->GetPath(pszPath, cchPath, 0, dwFlags);
    }

    return hr;
}

void CByUsageHiddenData::LoadFromShellLink(IShellLink *psl)
{
    ASSERT(_pwszTargetPath == NULL);

    HRESULT hr;
    TCHAR szPath[MAX_PATH];
    szPath[0] = TEXT('\0');

    hr = _GetPathOrDarwinID(psl, szPath, ARRAYSIZE(szPath), SLGP_RAWPATH);
    if (SUCCEEDED(hr))
    {
        SHStrDup(szPath, &_pwszTargetPath);
    }

    hr = psl->GetHotkey(&_wHotKey);
}

//****************************************************************************

ByUsageUI::ByUsageUI() : _byUsage(this),
    // We want to log execs as if they were launched by the Start Menu
    SFTBarHost(HOSTF_FIREUEMEVENTS |
               HOSTF_CANDELETE |
               HOSTF_CANRENAME)
{
    _iThemePart = SPP_PROGLIST;
    _iThemePartSep = SPP_PROGLISTSEPARATOR;
}

ByUsage::ByUsage(ByUsageUI *pByUsageUI)
{
    _pByUsageUI = pByUsageUI;

    GetStartTime(&_ftStartTime);

    _pidlBrowser = ILCreateFromPath(TEXT("shell:::{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}"));
    _pidlEmail   = ILCreateFromPath(TEXT("shell:::{2559a1f5-21d7-11d4-bdaf-00c04f60b9f0}"));
}

static ByUsageUI *g_pByUsageUI = nullptr;

SFTBarHost *ByUsage_CreateInstance()
{
    g_pByUsageUI = new ByUsageUI();
    return g_pByUsageUI;
}

template <class T>
int DPA_ILFreeCB(T self, LPVOID)
{
    CoTaskMemFree(self);
    return 1;
}

ByUsage::~ByUsage()
{
    if (_fUEMRegistered)
    {
        // Unregister with UEM DB if necessary
        UARegisterNotify(UANotifyCB, this, 0);
    }

	_dpaNew.DestroyCallback(DPA_ILFreeCB, nullptr);

    // Must clear the pinned items before releasing the MenuCache, 
    // as the pinned items point to AppInfo items in the cache.
    _rtPinned.Reset();

    if (_pMenuCache)
    {
        // Clean up the Menu cache properly.
        if (SUCCEEDED(_pMenuCache->LockPopup()))
        {
            _pMenuCache->UnregisterNotifyAll();
            _pMenuCache->AttachUI(nullptr);
            _pMenuCache->UnlockPopup();
        }
        _pMenuCache->Release();
    }

    ILFree(_pidlBrowser);
    ILFree(_pidlEmail);
    ATOMICRELEASE(_psmpin);

    if (_pdirDesktop)
    {
        _pdirDesktop->Release();
    }
}

// d:\longhorn\Shell\inc\idllib.h
PCUITEMID_CHILD _SHILMakeChild(const void *pv)
{
    PCUITEMID_CHILD pidl = reinterpret_cast<PCITEMID_CHILD>(pv);
    //RIP(ILIsChild(reinterpret_cast<PCUIDLIST_RELATIVE>(pidl))); // 178
    return pidl;
}

// {06C59536-1C66-4301-8387-82FBA3530E8D}
static const GUID TOID_STARTMENUCACHE =
{ 0x6c59536, 0x1c66, 0x4301, { 0x83, 0x87, 0x82, 0xfb, 0xa3, 0x53, 0xe, 0x8d } };


/*
 *  Background cache creation stuff...
 */
class CCreateMenuItemCacheTask : public CRunnableTask
{
    CMenuItemsCache *_pMenuCache;
    IShellTaskScheduler *_pScheduler;

public:
    CCreateMenuItemCacheTask(CMenuItemsCache* pMenuCache, IShellTaskScheduler* pScheduler)
        : CRunnableTask(RTF_DEFAULT)
        , _pMenuCache(pMenuCache)
        , _pScheduler(pScheduler)
    {
        _pMenuCache->AddRef();
        if (_pScheduler)
        {
            _pScheduler->AddRef();
        }
    }

    ~CCreateMenuItemCacheTask() override
    {
        IUnknown_SafeReleaseAndNullPtr(&_pMenuCache);
        IUnknown_SafeReleaseAndNullPtr(&_pScheduler);
    }

    static void DummyCallBack(LPVOID pvData, LPVOID pvHint, INT iIconIndex, INT iOpenIconIndex) {}

    STDMETHODIMP InternalResumeRT() override
    {
        _pMenuCache->DelayGetFileCreationTimes();
        _pMenuCache->DelayGetDarwinInfo();

        if (SUCCEEDED(_pMenuCache->LockPopup()))
        {
            if (_pMenuCache->InitDesktopFolder())
            {
                _pMenuCache->InitCache();
                _pMenuCache->UpdateCache();
                _pMenuCache->StartEnum();

                CByUsageHiddenData hd;          // construct once
                while (TRUE)
                {
                    CByUsageShortcut *pscut = _pMenuCache->GetNextShortcut();
                    if (!pscut)
                        break;

                    hd.Get(pscut->RelativePidl(), 2);
                    if (hd._wHotKey)
                    {
                        Tray_RegisterHotKey(hd._wHotKey, pscut->ParentPidl(), _SHILMakeChild(pscut->RelativePidl()));
                    }

                    // Pre-load the icons in the cache
                    int iIndex;
                    SHMapIDListToSystemImageListIndexAsync(_pScheduler, pscut->ParentFolder(),
                        _SHILMakeChild(pscut->RelativePidl()), DummyCallBack,
                        0, 0, &iIndex, 0);

                    // Register Darwin shortcut so that they can be grayed out if not installed
                    // and so we can map them to local paths as necessary
                    if (hd._pwszTargetPath && IsDarwinPath(hd._pwszTargetPath))
                    {
                        SHRegisterDarwinLink(pscut->CreateFullPidl(),
                            hd._pwszTargetPath + 1 /* Exclude the Darwin marker! */,
                            FALSE /* Don't update the Darwin state now, we'll do it later */);
                    }
                    hd.Clear();
                }
                IUnknown_SafeReleaseAndNullPtr(&_pMenuCache->_pdirDesktop);
            }
            _pMenuCache->UnlockPopup();
        }

        // Now determine all new items
        // Note: this is safe to do after the Unlock because we never remove anything from the _dpaAppInfo
        _pMenuCache->GetFileCreationTimes();

        _pMenuCache->AllowGetDarwinInfo();
        SHReValidateDarwinCache();

        // Refreshing Darwin shortcuts must be done under the popup lock
        // to avoid another thread doing a re-enumeration while we are
        // studying the dpa.  Keep SHReValidateDarwinCache outside of the
        // lock since it is slow.  (All we do is query the cache that
        // SHReValidateDarwinCache created for us.)
        if (SUCCEEDED(_pMenuCache->LockPopup()))
        {
            _pMenuCache->RefreshCachedDarwinShortcuts();
            _pMenuCache->UnlockPopup();
        }

        return S_OK;
    }
};

BOOL ByUsage::s_fDoneCreateMenuItemCachTask;

HRESULT ByUsage::Initialize()
{
#ifdef DEAD_CODE
    HRESULT hr;

    hr = CoCreateInstanceHook(CLSID_StartMenuPin, NULL, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&_psmpin));
    if (FAILED(hr))
    {
        return hr;
    }

    if (!(_pdirDesktop = CByUsageDir::CreateDesktop())) {
        return E_OUTOFMEMORY;
    }

    // Use already initialized MenuCache if available
    if (g_pMenuCache)
    {
        _pMenuCache = g_pMenuCache;
        _pMenuCache->AttachUI(_pByUsageUI);
        g_pMenuCache = NULL; // We take ownership here.
    }
    else
    {
        hr = CMenuItemsCache::ReCreateMenuItemsCache(_pByUsageUI, &_ftStartTime, &_pMenuCache);
        if (FAILED(hr))
        {
            return hr;
        }
    }


    _ulPinChange = -1;              // Force first query to re-enumerate

    _dpaNew = NULL;

    if (_pByUsageUI)
    {
        _hwnd = _pByUsageUI->_hwnd;

        //
        //  Register for the "pin list change" event.  This is an extended
        //  event (hence global), so listen in a location that contains
        //  no objects so the system doesn't waste time sending
        //  us stuff we don't care about.  Our choice: _pidlBrowser.
        //  It's not even a folder, so it can't contain any objects!
        //
        ASSERT(!_pMenuCache->IsLocked());
        _pByUsageUI->RegisterNotify(NOTIFY_PINCHANGE, SHCNE_EXTENDED_EVENT, _pidlBrowser, FALSE);
    }

    return S_OK;
#else
    CMenuItemsCache *v3; // eax
    CMenuItemsCache **p_pMenuCache; // edi

    HRESULT hr = CoCreateInstanceHook(CLSID_StartMenuPin, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&this->_psmpin));
    if (hr >= 0)
    {
        _pdirDesktop = CByUsageDir::CreateDesktop();
        if (_pdirDesktop)
        {
            //v3 = (CMenuItemsCache *)InterlockedExchange(&g_pMenuCache, 0);
            v3 = (CMenuItemsCache *)InterlockedExchangePointer((void *volatile *)&g_pMenuCache, 0);
            p_pMenuCache = &this->_pMenuCache;
            if (v3)
            {
                *p_pMenuCache = v3;
                v3->AttachUI(this->_pByUsageUI);
            }
            else
            {
                hr = CMenuItemsCache::ReCreateMenuItemsCache(_pByUsageUI, &_ftStartTime, &_pMenuCache);
            }

            if (hr >= 0)
            {
                this->_ulPinChange = -1;
                this->_dpaNew = nullptr;
                if (_pByUsageUI)
                {
                    this->_hwnd = _pByUsageUI->_hwnd;

                    ASSERT(!_pMenuCache->IsLocked()); // 1390
                    _pByUsageUI->RegisterNotify(NOTIFY_PINCHANGE, SHCNE_EXTENDED_EVENT, _pidlBrowser, 0);
                }
            }

            if (!s_fDoneCreateMenuItemCachTask)
            {
                IShellTaskScheduler *psched;
                if (SUCCEEDED(CoCreateInstance(CLSID_SharedTaskScheduler, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&psched))))
                {
                    CCreateMenuItemCacheTask *pTask = new CCreateMenuItemCacheTask(_pMenuCache, psched);
                    if (pTask)
                    {
                        hr = psched->AddTask(pTask, TOID_STARTMENUCACHE, 0, 0x10000000);
                        pTask->Release();
                        s_fDoneCreateMenuItemCachTask = 1;
                    }
                    psched->Release();
                }
            }
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }
    }
    return hr;
#endif
}

void CMenuItemsCache::_InitStringList(HKEY hk, LPCTSTR pszSub, CDPA<TCHAR, CTContainer_PolicyUnOwned<TCHAR>> *pdpa)
{
    ASSERT(static_cast<HDPA>(pdpa));

    LONG lRc;
    DWORD cb = 0;
    lRc = RegQueryValueEx(hk, pszSub, NULL, NULL, NULL, &cb);
    if (lRc == ERROR_SUCCESS)
    {
        // Add an extra TCHAR just to be extra-safe.  That way, we don't
        // barf if there is a non-null-terminated string in the registry.
        cb += sizeof(TCHAR);
        LPTSTR pszKillList = (LPTSTR)LocalAlloc(LPTR, cb);
        if (pszKillList)
        {
            lRc = SHGetValue(hk, NULL, pszSub, NULL, pszKillList, &cb);
            if (lRc == ERROR_SUCCESS)
            {
                // A semicolon-separated list of application names.
                LPTSTR psz = pszKillList;
                LPTSTR pszSemi;

                while ((pszSemi = StrChr(psz, TEXT(';'))) != NULL)
                {
                    *pszSemi = TEXT('\0');
                    if (*psz)
                    {
                        AppendString(pdpa, psz);
                    }
                    psz = pszSemi + 1;
                }
                if (*psz)
                {
                    AppendString(pdpa, psz);
                }
            }
            LocalFree(pszKillList);
        }
    }
}

//
//  Fill the kill list with the programs that should be ignored
//  should they be encountered in the Start Menu or elsewhere.
//
#define REGSTR_PATH_FILEASSOCIATION TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileAssociation")

void CMenuItemsCache::_InitKillList()
{
    HKEY hk;
    LONG lRc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, REGSTR_PATH_FILEASSOCIATION, 0,
                            KEY_READ, &hk);
    if (lRc == ERROR_SUCCESS)
    {
        _InitStringList(hk, TEXT("AddRemoveApps"), &_dpaKill);
        _InitStringList(hk, TEXT("AddRemoveNames"), &_dpaKillLink);
        RegCloseKey(hk);
    }
}

STDAPI_(UINT) ILGetSizeAndDepth(LPCITEMIDLIST pidl, DWORD *pdwDepth)
{
    DWORD dwDepth = 0;
    UINT cbTotal = 0;
    if (pidl)
    {
        //VALIDATE_PIDL(pidl);
        cbTotal += sizeof(pidl->mkid.cb);       // Null terminator
        while (pidl->mkid.cb)
        {
            cbTotal += pidl->mkid.cb;
            pidl = _ILNext(pidl);
            dwDepth++;
        }
    }

    if (pdwDepth)
        *pdwDepth = dwDepth;

    return cbTotal;
}

HRESULT __stdcall SHCompareIDs(
    IShellFolder *psf,
    DWORD a2,
    LPCITEMIDLIST pidl1,
    LPCITEMIDLIST pidl2,
    HRESULT *phr)
{
    HRESULT hr; // ebx MAPDST
    unsigned int v7; // esi
    int v8; // eax
    DWORD v12; // [esp+10h] [ebp-20h] BYREF
    DWORD v13; // [esp+14h] [ebp-1Ch] BYREF
    //CPPEH_RECORD ms_exc; // [esp+18h] [ebp-18h]
    IShellFolder *psf2; // [esp+38h] [ebp+8h] SPLIT BYREF

    //if (!IsValidPIDL(pidl1))
    //{
    //    CcshellDebugMsgA(2, 0, "invalid PIDL pointer - %#08lx", (char)pidl1);
    //    if (CcshellAssertFailedW(L"d:\\longhorn\\shell\\lib\\idllib.cpp", 568, L"IS_VALID_PIDL(pidl1)", 0))
    //    {
    //        AttachUserModeDebugger();
    //        do
    //        {
    //            __debugbreak();
    //            ms_exc.registration.TryLevel = -2;
    //        } while (dword_108B63C);
    //    }
    //}

    //if (!IsValidPIDL(pidl2))
    //{
    //    CcshellDebugMsgA(2, 0, "invalid PIDL pointer - %#08lx", (char)pidl2);
    //    if (CcshellAssertFailedW(L"d:\\longhorn\\shell\\lib\\idllib.cpp", 569, L"IS_VALID_PIDL(pidl2)", 0))
    //    {
    //        AttachUserModeDebugger();
    //        do
    //        {
    //            __debugbreak();
    //            ms_exc.registration.TryLevel = -2;
    //        } while (dword_108B638);
    //    }
    //}

    hr = 0x80004005;
    if (pidl1 == pidl2)
    {
        hr = 0;
    }
    else
    {
        v7 = ILGetSizeAndDepth(pidl1, &v13);
        v8 = ILGetSizeAndDepth(pidl2, &v12);
        if (v13 == v12 && v7 == v8 && !memcmp(pidl1, pidl2, v7))
            hr = 0;
        if (hr < 0)
        {
            if ((a2 & 0xFFFF0000) != 0)
            {
                if (psf->QueryInterface(IID_IShellFolder2, (void **)&psf2) < 0)
                    a2 = (unsigned __int16)a2;
                else
                    psf2->Release();
            }

            hr = psf->CompareIDs(a2, pidl1, pidl2);
            if (hr < 0)
            {
                //CcshellDebugMsgW(2, "SHCompareIDs failed, hr=0x%08X", hr);
                return hr;
            }
        }
    }
    if (phr)
        *phr = (__int16)hr;
    return hr;
}

//****************************************************************************
//
//  Filling the ByUsageShortcutList
//

int CALLBACK PidlSortCallback(LPITEMIDLIST pidl1, LPITEMIDLIST pidl2, IShellFolder *psf)
{
#ifdef DEAD_CODE
    HRESULT hr = psf->CompareIDs(0, pidl1, pidl2);

    // We got them from the ShellFolder; they should still be valid!
    ASSERT(SUCCEEDED(hr));

    return ShortFromResult(hr);
#else
    HRESULT hr;
    SHCompareIDs(psf, 0x10000000, pidl1, pidl2, &hr);
    return hr;
#endif
}

HRESULT AddMenuItemsCacheTask(IShellTaskScheduler* pSystemScheduler, BOOL fKeepCacheWhenFinished)
{
    HRESULT hr;

    CMenuItemsCache *pMenuCache = new CMenuItemsCache;

    FILETIME ftStart;
    // Initialize with something.
    GetStartTime(&ftStart);

    if (pMenuCache)
    {
        hr = pMenuCache->Initialize(nullptr, &ftStart);
        if (fKeepCacheWhenFinished)
        {
            g_pMenuCache = pMenuCache;
            g_pMenuCache->AddRef();
        }
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }

    if (SUCCEEDED(hr))
    {
        CCreateMenuItemCacheTask *pTask = new CCreateMenuItemCacheTask(pMenuCache, pSystemScheduler);

        if (pTask)
        {
            hr = pSystemScheduler->AddTask(pTask, TOID_STARTMENUCACHE, 0, ITSAT_DEFAULT_PRIORITY);
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }
    }

    return hr;
}

DWORD WINAPI CMenuItemsCache::ReInitCacheThreadProc(void *pv)
{
#ifdef DEAD_CODE
    HRESULT hr = SHCoInitialize();

    if (SUCCEEDED(hr))
    {
        CMenuItemsCache *pMenuCache = reinterpret_cast<CMenuItemsCache *>(pv);
        pMenuCache->DelayGetFileCreationTimes();
        pMenuCache->LockPopup();
        pMenuCache->InitCache();
        pMenuCache->UpdateCache();
        pMenuCache->UnlockPopup();

        // Now determine all new items
        // Note: this is safe to do after the Unlock because we never remove anything from the _dpaAppInfo
        pMenuCache->GetFileCreationTimes();
        pMenuCache->Release();
    }
    SHCoUninitialize(hr);

    return 0;
#else
    HRESULT hr = SHCoInitialize();

    if (SUCCEEDED(hr))
    {
        CMenuItemsCache* pMenuCache = static_cast<CMenuItemsCache*>(pv);
        pMenuCache->DelayGetFileCreationTimes();
        if (SUCCEEDED(pMenuCache->LockPopup()))
        {
            if (pMenuCache->InitDesktopFolder())
            {
                pMenuCache->InitCache();
                pMenuCache->UpdateCache();
                IUnknown_SafeReleaseAndNullPtr(&pMenuCache->_pdirDesktop);
            }
            pMenuCache->UnlockPopup();
        }

        pMenuCache->GetFileCreationTimes();
        pMenuCache->Release();
    }

    SHCoUninitialize(hr);
    return 0;
#endif
}

HRESULT CMenuItemsCache::ReCreateMenuItemsCache(ByUsageUI *pbuUI, FILETIME *ftOSInstall, CMenuItemsCache **ppMenuCache)
{
#ifdef DEAD_CODE
    HRESULT hr = E_OUTOFMEMORY;
    CMenuItemsCache *pMenuCache;

    // Create a CMenuItemsCache with ref count 1.
    pMenuCache = new CMenuItemsCache;
    if (pMenuCache)
    {
        hr = pMenuCache->Initialize(pbuUI, ftOSInstall);
    }

    if (SUCCEEDED(hr))
    {
        pMenuCache->AddRef();
        if (!SHQueueUserWorkItem(ReInitCacheThreadProc, pMenuCache, 0, 0, NULL, NULL, 0))
        {
            // No big deal if we fail here, we'll get another chance at enumerating later.
            pMenuCache->Release();
        }
        *ppMenuCache = pMenuCache;
    }
    return hr;
#else
    HRESULT hr = E_OUTOFMEMORY;

    CMenuItemsCache* pMenuCache = new CMenuItemsCache;
    if (pMenuCache)
    {
        hr = pMenuCache->Initialize(pbuUI, ftOSInstall);
        if (SUCCEEDED(hr))
        {
            pMenuCache->AddRef();
            if (!QueueUserWorkItem(ReInitCacheThreadProc, pMenuCache, 0))
            {
                pMenuCache->Release();
            }
            *ppMenuCache = pMenuCache;
        }
        else
        {
            pMenuCache->Release();
        }
    }
    return hr;
#endif
}


HRESULT CMenuItemsCache::GetFileCreationTimes()
{
    if (CompareFileTime(&_ftOldApps, &c_ftNever) != 0)
    {
        Lock();

        // Get all file creation times for our app list.
        _dpaAppInfo.EnumCallbackEx(CByUsageAppInfo::EnumGetFileCreationTime, this);

        Unlock();

        // From now on, we will be checkin newness when we create the app object.
        _fCheckNew = TRUE;
    }
    return S_OK;
}

//#define DEAD_CODE

//
//  Enumerate the contents of the folder specified by psfParent.
//  pidlParent represents the location of psfParent.
//
//  Note that we cannot do a depth-first walk into subfolders because
//  many (most?) machines have a timeout on FindFirst handles; if you don't
//  call FindNextFile for a few minutes, they secretly do a FindClose for
//  you on the assumption that you are a bad app that leaked a handle.
//  (There is also a valid DOS-compatibility reason for this behavior:
//  The DOS FindFirstFile API doesn't have a FindClose, so the server can
//  never tell if you are finished or not, so it has to guess that if you
//  don't do a FindNext for a long time, you're probably finished.)
//
//  So we have to save all the folders we find into a DPA and then walk
//  the folders after we close the enumeration.
//

void CMenuItemsCache::_FillFolderCache(CByUsageDir *pdir, CByUsageRoot *prt)
{
#ifdef DEAD_CODE
    // Caller should've initialized us
    ASSERT(prt->_sl);

    //
    // Note that we must use a namespace walk instead of FindFirst/FindNext,
    // because there might be folder shortcuts in the user's Start Menu.
    //

    // We do not specify SHCONTF_INCLUDEHIDDEN, so hidden objects are
    // automatically excluded
    IEnumIDList *peidl;
    if (S_OK == pdir->Folder()->EnumObjects(NULL, SHCONTF_FOLDERS |
                                              SHCONTF_NONFOLDERS, &peidl))
    {
        CDPAPidl dpaDirs;
        if (dpaDirs.Create(4))
        {
            CDPAPidl dpaFiles;
            if (dpaFiles.Create(4))
            {
                LPITEMIDLIST pidl;

                while (peidl->Next(1, &pidl, NULL) == S_OK)
                {
                    // _IsExcludedDirectory carse about SFGAO_FILESYSTEM and SFGAO_LINK
                    DWORD dwAttributes = SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_LINK;
                    if (SUCCEEDED(pdir->Folder()->GetAttributesOf(1, (LPCITEMIDLIST*)(&pidl),
                                                                  &dwAttributes)))
                    {
                        if (dwAttributes & SFGAO_FOLDER)
                        {
                            if (_IsExcludedDirectory(pdir->Folder(), pidl, dwAttributes) ||
                                dpaDirs.AppendPtr(pidl) < 0)
                            {
                                ILFree(pidl);
                            }
                        }
                        else
                        {
                            if (dpaFiles.AppendPtr(pidl) < 0)
                            {
                                ILFree(pidl);
                            }
                        }
                    }
                }

                dpaDirs.SortEx(PidlSortCallback, pdir->Folder());

                if (dpaFiles.GetPtrCount() > 0)
                {
                    dpaFiles.SortEx(PidlSortCallback, pdir->Folder());

                    //
                    //  Now merge the enumerated items with the ones
                    //  in the cache.
                    //
                    _MergeIntoFolderCache(prt, pdir, dpaFiles);
                }
                dpaFiles.DestroyCallback(DPA_ILFreeCB, NULL);
            }

            // Must release now to force the FindClose to happen
            peidl->Release();

            // Now go back and handle all the folders we collected
            ENUMFOLDERINFO info;
            info.self = this;
            info.pdir = pdir;
            info.prt = prt;

            dpaDirs.DestroyCallbackEx(FolderEnumCallback, &info);
        }
    }
#else
    CDPA<UNALIGNED ITEMIDLIST, CTContainer_PolicyUnOwned<UNALIGNED ITEMIDLIST>>* p_dpaDirs; // ecx
    ENUMFOLDERINFO info; // [esp+8h] [ebp-20h] BYREF
    IEnumIDList* peidl; // [esp+18h] [ebp-10h] BYREF
    DWORD dwAttributes; // [esp+1Ch] [ebp-Ch] BYREF
    CDPAPidl dpaDirs; // [esp+20h] [ebp-8h] BYREF
    CDPAPidl dpaFiles; // [esp+24h] [ebp-4h] BYREF
    LPITEMIDLIST pidl; // [esp+30h] [ebp+8h] SPLIT BYREF

    if (S_OK == pdir->Folder()->EnumObjects(nullptr, 0x60, &peidl))
    {
        dpaDirs = nullptr;
        if (dpaDirs.Create(4))
        {
            dpaFiles = nullptr;
            if (dpaFiles.Create(4))
            {
                if (!peidl->Next(1, &pidl, nullptr))
                {
                    do
                    {
                        dwAttributes = 0x60010000;
                        if (pdir->Folder()->GetAttributesOf(1u, (LPCITEMIDLIST*)&pidl, &dwAttributes) >= 0)
                        {
                            if ((dwAttributes & SFGAO_FOLDER) != 0)
                            {
                                if (_IsExcludedDirectory(pdir->Pidl(), pdir->Folder(), pidl, dwAttributes))
                                {
                                LABEL_20:
                                    ILFree(pidl);
                                    continue;
                                }
                                p_dpaDirs = &dpaDirs;
                            }
                            else
                            {
                                p_dpaDirs = &dpaFiles;
                            }

                            if (p_dpaDirs->AppendPtr(pidl) < 0)
                            {
								goto LABEL_20;
                            }
                        }
                    }
                    while (!peidl->Next(1u, &pidl, nullptr));
                }
                dpaDirs.SortEx(PidlSortCallback, pdir->Folder());

                if (dpaFiles.GetPtrCount() > 0)
                {
                    dpaFiles.SortEx(PidlSortCallback, pdir->Folder());
                    _MergeIntoFolderCache(prt, pdir, &dpaFiles);
                }
                dpaFiles.DestroyCallback(DPA_ILFreeCB, nullptr);
            }

            peidl->Release();

            info.self = this;
            info.prt = prt;
            info.pdir = pdir;
            dpaDirs.DestroyCallbackEx(FolderEnumCallback, &info);
        }
    }
#endif
}

BOOL CMenuItemsCache::FolderEnumCallback(LPITEMIDLIST pidl, ENUMFOLDERINFO* pinfo)
{
    CByUsageDir* pdir = CByUsageDir::Create(pinfo->pdir, pidl);
    if (pdir)
    {
        pinfo->self->_FillFolderCache(pdir, pinfo->prt);
        pdir->Release();
    }
    ILFree(pidl);
    return TRUE;
}

//
//  Returns the next element in prt->_slOld that still belongs to the
//  directory "pdir", or NULL if no more.
//
CByUsageShortcut* CMenuItemsCache::_NextFromCacheInDir(CByUsageRoot* prt, CByUsageDir* pdir)
{
    if (prt->_iOld < prt->_cOld)
    {
        CByUsageShortcut* pscut = prt->_slOld.FastGetPtr(prt->_iOld);
        if (pscut->Dir() == pdir)
        {
            prt->_iOld++;
            return pscut;
        }
    }
    return nullptr;
}

void CMenuItemsCache::_MergeIntoFolderCache(CByUsageRoot *prt, CByUsageDir *pdir, CDPAPidl* pdpaFiles)
{
    //
    //  Look at prt->_slOld to see if we have cached information about
    //  this directory already.
    //
    //  If we find directories that are less than us, skip over them.
    //  These correspond to directories that have been deleted.
    //
    //  For example, if we are "D" and we run across directories
    //  "B" and "C" in the old cache, that means that directories "B"
    //  and "C" were deleted and we should continue scanning until we
    //  find "D" (or maybe we find "E" and stop since E > D).
    //
    //
    CByUsageDir *pdirPrev = NULL;

    while (prt->_iOld < prt->_cOld)
    {
        CByUsageDir *pdirT = prt->_slOld.FastGetPtr(prt->_iOld)->Dir();
        HRESULT hr = _pdirDesktop->Folder()->CompareIDs(0, pdirT->Pidl(), pdir->Pidl());
        if (hr == ResultFromShort(0))
        {
            pdirPrev = pdirT;
            break;
        }
        else if (FAILED(hr) || ShortFromResult(hr) < 0)
        {
            //
            //  Skip over this directory
            //
            while (_NextFromCacheInDir(prt, pdirT)) { }
        }
        else
        {
            break;
        }
    }

    if (pdirPrev)
    {
        //
        //  If we have a cached previous directory, then recycle him.
        //  This keeps us from creating lots of copies of the same IShellFolder.
        //  It is also essential that all entries from the same directory
        //  have the same pdir; that's how _NextFromCacheInDir knows when
        //  to stop.
        //
        pdir = pdirPrev;

        //
        //  Make sure that this IShellFolder supports SHCIDS_ALLFIELDS.
        //  If not, then we just have to assume that they all changed.
        //
        IShellFolder2 *psf2;
        if (SUCCEEDED(pdir->Folder()->QueryInterface(IID_PPV_ARGS(&psf2))))
        {
            psf2->Release();
        }
        else
        {
            pdirPrev = NULL;
        }
    }

    //
    //  Now add all the items in dpaFiles to prt->_sl.  If we find a match
    //  in prt->_slOld, then use that information instead of hitting the disk.
    //
    int iNew;
    CByUsageShortcut *pscutNext = _NextFromCacheInDir(prt, pdirPrev);
    for (iNew = 0; iNew < pdpaFiles->GetPtrCount(); iNew++)
    {
        LPITEMIDLIST pidl = pdpaFiles->FastGetPtr(iNew);

        // Look for a match in the cache.
        HRESULT hr = S_FALSE;
        while (pscutNext &&
               (FAILED(hr = pdir->Folder()->CompareIDs(SHCIDS_ALLFIELDS, pscutNext->RelativePidl(), pidl)) ||
                ShortFromResult(hr) < 0))
        {
            pscutNext = _NextFromCacheInDir(prt, pdirPrev);
        }

        // pscutNext, if non-NULL, is the item that made us stop searching.
        // If hr == S_OK, then it was a match and we should use the data
        // from the cache.  Otherwise, we have a new item and should
        // fill it in the slow way.
        if (hr == ResultFromShort(0))
        {
            // A match from the cache; move it over
            _TransferShortcutToCache(prt, pscutNext);
            pscutNext = _NextFromCacheInDir(prt, pdirPrev);
        }
        else
        {
            // Brand new item, fill in from scratch
            _AddShortcutToCache(pdir, pidl, &prt->_sl);
            pdpaFiles->FastGetPtr(iNew) = NULL; // took ownership
        }
    }
}

//****************************************************************************

bool CMenuItemsCache::_SetInterestingLink(CByUsageShortcut *pscut)
{
    bool fInteresting = true;
    if (pscut->App() && !_PathIsInterestingExe(pscut->App()->GetAppPath()))
    {
        fInteresting = false;
    }
    else if (!_IsInterestingDirectory(pscut->Dir()))
    {
        fInteresting = false;
    }
    else
    {
        LPTSTR pszDisplayName = _DisplayNameOf(pscut->ParentFolder(), pscut->RelativePidl(), SHGDN_NORMAL | SHGDN_INFOLDER);
        if (pszDisplayName)
        {
            // SFGDN_INFOLDER should've returned a relative path
            ASSERT(pszDisplayName == PathFindFileName(pszDisplayName));

            int i;
            for (i = 0; i < _dpaKillLink.GetPtrCount(); i++)
            {
                if (StrStrI(pszDisplayName, _dpaKillLink.GetPtr(i)) != NULL)
                {
                    fInteresting = false;
                    break;
                }
            }
            SHFree(pszDisplayName);
        }
    }

    pscut->SetInteresting(fInteresting);
    return fInteresting;
}

BOOL CMenuItemsCache::_PathIsInterestingExe(LPCTSTR pszPath)
{
    //
    //  Darwin shortcuts are always interesting.
    //
    if (IsDarwinPath(pszPath))
    {
        return TRUE;
    }

    LPCTSTR pszExt = PathFindExtension(pszPath);

    //
    //  *.msc files are also always interesting.  They aren't
    //  strictly-speaking EXEs, but they act like EXEs and administrators
    //  really use them a lot.
    //
    if (StrCmpICW(pszExt, TEXT(".msc")) == 0)
    {
        return TRUE;
    }

    return StrCmpICW(pszExt, TEXT(".exe")) == 0 && !_IsExcludedExe(pszPath);
}


BOOL CMenuItemsCache::_IsExcludedExe(LPCTSTR pszPath)
{
#ifdef DEAD_CODE
    pszPath = PathFindFileName(pszPath);

    int i;
    for (i = 0; i < _dpaKill.GetPtrCount(); i++)
    {
        if (StrCmpI(pszPath, _dpaKill.GetPtr(i)) == 0)
        {
            return TRUE;
        }
    }

    HKEY hk;
    BOOL fRc = FALSE;

    if (SUCCEEDED(_pqa->Init(ASSOCF_OPEN_BYEXENAME, pszPath, NULL, NULL)) &&
        SUCCEEDED(_pqa->GetKey(0, ASSOCKEY_APP, NULL, &hk)))
    {
        fRc = ERROR_SUCCESS == SHQueryValueEx(hk, TEXT("NoStartPage"), NULL, NULL, NULL, NULL);
        RegCloseKey(hk);
    }

    return fRc;
#else
    pszPath = PathFindFileName(pszPath);

    for (int i = 0; i < _dpaKill.GetPtrCount(); i++)
    {
        if (StrCmpI(pszPath, _dpaKill.GetPtr(i)) == 0)
        {
            return 1;
        }
    }

    BOOL fRc = FALSE;
    IQueryAssociations* pqa = (IQueryAssociations *)InterlockedExchangePointer((volatile LPVOID*)&this->_pqa, 0);
    if (pqa || (AssocCreate(CLSID_QueryAssociations, IID_PPV_ARGS(&pqa)), pqa))
    {
        HKEY hkey;
        if (pqa->Init(2, pszPath, 0, 0) >= 0 && pqa->GetKey(0, ASSOCKEY_APP, 0, &hkey) >= 0)
        {
            fRc = SHQueryValueEx(hkey, L"NoStartPage", 0, 0, 0, 0) == 0;
            RegCloseKey(hkey);
        }

        if (InterlockedCompareExchangePointer((volatile LPVOID*)&_pqa, pqa, NULL))
            pqa->Release();
    }
    return fRc;
#endif
}


HRESULT ByUsage::_GetShortcutExeTarget(IShellFolder *psf, LPCITEMIDLIST pidl, LPTSTR pszPath, UINT cchPath)
{
    HRESULT hr;
    IShellLink *psl;

    hr = psf->GetUIObjectOf(_hwnd, 1, &pidl, IID_IShellLink, nullptr, (void**)&psl);

    if (SUCCEEDED(hr))
    {
        hr = psl->GetPath(pszPath, cchPath, 0, 0);
        psl->Release();
    }
    return hr;
}

void _GetUAInfo(const GUID *pguidGrp, LPCWSTR pszPath, UEMINFO *pueiOut)
{
    ZeroMemory(pueiOut, sizeof(UEMINFO));
    pueiOut->cbSize = sizeof(UEMINFO);
    pueiOut->dwMask = 0x11;

    UAQueryEntry(pguidGrp, pszPath, pueiOut);

    if (FILETIMEtoInt64(pueiOut->ftExecute) == 0)
    {
        pueiOut->R = 0.0f;
    }
}

void _GetUEMInfo(const GUID *pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam, UEMINFO *pueiOut)
{
    ZeroMemory(pueiOut, sizeof(UEMINFO));
    pueiOut->cbSize = sizeof(UEMINFO);
    pueiOut->dwMask = UEIM_HIT | UEIM_FILETIME;

    //
    // If this call fails (app / pidl was never run), then we'll
    // just use the zeros we pre-initialized with.
    //
    UEMQueryEvent(pguidGrp, eCmd, wParam, lParam, pueiOut);

    //
    // The UEM code invents a default usage count if the shortcut
    // was never used.  We don't want that.
    //
    if (FILETIMEtoInt64(pueiOut->ftExecute) == 0)
    {
        pueiOut->cLaunches = 0;
    }
}

//
//  Returns S_OK if the item changed, S_FALSE if the item stayed the same,
//  or an error code
//
HRESULT CMenuItemsCache::_UpdateMSIPath(CByUsageShortcut *pscut)
{
    HRESULT hr = S_FALSE;       // Assume nothing happened

    if (pscut->IsDarwin())
    {
        CByUsageHiddenData hd;
        hd.Get(pscut->RelativePidl(), CByUsageHiddenData::BUHD_ALL);
        if (hd.UpdateMSIPath() == S_OK)
        {
            // Redirect to the new target (user may have
            // uninstalled then reinstalled to a new location)
            CByUsageAppInfo *papp = GetAppInfoFromHiddenData(&hd);
            pscut->SetApp(papp);
            if (papp) papp->Release();

            if (pscut->UpdateRelativePidl(&hd))
            {
                hr = S_OK;          // We changed stuff
            }
            else
            {
                hr = E_OUTOFMEMORY; // Couldn't update the relative pidl
            }
        }
        hd.Clear();
    }

    return hr;
}

//
//  Take pscut (which is the CByUsageShortcut most recently enumerated from
//  the old cache) and move it to the new cache.  NULL out the entry in the
//  old cache so that DPADELETEANDDESTROY(prt->_slOld) won't free it.
//
void CMenuItemsCache::_TransferShortcutToCache(CByUsageRoot *prt, CByUsageShortcut *pscut)
{
    ASSERT(pscut);
    ASSERT(pscut == prt->_slOld.FastGetPtr(prt->_iOld - 1));
    if (SUCCEEDED(_UpdateMSIPath(pscut)) &&
        prt->_sl.AppendPtr(pscut) >= 0) {
        // Take ownership
        prt->_slOld.FastGetPtr(prt->_iOld - 1) = NULL;
    }
}

CByUsageAppInfo *CMenuItemsCache::GetAppInfoFromHiddenData(CByUsageHiddenData *phd)
{
    CByUsageAppInfo *papp = NULL;
    bool fIgnoreTimestamp = false;

    TCHAR szPath[MAX_PATH];
    LPTSTR pszPath = szPath;
    szPath[0] = TEXT('\0');

    if (phd->_pwszMSIPath && phd->_pwszMSIPath[0])
    {
        pszPath = phd->_pwszMSIPath;

        // When MSI installs an app, the timestamp is applies to the app
        // is the timestamp on the source media, *not* the time the user
        // user installed the app.  So ignore the timestamp entirely since
        // it's useless information (and in fact makes us think the app
        // is older than it really is).
        fIgnoreTimestamp = true;
    }
    else if (phd->_pwszTargetPath)
    {
        if (IsDarwinPath(phd->_pwszTargetPath))
        {
            pszPath = phd->_pwszTargetPath;
        }
        else
        {
            //
            //  Need to expand the path because it may contain environment
            //  variables.
            //
            ExpandEnvironmentStrings(phd->_pwszTargetPath, szPath, ARRAYSIZE(szPath));
        }
    }

    return GetAppInfo(pszPath, fIgnoreTimestamp);
}

CByUsageShortcut *CMenuItemsCache::CreateShortcutFromHiddenData(CByUsageDir *pdir, LPCITEMIDLIST pidl, CByUsageHiddenData *phd, BOOL fForce)
{
    CByUsageAppInfo *papp = GetAppInfoFromHiddenData(phd);
    bool fDarwin = phd->_pwszTargetPath && IsDarwinPath(phd->_pwszTargetPath);
    CByUsageShortcut *pscut = CByUsageShortcut::Create(pdir, pidl, papp, fDarwin, fForce);
    if (papp) papp->Release();
    return pscut;
}


void CMenuItemsCache::_AddShortcutToCache(CByUsageDir *pdir, LPITEMIDLIST pidl, ByUsageShortcutList* pslFiles)
{
#ifdef DEAD_CODE
    HRESULT hr;
    CByUsageHiddenData hd;

    if (pidl)
    {
        //
        //  Juice-up this pidl with cool info about the shortcut target
        //
        IShellLink *psl;
        hr = pdir->Folder()->GetUIObjectOf(NULL, 1, const_cast<LPCITEMIDLIST *>(&pidl),
                                           IID_IShellLink, nullptr, (void**)&psl);
        if (SUCCEEDED(hr))
        {
            hd.LoadFromShellLink(psl);

            psl->Release();

            if (hd._pwszTargetPath && IsDarwinPath(hd._pwszTargetPath))
            {
                SHRegisterDarwinLink(ILCombine(pdir->Pidl(), pidl),
                                     hd._pwszTargetPath +1 /* Exclude the Darwin marker! */,
                                     _fCheckDarwin);

                SHParseDarwinIDFromCacheW(hd._pwszTargetPath+1, &hd._pwszMSIPath);
            }

            // ByUsageHiddenData::Set frees the source pidl on failure
            pidl = hd.Set(pidl);

        }
    }

    if (pidl)
    {
        CByUsageShortcut *pscut = CreateShortcutFromHiddenData(pdir, pidl, &hd);

        if (pscut)
        {
            if (slFiles.AppendPtr(pscut) >= 0)
            {
                _SetInterestingLink(pscut);
            }
            else
            {
                // Couldn't append; oh well
                delete pscut;       // "delete" can handle NULL pointer
            }
        }

        ILFree(pidl);
    }
    hd.Clear();
#else
	HRESULT hr;
    CByUsageHiddenData hd;

    if (pidl)
    {
        IShellLink *psl;
        hr = pdir->Folder()->GetUIObjectOf(NULL, 1, const_cast<LPCITEMIDLIST *>(&pidl), IID_IShellLink, NULL, (void **)&psl);
        if (SUCCEEDED(hr))
        {
            hd.LoadFromShellLink(psl);
            psl->Release();
            if (hd._pwszTargetPath && IsDarwinPath(hd._pwszTargetPath))
            {
                SHRegisterDarwinLink(ILCombine(pdir->Pidl(), pidl), hd._pwszTargetPath + 1, this->_fCheckDarwin);
                SHParseDarwinIDFromCacheW(hd._pwszTargetPath + 1, &hd._pwszMSIPath);
            }

            pidl = hd.Set(pidl);
        }

        if (pidl)
        {
            CByUsageShortcut *pscut = CreateShortcutFromHiddenData(pdir, pidl, &hd);
            if (pscut)
            {
                if (pslFiles->AppendPtr(pscut) >= 0)
                {
                    _SetInterestingLink(pscut);
                }
                else
                {
                    delete pscut;
                }
            }

            ILFree(pidl);
        }
    }
    hd.Clear();
#endif
}

//
//  Find an entry in the AppInfo list that matches this application.
//  If not found, create a new entry.  In either case, bump the
//  reference count and return the item.
//
CByUsageAppInfo* CMenuItemsCache::GetAppInfo(LPTSTR pszAppPath, bool fIgnoreTimestamp)
{
    Lock();

    CByUsageAppInfo *pappBlank = NULL;

    int i;
    for (i = _dpaAppInfo.GetPtrCount() - 1; i >= 0; i--)
    {
        CByUsageAppInfo *papp = _dpaAppInfo.FastGetPtr(i);
        if (papp->IsBlank())
        {   // Remember that we found a blank entry we can recycle
            pappBlank = papp;
        }
        else if (lstrcmpi(papp->_pszAppPath, pszAppPath) == 0)
        {
            papp->AddRef();
            Unlock();
            return papp;
        }
    }

    // Not found in the list.  Try to recycle a blank entry.

    if (!pappBlank)
    {
        // No blank entries found; must make a new one.
        pappBlank = CByUsageAppInfo::Create();
        if (pappBlank && _dpaAppInfo.AppendPtr(pappBlank) < 0)
        {
            delete pappBlank;
            pappBlank = NULL;
        }
    }

    if (pappBlank && pappBlank->Initialize(pszAppPath, this, _fCheckNew, fIgnoreTimestamp))
    {
        ASSERT(pappBlank->IsBlank());
        pappBlank->AddRef();
    }
    else
    {
        pappBlank = NULL;
    }

    Unlock();
    return pappBlank;
}

    // A shortcut is new if...
    //
    //  the shortcut is newly created, and
    //  the target is newly created, and
    //  neither the shortcut nor the target has been run "in an interesting
    //  way".
    //
    // An "interesting way" is "more than one hour after the shortcut/target
    // was created."
    //
    // Note that we test the easiest things first, to avoid hitting
    // the disk too much.

bool ByUsage::_IsShortcutNew(CByUsageShortcut *pscut, CByUsageAppInfo *papp, const UEMINFO *puei)
{
    //
    //  Shortcut is new if...
    //
    //  It was run less than an hour after the app was installed.
    //  It was created relatively recently.
    //
    //
    bool fNew = FILETIMEtoInt64(puei->ftExecute) < FILETIMEtoInt64(papp->_ftCreated) + FT_NEWAPPGRACEPERIOD() &&
                _pMenuCache->IsNewlyCreated(&pscut->GetCreatedTime());

    return fNew;
}

//****************************************************************************


// See how many pinned items there are, so we can tell our dad
// how big we want to be.

void ByUsage::PrePopulate()
{
    _FillPinnedItemsCache();
    _NotifyDesiredSize();
}

//
//  Enumerating out of cache.
//
void ByUsage::EnumFolderFromCache()
{
#ifdef DEAD_CODE
    if (SHRestricted(REST_NOSMMFUPROGRAMS)) //If we don't need MFU list,...
        return;                            // don't enumerate this!

    _pMenuCache->StartEnum();

    LPITEMIDLIST pidlDesktop, pidlCommonDesktop;
    (void)SHGetSpecialFolderLocation(NULL, CSIDL_DESKTOPDIRECTORY, &pidlDesktop);
    (void)SHGetSpecialFolderLocation(NULL, CSIDL_COMMON_DESKTOPDIRECTORY, &pidlCommonDesktop);

    while (TRUE)
    {
        CByUsageShortcut *pscut = _pMenuCache->GetNextShortcut();

        if (!pscut)
            break;

        if (!pscut->IsInteresting())
            continue;

        // Find out if the item is on the desktop, because we don't track new items on the desktop.
        BOOL fIsDesktop = FALSE;
        if ((pidlDesktop && ILIsEqual(pscut->ParentPidl(), pidlDesktop)) ||
            (pidlCommonDesktop && ILIsEqual(pscut->ParentPidl(), pidlCommonDesktop)))
        {
            fIsDesktop = TRUE;
            pscut->SetNew(FALSE);
        }


        CByUsageAppInfo *papp = pscut->App();

        if (papp)
        {
            // Now enumerate the item itself.  Enumerating an item consists
            // of extracting its UEM data, updating the totals, and possibly
            // marking ourselves as the "best" representative of the associated
            // application.
            //
            //
            UEMINFO uei;
            pscut->GetUEMInfo(&uei);

            // See if this shortcut is still new.  If the app is no longer new,
            // then there's no point in keeping track of the shortcut's new-ness.

            if (pscut->IsNew() && papp->_fNew)
            {
                pscut->SetNew(_IsShortcutNew(pscut, papp, &uei));
            }

            //
            //  Maybe we are the "best"...  Note that we win ties.
            //  This ensures that even if an app is never run, *somebody*
            //  will be chosen as the "best".
            //
            if (CompareUEMInfo(&uei, &papp->_ueiBest) <= 0)
            {
                papp->_ueiBest = uei;
                papp->_pscutBest = pscut;
                if (!fIsDesktop)
                {
                    // Best Start Menu (i.e., non-desktop) item
                    papp->_pscutBestSM = pscut;
                }
            }

            //  Include this file's UEM info in the total
            papp->CombineUEMInfo(&uei, pscut->IsNew(), fIsDesktop);
        }
    }
    _pMenuCache->EndEnum();
    ILFree(pidlCommonDesktop);
    ILFree(pidlDesktop);
#else
    CMenuItemsCache *pMenuCache; // eax
    CMenuItemsCache *i; // ecx
    CByUsageShortcut *pscut; // eax MAPDST

    if (!SHRestricted(REST_NOSMMFUPROGRAMS))
    {
        _pMenuCache->StartEnum();

        LPITEMIDLIST pidlDesktop = 0;
        LPITEMIDLIST pidlCommonDesktop = 0;
        (void)SHGetSpecialFolderLocation(0, CSIDL_DESKTOPDIRECTORY, &pidlDesktop);
        (void)SHGetSpecialFolderLocation(0, CSIDL_COMMON_DESKTOPDIRECTORY, &pidlCommonDesktop);

        for (i = _pMenuCache; ; i = this->_pMenuCache)
        {
            pscut = i->GetNextShortcut();
            if (!pscut)
                break;

            if (pscut->IsInteresting())
            {
                BOOL fIsDesktop = FALSE;
                if ((pidlDesktop && ILIsEqual(pscut->ParentPidl(), pidlDesktop)) ||
                    (pidlCommonDesktop && ILIsEqual(pscut->ParentPidl(), pidlCommonDesktop)))
                {
                    fIsDesktop = TRUE;
                    pscut->SetNew(FALSE);
                }

                CByUsageAppInfo *papp = pscut->App();

                UEMINFO uei;
                if (papp && pscut->GetUAInfo(&uei) >= 0)
                {
#ifdef DEBUG
                    PRINT_UEMINFO(uei);
#endif
                    if (pscut->IsNew() && papp->_fNew)
                    {
                        pscut->SetNew(_IsShortcutNew(pscut, papp, &uei));
                    }

                    if (CompareUAInfo(&uei, &papp->_ueiBest) <= 0)
                    {
                        if (0.0f != uei.R || !papp->_pscutBest || papp->_pscutBest->field_17 || !pscut->field_17)
                        {
                            papp->_ueiBest = uei;
                            papp->_pscutBest = pscut;
                            if (CompareFileTime(&papp->_ueiBest.ftExecute, &field_5C) > 0)
                            {
								field_5C = papp->_ueiBest.ftExecute;
                            }
                        }

                        if (!fIsDesktop)
                            papp->_pscutBestSM = pscut;
                    }

                    papp->CombineUAInfo(&uei, pscut->IsNew(), fIsDesktop, 0);
                }
            }
        }

        ILFree(pidlCommonDesktop);
        ILFree(pidlDesktop);
    }
#endif
}

BOOL IsPidlInDPA(LPCITEMIDLIST pidl, const CDPAPidl& dpa)
{
    int i;
    for (i = dpa.GetPtrCount()-1; i >= 0; i--)
    {
        if (ILIsEqual(pidl, dpa.FastGetPtr(i)))
        {
            return TRUE;
        }
    }
    return FALSE;
}

void __stdcall ByUsage::_AddNewAppPidlAndParents(CDPAPidl *pdpa, ITEMIDLIST_ABSOLUTE *pidl)
{
    ITEMIDLIST_ABSOLUTE *v2; // ebx
    ITEMIDLIST_RELATIVE *v3; // esi

    v2 = pidl;
    if (pidl)
    {
        do
        {
            v3 = 0;
            if (pdpa->AppendPtr(v2, 0) >= 0)
            {
                v3 = ILCloneParent(v2);
                v2 = 0;
                if (ILIsEmpty(v3) || IsPidlInDPA(v3, *pdpa))
                {
                    ILFree(v3);
                    v3 = 0;
                }
            }
            ILFree(v2);
            v2 = v3;
        } while (v3);
    }
}

BOOL ByUsage::_AfterEnumCB(CByUsageAppInfo *papp, AFTERENUMINFO *paei)
{
#ifdef DEAD_CODE
    // A ByUsageAppInfo doesn't exist unless there's a CByUsageShortcut
    // that references it or it is pinned...

    if (!papp->IsBlank() && papp->_pscutBest)
    {
        UEMINFO uei;
        papp->GetUAInfo(&uei);
        papp->CombineUAInfo(&uei, papp->_IsUAINFONew(&uei));

        // A file counts on the list only if it has been used
        // and is not pinned.  (Pinned items are added to the list
        // elsewhere.)
        //
        // Note that "new" apps are *not* placed on the list until
        // they are used.  ("new" apps are highlighted on the
        // Start Menu.)

        if (!papp->_fPinned &&
            papp->_ueiTotal.cHit && FILETIMEtoInt64(papp->_ueiTotal.ftExecute))
        {

            CByUsageItem *pitem = papp->CreateByUsageItem();
            if (pitem)
            {
                LPITEMIDLIST pidl = pitem->CreateFullPidl();
                if (paei->self->_pByUsageUI)
                {
                    paei->self->_pByUsageUI->AddItem(pitem, NULL, pidl);
                }
                ILFree(pidl);
            }
        }
        else
        {
        }


#if 0
        //
        //  If you enable this code, then holding down Ctrl and Alt
        //  will cause us to pick a program to be new.  This is for
        //  testing the "new apps" balloon tip.
        //
#define DEBUG_ForceNewApp() \
        (paei->dpaNew && paei->dpaNew.GetPtrCount() == 0 && \
         GetAsyncKeyState(VK_CONTROL) < 0 && GetAsyncKeyState(VK_MENU) < 0)
#else
#define DEBUG_ForceNewApp() FALSE
#endif

        //
        //  Must also check _pscutBestSM because if an app is represented
        //  only on the desktop and not on the start menu, then
        //  _pscutBestSM will be NULL.
        //
        if (paei->dpaNew && (papp->IsNew() || DEBUG_ForceNewApp()) && papp->_pscutBestSM)
        {
            // NTRAID:193226 We mistakenly treat apps on the desktop
            // as if they were "new".
            // we should only care about apps in the start menu
            LPITEMIDLIST pidl = papp->_pscutBestSM->CreateFullPidl();
            while (pidl)
            {
                LPITEMIDLIST pidlParent = NULL;

                if (paei->dpaNew.AppendPtr(pidl) >= 0)
                {
                    pidlParent = ILClone(pidl);
                    pidl = NULL; // ownership of pidl transferred to DPA
                    if (!ILRemoveLastID(pidlParent) || ILIsEmpty(pidlParent) || IsPidlInDPA(pidlParent, paei->dpaNew))
                    {
                        // If failure or if we already have it in the list
                        ILFree(pidlParent);
                        pidlParent = NULL;
                    }

                    // Remember the creation time of the most recent app
                    if (CompareFileTime(&paei->self->_ftNewestApp, &papp->GetCreatedTime()) < 0)
                    {
                        paei->self->_ftNewestApp = papp->GetCreatedTime();
                    }

                    // If the shortcut is even newer, then use that.
                    // This happens in the "freshly installed Darwin app"
                    // case, because Darwin is kinda reluctant to tell
                    // us where the EXE is so all we have to go on is
                    // the shortcut.

                    if (CompareFileTime(&paei->self->_ftNewestApp, &papp->_pscutBestSM->GetCreatedTime()) < 0)
                    {
                        paei->self->_ftNewestApp = papp->_pscutBestSM->GetCreatedTime();
                    }


                }
                ILFree(pidl);

                // Now add the parent to the list also.
                pidl = pidlParent;
            }
        }
    }
    return TRUE;
#else
    if (!papp->IsBlank() && papp->_pscutBest)
    {
        UEMINFO puei;
        papp->GetUAInfo(&puei);
        papp->CombineUAInfo(&puei, papp->_IsUAINFONew(&puei), 0, 1);
        
        if (CompareFileTime(&puei.ftExecute, &paei->self->field_5C) > 0)
            paei->self->field_5C = puei.ftExecute;
        
        if (!papp->_fPinned)
        {
            if (FILETIMEtoInt64(papp->_ueiTotal.ftExecute))
            {
                CByUsageItem* pitem = papp->CreateCByUsageItem();
                if (pitem)
                {
                    if (paei->self->_pByUsageUI)
                    {
                        paei->self->_pByUsageUI->AddItem(pitem);
                        ++paei->field_8;
                    }
                    pitem->Release();
                }
            }
        }

        if (paei->dpaNew && papp->_fNew && !papp->_fPinned && papp->_pscutBestSM)
        {
            //CcshellDebugMsgW(32, "%p.app.new(%s)", papp, (const char *)papp->_pszAppPath);
            ITEMIDLIST_ABSOLUTE* v6 = ILCombine(papp->_pscutBestSM->ParentPidl(), papp->_pscutBestSM->RelativePidl());
            _AddNewAppPidlAndParents(&paei->dpaNew, paei->self->_pMenuCache->GetPerUserVersionOfSharedItem(v6));
            _AddNewAppPidlAndParents(&paei->dpaNew, v6);

            if (CompareFileTime(&paei->self->_ftNewestApp, &papp->_ftCreated) < 0)
                paei->self->_ftNewestApp = papp->_ftCreated;
            if (CompareFileTime(&paei->self->_ftNewestApp, &papp->_pscutBestSM->GetCreatedTime()) < 0)
                paei->self->_ftNewestApp = papp->_pscutBestSM->GetCreatedTime();
        }
    }
    return 1;
#endif
}

BOOL ByUsage::IsSpecialPinnedPidl(LPCITEMIDLIST pidl)
{
    return _pdirDesktop->Folder()->CompareIDs(SHCIDS_CANONICALONLY, pidl, _pidlEmail) == S_OK ||
           _pdirDesktop->Folder()->CompareIDs(SHCIDS_CANONICALONLY, pidl, _pidlBrowser) == S_OK;
}

BOOL ByUsage::IsSpecialPinnedItem(CByUsageItem *pitem)
{
    return IsSpecialPinnedPidl(pitem->RelativePidl());
}

//
//  For each app we found, add it to the list.
//
void ByUsage::AfterEnumItems()
{
    //
    //  First, all pinned items are enumerated unconditionally.
    //
    if (_rtPinned._sl && _rtPinned._sl.GetPtrCount())
    {
        int i;
        for (i = 0; i < _rtPinned._sl.GetPtrCount(); i++)
        {
            CByUsageShortcut *pscut = _rtPinned._sl.FastGetPtr(i);
            CByUsageItem *pitem = pscut->CreatePinnedItem(i);
            if (pitem)
            {
                // Pinned items are relative to the desktop, so we can
                // save ourselves an ILClone because the relative pidl
                // is equal to the absolute pidl.
                ASSERT(pitem->Dir() == _pdirDesktop);

                //
                // Special handling for E-mail and Internet pinned items
                //
                if (IsSpecialPinnedItem(pitem))
                {
                    pitem->EnableSubtitle();
                }

                if (_pByUsageUI)
                    _pByUsageUI->AddItem(pitem/*, NULL, pscut->RelativePidl()*/);

                pitem->Release();
            }
        }
    }

    //
    //  Now add the separator after the pinned items.
    //
    CByUsageItem *pitem = CByUsageItem::CreateSeparator();
    if (pitem)
    {
        if (_pByUsageUI)
        {
            _pByUsageUI->AddItem(pitem/*, NULL, NULL*/);
        }
        pitem->Release();
    }

    //
    //  Now walk through all the regular items.
    //
    //  PERF: Can skip this if _cMFUDesired==0 and "highlight new apps" is off
    //
    AFTERENUMINFO aei;
    aei.self = this;
    aei.dpaNew.Create(4);       // Will check failure in callback

    ByUsageAppInfoList *pdpaAppInfo = _pMenuCache->GetAppList();
    pdpaAppInfo->EnumCallbackEx(_AfterEnumCB, &aei);

#ifdef DEAD_CODE
    // Now that we have the official list of new items, tell the
    // foreground thread to pick it up.  We don't update the master
    // copy in-place for three reasons.
    //
    //  1.  It generates contention since both the foreground and
    //      background threads would be accessing it simultaneously.
    //      This means more critical sections (yuck).
    //  2.  It means that items that were new and are still new have
    //      a brief period where they are no longer new because we
    //      are rebuilding the list.
    //  3.  By having only one thread access the master copy, we avoid
    //      synchronization issues.

    if (aei.dpaNew && _pByUsageUI && _pByUsageUI->_hwnd && SendNotifyMessage(_pByUsageUI->_hwnd, BUM_SETNEWITEMS, 0, (LPARAM)(HDPA)aei.dpaNew))
    {
        aei.dpaNew.Detach();       // Successfully delivered
    }

    //  If we were unable to deliver the new HDPA, then destroy it here
    //  so we don't leak.
    if (aei.dpaNew)
    {
        aei.dpaNew.DestroyCallback(DPA_ILFreeCB, NULL);
    }
#else
    if (aei.dpaNew)
    {
        if (_pByUsageUI)
        {
            if (_pByUsageUI->_hwnd)
            {
                if (SendNotifyMessageW(_pByUsageUI->_hwnd, 0x8000u, 0, (LPARAM)(HDPA)aei.dpaNew))
                {
                    aei.dpaNew.Detach();
                }
            }
        }

        if (aei.dpaNew)
        {
            aei.dpaNew.DestroyCallback(DPA_ILFreeCB, NULL);
        }
    }
#endif

    if (!_fUEMRegistered)
    {
        // Register with UEM DB if we haven't done it yet
        ASSERT(!_pMenuCache->IsLocked());
        _fUEMRegistered = SUCCEEDED(UARegisterNotify(UANotifyCB, static_cast<void *>(this), 1));
    }
}

int ByUsage::UANotifyCB(void* param, const GUID* pguidGrp, const WCHAR*, UAEVENT eCmd)
{
#ifdef DEAD_CODE
    // Refresh our list whenever a new app is started.
    // or when the session changes (because that changes all the usage counts)
    printf("UANotifyCB: %d %x\n", eCmd, param);
    switch (eCmd)
    {
        case UAE_LAUNCH:
        case UAE_TIME:
        {
            // @MOD: Use a global pointer for the ByUsage UI instead of using UserAssist event param.
            // The former seems to be point to an invalid object sometimes.
            if (g_pByUsageUI)
            {
                g_pByUsageUI->Invalidate();
                g_pByUsageUI->StartRefreshTimer();
            }
            break;
        }
    }

    return 0;
#else
    ByUsage *pbu = reinterpret_cast<ByUsage *>(param);
    switch (eCmd)
    {
        case UAE_LAUNCH:
        case UAE_SESSION:
        {
            if (IsEqualGUID(*pguidGrp, UAIID_APPLICATIONS) || IsEqualGUID(*pguidGrp, UAIID_AUTOMATIC))
            {
                if (pbu && pbu->_pByUsageUI)
                {
                    pbu->_pByUsageUI->Invalidate();
                    pbu->_pByUsageUI->StartRefreshTimer();
                }
            }
            break;
        }
        default:
            // Do nothing
            ;
    }
	return 0;
#endif
}

template <class T>
int DPA_LocalFreeCB(T self, LPVOID)
{
    LocalFree(self);
    return TRUE;
}

BOOL CreateExcludedDirectoriesDPA(const int rgcsidlExclude[], CDPA<TCHAR, CTContainer_PolicyUnOwned<TCHAR>> *pdpaExclude)
{
    if (pdpaExclude)
    {
        pdpaExclude->EnumCallback(DPA_LocalFreeCB, NULL);
        if (pdpaExclude)
        {
            pdpaExclude->DeleteAllPtrs();
        }
    }
    else if (!pdpaExclude->Create(4))
    {
        return FALSE;
    }

    ASSERT(*pdpaExclude);
    ASSERT(pdpaExclude->GetPtrCount() == 0);

    int i = 0;
    while (rgcsidlExclude[i] != -1)
    {
        TCHAR szPath[MAX_PATH];
        // Note: This call can legitimately fail if the corresponding
        // folder does not exist, so don't get upset.  Less work for us!
        if (SUCCEEDED(SHGetFolderPath(NULL, rgcsidlExclude[i], NULL, SHGFP_TYPE_CURRENT, szPath)))
        {
            AppendString(pdpaExclude, szPath);
        }
        i++;
    }

    return TRUE;
}

BOOL CMenuItemsCache::_GetExcludedDirectories()
{
    //
    //  The directories we exclude from enumeration - Shortcuts in these
    //  folders are never candidates for inclusion.
    //
    static const int c_rgcsidlUninterestingDirectories[] = {
        CSIDL_ALTSTARTUP,
        CSIDL_STARTUP,
        CSIDL_COMMON_ALTSTARTUP,
        CSIDL_COMMON_STARTUP,
        -1          // End marker
    };

    return CreateExcludedDirectoriesDPA(c_rgcsidlUninterestingDirectories, &_dpaNotInteresting);
}

BOOL CMenuItemsCache::_IsExcludedDirectory(LPCITEMIDLIST a2, IShellFolder *psf, LPCITEMIDLIST pidl, DWORD dwAttributes)
{
#ifdef DEAD_CODE
    if (_enumfl & ENUMFL_NORECURSE)
        return TRUE;

    if (!(dwAttributes & SFGAO_FILESYSTEM))
        return TRUE;

    // SFGAO_LINK | SFGAO_FOLDER = folder shortcut.
    // We want to exclude those because we can get blocked
    // on network stuff
    if (dwAttributes & SFGAO_LINK)
        return TRUE;

    return FALSE;
#else
    WCHAR pszPath[260]; // [esp+Ch] [ebp-414h] BYREF
    WCHAR String2[260]; // [esp+214h] [ebp-20Ch] BYREF
    return (this->_enumfl & 1) != 0
        || (dwAttributes & 0x40000000) == 0
        || (dwAttributes & 0x10000) != 0
        || !SHGetPathFromIDListW(a2, pszPath)
        || DisplayNameOf(psf, pidl, 0x8000, String2, 260u) < 0
        || !lstrcmpiW(pszPath, String2)
        || !PathIsPrefixW(pszPath, String2);
#endif
}

BOOL CMenuItemsCache::_IsInterestingDirectory(CByUsageDir *pdir)
{
    STRRET str;
    TCHAR szPath[MAX_PATH];
    if (SUCCEEDED(_pdirDesktop->Folder()->GetDisplayNameOf(pdir->Pidl(), SHGDN_FORPARSING, &str)) &&
        SUCCEEDED(StrRetToBuf(&str, pdir->Pidl(), szPath, ARRAYSIZE(szPath))))
    {
        int i;
        for (i = _dpaNotInteresting.GetPtrCount() - 1; i >= 0; i--)
        {
            if (lstrcmpi(_dpaNotInteresting.FastGetPtr(i), szPath) == 0)
            {
                return FALSE;
            }
        }
    }
    return TRUE;
}

void ByUsage::OnPinListChange()
{
    _pByUsageUI->Invalidate();
    PostMessage(_pByUsageUI->_hwnd, ByUsageUI::SFTBM_REFRESH, TRUE, 0);
}

void ByUsage::OnChangeNotify(UINT id, LONG lEvent, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
{
    if (id == NOTIFY_PINCHANGE)
    {
        if (lEvent == SHCNE_EXTENDED_EVENT && pidl1)
        {
            SHChangeDWORDAsIDList *pdwidl = (SHChangeDWORDAsIDList *)pidl1;
            if (pdwidl->dwItem1 == SHCNEE_PINLISTCHANGED)
            {
                OnPinListChange();
            }
        }
    }
    else if (_pMenuCache)
    {
        _pMenuCache->OnChangeNotify(id, lEvent, pidl1, pidl2);
    }
}


void CMenuItemsCache::OnChangeNotify(UINT id, LONG lEvent, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2)
{
    ASSERT(id < min(MAXNOTIFY, NUM_PROGLIST_ROOTS));

    if (id < NUM_PROGLIST_ROOTS)
    {
        _rgrt[id].SetNeedRefresh();
        _fIsCacheUpToDate = FALSE;
        // Once we get one notification, there's no point in listening to further
        // notifications until our next enumeration.  This keeps us from churning
        // while Winstones are running.
        if (_pByUsageUI)
        {
            ASSERT(!IsLocked());
            _pByUsageUI->UnregisterNotify(id);
            _rgrt[id].ClearRegistered();
            _pByUsageUI->Invalidate();
            _pByUsageUI->RefreshNow();
        }
    }
}

void CMenuItemsCache::UnregisterNotifyAll()
{
    if (_pByUsageUI)
    {
        UINT id;
        for (id = 0; id < NUM_PROGLIST_ROOTS; id++)
        {
            _rgrt[id].ClearRegistered();
            _pByUsageUI->UnregisterNotify(id);
        }
    }
}

inline LRESULT ByUsage::_OnNotify(LPNMHDR pnm)
{
    switch (pnm->code)
    {
    case SMN_MODIFYSMINFO:
        return _ModifySMInfo(CONTAINING_RECORD(pnm, SMNMMODIFYSMINFO, hdr));
    }
    return 0;
}

//
//  We need this message to avoid a race condition between the background
//  thread (the enumerator) and the foreground thread.  So the rule is
//  that only the foreground thread is allowd to mess with _dpaNew.
//  The background thread collects the information it wants into a
//  separate DPA and hands it to us on the foreground thread, where we
//  can safely set it into _dpaNew without encountering a race condition.
//
inline LRESULT ByUsage::_OnSetNewItems(HDPA hdpaNew)
{
    CDPAPidl dpaNew(hdpaNew);

    //
    //  Most of the time, there are no new apps and there were no new apps
    //  last time either.  Short-circuit this case...
    //
    int cNew = _dpaNew ? _dpaNew.GetPtrCount() : 0;

    if (cNew == 0 && dpaNew.GetPtrCount() == 0)
    {
        // Both old and new are empty.  We're finished.
        // (Since we own dpaNew, free it to avoid a memory leak.)
        dpaNew.DestroyCallback(ILFreeCallback, NULL);
        return 0;
    }

    //  Now swap the new DPA in

    if (_dpaNew)
    {
        _dpaNew.DestroyCallback(ILFreeCallback, NULL);
    }
    _dpaNew.Attach(hdpaNew);

    // Tell our dad that we can identify new items
    // Also tell him the timestamp of the most recent app
    // (so he can tell whether or not to restart the "offer new apps" counter)
    SMNMHAVENEWITEMS nmhni;
    nmhni.ftNewestApp = _ftNewestApp;
    _SendNotify(_pByUsageUI->_hwnd, SMN_HAVENEWITEMS, &nmhni.hdr);

    return 0;
}

LRESULT ByUsage::OnWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_NOTIFY:
        return _OnNotify(reinterpret_cast<LPNMHDR>(lParam));

    case BUM_SETNEWITEMS:
        return _OnSetNewItems(reinterpret_cast<HDPA>(lParam));

    case WM_SETTINGCHANGE:
        static const TCHAR c_szClients[] = TEXT("Software\\Clients");
        if ((wParam == 0 && lParam == 0) ||     // wildcard
            (lParam && StrCmpNIC((LPCTSTR)lParam, c_szClients, ARRAYSIZE(c_szClients) - 1) == 0)) // client change
        {
            _pByUsageUI->ForceChange();         // even though the pidls didn't change, their targets did
            _ulPinChange = -1;                  // Force reload even if list didn't change
            OnPinListChange();                  // reload the pin list (since a client changed)
        }
        break;
    }

    // Else fall back to parent implementation
    return _pByUsageUI->SFTBarHost::OnWndProc(hwnd, uMsg, wParam, lParam);
}

LRESULT ByUsage::_ModifySMInfo(PSMNMMODIFYSMINFO pmsi)
{
    LPSMDATA psmd = pmsi->psmd;

    // Do this only if there is a ShellFolder.  We don't want to fault
    // on the static menu items.
    if ((psmd->dwMask & SMDM_SHELLFOLDER) && _dpaNew)
    {

        // NTRAID:135699: this needs big-time optimization
        // E.g., remember the previous folder if there was nothing found

        LPITEMIDLIST pidl = NULL;

        IAugmentedShellFolder2* pasf2;
        if (SUCCEEDED(psmd->psf->QueryInterface(IID_PPV_ARGS(&pasf2))))
        {
            LPITEMIDLIST pidlFolder;
            LPITEMIDLIST pidlItem;
            if (SUCCEEDED(pasf2->UnWrapIDList(psmd->pidlItem, 1, NULL, &pidlFolder, &pidlItem, NULL)))
            {
                pidl = ILCombine(pidlFolder, pidlItem);
                ILFree(pidlFolder);
                ILFree(pidlItem);
            }
            pasf2->Release();
        }

        if (!pidl)
        {
            pidl = ILCombine(psmd->pidlFolder, psmd->pidlItem);
        }

        if (pidl)
        {
            if (IsPidlInDPA(pidl, _dpaNew))
            {
                // Designers say: New items should never be demoted
                pmsi->psminfo->dwFlags |= SMIF_NEW;
                pmsi->psminfo->dwFlags &= ~SMIF_DEMOTED;
            }
            ILFree(pidl);
        }
    }
    return 0;
}

void ByUsage::_FillPinnedItemsCache()
{
#ifdef DEAD_CODE
    if (SHRestricted(REST_NOSMPINNEDLIST))   //If no pinned list is allowed,.....
        return;                             //....there is nothing to do!

    ULONG ulPinChange;
    _psmpin->GetChangeCount(&ulPinChange);
    if (_ulPinChange == ulPinChange)
    {
        // No change in pin list; do not need to reload
        return;
    }

    _ulPinChange = ulPinChange;
    _rtPinned.Reset();
    if (_rtPinned._sl.Create(4))
    {
        IEnumFullIDList *penum;

        if (SUCCEEDED(_psmpin->EnumObjects(&penum)))
        {
            LPITEMIDLIST pidl;
            while (penum->Next(1, &pidl, NULL) == S_OK)
            {
                IShellLink *psl;
                HRESULT hr;
                CByUsageHiddenData hd;

                //
                //  If we have a shortcut, do bookkeeping based on the shortcut
                //  target.  Otherwise do it based on the pinned object itself.
                //  Note that we do not go through _PathIsInterestingExe
                //  because all pinned items are interesting.

                hr = SHGetUIObjectFromFullPIDL(pidl, NULL, IID_PPV_ARGS(&psl));
                if (SUCCEEDED(hr))
                {
                    hd.LoadFromShellLink(psl);
                    psl->Release();

                    // We do not need to SHRegisterDarwinLink because the only
                    // reason for getting the MSI path is so pinned items can
                    // prevent items on the Start Menu from appearing in the MFU.
                    // So let the shortcut on the Start Menu do the registration.
                    // (If there is none, then that's even better - no work to do!)
                    hd.UpdateMSIPath();
                }

                if (FAILED(hr))
                {
                    hr = DisplayNameOfAsOLESTR(_pdirDesktop->Folder(), pidl, SHGDN_FORPARSING, &hd._pwszTargetPath);
                }

                //
                //  If we were able to figure out what the pinned object is,
                //  use that information to block the app from also appearing
                //  in the MFU.
                //
                //  Inability to identify the pinned
                //  object is not grounds for rejection.  A pinned items is
                //  of great sentimental value to the user.
                //
                if (FAILED(hr))
                {
                    ASSERT(hd.IsClear());
                }

                CByUsageShortcut *pscut = _pMenuCache->CreateShortcutFromHiddenData(_pdirDesktop, pidl, &hd, TRUE);
                if (pscut)
                {
                    if (_rtPinned._sl.AppendPtr(pscut) >= 0)
                    {
                        pscut->SetInteresting(true);  // Pinned items are always interesting
                        if (IsSpecialPinnedPidl(pidl))
                        {
                            CByUsageAppInfo *papp = _pMenuCache->GetAppInfoFromSpecialPidl(pidl);
                            pscut->SetApp(papp);
                            if (papp) papp->Release();
                        }
                    }
                    else
                    {
                        // Couldn't append; oh well
                        delete pscut;       // "delete" can handle NULL pointer
                    }
                }
                hd.Clear();
                ILFree(pidl);
            }
            penum->Release();
        }
    }
#else
    ULONG ulPinChange; // [esp+24h] [ebp-28h] BYREF
    IEnumFullIDList *penum; // [esp+28h] [ebp-24h] BYREF
    LPITEMIDLIST pidl; // [esp+30h] [ebp-1Ch] BYREF
    //CPPEH_RECORD ms_exc; // [esp+34h] [ebp-18h]

    if (!SHRestricted(REST_NOSMPINNEDLIST))
    {  
        _psmpin->GetChangeCount(&ulPinChange);
        if (this->_ulPinChange != ulPinChange)
        {
            this->_ulPinChange = ulPinChange;
            _rtPinned.Reset();            
            if (_rtPinned._sl.Create(4) && SUCCEEDED(_psmpin->EnumObjects(&penum)))
            {
                while (!penum->Next(1, &pidl, 0))
                {
                    IShellLink *psl;
                    HRESULT hr;
                    CByUsageHiddenData hd;
                    
                    hr = SHGetUIObjectFromFullPIDL(pidl, 0, IID_PPV_ARGS(&psl));
                    if (hr >= 0)
                    {
                        hd.LoadFromShellLink(psl);
                        psl->Release();
                        hd.UpdateMSIPath();
                    }
                    else
                    {
                        hd._pwszTargetPath = _DisplayNameOf(this->_pdirDesktop->Folder(), pidl, 0x8000u);
                        if (hd._pwszTargetPath)
                        {
                            hr = 0;
                        }

                        if (FAILED(hr))
                        {
                            ASSERT(hd.IsClear()); // 2981
                        }
                    }

                    CByUsageShortcut* pscut = _pMenuCache->CreateShortcutFromHiddenData(this->_pdirDesktop, pidl, &hd, 1);
                    if (pscut)
                    {
                        if (_rtPinned._sl.AppendPtr(pscut) < 0)
                        {
                            delete pscut;
                        }
                        else
                        {
							pscut->SetInteresting(true);
                            if (IsSpecialPinnedPidl(pidl))
                            {
                                CByUsageAppInfo* papp = _pMenuCache->GetAppInfoFromSpecialPidl(_SHILMakeChild(pidl));
                                pscut->SetApp(papp);
                                if (papp)
                                {
                                    //CByUsageAppInfo::DecrementUsage(papp);
                                }
                            }
                        }
                    }
                    hd.Clear();
                    ILFree(pidl);
                }
                penum->Release();
            }
            //SHTracePerf(&ShellTraceId_StartMenu_PinItemToMenu_Stop);
        }
    }
#endif
}

IAssociationElement* GetAssociationElementFromSpecialPidl(IShellFolder *psf, LPCITEMIDLIST pidlItem)
{
    IAssociationElement *pae = NULL;

    // There is no way to get the IAssociationElement directly, so
    // we get the IExtractIcon and then ask him for the IAssociationElement.
    IExtractIcon *pxi;
    if (SUCCEEDED(psf->GetUIObjectOf(NULL, 1, &pidlItem, IID_IExtractIcon, nullptr, (void**)&pxi)))
    {
        IUnknown_QueryService(pxi, IID_IAssociationElement, IID_PPV_ARGS(&pae));
        pxi->Release();
    }

    return pae;
}

//
//  On success, the returned ByUsageAppInfo has been AddRef()d
//
CByUsageAppInfo *CMenuItemsCache::GetAppInfoFromSpecialPidl(LPCITEMIDLIST pidl)
{
#ifdef DEAD_CODE
    CByUsageAppInfo *papp = NULL;

    IAssociationElement *pae = GetAssociationElementFromSpecialPidl(_pdirDesktop->Folder(), pidl);
    if (pae)
    {
        LPWSTR pszData;
        if (SUCCEEDED(pae->QueryString(AQVS_APPLICATION_PATH, L"open", &pszData)))
        {
            //
            //  HACK!  Outlook puts the short file name in the registry.
            //  Convert to long file name (if it won't cost too much) so
            //  people who select Outlook as their default mail client
            //  won't get a dup copy in the MFU.
            //
            LPTSTR pszPath = pszData;
            TCHAR szLFN[MAX_PATH];
            if (!PathIsNetworkPath(pszData))
            {
                DWORD dwLen = GetLongPathName(pszData, szLFN, ARRAYSIZE(szLFN));
                if (dwLen && dwLen < ARRAYSIZE(szLFN))
                {
                    pszPath = szLFN;
                }
            }

            papp = GetAppInfo(pszPath, true);
            SHFree(pszData);
        }
        pae->Release();
    }
    return papp;
#else
    int v2; // ebx MAPDST
    CByUsageAppInfo *papp; // edi
    IAssociationElement *pae; // eax MAPDST
    LPWSTR pszPath; // edi
    DWORD dwLen; // eax
    IShellFolder *psf; // [esp+Ch] [ebp-214h] BYREF
    LPWSTR pszData; // [esp+10h] [ebp-210h] BYREF
    WCHAR szLFN[260]; // [esp+14h] [ebp-20Ch] BYREF

    papp = 0;
    if (SHGetDesktopFolder(&psf) >= 0)
    {
        pae = GetAssociationElementFromSpecialPidl(psf, pidl);
        if (pae)
        {
            if (SUCCEEDED(pae->QueryString(AQVS_APPLICATION_PATH, L"open", &pszData)))
            {
                pszPath = pszData;
                if (!PathIsNetworkPathW(pszData))
                {
                    dwLen = GetLongPathNameW(pszData, szLFN, 260u);
                    if (dwLen)
                    {
                        if (dwLen < 260)
                        {
                            pszPath = szLFN;
                        }
                    }
                }
                papp = CMenuItemsCache::GetAppInfo(pszPath, 1);
                CoTaskMemFree(pszData);
            }
            pae->Release();
        }
        psf->Release();
    }
    return papp;
#endif
}

void ByUsage::_EnumPinnedItemsFromCache()
{
    if (_rtPinned._sl)
    {
        int i;
        for (i = 0; i < _rtPinned._sl.GetPtrCount(); i++)
        {
            CByUsageShortcut *pscut = _rtPinned._sl.FastGetPtr(i);


            // Enumerating a pinned item consists of marking the corresponding
            // application as "I am pinned, do not mess with me!"

            CByUsageAppInfo *papp = pscut->App();

            if (papp)
            {
                papp->_fPinned = TRUE;

            }
        }
    }
}

//const struct CMenuItemsCache::ROOTFOLDERINFO CMenuItemsCache::c_rgrfi[] = {
//    { CSIDL_STARTMENU,               ENUMFL_RECURSE | ENUMFL_CHECKNEW | ENUMFL_ISSTARTMENU },
//    { CSIDL_PROGRAMS,                ENUMFL_RECURSE | ENUMFL_CHECKNEW | ENUMFL_CHECKISCHILDOFPREVIOUS },
//    { CSIDL_COMMON_STARTMENU,        ENUMFL_RECURSE | ENUMFL_CHECKNEW | ENUMFL_ISSTARTMENU },
//    { CSIDL_COMMON_PROGRAMS,         ENUMFL_RECURSE | ENUMFL_CHECKNEW | ENUMFL_CHECKISCHILDOFPREVIOUS },
//    { CSIDL_DESKTOPDIRECTORY,        ENUMFL_NORECURSE | ENUMFL_NOCHECKNEW },
//    { CSIDL_COMMON_DESKTOPDIRECTORY, ENUMFL_NORECURSE | ENUMFL_NOCHECKNEW },  // The limit for register notify is currently 5 (slots 0 through 4)    
//                                                                            // Changing this requires changing ByUsageUI::SFTHOST_MAXNOTIFY
//};

const struct CMenuItemsCache::ROOTFOLDERINFO CMenuItemsCache::c_rgrfi[] =
{
    { FOLDERID_StartMenu,           0x0,        8u },
    { FOLDERID_Programs,            0x0,        4u },
    { FOLDERID_CommonStartMenu,     0x0,        8u },
    { FOLDERID_CommonPrograms,      0x0,        4u },
    { FOLDERID_Desktop,             0x1000,     3u },
    { FOLDERID_PublicDesktop,       0x1000,     3u },
    { FOLDERID_GameTasks,           0x0,        2u }
};

BOOL CMenuItemsCache::InitDesktopFolder()
{
    if (_pdirDesktop)
        return 0;
    
    _pdirDesktop = CByUsageDir::CreateDesktop();
    return _pdirDesktop != 0;
}

//
//  Here's where we decide all the things that should be enumerated
//  in the "My Programs" list.
//
void ByUsage::EnumItems()
{
#ifdef DEAD_CODE
    _FillPinnedItemsCache();
    _NotifyDesiredSize();


    _pMenuCache->LockPopup();
    _pMenuCache->InitCache();

    BOOL fNeedUpdateDarwin = !_pMenuCache->IsCacheUpToDate();

    // Note!  UpdateCache() must occur before _EnumPinnedItemsFromCache()
    // because UpdateCache() resets _fPinned.
    _pMenuCache->UpdateCache();

    if (fNeedUpdateDarwin)
    {
        SHReValidateDarwinCache();
    }

    _pMenuCache->RefreshDarwinShortcuts(&_rtPinned);
    _EnumPinnedItemsFromCache();
    EnumFolderFromCache();

    // Finished collecting data; do some postprocessing...
    AfterEnumItems();

    // Do not unlock before this point, as AfterEnumItems depends on the cache to stay put.
    _pMenuCache->UnlockPopup();
#else
    //SHTracePerf(&ShellTraceId_StartMenu_Fill_MenuCache_Start);
    
    _NotifyDesiredSize();

    if (_pMenuCache->LockPopup() < 0)
        return;
    
    if (_pMenuCache->InitDesktopFolder())
    {
        _pMenuCache->InitCache();

        BOOL fNeedUpdateDarwin = _pMenuCache->IsCacheUpToDate();

        _pMenuCache->UpdateCache();
        
        if (!fNeedUpdateDarwin)
            SHReValidateDarwinCache();

        _pMenuCache->RefreshDarwinShortcuts(&this->_rtPinned);
        _EnumPinnedItemsFromCache();
        EnumFolderFromCache();
        AfterEnumItems();
        IUnknown_SafeReleaseAndNullPtr(&this->_pMenuCache->_pdirDesktop);
    }
    _pMenuCache->UnlockPopup();
    //SHTracePerf(&ShellTraceId_StartMenu_Fill_MenuCache_Stop);
#endif
}

void ByUsage::_NotifyDesiredSize()
{
    if (_pByUsageUI)
    {
        int cPinned = _rtPinned._sl ? _rtPinned._sl.GetPtrCount() : 0;

        int cNormal;
        DWORD cb = sizeof(cNormal);
        if (SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTPANE_SETTINGS,
                       REGSTR_VAL_DV2_MINMFU, NULL, &cNormal, &cb) != ERROR_SUCCESS)
        {
            cNormal = REGSTR_VAL_DV2_MINMFU_DEFAULT;
        }

        _cMFUDesired = cNormal;
        _pByUsageUI->SetDesiredSize(cPinned, cNormal);
    }
}


//****************************************************************************
// CMenuItemsCache

CMenuItemsCache::CMenuItemsCache()
    : _cref(1)
{
    CByUsageRoot *rgrt; // eax

    int i; // edx
    rgrt = this->_rgrt;
    for (i = 6; i >= 0; --i)
    {
        rgrt->_sl.Detach();
        rgrt->_slOld.Detach();
        ++rgrt;
    }

    InitializeCriticalSection(&_csInUse);
}

LONG CMenuItemsCache::AddRef()
{
    return InterlockedIncrement(&_cref);
}

LONG CMenuItemsCache::Release()
{
    ASSERT( 0 != _cref );
    LONG cRef = InterlockedDecrement(&_cref);
    if ( 0 == cRef )
    {
        delete this;
    }
    return cRef;
}

HRESULT CMenuItemsCache::Initialize(ByUsageUI *pbuUI, FILETIME * pftOSInstall)
{
#ifdef DEAD_CODE
    HRESULT hr = S_OK;

    // Must do this before any of the operations that can fail
    // because we unconditionally call DeleteCriticalSection in destructor

    //_fCSInited = InitializeCriticalSectionAndSpinCount(&_csInUse, 0);

    //if (!_fCSInited)
    //{
    //    return E_OUTOFMEMORY;
    //}

    hr = AssocCreate(CLSID_QueryAssociations, IID_PPV_ARGS(&_pqa));
    if (FAILED(hr))
    {
        return hr;
    }

    _pByUsageUI = pbuUI;

    _ftOldApps = *pftOSInstall;

    _pdirDesktop = CByUsageDir::CreateDesktop();
    
    if (!_dpaAppInfo.Create(4))
    {
        hr = E_OUTOFMEMORY;
    }

    if (!_GetExcludedDirectories())
    {
        hr = E_OUTOFMEMORY;
    }

    if (!_dpaKill.Create(4) ||
        !_dpaKillLink.Create(4)) {
        return E_OUTOFMEMORY;
    }

    _InitKillList();

    _hPopupReady = CreateMutex(NULL, FALSE, NULL);
    if (!_hPopupReady)
    {
        return E_OUTOFMEMORY;
    }

    // By default, we want to check applications for newness.
    _fCheckNew = TRUE;

    return hr;
#else
    HRESULT hr; // ebx
    // eax

    this->_pByUsageUI = pbuUI;
    this->_ftOldApps = *pftOSInstall;
    hr = 0x8007000E;
    if (_dpaAppInfo.Create(4))
    {
        if (_dpaKill.Create(4))
        {
            if (_dpaKillLink.Create(4))
            {
                //if (CDPA_Base<IUnknown>::Create(4))
                {
                    if (CMenuItemsCache::_GetExcludedDirectories())
                    {
                        CMenuItemsCache::_InitKillList();
                        //CMenuItemsCache::_InitNoKillList(this);
                        this->_hPopupReady = CreateMutexW(0, 0, 0);
                        if (_hPopupReady)
                        {
                            this->_fCheckNew = 1;
                            return 0;
                        }
                    }
                }
            }
        }
    }
    return hr;
#endif
}
HRESULT CMenuItemsCache::AttachUI(ByUsageUI *pbuUI)
{
    // We do not AddRef here so that destruction always happens on the same thread that created the object
    // but beware of lifetime issues: we need to synchronize attachUI/detachUI operations with LockPopup and UnlockPopup.

    if (SUCCEEDED(LockPopup()))
    {
        _pByUsageUI = pbuUI;
        UnlockPopup();
    }
    return S_OK;
}

CMenuItemsCache::~CMenuItemsCache()
{
    if (_fIsCacheUpToDate)
    {
        _SaveCache();
    }


    _dpaNotInteresting.DestroyCallback(DPA_LocalFreeCB, NULL);
    _dpaKill.DestroyCallback(DPA_LocalFreeCB, NULL);
    _dpaKillLink.DestroyCallback(DPA_LocalFreeCB, NULL);


    // Must delete the roots before destroying _dpaAppInfo.
    int i;
    for (i = 0; i < ARRAYSIZE(_rgrt); i++)
    {
        _rgrt[i].Reset();
    }

    IUnknown_SafeReleaseAndNullPtr(&_pqa);
    _dpaAppInfo.DestroyCallback(DPA_DeleteCB, NULL);

    if (_pdirDesktop)
    {
        _pdirDesktop->Release();
    }

    if (_hPopupReady)
    {
        CloseHandle(_hPopupReady);
    }

    DeleteCriticalSection(&_csInUse);
}


BOOL CMenuItemsCache::_ShouldProcessRoot(int iRoot)
{
    BOOL fRet = TRUE;

    if (!_rgrt[iRoot]._pidl)
    {
        fRet = FALSE;
    }
    else if ((c_rgrfi[iRoot]._enumfl & ENUMFL_CHECKISCHILDOFPREVIOUS) && !SHRestricted(REST_NOSTARTMENUSUBFOLDERS) )
    {
        ASSERT(iRoot >= 1);
        if (_rgrt[iRoot-1]._pidl && ILIsParent(_rgrt[iRoot-1]._pidl, _rgrt[iRoot]._pidl, FALSE))
        {
            fRet = FALSE;
        }
    }
    return fRet;
}

//****************************************************************************
//
//  The format of the ProgramsCache is as follows:
//
//  [DWORD] dwVersion
//
//      If the version is wrong, then ignore.  Not worth trying to design
//      a persistence format that is forward-compatible since it's just
//      a cache.
//
//      Don't be stingy about incrementing the dwVersion.  We've got room
//      for four billion revs.

#define PROGLIST_VERSION 9

//
//
//  For each special folder we persist:
//
//      [BYTE] CSIDL_xxx (as a sanity check)
//
//      Followed by a sequence of segments; either...
//
//          [BYTE] 0x00 -- Change directory
//          [pidl] directory (relative to CSIDL_xxx)
//
//      or
//
//          [BYTE] 0x01 -- Add shortcut
//          [pidl] item (relative to current directory)
//
//      or
//
//          [BYTE] 0x02 -- end
//

ITEMID_CHILD *_ILMakeAlignedChild(const ITEMIDLIST_RELATIVE *pidl)
{
    //if (((unsigned __int8)pidl & 3) != 0
    //    && CcshellRipW(L"d:\\longhorn\\shell\\explorer\\desktop2\\proglist.cpp", 955, L"ILIsAligned(pidl)", 0))
    //{
    //    AttachUserModeDebugger();
    //    do
    //        __debugbreak();
    //    while (`_ILMakeAlignedChild'::`7': : gAlwaysAssert);
    //}
    return (ITEMID_CHILD *)pidl;
}

#define CACHE_CHDIR     0
#define CACHE_ITEM      1
#define CACHE_END       2

BOOL CMenuItemsCache::InitCache()
{
#ifdef DEAD_CODE
    COMPILETIME_ASSERT(ARRAYSIZE(c_rgrfi) == NUM_PROGLIST_ROOTS);

    // Make sure we don't use more than MAXNOTIFY notify slots for the cache
    COMPILETIME_ASSERT(NUM_PROGLIST_ROOTS <= MAXNOTIFY);

    if (_fIsInited)
        return TRUE;

    BOOL fSuccess = FALSE;
    int irfi;

    IStream *pstm = SHOpenRegStream2(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS, REGSTR_VAL_PROGLIST, STGM_READ);
    if (pstm)
    {
        CByUsageDir *pdirRoot = NULL;
        CByUsageDir *pdir = NULL;

        DWORD dwVersion;
        if (FAILED(IStream_Read(pstm, &dwVersion, sizeof(dwVersion))) ||
            dwVersion != PROGLIST_VERSION)
        {
            goto panic;
        }

        for (irfi = 0; irfi < ARRAYSIZE(c_rgrfi); irfi++)
        {
            CByUsageRoot *prt = &_rgrt[irfi];

            // If SHGetSpecialFolderLocation fails, it could mean that
            // the directory was recently restricted.  We *could* just
            // skip over this block and go to the next csidl, but that
            // would be actual work, and this is just a cache, so we may
            // as well just panic and re-enumerate from scratch.
            //
            if (FAILED(SHGetSpecialFolderLocation(NULL, c_rgrfi[irfi]._csidl, &prt->_pidl)))
            {
                goto panic;
            }

            if (!_ShouldProcessRoot(irfi))
                continue;

            if (!prt->_sl.Create(4))
            {
                goto panic;
            }

            BYTE csidl;
            if (FAILED(IStream_Read(pstm, &csidl, sizeof(csidl))) ||
                csidl != c_rgrfi[irfi]._csidl)
            {
                goto panic;
            }

            pdirRoot = CByUsageDir::Create(_pdirDesktop, prt->_pidl);

            if (!pdirRoot)
            {
                goto panic;
            }


            BYTE bCmd;
            do
            {
                LPITEMIDLIST pidl;

                if (FAILED(IStream_Read(pstm, &bCmd, sizeof(bCmd))))
                {
                    goto panic;
                }

                switch (bCmd)
                {
                    case CACHE_CHDIR:
                        // Toss the old directory
                        if (pdir)
                        {
                            pdir->Release();
                            pdir = NULL;
                        }

                        // Figure out where the new directory is
                        if (FAILED(IStream_ReadPidl(pstm, &pidl)))
                        {
                            goto panic;
                        }

                        // and create it
                        pdir = CByUsageDir::Create(pdirRoot, pidl);
                        ILFree(pidl);

                        if (!pdir)
                        {
                            goto panic;
                        }
                        break;

                    case CACHE_ITEM:
                    {
                        // Must set a directory befor creating an item
                        if (!pdir)
                        {
                            goto panic;
                        }

                        // Get the new item
                        if (FAILED(IStream_ReadPidl(pstm, &pidl)))
                        {
                            goto panic;
                        }

                        // Create it
                        CByUsageShortcut *pscut = _CreateFromCachedPidl(prt, pdir, pidl);
                        ILFree(pidl);
                        if (!pscut)
                        {
                            goto panic;
                        }
                    }
                    break;

                    case CACHE_END:
                        break;

                    default:
                        goto panic;
                }
            } while (bCmd != CACHE_END);

            pdirRoot->Release();
            pdirRoot = NULL;
            if (pdir)
            {
                pdir->Release();
                pdir = NULL;
            }

            prt->SetNeedRefresh();
        }

        fSuccess = TRUE;

    panic:
        if (!fSuccess)
        {
            for (irfi = 0; irfi < ARRAYSIZE(c_rgrfi); irfi++)
            {
                _rgrt[irfi].Reset();
            }
        }

        if (pdirRoot)
        {
            pdirRoot->Release();
        }

        if (pdir)
        {
            pdir->Release();
        }

        pstm->Release();
    }

    _fIsInited = TRUE;

    return fSuccess;
#else
    CMenuItemsCache *v1; // esi
    bool v2; // zf
    int result; // eax
    IStream *v4; // eax
    const CMenuItemsCache::ROOTFOLDERINFO *v5; // edi
    LPITEMIDLIST *p_pidl; // ebx
    struct CByUsageRoot *v7; // esi
    int v8; // edi
    const ITEMIDLIST *v9; // eax
    DWORD pv; // [esp+8h] [ebp-44h] BYREF
    const CMenuItemsCache::ROOTFOLDERINFO *v11; // [esp+Ch] [ebp-40h]
    int v12; // [esp+10h] [ebp-3Ch]
    int iRoot; // [esp+14h] [ebp-38h]
    unsigned int v14; // [esp+18h] [ebp-34h]
    IStream *pstm; // [esp+1Ch] [ebp-30h]
    CByUsageDir *pdir; // [esp+20h] [ebp-2Ch]
    LPITEMIDLIST ppidlOut; // [esp+24h] [ebp-28h] BYREF
    CMenuItemsCache *v18; // [esp+28h] [ebp-24h]
    struct CByUsageRoot *rgrt; // [esp+2Ch] [ebp-20h]
    CByUsageDir *v20; // [esp+30h] [ebp-1Ch]
    BYTE v21; // [esp+37h] [ebp-15h] BYREF
    KNOWNFOLDERID v22; // [esp+38h] [ebp-14h] BYREF

    v1 = this;
    v2 = this->_fIsInited == 0;
    v18 = this;
    if (!v2)
        return 1;
    v12 = 0;
    v4 = SHOpenRegStream2W(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartPage",
        L"ProgramsCache",
        0);
    pstm = v4;
    if (v4)
    {
        pdir = 0;
        v20 = 0;
        if (IStream_Read(v4, &pv, 4u) >= 0 && pv == PROGLIST_VERSION)
        {
            v5 = CMenuItemsCache::c_rgrfi;
            iRoot = 0;
            v11 = CMenuItemsCache::c_rgrfi;
            rgrt = v1->_rgrt;
            v14 = 0;
            while (1)
            {
                p_pidl = &rgrt->_pidl;
                if (SHGetKnownFolderIDList(v5->_kfid, v5->_dwFlags, 0, &rgrt->_pidl) < 0)
                    break;
                if (CMenuItemsCache::_ShouldProcessRoot(iRoot))
                {
                    if (rgrt->_sl.Create(4))
                    {
                        if (IStream_Read(pstm, &v22, 16u) >= 0 && !memcmp(&v22, v5, 0x10u))
                        {
                            pdir = CByUsageDir::Create(v18->_pdirDesktop, *p_pidl);
                            if (pdir)
                            {
                                while (IStream_Read(pstm, &v21, 1u) >= 0)
                                {
                                    if (v21)
                                    {
                                        if (v21 == 1)
                                        {
                                            if (!v20 || IStream_ReadPidl(pstm, &ppidlOut) < 0)
                                                goto LABEL_17;
                                            CMenuItemsCache::_CreateFromCachedPidl(rgrt, v20, _ILMakeAlignedChild(ppidlOut));
                                            ILFree(ppidlOut);
                                        }
                                        else if (v21 != 2)
                                        {
                                            goto LABEL_17;
                                        }
                                    }
                                    else
                                    {
                                        if (v20)
                                        {
                                            v20->Release();
                                            v20 = 0;
                                        }
                                        if (IStream_ReadPidl(pstm, &ppidlOut) < 0)
                                            goto LABEL_17;
                                        v20 = CByUsageDir::Create(pdir, ppidlOut);
                                        ILFree(ppidlOut);
                                        if (!v20)
                                            goto LABEL_17;
                                    }
                                    if (v21 == 2)
                                    {
                                        pdir->Release();
                                        pdir = 0;
                                        if (v20)
                                        {
                                            v20->Release();
                                            v20 = 0;
                                        }
                                        rgrt->_fNeedRefresh = 1;
                                        goto LABEL_36;
                                    }
                                }
                            }
                        }
                    }
                    break;
                }
            LABEL_36:
                v14 += 24;
                ++iRoot;
                ++rgrt;
                v5 = ++v11;
                if (v14 >= 0xA8)
                {
                    v12 = 1;
                    goto LABEL_21;
                }
            }
        }
    LABEL_17:
        v7 = v18->_rgrt;
        v8 = 7;
        do
        {
            v7++->Reset();
            //CByUsageRoot::Reset(v7++);
            --v8;
        } while (v8);

        if (pdir)
            pdir->Release();
    LABEL_21:
        if (v20)
            v20->Release();
        pstm->Release();
        v1 = v18;
    }
    result = v12;
    v1->_fIsInited = 1;
    return result;
#endif
}

GUID POLID_NoStartMenuSubFolders =
{
  2019720721u,
  60544u,
  18851u,
  { 179u, 183u, 106u, 190u, 115u, 150u, 147u, 252u }
};

HRESULT CMenuItemsCache::UpdateCache()
{
#ifdef DEAD_CODE
    FILETIME ft;
    // Apps are "new" only if installed less than 1 week ago.
    // They also must postdate the user's first use of the new Start Menu.
    GetSystemTimeAsFileTime(&ft);
    DecrementFILETIME(&ft, FT_ONEDAY * 7);

    // _ftOldApps is the more recent of OS install time, or last week.
    if (CompareFileTime(&ft, &_ftOldApps) >= 0)
    {
        _ftOldApps = ft;
    }

    _dpaAppInfo.EnumCallbackEx(CByUsageAppInfo::EnumResetCB, this);

    if(!SHRestricted(REST_NOSMMFUPROGRAMS))
    {
        int i;
        for (i = 0; i < ARRAYSIZE(c_rgrfi); i++)
        {
            CByUsageRoot *prt = &_rgrt[i];
            int csidl = c_rgrfi[i]._csidl;
            _enumfl = c_rgrfi[i]._enumfl;

            if (!prt->_pidl)
            {
                (void)SHGetSpecialFolderLocation(NULL, csidl, &prt->_pidl);     // void cast to keep prefast happy
                prt->SetNeedRefresh();
            }

            if (!_ShouldProcessRoot(i))
                continue;


            // Restrictions might deny recursing into subfolders
            if ((_enumfl & ENUMFL_ISSTARTMENU) && SHRestricted(REST_NOSTARTMENUSUBFOLDERS))
            {
                _enumfl &= ~ENUMFL_RECURSE;
                _enumfl |= ENUMFL_NORECURSE;
            }

            // Fill the cache if it is stale

            LPITEMIDLIST pidl;
            if (!IsRestrictedCsidl(csidl) &&
                SUCCEEDED(SHGetSpecialFolderLocation(NULL, csidl, &pidl)))
            {
                if (prt->_pidl == NULL || !ILIsEqual(prt->_pidl, pidl) ||
                    prt->NeedsRefresh() || prt->NeedsRegister())
                {
                    if (!prt->_pidl || prt->NeedsRefresh())
                    {
                        prt->ClearNeedRefresh();
                        ASSERT(prt->_slOld == NULL);
                        prt->_slOld = prt->_sl;
                        prt->_cOld = prt->_slOld ? prt->_slOld.GetPtrCount() : 0;
                        prt->_iOld = 0;

                        // Free previous pidl
                        ILFree(prt->_pidl);
                        prt->_pidl = NULL;

                        if (prt->_sl.Create(4))
                        {
                            CByUsageDir *pdir = CByUsageDir::Create(_pdirDesktop, pidl);
                            if (pdir)
                            {
                                prt->_pidl = pidl;  // Take ownership
                                pidl = NULL;        // So ILFree won't nuke it
                                _FillFolderCache(pdir, prt);
                                pdir->Release();
                            }
                        }
                        DPADELETEANDDESTROY(prt->_slOld);
                    }
                    if (_pByUsageUI && prt->NeedsRegister() && prt->_pidl)
                    {
                        ASSERT(i < ByUsageUI::SFTHOST_MAXNOTIFY);
                        prt->SetRegistered();
                        ASSERT(!IsLocked());
                        _pByUsageUI->RegisterNotify(i, SHCNE_DISKEVENTS, prt->_pidl, TRUE);
                    }
                }
                ILFree(pidl);

            }
            else
            {
                // Special folder doesn't exist; erase the file list
                prt->Reset();
            }
        } // for loop!

    } // Restriction!
    _fIsCacheUpToDate = TRUE;
    return S_OK;
#else
    struct CByUsageRoot *prt; // ebx MAPDST
    DWORD v3; // ecx
    LPITEMIDLIST *p_pidl; // esi
    struct _DPA *m_hdpa; // ecx
    struct CByUsageShortcutList *p_slOld; // edi
    int cp; // eax
    struct CByUsageDir *pdir; // eax
    struct _DPA *v10; // eax
    DWORD dwFlags; // [esp+18h] [ebp-40h] SPLIT
    LPITEMIDLIST pidl; // [esp+20h] [ebp-38h] BYREF
    unsigned int i; // [esp+24h] [ebp-34h]
    KNOWNFOLDERID kfid; // [esp+2Ch] [ebp-2Ch] BYREF
    //CPPEH_RECORD ms_exc; // [esp+40h] [ebp-18h]

    FILETIME ft; // [esp+10h] [ebp-48h] BYREF
    GetSystemTimeAsFileTime(&ft);
    //SystemTimeAsFileTime = (FILETIME)(_FILETIMEtoInt64(&SystemTimeAsFileTime) - 6048000000000LL);
	DecrementFILETIME(&ft, 6048000000000LL);
    if (CompareFileTime(&ft, &this->_ftOldApps) >= 0)
        this->_ftOldApps = ft;

    _dpaAppInfo.EnumCallbackEx(CByUsageAppInfo::EnumResetCB, this);

    if (!SHRestricted(REST_NOSMMFUPROGRAMS))
    {
        for (i = 0; i < 7; ++i)
        {
            prt = &this->_rgrt[i];
            kfid = CMenuItemsCache::c_rgrfi[i]._kfid;
            v3 = CMenuItemsCache::c_rgrfi[i]._dwFlags;
            dwFlags = v3;
            this->_enumfl = CMenuItemsCache::c_rgrfi[i]._enumfl;
            p_pidl = &prt->_pidl;
            if (!prt->_pidl)
            {
                SHGetKnownFolderIDList(kfid, v3, 0, &prt->_pidl);
                prt->_fNeedRefresh = 1;
            }
            if (CMenuItemsCache::_ShouldProcessRoot(i))
            {
                if ((this->_enumfl & 8) != 0 && SHWindowsPolicy(POLID_NoStartMenuSubFolders))
                    this->_enumfl |= 1u;
                if (CMenuItemsCache::IsRestrictedKfid(kfid) || SHGetKnownFolderIDList(kfid, dwFlags, 0, &pidl) < 0)
                {
                    prt->Reset();
                }
                else
                {
                    if (!*p_pidl || !ILIsEqual(*p_pidl, pidl) || prt->_fNeedRefresh || !prt->_fRegistered)
                    {
                        if (!*p_pidl || prt->_fNeedRefresh)
                        {
                            prt->_fNeedRefresh = 0;

                            //if (prt->_slOld.m_hdpa
                            //    && CcshellAssertFailedW(
                            //        L"d:\\longhorn\\shell\\explorer\\desktop2\\proglist.cpp",
                            //        3678,
                            //        L"prt->_slOld == NULL",
                            //        0))
                            //{
                            //    AttachUserModeDebugger();
                            //    do
                            //    {
                            //        __debugbreak();
                            //        ms_exc.registration.TryLevel = -2;
                            //    } while (dword_108B8D0);
                            //}

                            //m_hdpa = prt->_sl.m_hdpa;
                            //prt->_sl.m_hdpa = 0;
                            //p_slOld = &prt->_slOld;
                            //CDPA_Base<CByUsageShortcut>::Attach((CDPA_Base<class CByUsageShortcut> *) & prt->_slOld, m_hdpa);
                            prt->_slOld.Attach(prt->_sl.Detach());

                            //if (prt->_slOld.m_hdpa)
                            //    cp = p_slOld->m_hdpa->cp;
                            //else
                            //    cp = 0;

                            prt->_cOld = prt->_slOld ? prt->_slOld.GetPtrCount() : 0;
                            prt->_iOld = 0;
                            p_pidl = &prt->_pidl;
                            ILFree(prt->_pidl);
                            prt->_pidl = 0;
                            if (prt->_sl.Create(4))// prt->_sl.Create(4)
                            {
                                pdir = CByUsageDir::Create(this->_pdirDesktop, pidl);
                                dwFlags = (DWORD)pdir;
                                if (pdir)
                                {
                                    *p_pidl = pidl;
                                    pidl = 0;
                                    CMenuItemsCache::_FillFolderCache(pdir, prt);// _FillFolderCache(pdir, prt);
									pdir->Release();
                                }
                            }
                            else
                            {
                                //v10 = p_slOld->m_hdpa;
                                //p_slOld->m_hdpa = 0;
                                //CDPA_Base<CByUsageShortcut>::Attach((CDPA_Base<class CByUsageShortcut> *)prt, v10);
								prt->_sl.Attach(prt->_slOld.Detach());
                            }
							prt->_slOld.DestroyCallback(DPA_DeleteCB<CByUsageShortcut>, 0);
                            //CDPA<unsigned short>::DestroyCallback(DPA_DeleteCB<CByUsageShortcut>, 0);
                        }

                        if (this->_pByUsageUI && !prt->_fRegistered && *p_pidl)
                        {
                            //if ((int)i >= 9
                            //    && CcshellAssertFailedW(
                            //        L"d:\\longhorn\\shell\\explorer\\desktop2\\proglist.cpp",
                            //        3707,
                            //        L"i < ByUsageUI::SFTHOST_MAXNOTIFY",
                            //        0))
                            //{
                            //    AttachUserModeDebugger();
                            //    do
                            //    {
                            //        __debugbreak();
                            //        ms_exc.registration.TryLevel = -2;
                            //    } while (dword_108B8CC);
                            //}

                            prt->_fRegistered = 1;

                            //if (CMenuItemsCache::IsLocked(this)
                            //    && CcshellAssertFailedW(
                            //        L"d:\\longhorn\\shell\\explorer\\desktop2\\proglist.cpp",
                            //        3709,
                            //        L"!IsLocked()",
                            //        0))
                            //{
                            //    AttachUserModeDebugger();
                            //    do
                            //    {
                            //        __debugbreak();
                            //        ms_exc.registration.TryLevel = -2;
                            //    } while (dword_108B8C8);
                            //}

                            _pByUsageUI->RegisterNotify(
                                i,
                                SHCNE_DISKEVENTS,
                                prt->_pidl,
                                1);
                        }
                    }
                    ILFree(pidl);
                }
            }
        }
    }
    this->_fIsCacheUpToDate = 1;
    return 0;
#endif
}

void CMenuItemsCache::RefreshDarwinShortcuts(CByUsageRoot *prt)
{
    if (prt->_sl)
    {
        int j = prt->_sl.GetPtrCount();
        while (--j >= 0)
        {
            CByUsageShortcut *pscut = prt->_sl.FastGetPtr(j);
            if (FAILED(_UpdateMSIPath(pscut)))
            {
                prt->_sl.DeletePtr(j);  // remove the bad shortcut so we don't fault
            }
        }
    }
}

void CMenuItemsCache::RefreshCachedDarwinShortcuts()
{
    if(!SHRestricted(REST_NOSMMFUPROGRAMS))
    {
        Lock();

        for (int i = 0; i < ARRAYSIZE(c_rgrfi); i++)
        {
            RefreshDarwinShortcuts(&_rgrt[i]);
        }
        Unlock();
    }
}


CByUsageShortcut *CMenuItemsCache::_CreateFromCachedPidl(CByUsageRoot *prt, CByUsageDir *pdir, LPITEMIDLIST pidl)
{
    CByUsageHiddenData hd;
    UINT buhd = CByUsageHiddenData::BUHD_TARGETPATH | CByUsageHiddenData::BUHD_MSIPATH;
    hd.Get(pidl, buhd);

    CByUsageShortcut *pscut = CreateShortcutFromHiddenData(pdir, pidl, &hd);
    if (pscut)
    {
        if (prt->_sl.AppendPtr(pscut) >= 0)
        {
            _SetInterestingLink(pscut);
        }
        else
        {
            // Couldn't append; oh well
            delete pscut;       // "delete" can handle NULL pointer
        }
    }

    hd.Clear();

    return pscut;
}

HRESULT IStream_WriteByte(IStream *pstm, BYTE b)
{
    return IStream_Write(pstm, &b, sizeof(b));
}

#ifdef DEBUG
//
//  Like ILIsParent, but defaults to TRUE if we don't have enough memory
//  to determine for sure.  (ILIsParent defaults to FALSE on error.)
//
BOOL ILIsProbablyParent(LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidlChild)
{
    BOOL fRc = TRUE;
    LPITEMIDLIST pidlT = ILClone(pidlChild);
    if (pidlT)
    {

        // Truncate pidlT to the same depth as pidlParent.
        LPCITEMIDLIST pidlParentT = pidlParent;
        LPITEMIDLIST pidlChildT = pidlT;
        while (!ILIsEmpty(pidlParentT))
        {
            pidlChildT = _ILNext(pidlChildT);
            pidlParentT = _ILNext(pidlParentT);
        }

        pidlChildT->mkid.cb = 0;

        // Okay, at this point pidlT should equal pidlParent.
        IShellFolder *psfDesktop;
        if (SUCCEEDED(SHGetDesktopFolder(&psfDesktop)))
        {
            HRESULT hr = psfDesktop->CompareIDs(0, pidlT, pidlParent);
            if (SUCCEEDED(hr) && ShortFromResult(hr) != 0)
            {
                // Definitely, conclusively different.
                fRc = FALSE;
            }
            psfDesktop->Release();
        }

        ILFree(pidlT);
    }
    return fRc;
}
#endif

inline LPITEMIDLIST ILFindKnownChild(LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidlChild)
{
#ifdef DEBUG
    // ILIsParent will give wrong answers in low-memory situations
    // (which testers like to simulate) so we roll our own.
    // ASSERT(ILIsParent(pidlParent, pidlChild, FALSE));
    ASSERT(ILIsProbablyParent(pidlParent, pidlChild));
#endif

    while (!ILIsEmpty(pidlParent))
    {
        pidlChild = _ILNext(pidlChild);
        pidlParent = _ILNext(pidlParent);
    }
    return const_cast<LPITEMIDLIST>(pidlChild);
}

void CMenuItemsCache::_SaveCache()
{
    int irfi;
    BOOL fSuccess = FALSE;

    IStream *pstm = SHOpenRegStream2(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS, REGSTR_VAL_PROGLIST, STGM_WRITE);
    if (pstm)
    {
        DWORD dwVersion = PROGLIST_VERSION;
        if (FAILED(IStream_Write(pstm, &dwVersion, sizeof(dwVersion))))
        {
            goto panic;
        }

        for (irfi = 0; irfi < ARRAYSIZE(c_rgrfi); irfi++)
        {
            if (!_ShouldProcessRoot(irfi))
                continue;

            CByUsageRoot *prt = &_rgrt[irfi];

            if (FAILED(IStream_Write(pstm, c_rgrfi, 0x10)))
            {
                goto panic;
            }

            if (prt->_sl && prt->_pidl)
            {
                int i;
                CByUsageDir *pdir = NULL;
                for (i = 0; i < prt->_sl.GetPtrCount(); i++)
                {
                    CByUsageShortcut *pscut = prt->_sl.FastGetPtr(i);

                    // If the directory changed, write out a chdir entry
                    if (pdir != pscut->Dir())
                    {
                        pdir = pscut->Dir();

                        // Write the new directory
                        if (FAILED(IStream_WriteByte(pstm, CACHE_CHDIR)) ||
                            FAILED(IStream_WritePidl(pstm, ILFindKnownChild(prt->_pidl, pdir->Pidl()))))
                        {
                            goto panic;
                        }
                    }

                    // Now write out the shortcut
                    if (FAILED(IStream_WriteByte(pstm, CACHE_ITEM)) ||
                        FAILED(IStream_WritePidl(pstm, pscut->RelativePidl())))
                    {
                        goto panic;
                    }
                }
            }

            // Now write out the terminator
            if (FAILED(IStream_WriteByte(pstm, CACHE_END)))
            {
                goto panic;
            }

        }

        fSuccess = TRUE;

    panic:
        pstm->Release();

        if (!fSuccess)
        {
            SHDeleteValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS, REGSTR_VAL_PROGLIST);
        }
    }
}


void CMenuItemsCache::StartEnum()
{
    _iCurrentRoot = 0;
    _iCurrentIndex = 0;
}

void CMenuItemsCache::EndEnum()
{
}

CByUsageShortcut *CMenuItemsCache::GetNextShortcut()
{
#ifdef DEAD_CODE
    CByUsageShortcut *pscut = NULL;

    if (_iCurrentRoot < NUM_PROGLIST_ROOTS)
    {
        if (_rgrt[_iCurrentRoot]._sl && _iCurrentIndex < _rgrt[_iCurrentRoot]._sl.GetPtrCount())
        {
            pscut = _rgrt[_iCurrentRoot]._sl.FastGetPtr(_iCurrentIndex);
            _iCurrentIndex++;
        }
        else
        {
            // Go to next root
            _iCurrentIndex = 0;
            _iCurrentRoot++;
            pscut = GetNextShortcut();
        }
    }

    return pscut;
#else
    int iCurrentRoot; // esi
    CByUsageShortcut *pscut; // eax
    int iCurrentIndex; // edx

    while (1)
    {
        iCurrentRoot = this->_iCurrentRoot;
        pscut = 0;
        if (iCurrentRoot >= 7)
            break;

        if (this->_rgrt[iCurrentRoot]._sl)
        {
            iCurrentIndex = this->_iCurrentIndex;
            if (iCurrentIndex < this->_rgrt[iCurrentRoot]._sl.GetPtrCount())
            {
				pscut = this->_rgrt[iCurrentRoot]._sl.FastGetPtr(iCurrentIndex);
                this->_iCurrentIndex = iCurrentIndex + 1;
                return pscut;
            }
        }
        this->_iCurrentIndex = 0;
        this->_iCurrentRoot = iCurrentRoot + 1;
    }
    return pscut;
#endif
}

//****************************************************************************

void AppendString(CDPA<TCHAR, CTContainer_PolicyUnOwned<TCHAR>>* dpa, LPCTSTR psz)
{
    LPTSTR pszDup = StrDup(psz);
    if (pszDup && dpa->AppendPtr(pszDup) < 0)
    {
        LocalFree(pszDup);  // Append failed
    }
}

BOOL LocalFreeCallback(LPTSTR psz, LPVOID)
{
    LocalFree(psz);
    return TRUE;
}

BOOL ILFreeCallback(LPITEMIDLIST pidl, LPVOID)
{
    ILFree(pidl);
    return TRUE;
}

int ByUsage::CompareItems(PaneItem *p1, PaneItem *p2)
{
    //
    //  The separator comes before regular items.
    //
    if (p1->IsSeparator())
    {
        return -1;
    }

    if (p2->IsSeparator())
    {
        return +1;
    }

    CByUsageItem *pitem1 = static_cast<CByUsageItem *>(p1);
    CByUsageItem *pitem2 = static_cast<CByUsageItem *>(p2);

    return CompareUAInfo(&pitem1->_uei, &pitem2->_uei);
}

// Sort by most frequently used - break ties by most recently used
int ByUsage::CompareUAInfo(const UEMINFO *puei1, const UEMINFO *puei2)
{
    float flResult = puei2->R - puei1->R;
    if (flResult < 0.0f)
        return -1;
    if (flResult > 0.0f)
        return 1;
    return CompareFileTime(&puei2->ftExecute, &puei1->ftExecute);
}

LPITEMIDLIST ByUsage::GetFullPidl(PaneItem *p)
{
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);

    return pitem->CreateFullPidl();
}


HRESULT ByUsage::GetFolderAndPidl(PaneItem *p, IShellFolder **ppsfOut, LPCITEMIDLIST *ppidlOut)
{
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);

    // If a single-level child pidl, then we can short-circuit the
    // SHBindToFolderIDListParent
    if (_ILNext(pitem->_pidl)->mkid.cb == 0)
    {
        *ppsfOut = pitem->_pdir->Folder(); (*ppsfOut)->AddRef();
        *ppidlOut = pitem->_pidl;
        return S_OK;
    }
    else
    {
        // Multi-level child pidl
        return SHBindToFolderIDListParent(pitem->_pdir->Folder(), pitem->_pidl,
                    IID_PPV_ARGS(ppsfOut), ppidlOut);
    }
}

HRESULT ByUsage::ContextMenuDeleteItem(PaneItem *p, IContextMenu *pcm, CMINVOKECOMMANDINFOEX *pici)
{
    IShellFolder *psf;
    LPCITEMIDLIST pidlItem;
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);

    HRESULT hr = GetFolderAndPidl(pitem, &psf, &pidlItem);
    if (SUCCEEDED(hr))
    {
        // Unpin the item - we go directly to the IStartMenuPin because
        // the context menu handler might decide not to support pin/unpin
        // for this item because it doesn't satisfy some criteria or other.
        LPITEMIDLIST pidlFull = pitem->CreateFullPidl();
        if (pidlFull)
        {
            _psmpin->Modify(pidlFull, NULL); // delete from pin list
            ILFree(pidlFull);
        }

        // Set hit count for shortcut to zero
        UEMINFO uei;
        ZeroMemory(&uei, sizeof(UEMINFO));
        uei.cbSize = sizeof(UEMINFO);
        uei.dwMask = UEIM_HIT;
        uei.cLaunches = 0;

        _SetUEMPidlInfo(psf, pidlItem, &uei);

        // Set hit count for target app to zero
        TCHAR szPath[MAX_PATH];
        if (SUCCEEDED(_GetShortcutExeTarget(psf, pidlItem, szPath, ARRAYSIZE(szPath))))
        {
            _SetUEMPathInfo(szPath, &uei);
        }

        // Set hit count for Darwin target to zero
        CByUsageHiddenData hd;
        hd.Get(pidlItem, CByUsageHiddenData::BUHD_MSIPATH);
        if (hd._pwszMSIPath && hd._pwszMSIPath[0])
        {
            _SetUEMPathInfo(hd._pwszMSIPath, &uei);
        }
        hd.Clear();

        psf->Release();

        if (IsSpecialPinnedItem(pitem))
        {
            // EXEX-VISTA(isabella): Disabled temporarily.
            // c_tray.CreateStartButtonBalloon(0, IDS_STARTPANE_SPECIALITEMSTIP);
        }

        // If the item wasn't pinned, then all we did was dork some usage
        // counts, which does not trigger an automatic refresh.  So do a
        // manual one.
        _pByUsageUI->Invalidate();
        PostMessage(_pByUsageUI->_hwnd, ByUsageUI::SFTBM_REFRESH, TRUE, 0);

    }

    return hr;
}

HRESULT ByUsage::ContextMenuInvokeItem(PaneItem *pitem, IContextMenu *pcm, CMINVOKECOMMANDINFOEX *pici, LPCTSTR pszVerb)
{
    ASSERT(_pByUsageUI);

    HRESULT hr;
    if (StrCmpIC(pszVerb, TEXT("delete")) == 0)
    {
        hr = ContextMenuDeleteItem(pitem, pcm, pici);
    }
    else
    {
        // Don't need to refresh explicitly if the command is pin/unpin
        // because the changenotify will do it for us
        hr = _pByUsageUI->SFTBarHost::ContextMenuInvokeItem(pitem, pcm, pici, pszVerb);
    }

    return hr;
}

int ByUsage::ReadIconSize()
{
    COMPILETIME_ASSERT(SFTBarHost::ICONSIZE_SMALL == 0);
    COMPILETIME_ASSERT(SFTBarHost::ICONSIZE_LARGE == 1);
    return _SHRegGetBoolValueFromHKCUHKLM(REGSTR_EXPLORER_ADVANCED, REGSTR_VAL_DV2_LARGEICONS, TRUE /* default to large*/);
}

BOOL ByUsage::_IsPinnedExe(CByUsageItem* pitem, IShellFolder* psf, LPCITEMIDLIST pidlItem)
{
    if (!_IsPinned(pitem))
        return FALSE;

    BOOL fIsExe = FALSE;

    WCHAR* pszFileName;
    if (SUCCEEDED(DisplayNameOfAsString(psf, pidlItem, SHGDN_INFOLDER | SHGDN_FORPARSING, &pszFileName)))
    {
        const WCHAR* pszExt = PathFindExtensionW(pszFileName);
        fIsExe = StrCmpICW(pszExt, L".exe") == 0;
        CoTaskMemFree(pszFileName);
    }
    return fIsExe;
}

HRESULT ByUsage::ContextMenuRenameItem(PaneItem *p, LPCTSTR ptszNewName)
{
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);

    IShellFolder *psf;
    LPCITEMIDLIST pidlItem;
    HRESULT hr;

    hr = GetFolderAndPidl(pitem, &psf, &pidlItem);
    if (SUCCEEDED(hr))
    {
        if (_IsPinnedExe(pitem, psf, pidlItem))
        {
            // Renaming a pinned exe consists merely of changing the
            // display name inside the pidl.
            //
            // Note!  SetAltName frees the pidl on failure.

            LPITEMIDLIST pidlNew;
            if ((pidlNew = ILClone(pitem->RelativePidl())) &&
                (pidlNew = CByUsageHiddenData::SetAltName(pidlNew, ptszNewName)))
            {
                hr = _psmpin->Modify(pitem->RelativePidl(), pidlNew);
                if (SUCCEEDED(hr))
                {
                    pitem->SetRelativePidl(pidlNew);
                }
                else
                {
                    ILFree(pidlNew);
                }
            }
            else
            {
                hr = E_OUTOFMEMORY;
            }
        }
        else
        {
            LPITEMIDLIST pidlNew;
            hr = psf->SetNameOf(_hwnd, pidlItem, ptszNewName, SHGDN_INFOLDER, &pidlNew);

            //
            //  Warning!  SetNameOf can set pidlNew == NULL if the rename
            //  was handled by some means outside of the pidl (so the pidl
            //  is unchanged).  This means that the rename succeeded and
            //  we can keep using the old pidl.
            //

            if (SUCCEEDED(hr) && pidlNew)
            {
                //
                // The old Start Menu renames the UEM data when we rename
                // the shortcut, but we cannot guarantee that the old
                // Start Menu is around, so we do it ourselves.  Fortunately,
                // the old Start Menu does not attempt to move the data if
                // the hit count is zero, so if it gets moved twice, the
                // second person who does the move sees cHit=0 and skips
                // the operation.
                //
                UEMINFO uei;
                _GetUEMPidlInfo(psf, pidlItem, &uei);
                if (uei.cLaunches > 0)
                {
                    _SetUEMPidlInfo(psf, pidlNew, &uei);
                    uei.cLaunches = 0;
                    _SetUEMPidlInfo(psf, pidlItem, &uei);
                }

                //
                // Update the pitem with the new pidl.
                //
                if (_IsPinned(pitem))
                {
                    LPITEMIDLIST pidlDad = ILCloneParent(pitem->RelativePidl());
                    if (pidlDad)
                    {
                        LPITEMIDLIST pidlFullNew = ILCombine(pidlDad, pidlNew);
                        if (pidlFullNew)
                        {
                            _psmpin->Modify(pitem->RelativePidl(), pidlFullNew);
                            pitem->SetRelativePidl(pidlFullNew);    // takes ownership
                        }
                        ILFree(pidlDad);
                    }
                    ILFree(pidlNew);
                }
                else
                {
                    ASSERT(pidlItem == pitem->RelativePidl());
                    pitem->SetRelativePidl(pidlNew);
                }
            }
        }
        psf->Release();
    }

    return hr;
}


//
//  If asking for the display (not for parsing) name of a pinned EXE,
//  we need to return the "secret display name".  Otherwise, we can
//  use the default implementation.
//
LPTSTR ByUsage::DisplayNameOfItem(PaneItem *p, IShellFolder *psf, LPCITEMIDLIST pidlItem, _SHGDNF shgno)
{
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);

    LPTSTR pszName = NULL;

    // Only display (not for-parsing) names of EXEs need to be hooked.
    if (!(shgno & SHGDN_FORPARSING) && _IsPinnedExe(pitem, psf, pidlItem))
    {
        //
        //  EXEs get their name from the hidden data.
        //
        pszName = CByUsageHiddenData::GetAltName(pidlItem);
    }

    return pszName ? pszName
                   : _pByUsageUI->SFTBarHost::DisplayNameOfItem(p, psf, pidlItem, shgno);
}

//
//  "Internet" and "Email" get subtitles consisting of the friendly app name.
//
LPTSTR ByUsage::SubtitleOfItem(PaneItem *p, IShellFolder *psf, LPCITEMIDLIST pidlItem)
{
    ASSERT(p->HasSubtitle());

    LPTSTR pszName = NULL;

    IAssociationElement *pae = GetAssociationElementFromSpecialPidl(psf, pidlItem);
    if (pae)
    {
        // We detect error by looking at pszName
        pae->QueryString(AQS_FRIENDLYTYPENAME, NULL, &pszName);
        pae->Release();
    }

    return pszName ? pszName
                   : _pByUsageUI->SFTBarHost::SubtitleOfItem(p, psf, pidlItem);
}

HRESULT ByUsage::MovePinnedItem(PaneItem *p, int iInsert)
{
    CByUsageItem *pitem = static_cast<CByUsageItem *>(p);
    ASSERT(_IsPinned(pitem));

    return _psmpin->Modify(pitem->RelativePidl(), SMPIN_POS(iInsert));
}

//
//  For drag-drop purposes, we let you drop anything, not just EXEs.
//  We just reject slow media.
//
BOOL ByUsage::IsInsertable(IDataObject *pdto)
{
    return _psmpin->IsPinnable(pdto, SMPINNABLE_REJECTSLOWMEDIA) == S_OK;
}


HRESULT ByUsage::InsertPinnedItem(IDataObject *pdto, int iInsert)
{
    HRESULT hr = E_FAIL;

    LPITEMIDLIST pidlItem;
    if (IsPinnable(pdto, SMPINNABLE_REJECTSLOWMEDIA, &pidlItem) == S_OK)
    {
        if (SUCCEEDED(hr = _psmpin->Modify(NULL, pidlItem)) &&
            SUCCEEDED(hr = _psmpin->Modify(pidlItem, SMPIN_POS(iInsert))))
        {
            // Woo-hoo!
        }
        ILFree(pidlItem);
    }
    return hr;
}