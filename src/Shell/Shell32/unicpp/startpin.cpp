#include "startpinhelpers.h"
#include <new>

class CStartMenuPin;

//
//  PINENTRY - A single entry in the pin list
//
//  The _liPos/_cbLink point back into the CPinList._pstmLink
//
class PINENTRY {
public:
    LPITEMIDLIST    _pidl;
    IShellLink *    _psl;           // a live IShellLink
    LARGE_INTEGER   _liPos;         // location of the shell link inside the stream
    DWORD           _cbSize;        // size of the buffer pointed to by _liPos

    HRESULT UpdateShellLink();

    void FreeShellLink()
    {
        _cbSize = 0;
        ATOMICRELEASE(_psl);
    }

    void Destruct()
    {
        ILFree(_pidl);
        FreeShellLink();
    }

    static BOOL CALLBACK DestroyCallback(PINENTRY *self, LPVOID)
    {
        self->Destruct();
        return TRUE;
    }
};

//
//  CPinList
//

class CPinList
{
public:
    CPinList() : _dsaEntries(NULL), _pstmLink(NULL) { }

    ~CPinList()
    {
        ATOMICRELEASE(_pstmLink);
        if (_dsaEntries)
        {
            _dsaEntries.DestroyCallbackEx(PINENTRY::DestroyCallback, (void *)NULL);
        }
    }

    BOOL    Initialize() { return _dsaEntries.Create(4); }
    HRESULT Load(CStartMenuPin *psmpin);
    HRESULT Save(CStartMenuPin *psmpin);

    int AppendPidl(LPITEMIDLIST pidl)
    {
        PINENTRY entry = { pidl };
        return _dsaEntries.AppendItem(&entry);
    }

    PINENTRY *GetItemPtr(int i) { return _dsaEntries.GetItemPtr(i); }


    HRESULT SaveShellLink(PINENTRY *pentry, IStream *pstm);
    HRESULT LoadShellLink(PINENTRY *pentry, IShellLink **ppsl);
    HRESULT UpdateShellLink(PINENTRY *pentry) { return pentry->UpdateShellLink(); }

    PINENTRY *FindPidl(LPCITEMIDLIST pidl, int *pi);
    HRESULT ReplacePidl(LPCITEMIDLIST pidlOld, LPCITEMIDLIST pidlNew);

private:
    struct ILWRITEINFO {
        IStream *pstmPidlWrite;
        IStream *pstmLinkWrite;
        CPinList *ppl;
        HRESULT hr;
        LPITEMIDLIST rgpidl[20];    // Must match ARRAYSIZE(c_rgcsidlRelative)
    };
    static BOOL CALLBACK ILWriteCallback(PINENTRY *pentry, ILWRITEINFO *pwi);

    CDSA<PINENTRY>  _dsaEntries;    // The items themselves
    IStream *       _pstmLink;      // PINENTRY._liPos points into this stream

};

class ATL_NO_VTABLE CStartMenuPin
    : public IShellExtInit
    , public IContextMenu
    , public IStartMenuPin
    , public CObjectWithSite
    , public CComObjectRootEx<CComSingleThreadModel>
    , public CComCoClass<CStartMenuPin, &CLSID_StartMenuPin>
{
public:
    CStartMenuPin() : _pdtobj(NULL), _idmPinCmd(IDM_MAX), _pidl(NULL) {}
    ~CStartMenuPin();

BEGIN_COM_MAP(CStartMenuPin)
    COM_INTERFACE_ENTRY(IShellExtInit)
    // Need to use COM_INTERFACE_ENTRY_IID for the interfaces
    // that don't have an idl
    COM_INTERFACE_ENTRY_IID(IID_IContextMenu, IContextMenu)
    COM_INTERFACE_ENTRY(IStartMenuPin)
    COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

DECLARE_NO_REGISTRY()

    // *** IShellExtInit ***
    STDMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, IDataObject *pdto, HKEY hkProgID);

    // *** IContextMenu ***
    STDMETHODIMP  QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pwRes, LPSTR pszName, UINT cchMax);

    // *** IStartMenuPin ***
    STDMETHODIMP EnumObjects(IEnumIDList **ppenum);
    STDMETHODIMP Modify(LPCITEMIDLIST pidlFrom, LPCITEMIDLIST pidlTo);
    STDMETHODIMP GetChangeCount(ULONG *pulOut);
    STDMETHODIMP IsPinnable(IDataObject *pdtobj, DWORD dwFlags, OPTIONAL LPITEMIDLIST *ppidl);
    STDMETHODIMP Resolve(HWND hwnd, DWORD dwFlags, LPCITEMIDLIST pidl, LPITEMIDLIST *ppidlResolved);

    // *** IObjectWithSite ***
    // Inherited from CObjectWithSite

public:
    HRESULT SetChangeCount(ULONG ul);

protected:

    BOOL _IsAcceptableTarget(LPCTSTR pszPath, DWORD dwAttrib, DWORD dwFlags);

    enum {
        IDM_PIN =   0,
        IDM_UNPIN = 1,
        IDM_MAX,
    };

    // These "seem" backwards, but remember: If the item is pinned,
    // then the command is "unpin".  If the item is unpinned, then
    // the command is "pin".
    inline void _SetPinned() { _idmPinCmd = IDM_UNPIN; }
    inline void _SetUnpinned() { _idmPinCmd = IDM_PIN; }
    inline BOOL _IsPinned() const { return _idmPinCmd != IDM_PIN; }
    inline BOOL _DoPin() const { return _idmPinCmd == IDM_PIN; }
    inline BOOL _DoUnpin() const { return _idmPinCmd != IDM_PIN; }
    inline UINT _GetMenuStringID() const
    {
        COMPILETIME_ASSERT(IDS_STARTPIN_UNPINME == IDS_STARTPIN_PINME + IDM_UNPIN);
        return IDS_STARTPIN_PINME + _idmPinCmd;
    }

    HRESULT _ShouldAddMenu(UINT uFlags);
    HRESULT _InitPinRegStream();

protected:
    IDataObject *_pdtobj;
    UINT        _idmPinCmd;         // Which command did we add?

    LPITEMIDLIST _pidl;             // IContextMenu identity
};


static IStream *_OpenLinksRegStream()
{
    IStream* stream = NULL;
    PinReadStore(REGSTR_VAL_STARTFAVLINKS, &stream);
    return stream;
}

const LARGE_INTEGER c_li0 = { 0, 0 };

HRESULT IStream_GetPos(IStream *pstm, LARGE_INTEGER *pliPos)
{
    return pstm->Seek(c_li0, STREAM_SEEK_CUR, (ULARGE_INTEGER*)pliPos);
}

static HRESULT PinCopyStream(IStream *pstmFrom, IStream *pstmTo, DWORD cb)
{
    ULARGE_INTEGER uliToCopy, uliCopied;
    uliToCopy.QuadPart = cb;
    HRESULT hr = pstmFrom->CopyTo(pstmTo, uliToCopy, NULL, &uliCopied);
    if (SUCCEEDED(hr) && uliToCopy.QuadPart != uliCopied.QuadPart)
    {
        hr = E_FAIL;
    }
    return hr;
}

class ATL_NO_VTABLE CStartMenuPinEnum
    : public IEnumIDList
    , public CComObjectRootEx<CComSingleThreadModel>
    , public CComCoClass<CStartMenuPinEnum>
{
public:
    CStartMenuPinEnum() : _hrLastEnum(S_OK), _pstm(NULL) {}
    HRESULT InitializeStream();
    ~CStartMenuPinEnum()
    {
        ATOMICRELEASE(_pstm);
    }

BEGIN_COM_MAP(CStartMenuPinEnum)
    COM_INTERFACE_ENTRY(IEnumIDList)
END_COM_MAP()

    /// *** IEnumIDList ***
    STDMETHODIMP Next(ULONG celt, LPITEMIDLIST *rgelt, ULONG *pceltFetched);
    STDMETHODIMP Skip(ULONG celt);
    STDMETHODIMP Reset();
    STDMETHODIMP Clone(IEnumIDList **ppenum);

private:
    HRESULT _NextPidlFromStream(LPITEMIDLIST *ppidl);

private:
    HRESULT     _hrLastEnum;        // Result of last IEnumIDList::Next
    IStream *   _pstm;
};

CStartMenuPin::~CStartMenuPin()
{
    ILFree(_pidl);
    if (_pdtobj)
        _pdtobj->Release();
}


BOOL CStartMenuPin::_IsAcceptableTarget(LPCTSTR pszPath, DWORD dwAttrib, DWORD dwFlags)
{
    //  Regitems ("Internet" or "Email" for example) are acceptable
    //  provided we aren't restricted to EXEs only.
    if (!(dwAttrib & SFGAO_FILESYSTEM))
    {
        return !(dwFlags & SMPINNABLE_EXEONLY);
    }

    //  Otherwise, it's a file.
    // XP's link handler supplied the target through SID_LinkSite. A locally
    // composed menu gets the shortcut itself; inspect its target but pin the link.
    TCHAR szTarget[MAX_PATH] = {};
    bool hasLinkPath = false;
    if (PathIsLnk(pszPath) && (dwFlags & (SMPINNABLE_EXEONLY | SMPINNABLE_REJECTSLOWMEDIA)))
    {
        CComPtr<IShellLink> link;
        if (SUCCEEDED(LoadFromFileW(CLSID_ShellLink, pszPath, IID_PPV_ARG(IShellLink, &link))))
            hasLinkPath = link->GetPath(szTarget, ARRAYSIZE(szTarget), NULL, 0) == S_OK;
        if ((dwFlags & SMPINNABLE_EXEONLY) && !hasLinkPath)
            return FALSE;
    }

    //  If requested, reject non-EXEs.
    //  (Like the Start Menu, we treat MSC files as if they were EXEs)
    if (dwFlags & SMPINNABLE_EXEONLY)
    {
        LPCTSTR pszExt = PathFindExtension(hasLinkPath ? szTarget : pszPath);
        if (StrCmpIC(pszExt, TEXT(".EXE")) != 0 &&
            StrCmpIC(pszExt, TEXT(".MSC")) != 0)
        {
            return FALSE;
        }
    }

    //  If requested, reject slow media
    if (dwFlags & SMPINNABLE_REJECTSLOWMEDIA)
    {
        if (!_IsLocalHardDisk(pszPath))
        {
            return FALSE;
        }

        // If it's a shortcut, then apply the same rule to its target.
        if (hasLinkPath && !_IsLocalHardDisk(szTarget))
            return FALSE;
    }

    //  All tests pass!

    return TRUE;

}


HRESULT CStartMenuPin::IsPinnable(IDataObject *pdtobj, DWORD dwFlags, OPTIONAL LPITEMIDLIST *ppidl)
{
    HRESULT hr = S_FALSE;

    LPITEMIDLIST pidlRet = NULL;

    // Native SSF_STARTPANELON no longer describes ExplorerEx's XP panel.
    if (pdtobj && !SHRestricted(REST_NOSMPINNEDLIST))
    {
        pidlRet = PinDataObjectPidl(pdtobj);
        if (pidlRet)
        {
            DWORD dwAttr = SFGAO_FILESYSTEM;
            TCHAR szPath[MAX_PATH] = {};
            if (SUCCEEDED(SHGetNameAndFlags(pidlRet, SHGDN_FORPARSING,
                    szPath, ARRAYSIZE(szPath), &dwAttr)) &&
                _IsAcceptableTarget(szPath, dwAttr, dwFlags))
                hr = S_OK;
        }
    }

    // Return pidlRet only if the call succeeded and the caller requested it
    if (hr != S_OK || !ppidl)
    {
        ILFree(pidlRet);
        pidlRet = NULL;
    }

    if (ppidl)
    {
        *ppidl = pidlRet;
    }

    return hr;

}

// Returns S_OK if should add, S_FALSE if not

HRESULT CStartMenuPin::_ShouldAddMenu(UINT uFlags)
{
    ILFree(_pidl);
    _pidl = NULL;
    _idmPinCmd = IDM_MAX;
    // "Pin" is never a default verb
    if (uFlags & CMF_DEFAULTONLY)
        return S_FALSE;

    HRESULT hr;

    // The context menu appears only for fast media
    //
    // If extended verbs are disabled, then show the menu only for EXEs

    DWORD dwFlags = SMPINNABLE_REJECTSLOWMEDIA;
    if (!(uFlags & CMF_EXTENDEDVERBS))
    {
        dwFlags |= SMPINNABLE_EXEONLY;
    }

    hr = IsPinnable(_pdtobj, dwFlags, &_pidl);

    if (S_OK == hr)
    {
        //  If we are enclosed inside a shortcut, change our identity to the
        //  enclosing shortcut.

        IPersistFile *ppf;
        if (SUCCEEDED(IUnknown_QueryService(_punkSite, SID_LinkSite, IID_PPV_ARG(IPersistFile, &ppf))))
        {
            LPOLESTR pszFile = NULL;
            if (ppf->GetCurFile(&pszFile) == S_OK && pszFile)
            {
                LPITEMIDLIST pidl = ILCreateFromPath(pszFile);
                if (pidl)
                {
                    ILFree(_pidl);
                    LPITEMIDLIST logical = SHLogILFromFSIL(pidl);
                    if (logical) { ILFree(pidl); pidl = logical; }
                    _pidl = pidl;
                    hr = S_OK;
                }
                CoTaskMemFree(pszFile);
            }
            ppf->Release();
        }
    }

    return hr;
}

// IShellExtInit::Initialize
HRESULT CStartMenuPin::Initialize(LPCITEMIDLIST, IDataObject *pdtobj, HKEY)
{
    ILFree(_pidl);
    _pidl = NULL;
    _idmPinCmd = IDM_MAX;
    IUnknown_Set((IUnknown **)&_pdtobj, pdtobj);    // just grab this guy
    return S_OK;
}

// IContextMenu::QueryContextMenu

HRESULT CStartMenuPin::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (!hmenu || idCmdFirst > idCmdLast || idCmdLast - idCmdFirst < IDM_MAX - 1)
        return E_INVALIDARG;
    HRESULT hr = _ShouldAddMenu(uFlags);
    if (S_OK == hr)
    {
        _SetUnpinned();

        //  Determine whether this item is already on the Start Page or not.
        IEnumIDList *penum;
        hr = EnumObjects(&penum);
        if (SUCCEEDED(hr))
        {
            LPITEMIDLIST pidl;
            while ((hr = penum->Next(1, &pidl, NULL)) == S_OK)
            {
                BOOL bSame = ILIsEqual(pidl, _pidl);
                ILFree(pidl);
                if (bSame)
                {
                    _SetPinned();
                    break;
                }
            }
            penum->Release();
            if (FAILED(hr)) return hr;

            TCHAR szCommand[MAX_PATH];
            if (!LoadString(PIN_RESOURCES, _GetMenuStringID(), szCommand, ARRAYSIZE(szCommand)))
                return E_FAIL;
            if (!InsertMenu(hmenu, indexMenu, MF_STRING | MF_BYPOSITION,
                            idCmdFirst + _idmPinCmd, szCommand))
                return HRESULT_FROM_WIN32(GetLastError());
            hr = ResultFromShort(IDM_MAX);
        }
    }
    if (hr == S_FALSE) return S_OK; // No command IDs consumed.
    return hr;
}

const LPCTSTR c_rgpszVerb[] =
{
    TEXT("pin"),                    // IDM_PIN
    TEXT("unpin"),                  // IDM_UNPIN
};

// *** IContextMenu::InvokeCommand

HRESULT CStartMenuPin::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    if (!pici || pici->cbSize < sizeof(CMINVOKECOMMANDINFO) || !_pidl)
        return E_INVALIDARG;
    LPCMINVOKECOMMANDINFOEX picix = reinterpret_cast<LPCMINVOKECOMMANDINFOEX>(pici);
    HRESULT hr = E_INVALIDARG;
    UINT idmCmd;
    bool unicodeVerb = pici->cbSize >= sizeof(CMINVOKECOMMANDINFOEX) &&
        (pici->fMask & CMIC_MASK_UNICODE) && !IS_INTRESOURCE(picix->lpVerbW);

    if (!unicodeVerb && IS_INTRESOURCE(pici->lpVerb))
    {
        idmCmd = PtrToInt(pici->lpVerb);
    }
    else
    {
        // Convert the string to an ID (or out of range if invalid)
        LPCTSTR pszVerb;
#ifdef UNICODE
        WCHAR szVerb[MAX_PATH];
        if (unicodeVerb)
        {
            pszVerb = picix->lpVerbW;
        }
        else
        {
            if (!MultiByteToWideChar(CP_ACP, 0, pici->lpVerb, -1, szVerb, ARRAYSIZE(szVerb)))
                return E_INVALIDARG;
            pszVerb = szVerb;
        }
#else
        pszVerb = pici->lpVerb;
#endif
        for (idmCmd = 0; idmCmd < ARRAYSIZE(c_rgpszVerb); idmCmd++)
        {
            if (lstrcmpi(pszVerb, c_rgpszVerb[idmCmd]) == 0)
            {
                break;
            }
        }
    }

    if (idmCmd < IDM_MAX && idmCmd == _idmPinCmd)
    {
        if (_idmPinCmd == IDM_PIN)
        {
            hr = Modify(NULL, _pidl);
        }
        else
        {
            hr = Modify(_pidl, NULL);
        }
    }

    return hr;
}

// *** IContextMenu::GetCommandString

HRESULT CStartMenuPin::GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pwRes, LPSTR pszName, UINT cchMax)
{
    if ((uType & ~GCS_UNICODE) == GCS_VALIDATEA)
        return idCmd < ARRAYSIZE(c_rgpszVerb) ? S_OK : E_INVALIDARG;
    if (!pszName || !cchMax) return E_INVALIDARG;
    TCHAR szBuf[MAX_PATH];
    LPCTSTR pszResult = NULL;

    switch (uType & ~GCS_UNICODE)
    {
    case GCS_VERBA:
        if (idCmd < ARRAYSIZE(c_rgpszVerb))
        {
            pszResult = c_rgpszVerb[idCmd];
        }
        break;

    case GCS_HELPTEXTA:
        if (idCmd < ARRAYSIZE(c_rgpszVerb))
        {
            COMPILETIME_ASSERT(IDS_STARTPIN_PINME_HELP + IDM_UNPIN == IDS_STARTPIN_UNPINME_HELP);
            if (LoadString(PIN_RESOURCES, IDS_STARTPIN_PINME_HELP + (UINT)idCmd, szBuf, ARRAYSIZE(szBuf)))
            {
                pszResult = szBuf;
            }
        }
        break;
    }

    if (pszResult)
    {
        if (uType & GCS_UNICODE)
        {
            return StringCchCopyW(reinterpret_cast<LPWSTR>(pszName), cchMax, pszResult);
        }
        else
        {
            if (!WideCharToMultiByte(CP_ACP, 0, pszResult, -1, pszName, cchMax, NULL, NULL))
                return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

    return E_NOTIMPL;
}

PINENTRY *CPinList::FindPidl(LPCITEMIDLIST pidl, int *pi)
{
    for (int i = _dsaEntries.GetItemCount() - 1; i >= 0; i--)
    {
        PINENTRY *pentry = _dsaEntries.GetItemPtr(i);
        if (ILIsEqual(pentry->_pidl, pidl))
        {
            if (pi)
            {
                *pi = i;
            }
            return pentry;
        }
    }
    return NULL;
}

HRESULT CPinList::ReplacePidl(LPCITEMIDLIST pidlOld, LPCITEMIDLIST pidlNew)
{
    int i;
    PINENTRY *pentry = FindPidl(pidlOld, &i);
    if (pentry)
    {
        if (pidlNew == NULL)            // Delete
        {
            pentry->Destruct();
            _dsaEntries.DeleteItem(i);
            return S_OK;
        }
        else
        if (IS_INTRESOURCE(pidlNew))    // Move
        {
            // Move the pidl from i to iPos
            PINENTRY entry = *pentry;
            int iPos = ((int)(INT_PTR)pidlNew) - 1;
            if (iPos < 0 || iPos > _dsaEntries.GetItemCount())
                return E_INVALIDARG;
            if (i < iPos)
            {
                // Moving down; others move up
                iPos--;
                if (iPos == i) return S_OK;
                // Must use MoveMemory because the memory blocks overlap
                MoveMemory(_dsaEntries.GetItemPtr(i),
                           _dsaEntries.GetItemPtr(i+1),
                           sizeof(PINENTRY) * (iPos-i));
            }
            else if (i > iPos)
            {
                // Moving up; others move down
                // Must use MoveMemory because the memory blocks overlap
                MoveMemory(_dsaEntries.GetItemPtr(iPos+1),
                           _dsaEntries.GetItemPtr(iPos),
                           sizeof(PINENTRY) * (i-iPos));
            }
            _dsaEntries.SetItem(iPos, &entry);
            return S_OK;
        }
        else                            // Replace
        {
            if (FindPidl(pidlNew, NULL) && !ILIsEqual(pidlOld, pidlNew))
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            if (PinSetPidl(&pentry->_pidl, pidlNew))
            {
                // Failure to update the shell link is not fatal;
                // it just means we won't be able to repair the
                // shortcut if it breaks.
                pentry->UpdateShellLink();
                return S_OK;
            }
            else
            {
                return E_OUTOFMEMORY;
            }
        }
    }
    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

HRESULT CStartMenuPin::Modify(LPCITEMIDLIST pidlFrom, LPCITEMIDLIST pidlTo)
{
    CPinStoreLock lock;
    if ((pidlFrom && IS_INTRESOURCE(pidlFrom)) ||
        (!pidlFrom && (!pidlTo || IS_INTRESOURCE(pidlTo))))
        return E_INVALIDARG;
    HRESULT hr;

    if(SHRestricted(REST_NOSMPINNEDLIST))
        return E_ACCESSDENIED;

    // Remap pidls to logical pidls (change CSIDL_DESKTOPDIRECTORY
    // to CSIDL_DESKTOP, etc.) so we don't get faked out when people
    // access objects sometimes directly on the desktop and sometimes
    // via their full filesystem name.

    LPITEMIDLIST pidlFromFree = NULL;
    LPITEMIDLIST pidlToFree = NULL;

    if (!IS_INTRESOURCE(pidlFrom))
    {
        pidlFromFree = SHLogILFromFSIL(pidlFrom);
        if (pidlFromFree) {
            pidlFrom = pidlFromFree;
        }
    }

    if (!IS_INTRESOURCE(pidlTo))
    {
        pidlToFree = SHLogILFromFSIL(pidlTo);
        if (pidlToFree) {
            pidlTo = pidlToFree;
        }
    }

    CPinList pl;
    hr = pl.Load(this);
    if (SUCCEEDED(hr))
    {
        if (SUCCEEDED(hr))
        {
            if (pidlFrom)
            {
                hr = pl.ReplacePidl(pidlFrom, pidlTo);
            }
            else if (pidlTo)
            {
                if (pl.FindPidl(pidlTo, NULL))
                {
                    ILFree(pidlFromFree);
                    ILFree(pidlToFree);
                    return S_FALSE;
                }
                LPITEMIDLIST pidl = ILClone(pidlTo);
                if (pidl)
                {
                    int iPos = pl.AppendPidl(pidl);
                    if (iPos >= 0)
                    {
                        // Failure to update the shell link is not fatal;
                        // it just means we won't be able to repair the
                        // shortcut if it breaks.
                        pl.GetItemPtr(iPos)->UpdateShellLink();
                    }
                    else
                    {
                        ILFree(pidl);
                        hr = E_OUTOFMEMORY;
                    }
                }
                else
                {
                    hr = E_OUTOFMEMORY;
                }
            }
            else
            {
                // pidlFrom == pidlTo == NULL?  What does that mean?
                hr = E_INVALIDARG;
            }

            if (SUCCEEDED(hr))
            {
                hr = pl.Save(this);
            }
        }
    }

    ILFree(pidlFromFree);
    ILFree(pidlToFree);

    return hr;
}

//
//  Find the pidl on the pin list and resolve the shortcut that
//  tracks it.
//
//  Returns S_OK if the pidl changed and was resolved.
//  Returns S_FALSE if the pidl did not change.
//  Returns an error if the Resolve failed.
//

HRESULT CStartMenuPin::Resolve(HWND hwnd, DWORD dwFlags, LPCITEMIDLIST pidl, LPITEMIDLIST *ppidlResolved)
{
    if (!ppidlResolved) return E_POINTER;
    *ppidlResolved = NULL;
    if (IS_INTRESOURCE(pidl)) return E_INVALIDARG;
    CPinStoreLock lock;

    if(SHRestricted(REST_NOSMPINNEDLIST))
        return E_ACCESSDENIED;

    // Remap pidls to logical pidls (change CSIDL_DESKTOPDIRECTORY
    // to CSIDL_DESKTOP, etc.) so we don't get faked out when people
    // access objects sometimes directly on the desktop and sometimes
    // via their full filesystem name.

    LPITEMIDLIST pidlFree = SHLogILFromFSIL(pidl);
    if (pidlFree) {
        pidl = pidlFree;
    }

    CPinList pl;
    bool linkChanged = false;
    HRESULT hr = pl.Load(this);
    if (SUCCEEDED(hr))
    {
        PINENTRY *pentry = pl.FindPidl(pidl, NULL);
        if (pentry)
        {
            IShellLink *psl;
            hr =  pl.LoadShellLink(pentry, &psl);
            if (SUCCEEDED(hr))
            {
                hr = psl->Resolve(hwnd, dwFlags);
                if (hr == S_OK)
                {
                    IPersistStream *pps;
                    hr = psl->QueryInterface(IID_PPV_ARG(IPersistStream, &pps));
                    if (SUCCEEDED(hr))
                    {
                        if (pps->IsDirty() == S_OK)
                        {
                            linkChanged = true;
                            LPITEMIDLIST pidlNew = NULL;
                            hr = psl->GetIDList(&pidlNew);
                            if (SUCCEEDED(hr) && (hr != S_OK || !pidlNew))
                            {
                                // GetIDList returns S_FALSE on failure...
                                hr = E_FAIL;
                            }
                            if (SUCCEEDED(hr))
                            {
                                if (ILIsEqual(pentry->_pidl, pidlNew))
                                {
                                    ILFree(pidlNew);
                                    hr = S_FALSE;
                                }
                                else
                                {
                                    ILFree(pentry->_pidl);
                                    pentry->_pidl = pidlNew;
                                    *ppidlResolved = ILClone(pidlNew);
                                    hr = *ppidlResolved ? S_OK : E_OUTOFMEMORY;
                                }
                            }
                            else
                                ILFree(pidlNew);
                        }
                        else
                            hr = S_FALSE;
                        pps->Release();
                    }
                }
                else if (SUCCEEDED(hr))
                {
                    // S_FALSE means "cancelled by user"
                    hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                psl->Release();
            }
        }
        else
        {
            hr = HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        if (SUCCEEDED(hr) && linkChanged)
        {
            HRESULT hrSave = pl.Save(this);
            if (FAILED(hrSave)) hr = hrSave;
            if (FAILED(hr))
            {
                ILFree(*ppidlResolved);
                *ppidlResolved = NULL;
            }
        }

    }

    ILFree(pidlFree);

    return hr;
}

//
//  The target pidl has changed (or it's brand new).  Create an IShellLink
//  around it so we can resolve it later.
//
HRESULT PINENTRY::UpdateShellLink()
{
    ASSERT(_pidl);

    // Pitch the old link; it's useless now.
    FreeShellLink();

    HRESULT hr = SHCoCreateInstance(NULL, &CLSID_ShellLink, NULL, IID_PPV_ARG(IShellLink, &_psl));
    if (SUCCEEDED(hr))
    {
        hr = _psl->SetIDList(_pidl);
        if (FAILED(hr))
        {
            FreeShellLink();        // pitch it; it's no good
        }
    }
    return hr;
}

HRESULT CPinList::SaveShellLink(PINENTRY *pentry, IStream *pstm)
{
    HRESULT hr;
    if (pentry->_psl)
    {
        // It's still in the form of an IShellLink.
        // Save it to the stream, then go back and update the size information.
        LARGE_INTEGER liPos, liPosAfter;
        DWORD cbSize = 0;
        IPersistStream *pps = NULL;
        if (SUCCEEDED(hr = IStream_GetPos(pstm, &liPos)) &&
            // Write a dummy DWORD; we will come back and patch it up later
            SUCCEEDED(hr = IStream_Write(pstm, &cbSize, sizeof(cbSize))) &&
            SUCCEEDED(hr = pentry->_psl->QueryInterface(IID_PPV_ARG(IPersistStream, &pps))))
        {
            if (SUCCEEDED(hr = pps->Save(pstm, FALSE)) &&
                SUCCEEDED(hr = IStream_GetPos(pstm, &liPosAfter)) &&
                SUCCEEDED(hr = pstm->Seek(liPos, STREAM_SEEK_SET, NULL)))
            {
                LONGLONG cbLink = liPosAfter.QuadPart - liPos.QuadPart - sizeof(DWORD);
                if (cbLink < 0 || cbLink > MAXDWORD)
                {
                    pps->Release();
                    return STG_E_MEDIUMFULL;
                }
                cbSize = static_cast<DWORD>(cbLink);
                if (SUCCEEDED(hr = IStream_Write(pstm, &cbSize, sizeof(cbSize))) &&
                    SUCCEEDED(hr = pstm->Seek(liPosAfter, STREAM_SEEK_SET, NULL)))
                {
                    // Hooray!  All got saved okay
                }
            }
            pps->Release();
        }
    }
    else
    {
        // It's just a reference back into our parent stream; copy it
        if (SUCCEEDED(hr = IStream_Write(pstm, &pentry->_cbSize, sizeof(pentry->_cbSize))))
        {
            // If _cbSize == 0 then _pstmLink might be NULL, so guard against it
            if (pentry->_cbSize)
            {
                if (SUCCEEDED(hr = _pstmLink->Seek(pentry->_liPos, STREAM_SEEK_SET, NULL)) &&
                    SUCCEEDED(hr = PinCopyStream(_pstmLink, pstm, pentry->_cbSize)))
                {
                    // Hooray! All got saved okay
                }
            }
            else
            {
                // Entry was blank - nothing to do, vacuous success
            }
        }
    }
    return hr;
}

HRESULT CPinList::LoadShellLink(PINENTRY *pentry, IShellLink **ppsl)
{
    HRESULT hr;
    if (pentry->_psl)
    {
        hr = S_OK;              // We already have the link
    }
    else if (pentry->_cbSize == 0)
    {
        hr = E_FAIL;            // no link available
    }
    else
    {                           // gotta make it
        IPersistStream *pps;
        hr = SHCoCreateInstance(NULL, &CLSID_ShellLink, NULL, IID_PPV_ARG(IPersistStream, &pps));
        if (SUCCEEDED(hr))
        {
            CComPtr<IStream> bounded;
            hr = CreateStreamOnHGlobal(NULL, TRUE, &bounded);
            if (SUCCEEDED(hr)) hr = _pstmLink->Seek(pentry->_liPos, STREAM_SEEK_SET, NULL);
            if (SUCCEEDED(hr)) hr = PinCopyStream(_pstmLink, bounded, pentry->_cbSize);
            if (SUCCEEDED(hr)) hr = bounded->Seek(c_li0, STREAM_SEEK_SET, NULL);
            if (SUCCEEDED(hr)) hr = pps->Load(bounded);
            if (SUCCEEDED(hr)) hr = pps->QueryInterface(IID_PPV_ARG(IShellLink, &pentry->_psl));
            pps->Release();
        }
    }

    *ppsl = pentry->_psl;

    if (SUCCEEDED(hr))
    {
        pentry->_psl->AddRef();
        hr = S_OK;
    }

    return hr;
}


HRESULT CStartMenuPin::GetChangeCount(ULONG *pulOut)
{
    if (!pulOut) return E_POINTER;
    DWORD cb = sizeof(*pulOut), type = 0;
    if (SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
                   REGSTR_VAL_STARTFAVCHANGES, &type, pulOut, &cb) != ERROR_SUCCESS ||
        type != REG_DWORD || cb != sizeof(*pulOut))
    {
        *pulOut = 0;
    }

    return S_OK;
}

HRESULT CStartMenuPin::SetChangeCount(ULONG ulChange)
{
    return HRESULT_FROM_WIN32(SHSetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
               REGSTR_VAL_STARTFAVCHANGES, REG_DWORD, &ulChange,
               sizeof(ulChange)));
}

//
//  We scan this list in order, so if there is a CSIDL that is a subdirectory
//  of another CSIDL, we must put the subdirectory first.  For example,
//  CSIDL_PROGRAMS is typically a subdirectory of CSIDL_STARTMENU, so we
//  must put CSIDL_PROGRAMS first so we get the best match.
//
//  Furthermore, directories pinned items are more likely to be found in
//  should come before less popular directories.
//
const int c_rgcsidlRelative[] = {
    // Most common: Start Menu stuff
    CSIDL_PROGRAMS,                 // Programs must come before StartMenu
    CSIDL_STARTMENU,                // Programs must come before StartMenu

    // Next most common: My Documents stuff
    CSIDL_MYPICTURES,               // MyXxx must come before Personal
    CSIDL_MYMUSIC,                  // MyXxx must come before Personal
    CSIDL_MYVIDEO,                  // MyXxx must come before Personal
    CSIDL_PERSONAL,                 // MyXxx must come before Personal
    CSIDL_COMMON_PROGRAMS,          // Programs must come before StartMenu
    CSIDL_COMMON_STARTMENU,         // Programs must come before StartMenu

    // Next most common: Desktop stuff
    CSIDL_DESKTOPDIRECTORY,
    CSIDL_COMMON_DESKTOPDIRECTORY,

    // Next most common: Program files stuff
    CSIDL_PROGRAM_FILES_COMMON,     // ProgramFilesCommon must come before ProgramFiles
    CSIDL_PROGRAM_FILES,            // ProgramFilesCommon must come before ProgramFiles
    CSIDL_PROGRAM_FILES_COMMONX86,  // ProgramFilesCommon must come before ProgramFiles
    CSIDL_PROGRAM_FILESX86,         // ProgramFilesCommon must come before ProgramFiles

    // Other stuff (less common)
    CSIDL_APPDATA,
    CSIDL_COMMON_APPDATA,
    CSIDL_SYSTEM,
    CSIDL_SYSTEMX86,
    CSIDL_WINDOWS,
    CSIDL_PROFILE,                  // Must come after all other profile-relative directories
};

BOOL CALLBACK CPinList::ILWriteCallback(PINENTRY *pentry, ILWRITEINFO *pwi)
{
    BYTE csidl = CSIDL_DESKTOP;     // Assume nothing interesting
    LPITEMIDLIST pidlWrite = pentry->_pidl;  // Assume nothing interesting

    for (int i = 0; i < ARRAYSIZE(pwi->rgpidl); i++)
    {
        LPITEMIDLIST pidlT;
        if (pwi->rgpidl[i] &&
            (pidlT = ILFindChild(pwi->rgpidl[i], pentry->_pidl)))
        {
            csidl = (BYTE)c_rgcsidlRelative[i];
            pidlWrite = pidlT;
            break;
        }
    }

    if (SUCCEEDED(pwi->hr = IStream_Write(pwi->pstmPidlWrite, &csidl, sizeof(csidl))) &&
        SUCCEEDED(pwi->hr = PinWritePidl(pwi->pstmPidlWrite, pidlWrite)) &&
        SUCCEEDED(pwi->hr = pwi->ppl->SaveShellLink(pentry, pwi->pstmLinkWrite)))
    {
        // woo-hoo, all written successfully
    }

    return SUCCEEDED(pwi->hr);
}

#define CSIDL_END ((BYTE)0xFF)

HRESULT CPinList::Save(CStartMenuPin *psmpin)
{
    ILWRITEINFO wi = {};
    COMPILETIME_ASSERT(ARRAYSIZE(c_rgcsidlRelative) == ARRAYSIZE(wi.rgpidl));
    for (int i = 0; i < ARRAYSIZE(c_rgcsidlRelative); i++)
        SHGetSpecialFolderLocation(NULL, c_rgcsidlRelative[i], &wi.rgpidl[i]);

    // Serialize both streams before touching the last saved list. SHOpenRegStream
    // commits on Release and cannot report write failures, so use checked writes.
    wi.hr = CreateStreamOnHGlobal(NULL, TRUE, &wi.pstmPidlWrite);
    if (SUCCEEDED(wi.hr))
        wi.hr = CreateStreamOnHGlobal(NULL, TRUE, &wi.pstmLinkWrite);
    if (SUCCEEDED(wi.hr))
    {
        wi.ppl = this;
        _dsaEntries.EnumCallbackEx(ILWriteCallback, &wi);
        BYTE csidlEnd = CSIDL_END;
        if (SUCCEEDED(wi.hr))
            wi.hr = IStream_Write(wi.pstmPidlWrite, &csidlEnd, sizeof(csidlEnd));
        if (SUCCEEDED(wi.hr))
            wi.hr = PinStoreStream(REGSTR_VAL_STARTFAVLINKS, wi.pstmLinkWrite);
        if (SUCCEEDED(wi.hr))
            wi.hr = PinStoreStream(REGSTR_VAL_STARTFAVS, wi.pstmPidlWrite);
    }
    ATOMICRELEASE(wi.pstmLinkWrite);
    ATOMICRELEASE(wi.pstmPidlWrite);
    for (int i = 0; i < ARRAYSIZE(c_rgcsidlRelative); i++)
        ILFree(wi.rgpidl[i]);
    if (FAILED(wi.hr)) return wi.hr;

    ULONG ulChange;
    psmpin->GetChangeCount(&ulChange);
    wi.hr = psmpin->SetChangeCount(ulChange + 1);
    if (FAILED(wi.hr)) return wi.hr;

    // SHChangeDWORDAsIDList is a packed private shell notification payload.
    #pragma pack(push, 1)
    struct { USHORT cb; DWORD dwItem1; DWORD dwItem2; USHORT cbZero; } dwidl =
        { 10, SHCNEE_PINLISTCHANGED, 0, 0 };
    #pragma pack(pop)
    SHChangeNotify(SHCNE_EXTENDED_EVENT, SHCNF_FLUSH, &dwidl, NULL);
    return S_OK;
}

HRESULT CPinList::Load(CStartMenuPin *psmpin)
{
    HRESULT hr;

    if (Initialize())
    {
        IEnumIDList *penum;

        hr = psmpin->EnumObjects(&penum);
        if (SUCCEEDED(hr))
        {
            LPITEMIDLIST pidl;
            while ((hr = penum->Next(1, &pidl, NULL)) == S_OK)
            {
                if (AppendPidl(pidl) < 0)
                {
                    ILFree(pidl);
                    hr = E_OUTOFMEMORY;
                    break;
                }
            }
            penum->Release();
        }

        if (SUCCEEDED(hr))
        {
            //
            //  Now read the persisted shortcuts.
            //
            _pstmLink = _OpenLinksRegStream();
            if (_pstmLink)
            {
                STATSTG stat = {};
                hr = _pstmLink->Stat(&stat, STATFLAG_NONAME);
                for (int i = 0; SUCCEEDED(hr) && i < _dsaEntries.GetItemCount(); i++)
                {
                    PINENTRY *pentry = _dsaEntries.GetItemPtr(i);
                    LARGE_INTEGER liSeek = { 0, 0 };
                    if (SUCCEEDED(hr = IStream_Read(_pstmLink, &liSeek.LowPart, sizeof(liSeek.LowPart))) && // read size
                        SUCCEEDED(hr = IStream_GetPos(_pstmLink, &pentry->_liPos)) &&  // read current pos
                        SUCCEEDED(hr = (pentry->_liPos.QuadPart < 0 ||
                            static_cast<ULONGLONG>(pentry->_liPos.QuadPart) > stat.cbSize.QuadPart ||
                            liSeek.LowPart > stat.cbSize.QuadPart - pentry->_liPos.QuadPart)
                            ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK) &&
                        SUCCEEDED(hr = _pstmLink->Seek(liSeek, STREAM_SEEK_CUR, NULL))) // skip over link
                    {
                        pentry->_cbSize = liSeek.LowPart; // set this only on success
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // If we encountered an error,
            // then throw all the shortcuts away because they are
            // probably corrupted.
            if (FAILED(hr))
            {
                for (int i = 0; i < _dsaEntries.GetItemCount(); i++)
                {
                    _dsaEntries.GetItemPtr(i)->FreeShellLink();
                }
            }

            // Problems reading the persisted shortcuts are ignored
            // since they are merely advisory.
            hr = S_OK;
        }
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }
    return hr;
}

//
//  Reading a pidl from a stream is a dangerous proposition because
//  a corrupted pidl can cause a shell extension to go haywire.
//
//  A pinned item is stored in the stream in the form
//
//  [byte:csidl] [dword:cbPidl] [size_is(cbPidl):pidl]
//
//  With the special csidl = -1 indicating the end of the list.
//
//  We use a byte for the csidl so a corrupted stream won't accidentally
//  pass "CSIDL_FLAG_CREATE" as a csidl to SHGetSpecialFolderLocation.

HRESULT CStartMenuPinEnum::_NextPidlFromStream(LPITEMIDLIST *ppidl)
{
    BYTE csidl;
    HRESULT hr = IStream_Read(_pstm, &csidl, sizeof(csidl));
    if (SUCCEEDED(hr))
    {
        if (csidl == CSIDL_END)
        {
            hr = S_FALSE;     // end of enumeration
        }
        else
        {
            LPITEMIDLIST pidlRoot;
            hr = SHGetSpecialFolderLocation(NULL, csidl, &pidlRoot);
            if (SUCCEEDED(hr))
            {
                LPITEMIDLIST pidl;
                hr = PinReadPidl(_pstm, &pidl);
                if (SUCCEEDED(hr))
                {
                    *ppidl = ILCombine(pidlRoot, pidl);
                    hr = *ppidl ? S_OK : E_OUTOFMEMORY;
                    ILFree(pidl);
                }
                ILFree(pidlRoot);
            }
        }
    }

    return hr;
}

// *** IEnumIDList::Next

HRESULT CStartMenuPinEnum::InitializeStream()
{
    HRESULT hr = PinReadStore(REGSTR_VAL_STARTFAVS, &_pstm);
    if (FAILED(hr)) return hr;
    STATSTG stat = {};
    hr = _pstm->Stat(&stat, STATFLAG_NONAME);
    if (SUCCEEDED(hr) && !stat.cbSize.QuadPart) _hrLastEnum = S_FALSE;
    return hr;
}

HRESULT CStartMenuPinEnum::Next(ULONG celt, LPITEMIDLIST *rgelt, ULONG *pceltFetched)
{
    if (pceltFetched) *pceltFetched = 0;
    if (!rgelt || (!pceltFetched && celt != 1)) return E_POINTER;
    if (!celt) return S_OK;
    ULONG fetched = 0;
    while (fetched < celt)
    {
        rgelt[fetched] = NULL;
        if (_hrLastEnum != S_OK) break;
        _hrLastEnum = _NextPidlFromStream(&rgelt[fetched]);
        if (_hrLastEnum != S_OK) break;
        ++fetched;
    }
    if (pceltFetched) *pceltFetched = fetched;
    return FAILED(_hrLastEnum) ? _hrLastEnum : fetched == celt ? S_OK : S_FALSE;
}

// *** IEnumIDList::Skip

HRESULT CStartMenuPinEnum::Skip(ULONG celt)
{
    for (ULONG i = 0; i < celt; ++i)
    {
        LPITEMIDLIST pidl = NULL;
        HRESULT hr = Next(1, &pidl, NULL);
        ILFree(pidl);
        if (hr != S_OK) return hr;
    }
    return S_OK;
}

// *** IEnumIDList::Reset

HRESULT CStartMenuPinEnum::Reset()
{
    HRESULT hr = _pstm->Seek(c_li0, STREAM_SEEK_SET, NULL);
    STATSTG stat = {};
    if (SUCCEEDED(hr)) hr = _pstm->Stat(&stat, STATFLAG_NONAME);
    _hrLastEnum = FAILED(hr) ? hr : stat.cbSize.QuadPart ? S_OK : S_FALSE;
    return hr;
}


// *** IEnumIDList::Clone

STDMETHODIMP CStartMenuPinEnum::Clone(IEnumIDList **ppenum)
{
    if (!ppenum) return E_POINTER;
    *ppenum = NULL;
    // Like the other Shell32 classes, these objects do not own an ATL module.
    CComObjectNoLock<CStartMenuPinEnum>* clone =
        new (std::nothrow) CComObjectNoLock<CStartMenuPinEnum>;
    if (!clone) return E_OUTOFMEMORY;
    clone->AddRef();
    HRESULT hr = _pstm->Clone(&clone->_pstm);
    if (SUCCEEDED(hr))
    {
        clone->_hrLastEnum = _hrLastEnum;
        hr = clone->QueryInterface(IID_PPV_ARG(IEnumIDList, ppenum));
    }
    clone->Release();
    return hr;
}

// *** IStartMenuPin::EnumObjects

STDMETHODIMP CStartMenuPin::EnumObjects(IEnumIDList **ppenum)
{
    if (!ppenum) return E_POINTER;
    *ppenum = NULL;
    CPinStoreLock lock;
    HRESULT hr = _InitPinRegStream();
    if (FAILED(hr)) return hr;
    CComObjectNoLock<CStartMenuPinEnum>* enumerator =
        new (std::nothrow) CComObjectNoLock<CStartMenuPinEnum>;
    if (!enumerator) return E_OUTOFMEMORY;
    enumerator->AddRef();
    hr = enumerator->InitializeStream();
    if (SUCCEEDED(hr)) hr = enumerator->QueryInterface(IID_PPV_ARG(IEnumIDList, ppenum));
    enumerator->Release();
    return hr;
}

extern "C" HRESULT CStartMenuPin_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppunk)
{
    if (!ppunk) return E_POINTER;
    *ppunk = NULL;
    if (punkOuter) return CLASS_E_NOAGGREGATION;
    return CComCreator<CComObjectNoLock<CStartMenuPin>>::CreateInstance(NULL, riid, ppunk);
}

// *** IStartMenuPin::_InitPinRegStream
//
// If the pin list has not yet been created, then create a default one.
//

static LPCTSTR c_rgszDefaultPin[] = {
    TEXT("shell:::{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}"), // CLSID_AutoCMClientInet
    TEXT("shell:::{2559a1f5-21d7-11d4-bdaf-00c04f60b9f0}"), // CLSID_AutoCMClientMail
};

HRESULT CStartMenuPin::_InitPinRegStream()
{
    CPinStoreLock lock;
    // Existence, not emptiness, is authoritative: a user-cleared list stays empty.
    DWORD size = 0;
    LSTATUS error = SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
        REGSTR_VAL_STARTFAVS, NULL, NULL, &size);
    if (error == ERROR_SUCCESS || error == ERROR_MORE_DATA) return S_OK;
    if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
        return HRESULT_FROM_WIN32(error);

    CPinList pl;
    if (!pl.Initialize()) return E_OUTOFMEMORY;
    bool imported = false;
    // Only the first method is used, shared by the readable native variants.
    // Never route through the local activation hook or call native Modify.
    struct IReadablePinnedList : IUnknown
    {
        virtual HRESULT STDMETHODCALLTYPE EnumObjects(IEnumFullIDList**) = 0;
    };
    const IID* nativeIids[] = { &IID_IPinnedList3, &IID_IFlexibleTaskbarPinnedList,
        &IID_IPinnedList25, &IID_IPinnedList2 };
    for (const IID* iid : nativeIids)
    {
        CComPtr<IReadablePinnedList> native;
        HRESULT hr = CoCreateInstance(CLSID_StartMenuPin, NULL, CLSCTX_INPROC_SERVER,
            *iid, reinterpret_cast<void**>(&native));
        if (FAILED(hr)) continue;
        CComPtr<IEnumFullIDList> enumerator;
        hr = native->EnumObjects(&enumerator);
        if (FAILED(hr) || !enumerator) continue;
        imported = true; // A readable empty list is also an intentional list.
        for (;;)
        {
            PIDLIST_ABSOLUTE pidl = NULL;
            ULONG fetched = 0;
            hr = enumerator->Next(1, &pidl, &fetched);
            if (hr == S_FALSE && !fetched) { ILFree(pidl); break; }
            if (hr != S_OK || fetched != 1 || !pidl)
            {
                ILFree(pidl);
                return FAILED(hr) ? hr : E_UNEXPECTED;
            }
            LPITEMIDLIST logical = SHLogILFromFSIL(pidl);
            if (logical) { ILFree(pidl); pidl = logical; }
            int index = pl.AppendPidl(pidl);
            if (index < 0) { ILFree(pidl); return E_OUTOFMEMORY; }
            pl.GetItemPtr(index)->UpdateShellLink();
        }
        break;
    }
    if (!imported)
    {
        for (int i = 0; i < ARRAYSIZE(c_rgszDefaultPin); i++)
        {
            LPITEMIDLIST pidl = ILCreateFromPath(c_rgszDefaultPin[i]);
            if (!pidl) continue; // The XP default namespace object may not be installed.
            int index = pl.AppendPidl(pidl);
            if (index < 0) { ILFree(pidl); return E_OUTOFMEMORY; }
            pl.GetItemPtr(index)->UpdateShellLink();
        }
    }
    HRESULT hr = pl.Save(this);
    if (FAILED(hr)) return hr;
    DWORD complete = 1;
    return HRESULT_FROM_WIN32(SHSetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
        TEXT("MigrationComplete"), REG_DWORD, &complete, sizeof(complete)));
}
