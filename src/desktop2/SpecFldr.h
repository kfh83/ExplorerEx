//
//  Plug-in to enumerate a list of folders from a registry key
//
//
//  Property bag:
//
//      "Target"    - name of registry key to enumerate
//

#include "pch.h"
#include "sfthost.h"

class CKnownFolderInformation
{
public:
    ~CKnownFolderInformation()
    {
        ILFree(_pidl);
        CoTaskMemFree(_pszDispName);
        if (_hIcon)
        {
            DestroyIcon(_hIcon);
        }
    }

    LPITEMIDLIST _pidl;
    LPWSTR _pszDispName;
    HICON _hIcon;
};

class SpecialFolderList : public SFTBarHost
{
public:

    friend SFTBarHost *SpecList_CreateInstance()
        { return new SpecialFolderList(); }

    SpecialFolderList()
        : SFTBarHost(HOSTF_CANRENAME | HOSTF_REVALIDATE | HOSTF_CASCADEMENU)
    {
        field_18C = 0;
        _iThemePart = SPP_PLACESLIST;
        _iThemePartSep = SPP_PLACESLISTSEPARATOR;
    }

	// *** IOleCommandTarget ***
    STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut);

    class CLoadFullPidlTask : public CRunnableTask
    {
    public:
        CLoadFullPidlTask(HWND hwnd, DWORD dwIndex, int iIconSize)
            : CRunnableTask(RTF_DEFAULT)
            , _hwnd(hwnd)
            , _dwIndex(dwIndex)
            , _iIconSize(_iIconSize)
        {
        }

        // *** IRunnableTask ***
        STDMETHODIMP InternalResumeRT();

    private:
        HWND   _hwnd;
        DWORD  _dwIndex;
        int    _iIconSize;
    };

private:
    ~SpecialFolderList();

    HRESULT Initialize();
    void EnumItems();
    int CompareItems(PaneItem *p1, PaneItem *p2);
    HRESULT GetFolderAndPidl(PaneItem *pitem, IShellFolder **ppsfOut, PCITEMID_CHILD *ppidlOut);

    HRESULT GetFolderAndPidlForActivate(PaneItem *pitem, IShellFolder **ppsfOut, PCITEMID_CHILD *ppidlOut)
    {
		return GetFolderAndPidl(pitem, ppsfOut, ppidlOut);
    }
    HRESULT ContextMenuRenameItem(PaneItem *p, LPCTSTR ptszNewName);
    HRESULT GetCascadeMenu(PaneItem *pitem, IShellMenu **ppsm);
    int ReadIconSize() { return ICONSIZE_MEDIUM; }
    BOOL NeedBackgroundEnum() { return TRUE; }
    LPTSTR DisplayNameOfItem(PaneItem *p, IShellFolder *psf, LPCITEMIDLIST pidlItem, SHGDNF shgno);
    TCHAR GetItemAccelerator(PaneItem *pitem, int iItemStart);
    void OnChangeNotify(UINT id, LONG lEvent, LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2);
    BOOL IsBold(PaneItem *pitem);
    void GetItemInfoTip(PaneItem *pitem, LPTSTR pszText, DWORD cch);

    void UpdateImage(int iImage)
    {
        PostMessage(_hwnd, 0x40Au, 1u, 0);
    }

    void _NotifyHoverImage(int iImage)
    {
        VARIANT vt;
        vt.vt = VT_BYREF;
        vt.byref = _himl;
        IUnknown_QueryServiceExec(_punkSite, SID_SM_UserPane, &SID_SM_DV2ControlHost, 314, 0, &vt, NULL);

        vt.vt = VT_I4;
        vt.lVal = iImage;
        IUnknown_QueryServiceExec(_punkSite, SID_SM_UserPane, &SID_SM_DV2ControlHost, 313, 0, &vt, NULL);
    }

    HRESULT ContextMenuInvokeItem(PaneItem *p, IContextMenu *pcm, CMINVOKECOMMANDINFOEX *pici, LPCTSTR pszVerb);
    UINT AdjustDeleteMenuItem(PaneItem *pitem, UINT *puiFlags);
    LRESULT OnWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    int GetMinTextWidth();
    HRESULT OnItemUpdate(PaneItem *p, WPARAM wParam, LPARAM lParam);

    void OnItemUpdateComplete(LPARAM lParam, WPARAM)
    {
        CKnownFolderInformation *pkfi = reinterpret_cast<CKnownFolderInformation *>(lParam);
        if (pkfi)
        {
            delete pkfi;
        }
    }

    HRESULT _CreateFullListItem(class SpecialFolderListItem *pitem);
    HRESULT _CreateSimpleListItem(IKnownFolderManager *pkfm, class SpecialFolderListItem *pitem, DWORD a4);

private:
    static DWORD WINAPI _HasEnoughChildrenThreadProc(LPVOID pvData);

    UINT    _cNotify;

    int     field_188;
    int     field_18C;
    int     field_190;
};
