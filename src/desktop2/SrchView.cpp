#include "pch.h"

#include "SrchView.h"

#ifdef COMPILE_SRCHVIEW

#include "cabinet.h"
#include "propvarutil.h"
#include "SFTHost.h"

#pragma comment(lib, "propsys.lib")

#define GIT_COOKIEINVALID ((DWORD)-1) // Guessed value

CSearchOpenView::CSearchOpenView()
{
	field_98 = 0;
	DWORD cbData = sizeof(field_98);
	_SHRegGetValueFromHKCUHKLM(DV2_REGPATH, L"StartPanel_TopMatch", SRRF_RT_ANY, 0, &field_98, &cbData);
	if (field_98 != 2
		&& field_98 != 3
		&& field_98 >= 2
		&& field_98 != 4)
	{
		field_98 = 0;
	}

	field_AC = -2;
	dword9C = -1;
	cbData = sizeof(DWORD);
	SHGetValue(HKEY_CURRENT_USER, DV2_REGPATH, L"StartMenuIndexed", 0, &field_AC, &cbData);

	ASSERT(_pszGroupPrograms == NULL);			// 263
	ASSERT(_pszGroupInternet == NULL);			// 264
	ASSERT(_pqp == NULL);						// 265
	ASSERT(_pcf == NULL);						// 266
	ASSERT(_psmqc == NULL);						// 267
	ASSERT(_psmqs == NULL);						// 268
	ASSERT(_dwCookieSink == GIT_COOKIEINVALID); // 269
}

CSearchOpenView::~CSearchOpenView()
{
	if (_ppci)
	{
		delete _ppci;
	}

	LocalFree(_pszGroupPrograms);
	LocalFree(_pszGroupInternet);
	CoTaskMemFree(field_A8);

	ASSERT(_pqp == NULL);						// 278
	ASSERT(_pcf == NULL);						// 279
	ASSERT(_psmqc == NULL);						// 280

	_RevokeQuerySink();
	IUnknown_SafeReleaseAndNullPtr(&_psmqs);
}

HRESULT CSearchOpenView::QueryInterface(REFIID riid, void **ppvObj)
{
	static const QITAB qit[] =
	{
		QITABENT(CSearchOpenView, IServiceProvider),
		QITABENT(CSearchOpenView, IOleCommandTarget),
		QITABENT(CSearchOpenView, IObjectWithSite),
		QITABENT(CSearchOpenView, IExplorerBrowserEvents),
		QITABENT(CSearchOpenView, IDispatch),
		QITABENT(CSearchOpenView, INavigationOptions),
		QITABENTMULTI(CSearchOpenView, ICommDlgBrowser3, ICommDlgBrowser2, ICommDlgBrowser),
		QITABENT(CSearchOpenView, IContextMenuModifier),
		QITABENT(CSearchOpenView, IBrowserSettings),
		QITABENT(CSearchOpenView, IAccessibleProvider),
		{0},
	};

	return QISearch(this, qit, riid, ppvObj);
}

ULONG CSearchOpenView::AddRef()
{
	return CUnknown::AddRef();
}

ULONG CSearchOpenView::Release()
{
	return CUnknown::Release();
}

const IID IID_IAccessibleProvider = __uuidof(IAccessibleProvider);
const IID IID_IBrowserSettings = __uuidof(IBrowserSettings);

// Thanks to ep_taskbar by @amrsatrio for these GUIDs
DEFINE_GUID(SID_DefViewNavigationManager, 0x796597DA, 0x7A1B, 0x4978, 0x84, 0xCB, 0xFD, 0x33, 0xEE, 0x26, 0x43, 0xA6);
DEFINE_GUID(SID_SShellFolderViewCM, 0x5C65B0D2, 0xFC07, 0x48B6, 0xBA, 0x51, 0x45, 0x97, 0xDA, 0xC8, 0x47, 0x47);

HRESULT CSearchOpenView::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
	if (IsEqualGUID(guidService, SID_SM_OpenView)
		|| IsEqualGUID(guidService, SID_SExplorerBrowserFrame)
		|| IsEqualGUID(guidService, SID_DefViewNavigationManager)
		|| IsEqualGUID(guidService, SID_SShellFolderViewCM)
		|| IsEqualGUID(guidService, IID_IAccessibleProvider)
		|| IsEqualGUID(guidService, IID_IBrowserSettings))
	{
		HRESULT hr = QueryInterface(riid, ppvObject);
		if (SUCCEEDED(hr))
		{
			return hr;
		}
	}
	return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

HRESULT CSearchOpenView::QueryStatus(const GUID *pguidCmdGroup,
	ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

DEFINE_GUID(POLID_NoRun, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

HRESULT CSearchOpenView::Exec(const GUID *pguidCmdGroup,
	DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
	HRESULT hr = E_INVALIDARG;

	if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
	{
		switch (nCmdID)
		{
			case 300:
			{
				field_8C = 0;

				LPWSTR pszPath = NULL;
				hr = VariantToStringAlloc(*pvarargIn, &pszPath);
				if (SUCCEEDED(hr) && pszPath)
				{
					if (*pszPath)
					{
						if (!_peb)
						{
							hr = _RecreateBrowserObject();
						}

						if (SUCCEEDED(hr) && _peb)
						{
							field_84 = 0;
							field_88 = 0;
							field_94 = 0;

							LPWSTR pszPath1;
							hr = VariantToStringAlloc(*pvarargIn, &pszPath1);
							if (SUCCEEDED(hr))
							{
								if (field_AC == -2)
								{
									field_AC = -1;
									_AddCheckIndexerStaterTask();
								}

								PathStripToRoot(pszPath1);
								if (!PathIsRoot(pszPath1) || _fIsBrowsing || SHWindowsPolicy(POLID_NoRun))
								{
									hr = _UpdateSearchText(pszPath);
								}
								else
								{
									hr = AddPathCompletionTask(pszPath);
								}
								IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &SID_SM_DV2ControlHost, 302, 2, NULL, NULL);
								CoTaskMemFree(pszPath1);
							}
							field_90 = 1;
						}
					}
					else
					{
						_CancelNavigation();

						field_84 = 1;

						CoTaskMemFree(field_A8);
						field_A8 = NULL;

						if (field_90 != 0)
						{
							_RecreateBrowserObject();
						}

						_SwitchToMode(VIEWMODE_DEFAULT, 0);

						VARIANT vt;
						vt.vt = VT_BOOL;
						vt.iVal = VARIANT_TRUE;
						hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &SID_SM_DV2ControlHost, 302, -1, &vt, NULL);
					}
					CoTaskMemFree(pszPath);
				}
				break;
			}
			case 301:
			{
				hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 301, 0, NULL, NULL);
				if (FAILED(hr))
				{
					if (field_88 || _viewMode)
					{
						if (_pFolderView)
						{
							int iCurSel = _GetCurSel();
							if (iCurSel >= 0)
							{
								return _ActivateItem(iCurSel);
							}

							if (_viewMode != 1)
							{
								return hr;
							}
							return _ActivateItem(-1);
						}
					}
					else
					{
						field_8C = 1;
					}
				}
				break;
			}
			case 304:
			{
				if (_pFolderView && _viewMode == VIEWMODE_DEFAULT)
				{
					if ((pvarargIn->lVal & 0x20000) != 0 || IsChild(_hwnd, GetFocus()))
					{
						if (_GetItemCount())
						{
							if (field_88)
							{
								int iItem;
								_pFolderView->GetVisibleItem(-1, FALSE, &iItem);

								LPITEMIDLIST pidl;
								if (iItem >= 0 && SUCCEEDED(_pFolderView->Item(iItem, &pidl)))
								{
									_pFolderView->SelectAndPositionItems(1, (LPCITEMIDLIST *)&pidl, NULL, SVSI_SELECT);
									ILFree(pidl);
								}
							}
						}
					}
					else if ((pvarargIn->lVal & 0x40000) == 0)
					{
						IShellView2 *psv2;
						hr = _pFolderView->QueryInterface(IID_PPV_ARGS(&psv2));
						if (SUCCEEDED(hr))
						{
							hr = psv2->SelectAndPositionItem(NULL, SVSI_DESELECTOTHERS, NULL);
							psv2->Release();
						}
					}
				}
				break;
			}
			case 325:
			{
				hr = S_OK;
				pvarargOut->vt = VT_I4;
				pvarargOut->lVal = _GetCurSel();
				break;
			}
		}
	}
	return hr;
}

HRESULT CSearchOpenView::SetSite(IUnknown *punkSite)
{
	CObjectWithSite::SetSite(punkSite);
	if (!punkSite)
	{
		_ReleaseExplorerBrowser();
		IUnknown_SafeReleaseAndNullPtr(&_pFolderView);
		IUnknown_SafeReleaseAndNullPtr(&_psiaStartMenuProvider);
		IUnknown_SafeReleaseAndNullPtr(&_psched);
		IUnknown_SafeReleaseAndNullPtr(&_pqp);
		IUnknown_SafeReleaseAndNullPtr(&_pcf);

		if (_psmqc)
		{
			_psmqc->ReleaseCache();
		}

		IUnknown_SafeReleaseAndNullPtr(&_psmqc);
		_DisconnectShellView();
	}
	return S_OK;
}

HRESULT CSearchOpenView::OnNavigationPending(PCIDLIST_ABSOLUTE pidlFolder)
{
	return S_OK;
}

HRESULT CSearchOpenView::OnViewCreated(IShellView *psv)
{
	return S_OK;
}

HRESULT CSearchOpenView::OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder)
{
	_peb->SetOptions(EBO_NOTRAVELLOG | EBO_NAVIGATEONCE);

	if (_ppci && _viewMode == 1)
	{
		_FilterPathCompleteView(_ppci->_psz1);
	}
	return S_OK;
}

HRESULT CSearchOpenView::OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder)
{
	wprintf(L"CSearchOpenView::OnNavigationFailed\n");
	return S_OK;
}

HRESULT CSearchOpenView::GetTypeInfoCount(UINT *pctinfo)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames,
	UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return E_NOTIMPL;
}

extern FILETIME g_ftOpenBoxChar;
extern HANDLE g_SHPerfRegHandle;
extern DWORD g_nOpenBoxCharThreadId;

HRESULT CSearchOpenView::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult,
	EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (dispIdMember == 207)
	{
		_UpdateTopMatch(UTM_REASON_0);
	}
	else if (dispIdMember == 216)
	{
		/*if (g_SHPerfRegHandle)
		{
			if (EventEnabled(g_SHPerfRegHandle, &ShellTraceId_Explorer_StartPane_OpenBox_SearchReady_Info))
			{
				FILETIME ft; // [esp+4h] [ebp-8h] BYREF
				GetSystemTimeAsFileTime(&ft);
				INT64 v9 = FILETIMEtoInt64(g_ftOpenBoxChar);
				INT64 v10 = FILETIMEtoInt64(ft);
				HRESULT hr = ULongLongToUInt(v10 - v9, (unsigned int *)&dispIdMember);

				DWORD dwThreadId = GetCurrentThreadId();
				if (dwThreadId == InterlockedCompareExchange(&g_nOpenBoxCharThreadId, 0, dwThreadId))
				{
					if (hr >= 0)
						SHTracePerfSQMStreamOneImpl(&ShellTraceId_Explorer_StartPane_OpenBox_SearchReady_Info, 179, dispIdMember);
					else
						SHTracePerf(&ShellTraceId_Explorer_StartPane_OpenBox_SearchReady_Info);
				}
			}
		}*/

		field_84 = 1;
		_UpdateTopMatch(UTM_REASON_0);
		if (_viewMode == VIEWMODE_PATHCOMPLETE)
		{
			_UpdateScrolling();
		}
		if (field_94)
		{
			_RecreateBrowserObject();
		}
	}
	return S_OK;
}

HRESULT CSearchOpenView::CanNavigateToIDList(PCIDLIST_ABSOLUTE pidl1, PCIDLIST_ABSOLUTE pidl2)
{
	HRESULT hr = S_FALSE;

	if (_viewMode == VIEWMODE_DEFAULT && !field_7C)
		hr = S_OK;

	if (_viewMode == VIEWMODE_PATHCOMPLETE && ILIsEqual(pidl1, pidl2))
		hr = S_OK;

	field_7C = 0;
	return hr;
}

HRESULT CSearchOpenView::OnDefaultCommand(IShellView *ppshv)
{
	HRESULT hr = E_FAIL;

	if (_pFolderView)
	{
		if (dword9C >= 0)
		{
			hr = _ActivateItem(dword9C);
		}

		dword9C = _GetCurSel();
		if (dword9C >= 0)
		{
			hr = _ActivateItem(dword9C);
		}
	}
	return hr;
}

HRESULT CSearchOpenView::OnStateChange(IShellView *ppshv, ULONG uChange)
{
	HRESULT hr = S_OK;

	if (uChange == CDBOSC_KILLFOCUS)
	{
		HWND hwndFocus = GetFocus();
		if (!IsChild(_hwnd, hwndFocus))
		{
			IShellView2 *psv2;
			hr = ppshv->QueryInterface(IID_PPV_ARGS(&psv2));
			if (SUCCEEDED(hr))
			{
				hr = psv2->SelectAndPositionItem(NULL, SVSI_DESELECTOTHERS, NULL);
				psv2->Release();
			}
		}
	}
	return hr;
}

HRESULT CSearchOpenView::IncludeObject(IShellView *ppshv, PCUITEMID_CHILD pidl)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::Notify(IShellView *ppshv, DWORD dwNotifyType)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::GetDefaultMenuText(IShellView *ppshv, LPWSTR pszText, int cchMax)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::GetViewFlags(DWORD *pdwFlags)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::OnColumnClicked(IShellView *ppshv, int iColumn)
{
	return E_NOTIMPL;
}

HRESULT CSearchOpenView::GetCurrentFilter(LPWSTR pszFileSpec, int cchFileSpec)
{
	return E_NOTIMPL;
}

const PROPERTYKEY PKEY_ItemNameDisplay = { { 3072717104u, 18415u, 4122u, { 165u, 241u, 2u, 96u, 140u, 158u, 235u, 172u } }, 10u };

HRESULT CSearchOpenView::OnPreViewCreated(IShellView *ppshv)
{
	_fIsBrowsing = 0;

	IUnknown_SafeReleaseAndNullPtr(&_pFolderView);

	_DisconnectShellView();
	_ConnectShellView(ppshv);

	field_B0 = _viewMode == VIEWMODE_PATHCOMPLETE;

	HRESULT hr = ppshv->QueryInterface(IID_PPV_ARGS(&_pFolderView));
	if (SUCCEEDED(hr))
	{
		if (_hTheme)
		{
			IVisualProperties *pvp;
			hr = _pFolderView->QueryInterface(IID_PPV_ARGS(&pvp));
			if (SUCCEEDED(hr))
			{
				LPCWSTR pszTheme = IsCompositionActive() ? L"StartMenuComposited" : L"StartMenu";
				hr = pvp->SetTheme(pszTheme, NULL);
				pvp->Release();
			}
		}

		if (SUCCEEDED(hr))
		{
			RECT rc;
			GetClientRect(_hwnd, &rc);
			_SizeExplorerBrowser(rc.right, rc.bottom);

			FOLDERFLAGS folderFlags = _GetFolderFlags();
			hr = _pFolderView->SetCurrentFolderFlags(folderFlags, folderFlags);
			if (SUCCEEDED(hr))
			{
				hr = _pFolderView->SetCurrentViewMode(FVM_DETAILS);
				if (SUCCEEDED(hr))
				{
					IColumnManager *pcm;
					hr = _pFolderView->QueryInterface(IID_PPV_ARGS(&pcm));
					if (SUCCEEDED(hr))
					{
						hr = pcm->SetColumns(&PKEY_ItemNameDisplay, 1);
						pcm->Release();
					}
				}
			}
		}
	}

	_fIsBrowsing = FAILED(hr);

	if (_viewMode != VIEWMODE_PATHCOMPLETE)
	{
		_LimitViewResults(_pFolderView, 1);

		IFilterView *pfv;
		hr = ppshv->QueryInterface(IID_PPV_ARGS(&pfv));
		if (SUCCEEDED(hr))
		{
			hr = _FilterView(pfv);
			pfv->Release();
		}
	}

	_SwitchToMode(_viewMode, 1);
	return hr;
}

//class CContextMenuForwarder
//    : public IContextMenu3
//    , public IObjectWithSite
//    , public IShellExtInit
//{
//public:
//    CContextMenuForwarder(IUnknown *);
//
//    STDMETHODIMP QueryInterface(REFIID riid, void **ppv);
//    STDMETHODIMP_(ULONG) AddRef(void);
//    STDMETHODIMP_(ULONG) Release(void);
//
//    // *** IContextMenu ***
//    STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) { return _pcm->QueryContextMenu(hmenu, indexMenu, idCmdFirst, idCmdLast, uFlags); }
//    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO lpici) { return _pcm->InvokeCommand(lpici); }
//    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pwReserved, LPSTR pszName, UINT cchMax) { return _pcm->GetCommandString(idCmd, uType, pwReserved, pszName, cchMax); }
//
//    // *** IContextMenu2 ***
//    STDMETHODIMP HandleMenuMsg(UINT uMsg, WPARAM wParam, LPARAM lParam) { return _pcm2->HandleMenuMsg(uMsg, wParam, lParam); }
//
//    // *** IContextMenu3 ***
//    STDMETHODIMP HandleMenuMsg2(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *plResult) { return _pcm3->HandleMenuMsg2(uMsg, wParam, lParam, plResult); }
//
//    // *** IObjectWithSite ***
//    STDMETHOD(SetSite)(IUnknown *punkSite) { return _pows->SetSite(punkSite); }
//    STDMETHOD(GetSite)(REFIID riid, void **ppvSite) { return _pows->GetSite(riid, ppvSite); }
//
//    // *** IShellExtInit ***
//    HRESULT Initialize(const PIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID) { return _psei->Initialize(pidlFolder, pdtobj, hkeyProgID); }
//
//protected:
//    ~CContextMenuForwarder();
//
//private:
//    LONG _cRef;
//
//protected:
//
//    IUnknown *_punk;
//
//    IObjectWithSite *_pows;
//    IShellExtInit *_psei;
//    IContextMenu *_pcm;
//    IContextMenu2 *_pcm2;
//    IContextMenu3 *_pcm3;
//};
//
//CContextMenuForwarder::CContextMenuForwarder(IUnknown *punk) : _cRef(1)
//{
//    _punk = punk;
//    _punk->AddRef();
//
//    _punk->QueryInterface(IID_PPV_ARGS(&_pows));
//    _punk->QueryInterface(IID_PPV_ARGS(&_psei));
//    _punk->QueryInterface(IID_PPV_ARGS(&_pcm));
//    _punk->QueryInterface(IID_PPV_ARGS(&_pcm2));
//    _punk->QueryInterface(IID_PPV_ARGS(&_pcm3));
//
//}
//
//CContextMenuForwarder::~CContextMenuForwarder()
//{
//    if (_pows) _pows->Release();
//	if (_psei) _psei->Release();
//    if (_pcm)  _pcm->Release();
//    if (_pcm2) _pcm2->Release();
//    if (_pcm3) _pcm3->Release();
//    _punk->Release();
//}
//
//STDMETHODIMP CContextMenuForwarder::QueryInterface(REFIID riid, void **ppv)
//{
//    HRESULT hr = _punk->QueryInterface(riid, ppv);
//
//    if (SUCCEEDED(hr))
//    {
//        IUnknown *punkTmp = (IUnknown *)(*ppv);
//
//        static const QITAB qit[] = {
//            QITABENT(CContextMenuForwarder, IObjectWithSite),                     // IID_IObjectWithSite
//            QITABENT(CContextMenuForwarder, IContextMenu3),                       // IID_IContextMenu3
//            QITABENTMULTI(CContextMenuForwarder, IContextMenu2, IContextMenu3),   // IID_IContextMenu2
//            QITABENTMULTI(CContextMenuForwarder, IContextMenu, IContextMenu3),    // IID_IContextMenu
//			QITABENT(CContextMenuForwarder, IShellExtInit),                       // IID_IShellExtInit
//            { 0 },
//        };
//
//        HRESULT hrTmp = QISearch(this, qit, riid, ppv);
//
//        if (SUCCEEDED(hrTmp))
//        {
//            punkTmp->Release();
//        }
//        else
//        {
//            // RIPMSG(FALSE, "CContextMenuForwarder asked for an interface it doesn't support");
//            *ppv = NULL;
//            hr = E_NOINTERFACE;
//        }
//    }
//
//    return hr;
//}
//
//STDMETHODIMP_(ULONG) CContextMenuForwarder::AddRef(void)
//{
//    return InterlockedIncrement(&_cRef);
//}
//
//STDMETHODIMP_(ULONG) CContextMenuForwarder::Release(void)
//{
//    if (InterlockedDecrement(&_cRef))
//        return _cRef;
//
//    delete this;
//    return 0;
//}
//
//class CContextMenuWithoutVerbs : CContextMenuForwarder
//{
//public:
//    STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
//
//protected:
//    CContextMenuWithoutVerbs(IUnknown *punk, LPCWSTR pszVerbList);
//
//    friend HRESULT Create_ContextMenuWithoutVerbs(IUnknown *punk, LPCWSTR pszVerbList, REFIID riid, void **ppv);
//
//private:
//    LPCWSTR _pszVerbList;
//};
//
//CContextMenuWithoutVerbs::CContextMenuWithoutVerbs(IUnknown *punk, LPCWSTR pszVerbList) : CContextMenuForwarder(punk)
//{
//    _pszVerbList = pszVerbList; // no reference - this should be a pointer to the code segment
//}
//
//HRESULT Create_ContextMenuWithoutVerbs(IUnknown *punk, LPCWSTR pszVerbList, REFIID riid, void **ppv)
//{
//    HRESULT hr = E_OUTOFMEMORY;
//
//    *ppv = NULL;
//
//    if (pszVerbList)
//    {
//        CContextMenuWithoutVerbs *p = new CContextMenuWithoutVerbs(punk, pszVerbList);
//        if (p)
//        {
//            hr = p->QueryInterface(riid, ppv);
//            p->Release();
//        }
//    }
//
//    return hr;
//}
//
//HRESULT CContextMenuWithoutVerbs::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
//{
//    HRESULT hr; // ebx
//    const WCHAR *i; // esi
//
//    hr = CContextMenuForwarder::QueryContextMenu(hmenu, indexMenu, idCmdFirst, idCmdLast, uFlags);
//    if (hr >= 0)
//    {
//        for (i = this->_pszVerbList; *i; i += lstrlenW(i) + 1)
//            ContextMenu_DeleteCommandByName(_pcm, hmenu, idCmdFirst, i);
//    }
//    return hr;
//}

HRESULT CSearchOpenView::GetContextMenu(IContextMenu *pcmIn, IContextMenu **ppcmOut)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
	//return Create_ContextMenuWithoutVerbs(pcmIn, TEXT("rename"), IID_PPV_ARGS(ppcmOut));
}

// Requires CSrchViewAccStateWrapper which requires CAccessibleWrapperBase which we do not have yet.
HRESULT CSearchOpenView::CreateAccessibleObject(HWND hwnd, LONG idObject, REFIID riid, void **ppv)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CSearchOpenView::GetEnumerationTimeout(DWORD *pdwTimeout)
{
	*pdwTimeout = -1;
	return S_OK;
}

HRESULT CSearchOpenView::GetViewFlags(BROWSER_VIEW_FLAGS *pbvf)
{
	*pbvf = BVF_NOLINKOVERLAY;
	return S_OK;
}

#pragma region Resource String Helpers

// Thanks to ep_taskbar by @amrsatrio for these functions

DWORD GetLastErrorError()
{
	DWORD result = GetLastError();
	return result == ERROR_SUCCESS ? 1 : result;
}

HRESULT HRESULTFromLastErrorError()
{
	DWORD error = GetLastError();
	if (error != ERROR_SUCCESS && (int)error <= 0)
		return (HRESULT)GetLastErrorError();
	else
		return (HRESULT)((GetLastErrorError() & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000);
}

HRESULT ResourceStringFindAndSizeEx(HMODULE hInstance, UINT uId, WORD wLanguage, const WCHAR **ppch, WORD *plen)
{
	HRESULT hr;

	HRSRC hRsrc = FindResourceEx(hInstance, RT_STRING, MAKEINTRESOURCE((uId >> 4) + 1), wLanguage);
	if (hRsrc)
	{
		HGLOBAL hgRsrc = LoadResource(hInstance, hRsrc);
		if (hgRsrc)
		{
			WORD *pRsrc = (WORD *)LockResource(hgRsrc);
			if (pRsrc)
			{
				for (UINT i = uId & 0xF; i; --i)
					pRsrc += *pRsrc + 1;

				if (ppch)
				{
					WORD len = *pRsrc;
					*ppch = len ? (const WCHAR*)pRsrc + 1 : nullptr;
					*plen = len;
				}
				else if (plen)
				{
					*plen = *pRsrc;
				}

				hr = S_OK;
			}
			else
			{
				hr = E_FAIL;
			}
		}
		else
		{
			hr = HRESULTFromLastErrorError();
		}
	}
	else
	{
		hr = HRESULTFromLastErrorError();
	}

	return hr;
}

template <typename T>
HRESULT TResourceStringAllocCopyEx(
	HMODULE hInstance,
	UINT uId,
	WORD wLanguage,
	HRESULT(CALLBACK *pfnAlloc)(HANDLE, SIZE_T, T *),
	HANDLE hHeap,
	T *pt)
{
	*pt = nullptr;

	const WCHAR *rgch;
	WORD len;
	HRESULT hr = ResourceStringFindAndSizeEx(hInstance, uId, wLanguage, &rgch, &len);
	if (SUCCEEDED(hr))
	{
		SIZE_T elemSize = sizeof(*pt);

		SIZE_T cb = elemSize * len;
		T t;
		hr = pfnAlloc(hHeap, cb + elemSize, &t);
		if (SUCCEEDED(hr))
		{
			memcpy(t, rgch, cb);
			t[cb / elemSize] = 0;
			*pt = t;
		}
	}

	return hr;
}

int SysAllocCb(SIZE_T cb, WCHAR **ppsz)
{
	if (cb < sizeof(WCHAR))
		return E_INVALIDARG;
	WCHAR *psz = SysAllocStringByteLen(nullptr, (UINT)(cb - sizeof(WCHAR)));
	HRESULT hr = psz ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		*ppsz = psz;
	}
	return hr;
}

HRESULT CALLBACK ResourceStringAllocCopyExSysAlloc(HANDLE hHeap, SIZE_T cb, WCHAR **ppsz)
{
	return SysAllocCb(cb, ppsz);
}

HRESULT ResourceStringSysAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR **ppsz)
{
	return TResourceStringAllocCopyEx(hModule, uId, wLanguage, ResourceStringAllocCopyExSysAlloc, nullptr, ppsz);
}

HRESULT ResourceStringSysAllocCopy(HINSTANCE hModule, UINT uId, WCHAR **ppsz)
{
	return ResourceStringSysAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

template <typename T>
HRESULT TLocalAllocArrayEx(UINT uFlags, SIZE_T uBytes, T **out)
{
	T *p = (T *)LocalAlloc(uFlags, uBytes);
	HRESULT hr = p ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		*out = p;
	}
	return hr;
}

HRESULT CALLBACK _ResourceStringAllocCopyExLocalAlloc(HANDLE hHeap, SIZE_T cb, WCHAR **ppsz)
{
	return TLocalAllocArrayEx(0, cb, (BYTE **)ppsz);
}

HRESULT ResourceStringLocalAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR **ppsz)
{
	return TResourceStringAllocCopyEx(hModule, uId, wLanguage, _ResourceStringAllocCopyExLocalAlloc, nullptr, ppsz);
}

HRESULT ResourceStringLocalAllocCopy(HINSTANCE hModule, UINT uId, WCHAR **ppsz)
{
	return ResourceStringLocalAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

#pragma endregion

enum SCHEDULERFLAGS
{
	SCHF_DEFAULT = 0x0,
	SCHF_UITHREADS = 0x1,
	SCHF_NOADDREFLIBS = 0x2,
};

DEFINE_ENUM_FLAG_OPERATORS(SCHEDULERFLAGS);

MIDL_INTERFACE("a8272e00-a569-40d2-9dac-b7756fa092c4")
IShellTaskSchedulerSettings : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE SetWorkerThreadCountMax(int) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetWorkerThreadCountMax(int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetWorkerThreadPriority(int) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetFlags(SCHEDULERFLAGS, SCHEDULERFLAGS) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFlags(SCHEDULERFLAGS *) = 0;
};

// d:\\win7m2\\shell\\lib\\cul\\registryhelpers.cpp
LSTATUS __stdcall SHRegGetDWORDW(HKEY hkey, LPCWSTR pszSubKey, LPCWSTR pszValue, DWORD *pdwData)
{
	DWORD pcbData; // [esp+8h] [ebp-4h] BYREF

	//if (!hkey)
	//{
	//	CULAssertOutputDebugString((int)L"d:\\win7m2\\shell\\lib\\cul\\registryhelpers.cpp", 155, (int)L"hkey");
	//	__debugbreak();
	//}
	// 
	//if (!pdwData)
	//{
	//	CULAssertOutputDebugString((int)L"d:\\win7m2\\shell\\lib\\cul\\registryhelpers.cpp", 156, (int)L"pdwData");
	//	__debugbreak();
	//}
	pcbData = 4;
	LSTATUS lr = SHRegGetValueW(hkey, pszSubKey, pszValue, 16, 0, pdwData, &pcbData);
	//if (lr == ERROR_MORE_DATA)
	//{
	//	CULAssertOutputDebugString(
	//		(int)L"d:\\win7m2\\shell\\lib\\cul\\registryhelpers.cpp",
	//		161,
	//		(int)L"lr != ERROR_MORE_DATA");
	//	__debugbreak();
	//}
	if (lr > 0)
		return (unsigned __int16)lr | 0x80070000;
	return lr;
}

#define HYBRID_CODE

// 8099904b-0b91-4906-a89d-11de2bd8f737
DEFINE_GUID(CLSID_StartMenuPathCompleteQueryCache, 0x8099904B, 0x0B91, 0x4906, 0xA8, 0x9D, 0x11, 0xDE, 0x2B, 0xD8, 0xF7, 0x37);

// #define SIMULATE_QUERY_PARSER_FAILURE

HRESULT CSearchOpenView::Initialize(HWND hwnd)
{
	HTHEME hTheme;

	HRESULT hr = _CreateExplorerBrowser(hwnd);
	if (SUCCEEDED(hr) && (_hwnd = hwnd, (hTheme = _hTheme) != nullptr))
	{
		GetThemeMargins(hTheme, nullptr, SPP_SEARCHVIEW, 0, TMT_CONTENTMARGINS, nullptr, &_margins);
	}
	else
	{
		_margins.cxLeftWidth = SHGetSystemMetricsScaled(SM_CXEDGE) * 2;
		_margins.cxRightWidth = SHGetSystemMetricsScaled(SM_CXEDGE) * 2;
	}

	if (SUCCEEDED(hr))
	{
		field_C4 = 1;

		LPITEMIDLIST pidlStartMenuProvider;
		hr = SHParseDisplayName(L"shell:::{daf95313-e44d-46af-be1b-cbacea2c3065}", nullptr, &pidlStartMenuProvider, 0, nullptr);
		if (SUCCEEDED(hr))
		{
			hr = SHCreateShellItemArrayFromIDLists(1, (LPCITEMIDLIST*)&pidlStartMenuProvider, &_psiaStartMenuProvider);
			ILFree(pidlStartMenuProvider);
		}
		if (SUCCEEDED(hr))
		{
			LPITEMIDLIST pidlStartMenuAutoComplete;
			hr = SHParseDisplayName(L"shell:::{e345f35f-9397-435c-8f95-4e922c26259e}", nullptr, &pidlStartMenuAutoComplete, 0, nullptr);
			if (SUCCEEDED(hr))
			{
				hr = SHCreateShellItemArrayFromIDLists(1, (LPCITEMIDLIST*)&pidlStartMenuAutoComplete, &_psiaStartMenuAutoComplete);
				ILFree(pidlStartMenuAutoComplete);
			}
		}

		if (SUCCEEDED(hr))
		{
			hr = _peb->Advise(this, &_dwCookie);
		}
		if (SUCCEEDED(hr))
		{
			hr = CoCreateInstance(CLSID_ShellTaskScheduler, nullptr, CLSCTX_INPROC, IID_PPV_ARGS(&_psched));
		}
		if (SUCCEEDED(hr))
		{
			_psched->Status(ITSSFLAG_KILL_ON_DESTROY, -2);

			IShellTaskSchedulerSettings* pstss;
			if (SUCCEEDED(_psched->QueryInterface(IID_PPV_ARGS(&pstss))))
			{
				pstss->SetWorkerThreadCountMax(2);
				pstss->Release();
			}
			hr = CoCreateInstance(CLSID_StartMenuPathCompleteQueryCache, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_psmqc));
		}
		if (SUCCEEDED(hr))
		{
			hr = _psmqc->InitializeCache();
		}
		if (SUCCEEDED(hr))
		{
#ifdef SIMULATE_QUERY_PARSER_FAILURE && _DEBUG // EXEX-Vista(allison): maybe todo?
			if (!SHRegGetBoolUSValueW(REGSTR_PATH_STARTPANE_SETTINGS, L"Start_SimulateQueryParserFailure", FALSE, FALSE))
			{
				InitializeQueryParser(GetUserDefaultUILanguage(), &field_54, IID_PPV_ARGS(&_pqp));
			}
#endif
			return SHCreateConditionFactory(IID_PPV_ARGS(&_pcf));
		}
	}

	return hr;
}

TASKOWNERID stru_1015F74 = { 228903948u, 12084u, 19201u, { 162u, 153u, 110u, 250u, 217u, 237u, 114u, 28u } };

HRESULT CSearchOpenView::AddPathCompletionTask(LPCWSTR pszPath)
{
	LPWSTR v7;
	HRESULT hr = SHStrDup(pszPath, &v7);
	if (hr >= 0)
	{
		CPathCompletionTask *pct = new CPathCompletionTask(_punkSite, _peb, _hwnd, v7);
		if (pct)
		{
			_psched->RemoveTasks(stru_1015F74, 0, 0);
			hr = _psched->AddTask(pct, stru_1015F74, 0, 268435712);
			pct->Release();
		}
		else
		{
			CoTaskMemFree(v7);
		}
	}
	return hr;
}

DWORD CSearchOpenView::s_ExecuteCommandLine(LPVOID lpv)
{
	return 0; // EXEX-Vista(allison): TODO.
}

DWORD CSearchOpenView::s_ExecuteIDList(LPVOID lpv)
{
	return 0; // EXEX-Vista(allison): TODO.
}

LRESULT CSearchOpenView::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef CONFIRMED_WORKING_CODE
	LRESULT lres; // eax MAPDST

	LPWINDOWPOS lpwp = reinterpret_cast<LPWINDOWPOS>(lParam);

	CSearchOpenView *self = reinterpret_cast<CSearchOpenView *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	IUnknown *punk = NULL;
	if (self)
	{
		self->QueryInterface(IID_PPV_ARGS(&punk));
	}

	if (uMsg > 0x81)
	{
		switch (uMsg)
		{
			case 0x82u:
				lres = self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
				goto LABEL_39;
			case 0x401u:
				self->_PathCompleteUpdate((CPathCompleteInfo *)wParam, (LPITEMIDLIST)lParam);
				goto LABEL_35;
			case 0x402u:
				self->_UpdateIndexState(wParam);
				goto LABEL_35;
			case 0x403u:
				if (!StrCmp(self->field_A8, (PCWSTR)lParam))
				{
					self->_UpdateTopMatch(UTM_REASON_1);
					if (self->field_8C)
					{
						self->_CancelNavigation();
					}
				}
				goto LABEL_35;
		}
		goto LABEL_29;
	}
	if (uMsg != 0x81)
	{
		if (uMsg == 1)
		{
			lres = self->_OnCreate(hwnd, uMsg, wParam, lParam);
			goto LABEL_39;
		}
		if (uMsg == 0x14)
		{
			lres = self->_OnEraseBkGnd(hwnd, (HDC)wParam);
			goto LABEL_39;
		}
		if (uMsg != 0x47)
		{
			if (uMsg != 0x4E)
			{
				if (uMsg == 0x7B && (DWORD)lParam == 0xFFFFFFFF)
				{
					self->_DoKeyBoardContextMenu();
				LABEL_35:
					if (punk)
					{
						punk->Release();
					}
					return 0;
				}
				goto LABEL_29;
			}
			lres = self->_OnNotify(hwnd, 0x4Eu, wParam, lParam);
		LABEL_39:
			if (punk)
			{
				punk->Release();
			}
			return lres;
		}

		if (lParam && (lpwp->flags & SWP_NOSIZE) == 0)
		{
			lres = self->_OnSize(lpwp);
			goto LABEL_39;
		}
	LABEL_29:
		lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
		goto LABEL_39;
	}
	CSearchOpenView *psopv = new CSearchOpenView();
	if (!psopv)
	{
		goto LABEL_29;
	}
	SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)psopv);
	psopv->_hTheme = PaneDataFromCreateStruct(lParam)->hTheme;
	if (punk)
	{
		punk->Release();
	}
	return 1;
#else
	LRESULT lres; // eax MAPDST

	LPWINDOWPOS lpwp = reinterpret_cast<LPWINDOWPOS>(lParam);

	CSearchOpenView *self = reinterpret_cast<CSearchOpenView *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	IUnknown *punk = NULL;
	if (self)
	{
		self->QueryInterface(IID_PPV_ARGS(&punk));
	}

	if (uMsg > 0x81)
	{
		switch (uMsg)
		{
			case 0x82u:
				lres = self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
				goto LABEL_39;
			case 0x401u:
				self->_PathCompleteUpdate((CPathCompleteInfo*)wParam);
				goto LABEL_35;
			case 0x402u:
				self->_UpdateIndexState(wParam);
				goto LABEL_35;
			case 0x403u:
				if (!StrCmp(self->field_A8, (LPCWSTR)lParam))
				{
					self->_UpdateTopMatch(UTM_REASON_1);
					if (self->field_8C)
					{
						self->_CancelNavigation();
					}
				}
				goto LABEL_35;
		}
		goto LABEL_29;
	}

	if (uMsg == 0x81)
	{
		CSearchOpenView *psopv = new CSearchOpenView();
		if (!psopv)
		{
			goto LABEL_29;
		}
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)psopv);
		psopv->_hTheme = PaneDataFromCreateStruct(lParam)->hTheme;
		if (punk)
		{
			punk->Release();
		}
		return 1;
	}

	if (uMsg == 1)
	{
		lres = self->_OnCreate(hwnd, uMsg, wParam, lParam);
		goto LABEL_39;
	}

	if (uMsg == 0x14)
	{
		lres = self->_OnEraseBkGnd(hwnd, (HDC)wParam);
		goto LABEL_39;
	}

	if (uMsg == 0x47)
	{
		if (lParam && (lpwp->flags & SWP_NOSIZE) == 0)
		{
			lres = self->_OnSize(lpwp);
			goto LABEL_39;
		}
	LABEL_29:
		lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
		goto LABEL_39;
	}

	if (uMsg == WM_NOTIFY)
	{
		lres = self->_OnNotify(hwnd, uMsg, wParam, lParam);
	LABEL_39:
		if (punk)
		{
			punk->Release();
		}
		return lres;
	}

	if (uMsg == 0x7B && IS_WM_CONTEXTMENU_KEYBOARD(lParam))
	{
		self->_DoKeyBoardContextMenu();
	LABEL_35:
		if (punk)
		{
			punk->Release();
		}
		return 0;
	}
	goto LABEL_29;

#endif
}

LRESULT CSearchOpenView::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lres = 0;
	IUnknown_Set(&PaneDataFromCreateStruct(lParam)->punk, SAFECAST(this, IServiceProvider *));
	if (SUCCEEDED(Initialize(hwnd)))
	{
		return lres;
	}
	return -1;
}

LRESULT CSearchOpenView::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);
	if (pnm)
	{
		switch (pnm->code)
		{
			case 201u:
				HandleApplyRegion(hwnd, _hTheme, (SMNMAPPLYREGION*)pnm, SPP_SEARCHVIEW, 0);
				break;
			case 215u:
				return _OnSMNFindItem((SMNDIALOGMESSAGE*)pnm);
			case 223u:
				return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
			case NM_KILLFOCUS:
			{
				IShellView2* psv2;
				if (_pFolderView && SUCCEEDED(_pFolderView->QueryInterface(IID_PPV_ARGS(&psv2))))
				{
					psv2->SelectAndPositionItem(nullptr, SVSI_DESELECTOTHERS, nullptr);
					psv2->Release();
				}
				break;
			}
		}
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CSearchOpenView::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
	LRESULT lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
	if (this)
	{
		this->Release();
	}
	return lres;
}

MIDL_INTERFACE("dee61571-6f83-4a7b-b8c0-ca2ca6f76bf8")
IHitTestView : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE HitTestItem(POINT, int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetItemRect(int, RECT *) = 0;
	virtual HRESULT STDMETHODCALLTYPE MapRect(int, HWND, RECT *) = 0;
};

LRESULT CSearchOpenView::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
	// esi
	// eax
	IFolderView2 *v5; // edi
	int v6; // eax
	int v9; // eax
	WPARAM wParam; // eax
	IFolderView2 *v11; // ecx
	HWND Focus; // eax
	MSG *pmsg; // ecx
	MSG *v14; // esi
	IHitTestView *phtv; // [esp+10h] [ebp-20h] BYREF
	// [esp+14h] [ebp-1Ch]
	//CPPEH_RECORD ms_exc; // [esp+18h] [ebp-18h]
	int v18; // [esp+38h] [ebp+8h] SPLIT BYREF
	int v19; // [esp+38h] [ebp+8h] SPLIT BYREF
	int v20; // [esp+38h] [ebp+8h] SPLIT BYREF
	int iCurSel; // [esp+38h] [ebp+8h] SPLIT BYREF
	int iCurSel1; // [esp+38h] [ebp+8h] SPLIT BYREF

	LRESULT lRes = 0;
	switch (pdm->flags & 0xF)
	{
		case 0u:
		case 1u:
		{
			IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 310, 0, 0, 0);
			Focus = GetFocus();
			pmsg = pdm->pmsg;
			if (pmsg)
			{
				pmsg->hwnd = Focus;
				pdm->flags |= 0x1000u;
			}
			goto LABEL_34;	
		}
		case 2u:
		case 7u:
		{
			if (_pFolderView && _pFolderView->QueryInterface(IID_PPV_ARGS(&phtv)) >= 0)
			{
				v19 = -1;
				if (phtv->HitTestItem(pdm->pt, &v19) < 0 && (pdm->flags & 0x800) != 0)
				{
					_pFolderView->GetVisibleItem(-1, _viewMode == VIEWMODE_PATHCOMPLETE, &v20);
				}
				v9 = v19;
				pdm->itemID = v19;
				
				lRes = v9 >= 0;
				if ((pdm->flags & 0x10) != 0)
					dword9C = v9;

				phtv->Release();
			}
			return lRes;	
		}
		case 3u:
		case 4u:
		{
			v5 = _pFolderView;
			if (!v5)
				return lRes;
			v18 = -1;
			v5->GetVisibleItem(-1, pdm->flags & 4, &v18);
			goto LABEL_4;	
		}
		case 5u:
		{
			if (!_pFolderView)
				return lRes;
			iCurSel = _GetCurSel();
			wParam = pdm->pmsg->wParam;
			if (wParam == 38)
			{
				_pFolderView->GetVisibleItem(iCurSel, 1, &iCurSel);
			}
			else if (wParam == 40)
			{
				_pFolderView->GetVisibleItem(iCurSel, 0, &iCurSel);
			}
			else
			{
				iCurSel = -1;
			}
			LABEL_4:
				v6 = v18;
			pdm->itemID = v18;
			return v6 >= 0;	
		}
		case 6u:
		{
			if (field_88 ||_viewMode)
			{
				v11 = _pFolderView;
				if (v11)
				{
					iCurSel1 = -1;
					if ((pdm->flags & 0x400) != 0)
					{
						iCurSel1 = _GetCurSel();
					}
					else if (v11->QueryInterface(IID_PPV_ARGS( & phtv)) >= 0)
					{
						phtv->HitTestItem(pdm->pt, &iCurSel1);
						phtv->Release();
					}
					if (iCurSel1 >= 0 && _ActivateItem(iCurSel1) >= 0)
						LABEL_31:
					lRes = 1;
				}
			}
			else
			{
				field_8C = 1;
			}
			break;	
		}
		case 8u:
			goto LABEL_34;
		case 9u:
		case 0xAu:
			v14 = pdm->pmsg;
			if (!v14 || v14->message != 256 || v14->wParam != 32)
				goto LABEL_31;
		LABEL_34:
			lRes = 0;
			break;
		case 0xBu:
			return 0;
		default:
			//if (CcshellAssertFailedW(
			//	L"d:\\longhorn\\shell\\explorer\\desktop2\\srchview.cpp",
			//	2121,
			//	L"!\"Unknown SMNDM command\"",
			//	0))
			//{
			//	AttachUserModeDebugger();
			//	do
			//	{
			//		__debugbreak();
			//		ms_exc.registration.TryLevel = -2;
			//	} while (dword_108BA00);
			//}
			return lRes;
	}
	return lRes;
}

LRESULT CSearchOpenView::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
	LRESULT lres = _OnSMNFindItemWorker(pdm);
	if (lres)
	{
		if ((pdm->flags & 0x900) != 0)
		{
			LPITEMIDLIST pidl;
			if (_pFolderView && SUCCEEDED(_pFolderView->Item(pdm->itemID, &pidl)))
			{
				int v4 = (pdm->flags & 0xF) != 7 ? 1025 : 1;
				if ((pdm->flags & 0x100) != 0)
				{
					v4 |= 0x10;
				}
			
				_pFolderView->SelectAndPositionItems(1, (LPCITEMIDLIST *)&pidl, 0, v4 | 0x8);
				ILFree(pidl);
			
				int v5 = pdm->flags & 0xF;
				if (v5 == 5 || v5 == 3 || v5 == 4)
				{
					_UpdateOpenBoxText();
				}
				pdm->flags = pdm->flags & 0xFFFBF6FF | 0x40000;	
			}
		}
	}
	else
	{
		pdm->flags |= 0x4000u;
		pdm->pt.x = 0;
		pdm->pt.y = 0;

		int iCurSel = _GetCurSel();
		if (iCurSel > 0 )
		{
			IHitTestView *phtv;
			if (this->_pFolderView && this->_pFolderView->QueryInterface(IID_PPV_ARGS(&phtv)) >= 0)
			{
				RECT rc;
				if (phtv->GetItemRect(iCurSel, &rc) >= 0)
				{
					pdm->pt.x = (RECTWIDTH(rc)) / 2;
					pdm->pt.y = (RECTHEIGHT(rc)) / 2;
				}
				phtv->Release();	
			}
		}
	}
	return lres;
}

LRESULT CSearchOpenView::_OnSize(LPWINDOWPOS pwp)
{
	LRESULT lres = 0;
	if (_peb)
	{
		_SizeExplorerBrowser(pwp->cx, pwp->cy);
		lres = 1;
	}
	return lres;
}

int CSearchOpenView::_OnEraseBkGnd(HWND hwnd, HDC hdc)
{
	RECT rc;
	GetClientRect(hwnd, &rc);

	if (_hTheme)
	{
		if (IsCompositionActive())
		{
			SHFillRectClr(hdc, &rc, 0);
		}
		DrawThemeBackground(_hTheme, hdc, SPP_SEARCHVIEW, 0, &rc, nullptr);
	}
	else
	{
		SHFillRectClr(hdc, &rc, GetSysColor(COLOR_MENU));
	}
	return 1;
}

int CSearchOpenView::_GetCurSel()
{
	int iCurSel = -1;
	if (_pFolderView)
	{
		_pFolderView->GetSelectedItem(-1, &iCurSel);
	}
	return iCurSel;
}

int CSearchOpenView::_GetItemCount()
{
	HRESULT hr;
	int cItems = 0;
	if (_pFolderView)
	{
		hr = _pFolderView->ItemCount(SVGIO_ALLVIEW, &cItems);
	}
	return cItems;
}

struct EXECUTEINFO
{
	HWND field_0;
	LPVOID field_4;
	BOOL field_8;
};

HRESULT CSearchOpenView::_ActivateItem(int iItem)
{
	bool v3; // zf
	LPCWSTR v4; // eax
	BOOL v6; // eax
	EXECUTEINFO* v7; // edi
	BOOL v8; // eax
	HWND Parent; // eax
	NMHDR nm; // [esp+Ch] [ebp-454h] BYREF
	VARIANT varIn; // [esp+28h] [ebp-438h] BYREF
	LPWSTR pszDisplayName; // [esp+38h] [ebp-428h] BYREF
	IShellItem2 *psi; // [esp+3Ch] [ebp-424h] BYREF
	LPITEMIDLIST ppidl; // [esp+40h] [ebp-420h] BYREF
	int v16; // [esp+44h] [ebp-41Ch]
	LPWSTR v17; // [esp+44h] [ebp-41Ch] SPLIT BYREF
	HRESULT hr; // [esp+48h] [ebp-418h]
	WCHAR v19[260]; // [esp+4Ch] [ebp-414h] BYREF
	WCHAR szKeyWord[260]; // [esp+254h] [ebp-20Ch] BYREF

	//SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Result_Launch, 75);
	v3 = this->_viewMode == VIEWMODE_PATHCOMPLETE;
	this->field_8C = 0;
	ppidl = 0;
	memset(szKeyWord, 0, sizeof(szKeyWord));
	v16 = v3;
	if (iItem == -1)
	{
		(void)120; // Skopped telemetry StartMenu_Search_URL_Count
		hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenBox, &SID_SM_DV2ControlHost, 308, 0, nullptr, &varIn);
		if (hr < 0)
		{
			goto LABEL_38;
		}
		v4 = VariantToStringWithDefault(varIn, L"");
		hr = StringCchCopyW(v19, 260u, v4);
		/*SH*/ExpandEnvironmentStrings(v19, szKeyWord, 260);
		VariantClear(&varIn);
	}
	else
	{
		_InstrumentActivation(iItem);
		hr = _GetItemKeyWord(iItem, szKeyWord, 0x104u);
		if (hr < 0)
		{
			goto LABEL_6;
		}
	}

	v16 = 1;
	if (hr >= 0)
	{
	LABEL_15:
		v17 = 0;
		hr = SHStrDupW(szKeyWord, &v17);
		if (hr < 0)
			goto LABEL_38;

		EXECUTEINFO *v5 = new EXECUTEINFO;
		if (v5)
		{
			v5->field_0 = GetAncestor(this->_hwnd, 2u);
			v5->field_4 = v17;
			v6 = GetAsyncKeyState(VK_SHIFT) < 0 && GetAsyncKeyState(VK_CONTROL) < 0;
			v5->field_8 = v6;
			if (SHCreateThread(s_ExecuteCommandLine, v5, 8u, 0))
			{
				v17 = 0;
			}
			else
			{
				operator delete(v5);
				hr = 0x80004005;
			}
		}
		if (v17)
		{
			CoTaskMemFree(v17);
		}
		goto LABEL_34;
	}
LABEL_6:
	if (iItem >= 0)
	{
		hr = this->_pFolderView->GetItem(
			iItem,
			IID_IShellItem2,
			(void **)&psi);
		if (hr < 0)
			goto LABEL_38;
		hr = SHGetIDListFromObject((IUnknown *)psi, &ppidl);
		if (hr >= 0)
		{
			if (psi->GetDisplayName(SIGDN_FILESYSPATH, &pszDisplayName) < 0)
			{
				v16 = 0;
			}
			else
			{
				StringCchCopyW(szKeyWord, 260u, pszDisplayName);
				CoTaskMemFree(pszDisplayName);
			}
		}
		psi->Release();
	}
	if (hr < 0)
		goto LABEL_38;
	if (v16)
		goto LABEL_15;

	v7 = new EXECUTEINFO;
	if (v7)
	{
		v7->field_0 = GetAncestor(this->_hwnd, 2u);
		v7->field_4 = &ppidl->mkid.cb;
		v8 = GetAsyncKeyState(16) < 0 && GetAsyncKeyState(17) < 0;
		v7->field_8 = v8;
		if (!SHCreateThread(s_ExecuteIDList, v7, 8u, 0))
		{
			operator delete(v7);
			hr = 0x80004005;
			goto LABEL_38;
		}
		ppidl = 0;
	}
LABEL_34:
	if (hr >= 0 && this->_viewMode == VIEWMODE_PATHCOMPLETE && szKeyWord[0])
		_UpdateMRU(szKeyWord);
LABEL_38:
	Parent = GetParent(this->_hwnd);
	_SendNotify(Parent, 204u, &nm);
	ILFree(ppidl);
	return hr;
}

HRESULT CSearchOpenView::_AddCheckIndexerStaterTask()
{
	return S_OK; // EXEX-Vista(allison): TODO.
}

HRESULT CSearchOpenView::_AddCondition(IObjectArray *poa, REFPROPERTYKEY a3, CONDITION_OPERATION a4, LPCWSTR a5)
{
	return S_OK; // EXEX-Vista(allison): TODO.
}

struct SEARCHOPTIONINFO
{
	LPCWSTR pszKey;
	LPCWSTR pszValue;
	int fDefaultValue;
};

SEARCHOPTIONINFO c_rgsoi[] =
{
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\PrimaryProperties\\IndexedLocations",
		L"SearchOnly",
		FALSE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\PrimaryProperties\\UnindexedLocations",
		L"SearchOnly",
		TRUE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"SearchSubFolders",
		TRUE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"AutoWildCard",
		TRUE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"EnableNaturalQuerySyntax",
		FALSE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"WholeFileSystem",
		FALSE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"SystemFolders",
		FALSE
	},
	{
		L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Search\\Preferences",
		L"ArchivedFiles",
		FALSE
	}
};

int GetSearchOptionValue(DWORD so)
{
	return so < 8 && _SHRegGetBoolValueFromHKCUHKLM(c_rgsoi[so].pszKey, c_rgsoi[so].pszValue, c_rgsoi[so].fDefaultValue);
}

BOOL ShouldSearchIndexFullText()
{
	return GetSearchOptionValue(0) == 0;
}

#define COMPILE_ENUMOBJECTCOLLECTION

#ifdef COMPILE_ENUMOBJECTCOLLECTION

class CEnumerableObjectCollection
	: public IObjectCollection
	, public IEnumObjects
	, public IEnumUnknown
{
public:
	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void **ppv) override
	{
		static const QITAB qit[] =
		{
			QITABENT(CEnumerableObjectCollection, IObjectCollection),
			QITABENT(CEnumerableObjectCollection, IObjectArray),
			QITABENT(CEnumerableObjectCollection, IEnumObjects),
			QITABENT(CEnumerableObjectCollection, IEnumUnknown),
			{},
		};

		return QISearch(this, qit, riid, ppv);
	}

	STDMETHODIMP_(ULONG) AddRef() override
	{
		return InterlockedIncrement(&_cRef);
	}

	STDMETHODIMP_(ULONG) Release() override
	{
		LONG cRef = InterlockedDecrement(&_cRef);
		if (cRef == 0 && this)
		{
			delete this;
		}
		return cRef;
	}

	// *** IObjectArray ***
	STDMETHODIMP GetCount(UINT *a3) override
	{
		*a3 = _dpaObjects.GetPtrCount();
		return S_OK;
	}

	STDMETHODIMP GetAt(UINT iElement, REFIID riid, void **ppv) override
	{
		*ppv = NULL;

		HRESULT hr = E_INVALIDARG;

		IUnknown *punk = _dpaObjects.GetPtr(iElement);
		if (punk)
		{
			hr = punk->QueryInterface(riid, ppv);
		}
		return hr;
	}

	// *** IObjectCollection ***
	STDMETHODIMP AddObject(IUnknown *punk) override
	{
		HRESULT hr = E_OUTOFMEMORY;

		if (_dpaObjects || _dpaObjects.Create(_iGrowBy))
		{
			hr = _dpaObjects.AppendPtr(punk);
			if (SUCCEEDED(hr))
			{
				punk->AddRef();
			}
		}
		return hr;
	}

	STDMETHODIMP AddFromArray(IObjectArray *poaSource) override
	{
		UINT cSource;
		HRESULT hr = poaSource->GetCount(&cSource);
		for (UINT i = 0; SUCCEEDED(hr) && i < cSource; ++i)
		{
			IUnknown *punk;
			hr = poaSource->GetAt(i, IID_PPV_ARGS(&punk));
			if (SUCCEEDED(hr))
			{
				hr = AddObject(punk);
				punk->Release();
			}
		}
		return hr;
	}

	STDMETHODIMP RemoveObjectAt(UINT uiIndex) override
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP Clear() override
	{
		_iNext = 0;
		return S_OK;
	}

	// *** IEnumObjects ***
	STDMETHODIMP Next(ULONG celt, REFIID riid, void **rgelt, ULONG *pceltFetched) override
	{
		return E_NOTIMPL; // EXEX-Vista(allison): TODO.
	}

	STDMETHODIMP Skip(ULONG celt) override
	{
		return E_NOTIMPL;
	}

	STDMETHODIMP Reset() override
	{
		_iNext = 0;
		return S_OK;
	}

	STDMETHODIMP Clone(IEnumObjects **ppeo) override
	{
		return Clone(IID_PPV_ARGS(ppeo));
	}

	// *** IEnumUnknown ***
	STDMETHODIMP Next(unsigned long, IUnknown **, unsigned long *) override;
	STDMETHODIMP Clone(IEnumUnknown **ppeu) override
	{
		return Clone(IID_PPV_ARGS(ppeu));
	}

protected:
	HRESULT Clone(REFIID, void **);

public:
	HRESULT CreateInstance(int, _GUID &, void **);
	HRESULT CreateInstance(_GUID &, void **);

protected:
	CEnumerableObjectCollection(int, unsigned int);
	CEnumerableObjectCollection(const CEnumerableObjectCollection &other) = delete;
	~CEnumerableObjectCollection();

	HRESULT InitializeClone(CEnumerableObjectCollection *peocNew);

	LONG _cRef;
	int _iGrowBy;
	UINT _iNext;
	CDPA<IUnknown, CTContainer_PolicyRelease<IUnknown>> _dpaObjects;
};

#endif

STDAPI TextToCondition(const WCHAR *pszQuery, ICondition **ppc);

HRESULT CSearchOpenView::_CreateConditions(LPCWSTR pcszURL, ICondition **ppCondition)
{
#if 0
	// esi MAPDST
	PARSEDURL pu; // [esp+4h] [ebp-20h] BYREF

	*ppCondition = 0;

	IObjectArray *poa; // [esp+20h] [ebp-4h] BYREF
	HRESULT hr = CEnumUnknown::s_CreateInstance(IID_PPV_ARGS(&poa));
	if (hr >= 0)
	{
		hr = _AddCondition(poa, PKEY_ItemNameDisplay, COP_VALUE_STARTSWITH, pcszURL);
		if (hr < 0)
		{
		LABEL_19:
			poa->Release();
			return hr;
		}
		memset(&pu.pszProtocol, 0, 20);
		pu.cbSize = sizeof(pu);
		ParseURL(pcszURL, &pu);
		if (this->_fIsBrowsing)
		{
			goto LABEL_4;
		}
		if (pu.nScheme == 2 || pu.nScheme == 11 || !StrCmpNI(pcszURL, L"www.", 4))
		{
			hr = _AddCondition(poa, PKEY_ItemNameDisplay, COP_VALUE_STARTSWITH, pcszURL);
			if (hr < 0)
			{
				goto LABEL_19;
			}
			if (ShouldSearchIndexFullText())
			{
				hr = CSearchOpenView::_AddCondition(poa, PKEY_FullText, COP_VALUE_STARTSWITH, pcszURL);
			}
			if (hr < 0)
			{
				goto LABEL_19;
			}
			goto LABEL_18;
		}
		if (!this->_pqp)
		{
			goto LABEL_4;
		}
		hr = _Parse(poa, PKEY_ItemNameDisplay, pcszURL);
		if (hr < 0)
		{
			goto LABEL_4;
		}
		if (ShouldSearchIndexFullText())
		{
			hr = CSearchOpenView::_Parse(poa, PKEY_FullText, pcszURL);
		}
		if (hr < 0)
		{
		LABEL_4:
			hr = 1;
		}
	LABEL_18:
		hr = SHCreateAndOrConditionEx(1, poa, this->_pcf, IID_ICondition, ppCondition);
		goto LABEL_19;
	}
	return hr;
#endif
	if (!ppCondition)
		return E_POINTER;

	*ppCondition = nullptr;

	//// No input -> caller expects S_FALSE (sets empty text). Do not return S_OK without an out param.
	if (!pcszURL || !*pcszURL)
		return S_FALSE;

	//// Use existing helper that parses text into a structured ICondition.
	ICondition *pCond = nullptr;
	HRESULT hr = TextToCondition(pcszURL, &pCond);
	if (SUCCEEDED(hr))
	{
		if (pCond)
		{
			*ppCondition = pCond; // ownership: TextToCondition returns an AddRef'd pointer
		}
		else
		{
			// Defensive: don't return S_OK unless we actually provide a condition.
			hr = S_FALSE;
		}
	}

	return hr;
}

HRESULT CSearchOpenView::_CreateExplorerBrowser(HWND hwnd)
{
	HRESULT hr = CoCreateInstance(CLSID_ExplorerBrowser, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_peb));
	if (SUCCEEDED(hr))
	{
		IUnknown_SetSite(_peb, SAFECAST(this, IServiceProvider *));
		hr = _peb->SetOptions(EBO_NOWRAPPERWINDOW);
		if (SUCCEEDED(hr))
		{
			RECT rc = {};

			FOLDERSETTINGS fs;
			fs.ViewMode = FVM_DETAILS;
			fs.fFlags = FWF_NOBROWSERVIEWSTATE;
			hr = _peb->Initialize(hwnd, &rc, &fs);
			if (SUCCEEDED(hr))
			{
				GetClientRect(hwnd, &rc);
				_SizeExplorerBrowser(RECTWIDTH(rc), RECTHEIGHT(rc));
			}
		}
	}
	return hr;
}

HRESULT ResourceStringCoAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR **ppsz);

HRESULT CSearchOpenView::_FilterView(IFilterView* pfv)
{
	HRESULT hr = E_FAIL;
	if (field_A8)
	{
		ICondition* pc;
		hr = _CreateConditions(field_A8, &pc);
		if (SUCCEEDED(hr))
		{
			LPWSTR pszText;
			if (hr == S_FALSE && SUCCEEDED(ResourceStringCoAllocCopyEx(g_hinstCabinet, 7050u, 0, &pszText)))
			{
				_peb->SetEmptyText(pszText);
				CoTaskMemFree(pszText);
			}

			IFilterCondition* pfc;
			hr = SHCreateFilter(L"StartMenu_Search", nullptr, PKEY_ItemNameDisplay, FCT_DEFAULT, pc, IID_PPV_ARGS(&pfc));
			if (SUCCEEDED(hr))
			{
				hr = pfv->FilterByCondition(pfc);
				pfc->Release();
			}
			pc->Release();
		}
	}
	return hr;
}

HRESULT CSearchOpenView::_GetItemKeyWord(int itemIndex, LPWSTR pszDest, UINT cchDest)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CSearchOpenView::_GetQueryFactoryWithSink(REFIID riid, void **ppv)
{
	if (!ppv)
		return E_POINTER;

	// We don't have the real query-factory implementation or the GUID mapping.
	// Return S_FALSE (a non-failing result) and leave *ppv == nullptr so callers
	// continue without a custom query factory.
	*ppv = nullptr;

	OutputDebugStringW(L"_GetQueryFactoryWithSink: no factory available, returning S_FALSE\n");
	return S_FALSE;
}

HRESULT CSearchOpenView::_GetSelectedItemParsingName(LPWSTR pszName, UINT cchName)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CSearchOpenView::_InitPathCompletePidlAutoList(LPCWSTR pszPath, PIDLIST_ABSOLUTE *ppidl)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

#pragma region misc 

struct IItemFilter;

typedef enum tagFILTERIDLISTTYPE
{
	FIT_STACK = 1,
	FIT_FILTER = 2,
	FIT_GROUP = 3,
} FILTERIDLISTTYPE;

enum CONDITIONSOURCEFLAGS
{
	SOURCE_NONE = 0x0,
	SOURCE_VISIBLEIN = 0x1,
	SOURCE_AUTOLIST = 0x2,
	SOURCE_FILTERSTACKOPS = 0x4,
};

DEFINE_ENUM_FLAG_OPERATORS(CONDITIONSOURCEFLAGS);

MIDL_INTERFACE("711b2cfd-93d1-422b-bdf4-69be923f2449")
IShellFolder3 : IShellFolder2
{
	virtual HRESULT STDMETHODCALLTYPE CreateFilteredIDList(IFilterCondition *, FILTERIDLISTTYPE, IPropertyStore *, ITEMID_CHILD **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFilteredIDListType(PCUITEMID_CHILD, FILTERIDLISTTYPE *) = 0;
	virtual HRESULT STDMETHODCALLTYPE ModifyFilteredIDList(PCUITEMID_CHILD, IFilterCondition *, ITEMID_CHILD **) = 0;
	virtual HRESULT STDMETHODCALLTYPE ReparentFilteredIDList(PCUIDLIST_RELATIVE, ITEMIDLIST_RELATIVE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateStackedIDList(REFPROPERTYKEY, ITEMIDLIST_ABSOLUTE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackedKey(PROPERTYKEY *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackData(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE EnumObjectsEx(HWND, IBindCtx *, DWORD, IItemFilter *, IEnumIDList **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetConditions(REFPROPERTYKEY, CONDITIONSOURCEFLAGS, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetAutoListFlags(DWORD *) = 0;
};

DEFINE_GUID(CLSID_ScopeFactory, 0x6746C347, 0x576B, 0x4F73, 0x90, 0x12, 0xCD, 0xFE, 0xEA, 0x25, 0x1B, 0xC4); // 6746c347-576b-4f73-9012-cdfeea251bc4

enum SCOPE_ITEM_TYPE
{
	SI_TYPE_INVALID = 0,
	SI_TYPE_INCLUDE = 1,
	SI_TYPE_EXCLUDE = 2,
};

enum SCOPE_ITEM_DEPTH
{
	SI_DEPTH_INVALID = 0,
	SI_DEPTH_SHALLOW = 1,
	SI_DEPTH_DEEP = 2,
};

enum SCOPE_ITEM_FLAGS
{
	SI_FLAG_DEFAULT = 0x0,
	SI_FLAG_FORCEEXHAUSTIVE = 0x1,
	SI_FLAG_KNOWNFOLDER = 0x2,
	SI_FLAG_FASTPROPERTIESONLY = 0x4,
	SI_FLAG_FASTITEMSONLY = 0x8,
	SI_FLAG_NOINFOBAR = 0x10,
	SI_FLAG_USECHILDSCOPES = 0x20,
	SI_FLAG_FASTPROVIDERSONLY = 0x40,
	SI_FLAG_WINRTDATAMODEL = 0x80,
	SI_FLAG_USESCOPEITEMDEPTH = 0x100,
};

DEFINE_ENUM_FLAG_OPERATORS(SCOPE_ITEM_FLAGS);

MIDL_INTERFACE("54410b83-6787-4418-9735-5aaaabe83a9a")
IScopeFactory : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE CreateScope(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateScopeFromShellItemArray(IShellItemArray *, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateScopeFromIDLists(UINT, const ITEMIDLIST_ABSOLUTE *const *, SCOPE_ITEM_TYPE, SCOPE_ITEM_DEPTH, SCOPE_ITEM_FLAGS, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateScopeItemFromIDList(const ITEMIDLIST_ABSOLUTE *, SCOPE_ITEM_TYPE, SCOPE_ITEM_DEPTH, SCOPE_ITEM_FLAGS, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateScopeItemFromKnownFolder(REFGUID, SCOPE_ITEM_TYPE, SCOPE_ITEM_DEPTH, SCOPE_ITEM_FLAGS, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateScopeItemFromShellItem(IShellItem *, SCOPE_ITEM_TYPE, SCOPE_ITEM_DEPTH, SCOPE_ITEM_FLAGS, REFIID, void **) = 0;
};

struct IScopeItem;

enum SCOPE_FLATTEN_TYPE
{
	SF_DEFAULT = 0x0,
	SF_GROUP_EXCLUSIONS = 0x1000,
};

DEFINE_ENUM_FLAG_OPERATORS(SCOPE_FLATTEN_TYPE);

enum SCOPE_CMP_MODE
{
	SCMPM_DEFAULT = 0,
	SCMPM_TYPEANDDEPTH = 1,
};

MIDL_INTERFACE("655d1685-2bfd-4f7f-ad22-5ab61c8d8798")
IScope : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE AddScopeItem(IScopeItem *) = 0;
	virtual HRESULT STDMETHODCALLTYPE RemoveScopeItem(IScopeItem *) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsItemIncluded(IScopeItem *, int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetKnownScopeID(GUID *) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsEmpty() = 0;
	virtual HRESULT STDMETHODCALLTYPE Flatten(SCOPE_FLATTEN_TYPE, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsEqual(IScope *, SCOPE_CMP_MODE) = 0;
	virtual HRESULT STDMETHODCALLTYPE Clone(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE RemoveRootsMatchingUrlScheme(const WCHAR *) = 0;
};

class DECLSPEC_UUID("1C1800C1-3258-44C2-BE80-3DEADB6C5E39")
KindDescriptionFactory
{
};

MIDL_INTERFACE("fdada2fa-894d-47d8-ae78-adf1fd7f28df")
IKindDescriptionFactory : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE CreateKindDescription(const WCHAR *, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateKindList(IUnknown *, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateKindDescriptionForItem(IShellItem2 *, REFIID, void **) = 0;
};

DEFINE_GUID(CLSID_SearchIDListFactory, 0x6E682784, 0x1ECA, 0x4CF2, 0x98, 0x8D, 0x96, 0xB6, 0xE8, 0x9E, 0x9A, 0x4D); // 6e682784-1eca-4cf2-988d-96b6e89e9a4d

enum VIDFLAGS
{
	VIDF_DEFAULT = 0x0,
	VIDF_ENUMERABLE = 0x1,
	VIDF_VIRTUAL = 0x2,
};

DEFINE_ENUM_FLAG_OPERATORS(VIDFLAGS);

typedef enum tagVI_FOLDERTYPE
{
	VIFT_SEARCH = 0,
	VIFT_STACKED = 1,
} VI_FOLDERTYPE;

MIDL_INTERFACE("1d17905d-aeb6-4ac9-8cd5-2c103f186212")
IVisibleInDescription : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetCanonicalName(WCHAR **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDisplayName(WCHAR **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFlags(VIDFLAGS *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetProperties(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetItemOfTypeCondition(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDefaultFolderType(VI_FOLDERTYPE, GUID *) = 0;
	virtual HRESULT STDMETHODCALLTYPE IsEqual(IVisibleInDescription *, int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetCanonicalName(const WCHAR *) = 0;
};

MIDL_INTERFACE("d18f09e2-81c2-4217-a295-43b66fa4e495")
IVisibleInList : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetVisibleInCount(UINT *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetVisibleInAt(UINT, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetItemOfTypeCondition(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetProperties(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDefaultFolderType(VI_FOLDERTYPE, GUID *) = 0;
	virtual HRESULT STDMETHODCALLTYPE Intersect(IVisibleInList *, REFIID, void **) = 0;
};

struct IAutoListDescription;
struct IColumnList;

enum AUTOLISTFLAGS
{
	ALF_DEFAULT = 0x0,
	ALF_FORCEFULLTEXT = 0x2,
	ALF_SHOWHIDDEN = 0x4,
	ALF_SUBFOLDERS = 0x8,
	ALF_NO_ADVANCED_QUERY_SYNTAX = 0x20,
	ALF_FETCH_COLUMNS_IN_ENUM = 0x40,
	ALF_USE_SORTBY_IN_ENUM = 0x80,
	ALF_DYNAMIC_FILTERS = 0x100,
};

DEFINE_ENUM_FLAG_OPERATORS(AUTOLISTFLAGS);

struct AUTOLISTINIT
{
	WCHAR* pszParsingName;
	WCHAR* pszDisplayName;
	GUID ftid;
	AUTOLISTFLAGS dwAutoListFlags;
	DWORD dwFolderFlags;
	DWORD dwUITaskFlags;
	FOLDERLOGICALVIEWMODE flvm;
	int iIconSize;
	FOLDERLOGICALVIEWMODE flvmStack;
	int iIconSizeStack;
	PROPERTYKEY keyGroup;
	BOOL fSortGroupAscending;
	UINT cStackKeys;
	const PROPERTYKEY* rgStackKeys;
	UINT cSortColumns;
	const SORTCOLUMN* rgSortColumns;
	IColumnList* pcl;
	IVisibleInList* pvl;
	IScope* pscope;
	IShellItemArray* psiaSubQueries;
	ICondition* pc;
	IPropertyKeyStore* pksRelevance;
	UINT cRowsGroupSubset;
	IUnknown* punkStackData;
};

MIDL_INTERFACE("c0a6c367-c264-4385-a704-9088bdc3640e")
ISearchIDListFactory : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE CreateAutoList(const AUTOLISTINIT *, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE ParseRootSearchUrl(const WCHAR *, IBindCtx *, ITEMIDLIST_ABSOLUTE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateSearchIDListFromAutoList(IAutoListDescription *, INamedPropertyStore *, ITEMIDLIST_ABSOLUTE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateTransientVFolderIDList(IScope *, IShellItemArray *, ITEMIDLIST_ABSOLUTE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateSearchIDList(const WCHAR *, const WCHAR *, IScope *, IShellItemArray *, ITEMIDLIST_ABSOLUTE **) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateSearchWithNewFlags(IShellFolder3 *, AUTOLISTFLAGS, REFIID, void **) = 0;
};

struct IEnumUICommand;

enum GET_SCOPE_TYPE
{
	GST_DEFAULT = 0,
	GST_INCLUDE_SEARCHES = 1,
};

MIDL_INTERFACE("ffffffff-ffff-ffff-ffff-ffffffffffff")
IListDescription : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE GetDisplayName(WCHAR **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetParsingName(WCHAR **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackBy(UINT, PROPERTYKEY *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetGroupBy(PROPERTYKEY *, int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetSortColumn(UINT, PROPERTYKEY *, int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetViewMode(FOLDERLOGICALVIEWMODE *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetIconSize(int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetOverrideColumns(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetGroupSubsetRows(UINT *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetUICommands(IVisibleInList *, IUnknown *, IEnumUICommand **, IEnumUICommand **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFolderType(GUID *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetVisibleInList(REFIID, void **) = 0;
};

MIDL_INTERFACE("607c87f7-0696-4558-a368-de5e59cfe456")
IAutoListDescription : IListDescription
{
	virtual HRESULT STDMETHODCALLTYPE GetAutoListFlags(AUTOLISTFLAGS *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFolderFlags(DWORD *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetTaskFlags(DWORD *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackIconSize(int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetRelevanceProperties(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetConditions(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetScope(GET_SCOPE_TYPE, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetSubQueries(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackByData(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetStackViewMode(FOLDERLOGICALVIEWMODE *) = 0;
};

HRESULT SHCreateSingleKindList(const WCHAR *pszKind, REFIID riid, void **ppv)
{
	*ppv = nullptr;

	IKindDescriptionFactory *pkdf;
	HRESULT hr = CoCreateInstance(__uuidof(KindDescriptionFactory), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pkdf));
	if (SUCCEEDED(hr))
	{
		IVisibleInDescription *pvd;
		hr = pkdf->CreateKindDescription(pszKind, IID_PPV_ARGS(&pvd));
		if (SUCCEEDED(hr))
		{
			hr = pkdf->CreateKindList(pvd, riid, ppv);
			pvd->Release();
		}

		pkdf->Release();
	}

	return hr;
}

#pragma endregion

class CStartMenuQuerySink : public IStartMenuQuerySink
{
public:
	CStartMenuQuerySink(HWND hwnd);
	virtual ~CStartMenuQuerySink();

	//~ Begin IUnknown Interface
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;
	//~ End IUnknown Interface

	//~ Begin IStartMenuQuerySink Interface
	STDMETHODIMP SetStatusForQuery(const WCHAR *pszQuery, SM_QUERY_STATUS status) override;
	//~ End IStartMenuQuerySink Interface

	static HRESULT s_CreateInstance(HWND hwnd, REFIID riid, void **ppv);

private:
	LONG m_refCount;
	HWND m_hwnd;
	IUnknown *m_punkMarshal;
};

CStartMenuQuerySink::CStartMenuQuerySink(HWND hwnd)
	: m_refCount(1)
	, m_hwnd(hwnd)
	, m_punkMarshal(nullptr)
{
}

CStartMenuQuerySink::~CStartMenuQuerySink()
{
	IUnknown_SafeReleaseAndNullPtr(&m_punkMarshal);
}

HRESULT CStartMenuQuerySink::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] = {
		QITABENT(CStartMenuQuerySink, IStartMenuQuerySink),
		{},
	};

	return QISearch(this, qit, riid, ppvObject);
}

ULONG CStartMenuQuerySink::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

ULONG CStartMenuQuerySink::Release()
{
	ULONG refCount = InterlockedDecrement(&m_refCount);
	if (refCount == 0 && this)
		delete this;
	return refCount;
}

HRESULT CStartMenuQuerySink::SetStatusForQuery(const WCHAR *pszQuery, SM_QUERY_STATUS status)
{
	SendMessageW(m_hwnd, 0x403, 0, (LPARAM)pszQuery);
	return S_OK;
}

HRESULT CStartMenuQuerySink::s_CreateInstance(HWND hwnd, REFIID riid, void **ppv)
{
	*ppv = nullptr;
	HRESULT hr = E_OUTOFMEMORY;

	CStartMenuQuerySink *p = new CStartMenuQuerySink(hwnd);
	if (p)
	{
		hr = CoCreateFreeThreadedMarshaler(p, &p->m_punkMarshal);
		if (SUCCEEDED(hr))
		{
			hr = p->QueryInterface(riid, ppv);
		}

		p->Release();
	}

	return hr;
}

PROPERTYKEY PKEY_StartMenu_Group =
{
  {
	1272003389u,
	59019u,
	17644u,
	{ 137u, 238u, 118u, 17u, 120u, 157u, 64u, 112u }
  },
  100u
};

FOLDERTYPEID FOLDERTYPEID_Library =
{ 1269693544u, 50348u, 18198u, { 160u, 160u, 77u, 93u, 170u, 107u, 15u, 62u } };

HRESULT CSearchOpenView::_InitPidlAutoList(PIDLIST_ABSOLUTE *ppidl)
{
	*ppidl = 0;

	IScopeFactory *psf; // [esp+18h] [ebp-94h] BYREF
	HRESULT hr = CoCreateInstance(CLSID_ScopeFactory, 0, 1u, IID_PPV_ARGS(&psf));
	if (hr >= 0)
	{
		IScope *pscope;
		if (psf->CreateScopeFromShellItemArray(_psiaStartMenuProvider, IID_PPV_ARGS(&pscope)) != S_OK)
		{
			hr = 0x80004005;
		}
		else
		{
			IVisibleInList *pvl;
			hr = SHCreateSingleKindList(L"item", IID_PPV_ARGS(&pvl));
			if (hr >= 0)
			{
				IStartMenuQuerySink *psmqs;
				hr = CStartMenuQuerySink::s_CreateInstance(this->_hwnd, IID_PPV_ARGS(&psmqs));
				if (hr >= 0)
				{
					_RegisterQuerySink(psmqs);

					AUTOLISTINIT ali = { 0 };
					ali.ftid = FOLDERTYPEID_Documents; //FOLDERTYPEID_Library in vista, but doesnt exist anymore
					ali.dwFolderFlags = _GetFolderFlags();
					ali.keyGroup = PKEY_StartMenu_Group;
					ali.pscope = pscope;
					ali.pvl = pvl;
					ali.fSortGroupAscending = TRUE;
					ali.flvm = FLVM_FIRST;
					ali.iIconSize = 16;

					ResourceStringCoAllocCopyEx(g_hinstCabinet, 8246, 0, &ali.pszDisplayName);

					// Request optional query factory/sink. AUTOLISTINIT no longer has a field to hold it,
					// so accept the returned interface (if any) and release it immediately to avoid leaks.
					void *pvFactory = nullptr;
					hr = _GetQueryFactoryWithSink(IID_NULL, &pvFactory);
					// Treat S_FALSE (no factory) as non-fatal; callers continue when hr >= 0.
					if (pvFactory)
					{
						IUnknown *punkFactory = reinterpret_cast<IUnknown *>(pvFactory);
						punkFactory->Release();
						pvFactory = nullptr;
					}


					ISearchIDListFactory *psilf;
					hr = CoCreateInstance(CLSID_SearchIDListFactory, 0, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psilf));
					if (SUCCEEDED(hr))
					{
						IAutoListDescription* pald;
						hr = psilf->CreateAutoList(&ali, IID_PPV_ARGS(&pald));
						if (SUCCEEDED(hr))
						{
							hr = psilf->CreateSearchIDListFromAutoList(pald, 0, ppidl);
							if (SUCCEEDED(hr))
							{
								IUnknown_Set((IUnknown**)&_psmqs, psmqs);
							}
							pald->Release();
						}
						psilf->Release();
					}


					CoTaskMemFree(ali.pszDisplayName);
					if (FAILED(hr))
					{
						_RevokeQuerySink();
					}

					psmqs->Release();
				}
				pvl->Release();
			}
			pscope->Release();
		}
		psf->Release();
	}
	return hr;
}

HRESULT CSearchOpenView::_InitRegularAutoListItem(IShellItem **ppsi)
{
	*ppsi = NULL;

	IScopeFactory *psf = 0;
	HRESULT hr = CoCreateInstance(CLSID_ScopeFactory, 0, 1u, IID_PPV_ARGS(&psf));
	if (hr >= 0)
	{
		IScope *pscope = 0;
		if (psf->CreateScopeFromShellItemArray(_psiaStartMenuProvider, IID_PPV_ARGS(&pscope)) != S_OK)
		{
			hr = E_FAIL;
		}
		else
		{
			IVisibleInList *pvl = 0;
			hr = SHCreateSingleKindList(L"item", IID_PPV_ARGS(&pvl));
			if (hr >= 0)
			{
				IStartMenuQuerySink *psmqs = 0;
				hr = CStartMenuQuerySink::s_CreateInstance(_hwnd, IID_PPV_ARGS(&psmqs));
				if (hr >= 0)
				{
					_RegisterQuerySink(psmqs);

					AUTOLISTINIT ali = {};
					ali.ftid = FOLDERTYPEID_StartMenu;
					ali.dwFolderFlags = _GetFolderFlags();
					ali.pscope = pscope;
					ali.pvl = pvl;
					ResourceStringCoAllocCopyEx(g_hinstCabinet, 8246, 0, &ali.pszDisplayName);
					//hr = CSearchOpenView::_GetQueryFactoryWithSink(
					//	&_GUID_4467e729_4f73_49d4_afe7_5ccbffdeb8ae,
					//	(void **)&ali.field_78);
					if (hr >= 0)
					{
						ISearchIDListFactory *psilf = 0;
						hr = CoCreateInstance(CLSID_SearchIDListFactory, 0, 1u, IID_PPV_ARGS(&psilf));
						if (hr >= 0)
						{
							IAutoListDescription *pald = 0;
							hr = psilf->CreateAutoList(&ali, IID_PPV_ARGS(&pald));
							if (hr >= 0)
							{
								LPITEMIDLIST v13; // [esp+20h] [ebp-94h] SPLIT BYREF
								hr = psilf->CreateSearchIDListFromAutoList(pald, 0, &v13);
								if (hr >= 0)
								{
									hr = SHCreateItemFromIDList(v13, IID_PPV_ARGS(ppsi));
									if (hr >= 0)
									{
										IUnknown_Set((IUnknown **)&this->_psmqs, (IUnknown *)psmqs);
									}
									ILFree(v13);
								}
							}
							if (pald)
							{
								pald->Release();
							}
						}

						//(*(void(__stdcall **)(int))(*(_DWORD *)ali.field_78 + 8))(ali.field_78);
						if (psilf)
						{
							psilf->Release();
						}
					}

					CoTaskMemFree(ali.pszDisplayName);
					if (hr < 0)
					{
						_RevokeQuerySink();
					}
				}

				if (psmqs)
				{
					psmqs->Release();
				}
			}

			if (pvl)
			{
				pvl->Release();
			}
		}

		if (pscope)
		{
			pscope->Release();
		}
	}

	if (psf)
	{
		psf->Release();
	}
	return hr;
}

MIDL_INTERFACE("D92E45A1-C7C0-4380-BEA1-10BEB57D9610")
ILimitedItemsView : IUnknown
{
	STDMETHOD(LimitItemsToVisibleCount)(int) PURE;
	STDMETHOD(SetMaxItemCount)(UINT) PURE;
	STDMETHOD(GetMaxItemCount)(UINT *) PURE;
};

HRESULT CSearchOpenView::_LimitViewResults(IFolderView* pfv, int nMaxItems)
{
	ILimitedItemsView* pliv;
	HRESULT hr = pfv->QueryInterface(IID_PPV_ARGS(&pliv));
	if (SUCCEEDED(hr))
	{
		if (_hTheme || !nMaxItems)
		{
			pliv->LimitItemsToVisibleCount(nMaxItems);
			pliv->SetMaxItemCount(-1);
		}
		else
		{
			UINT nItems = 0;
			pliv->LimitItemsToVisibleCount(1);
			pliv->GetMaxItemCount(&nItems);
			pliv->SetMaxItemCount(nItems - 5);
			pliv->LimitItemsToVisibleCount(0);
		}
		pliv->Release();
	}
	return hr;
}

HRESULT CSearchOpenView::_Parse(IObjectArray *poa, REFPROPERTYKEY pkey, LPCWSTR psz)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CSearchOpenView::_RecreateBrowserObject()
{
	HRESULT hr = S_OK;

	if (field_84)
	{
		field_94 = 0;
		field_90 = 0;
		_ReleaseExplorerBrowser();

		if (_ppci)
		{
			delete _ppci;
			_ppci = NULL;
		}

		IUnknown_SafeReleaseAndNullPtr(&_pFolderView);
		_DisconnectShellView();
		hr = _CreateExplorerBrowser(_hwnd);

		// if the previous hr is FAILED(hr), or the "hr _peb->Advise(xxx)" is FAILED(hr), then release the browser
		//if (hr < 0 || (hr = _peb->Advise(static_cast<IExplorerBrowserEvents *>(this), &_dwCookie), hr < 0))
		//{
		//	_ReleaseExplorerBrowser();
		//}

		if (hr >= 0)
		{
			hr = _peb->Advise(static_cast<IExplorerBrowserEvents *>(this), &_dwCookie);
		}

		if (hr < 0)
		{
			_ReleaseExplorerBrowser();
		}
	}
	else
	{
		field_94 = 1;
	}
	return hr;
}

HRESULT MarshalToGIT(IUnknown *pUnk, REFIID riid, DWORD *pdwCookie)
{
	*pdwCookie = 0;

	IGlobalInterfaceTable *pgit;
	HRESULT hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pgit));
	if (SUCCEEDED(hr))
	{
		hr = pgit->RegisterInterfaceInGlobal(pUnk, riid, pdwCookie);
		pgit->Release();
	}
	return hr;
}

static const IID IID_IStartMenuQuerySink = __uuidof(IStartMenuQuerySink);

HRESULT CSearchOpenView::_RegisterQuerySink(IStartMenuQuerySink *psmqs)
{
	_RevokeQuerySink();
	return MarshalToGIT(psmqs, IID_IStartMenuQuerySink, &_dwCookieSink);
}

// Thanks to ep_taskbar by @amrsatrio for this interface
MIDL_INTERFACE("F729FC5E-8769-4F3E-BDB2-D7B50FD2275B")
IShellACLCustomMRU : IUnknown
{
	STDMETHOD(Initialize)(LPCWSTR, DWORD) PURE;
	STDMETHOD(AddMRUStringW)(LPCWSTR) PURE;
};

HRESULT CSearchOpenView::_UpdateMRU(LPCWSTR psz)
{
	WCHAR szMRUKey[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU";

	IShellACLCustomMRU* pacm;
	HRESULT hr = CoCreateInstance(CLSID_ACLCustomMRU, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pacm));
	if (SUCCEEDED(hr))
	{
		WCHAR szMRU[260];
		StringCchCopyW(szMRU, ARRAYSIZE(szMRU), psz);
		StringCchCatW(szMRU, ARRAYSIZE(szMRU), L"\\1");
		hr = pacm->Initialize(szMRUKey, 0x2 | 0x8 | 0x10);
		if (SUCCEEDED(hr))
		{
			hr = pacm->AddMRUStringW(szMRU);
		}
		pacm->Release();
	}
	return hr;
}

#pragma region Hybrid 7601 code

STDAPI GetMatchedFilterFromIDList(IShellFolder3 *psf3, PCUITEMID_CHILD pidlChild, FC_FLAGS fcFlags, IFilterCondition **ppfcOut)
{
	*ppfcOut = nullptr;
	HRESULT hr = E_FAIL;

	FILTERIDLISTTYPE fiType;
	if (SUCCEEDED(psf3->GetFilteredIDListType(pidlChild, &fiType)) && fiType == FIT_FILTER)
	{
		IFilterCondition *pfc;
		if (SUCCEEDED(psf3->BindToObject(pidlChild, nullptr, IID_PPV_ARGS(&pfc))))
		{
			FC_FLAGS fc;
			if (SUCCEEDED(pfc->GetTypeFlags(&fc)) && (fc & fcFlags) != 0)
			{
				*ppfcOut = pfc;
				hr = S_OK;
			}
			else
			{
				pfc->Release();
			}
		}
	}

	return hr;
}

STDAPI FindIDListbyFilterFlags(IUnknown *psi, FC_FLAGS fcFlags, UINT *puiFromLast, IFilterCondition **ppfc)
{
	if (puiFromLast)
		*puiFromLast = 0;
	if (ppfc)
		*ppfc = nullptr;

	IParentAndItem *ppai;
	HRESULT hr = psi->QueryInterface(IID_PPV_ARGS(&ppai));
	if (SUCCEEDED(hr))
	{
		IShellFolder *psfParent;
		hr = ppai->GetParentAndItem(nullptr, &psfParent, nullptr);
		if (SUCCEEDED(hr))
		{
			BOOL bFound = FALSE;

			ITEMIDLIST_ABSOLUTE *pidlWalk;
			hr = SHGetIDListFromObject(psi, &pidlWalk);
			if (SUCCEEDED(hr))
			{
				BOOL b = TRUE;
				for (UINT uiFromLast = 0; SUCCEEDED(hr) && !ILIsEmpty(pidlWalk) && !bFound; ++uiFromLast)
				{
					IShellFolder3 *psf3;
					PCUITEMID_CHILD pidlChild;
					if (b && psfParent)
					{
						pidlChild = ILFindLastID(pidlWalk);
						hr = psfParent->QueryInterface(IID_PPV_ARGS(&psf3));
						b = FALSE;
					}
					else
					{
						hr = SHBindToParent(pidlWalk, IID_PPV_ARGS(&psf3), &pidlChild);
					}

					if (SUCCEEDED(hr))
					{
						IFilterCondition *pfc;
						if (SUCCEEDED(GetMatchedFilterFromIDList(psf3, pidlChild, fcFlags, &pfc)))
						{
							bFound = TRUE;
							if (puiFromLast)
								*puiFromLast = uiFromLast;
							if (ppfc)
								IUnknown_Set((IUnknown **)ppfc, pfc);
							pfc->Release();
						}

						psf3->Release();
					}

					ILRemoveLastID(pidlWalk);
				}

				ILFree(pidlWalk);
			}

			psfParent->Release();

			if (SUCCEEDED(hr) && !bFound)
			{
				hr = E_FAIL;
			}
		}

		ppai->Release();
	}

	return hr;
}


HRESULT SHILClone(PCUIDLIST_RELATIVE pidl, ITEMIDLIST_RELATIVE **ppidlOut)
{
	if (pidl)
	{
		*ppidlOut = ILClone(pidl);
		return *ppidlOut ? S_OK : E_OUTOFMEMORY;
	}
	else
	{
		*ppidlOut = nullptr;
		return E_INVALIDARG;
	}
}


HRESULT SHILCloneFirst(PCUIDLIST_RELATIVE pidl, PITEMID_CHILD *ppidlOut)
{
	*ppidlOut = ILCloneFirst(pidl);
	return *ppidlOut ? S_OK : E_OUTOFMEMORY;
}

HRESULT SplitAtMatchedFilter(
	IShellItem *psi,
	const ITEMIDLIST_ABSOLUTE *pidl,
	UINT uiMatchFromLast,
	CDPA<ITEMID_CHILD, CTContainer_PolicyUnOwned<ITEMID_CHILD>> *pdpaRight,
	ITEMIDLIST_ABSOLUTE **ppidlLeft)
{
	*ppidlLeft = nullptr;
	HRESULT hr;

	ITEMIDLIST_ABSOLUTE *pidlLeft;
	if (psi)
		hr = SHGetIDListFromObject(psi, &pidlLeft);
	else
		hr = SHILClone(pidl, (ITEMIDLIST_RELATIVE **)&pidlLeft);

	if (SUCCEEDED(hr))
	{
		for (UINT i = 0; SUCCEEDED(hr) && i <= uiMatchFromLast; ++i)
		{
			ITEMID_CHILD *pidlLast;
			hr = SHILCloneFirst(ILFindLastID(pidlLeft), &pidlLast);
			if (SUCCEEDED(hr))
			{
				hr = pdpaRight->InsertPtr(0, pidlLast);
				if (SUCCEEDED(hr))
				{
					pidlLast = nullptr;
					hr = ILRemoveLastID(pidlLeft) ? S_OK : E_FAIL;
				}

				ILFree(pidlLast);
			}
		}

		if (SUCCEEDED(hr))
		{
			ITEMIDLIST_ABSOLUTE *temp = pidlLeft;
			pidlLeft = nullptr;
			*ppidlLeft = temp;
		}

		ILFree(pidlLeft);
	}

	return hr;
}

STDAPI SHILCombine(const ITEMIDLIST_ABSOLUTE *pidl1, const ITEMIDLIST_RELATIVE *pidl2, ITEMIDLIST_ABSOLUTE **ppidlOut)
{
	*ppidlOut = ILCombine(pidl1, pidl2);
	return *ppidlOut ? S_OK : E_OUTOFMEMORY;
}

HRESULT RebuildMatchedFilterIDList(
	const ITEMIDLIST_ABSOLUTE *pidlParent,
	CDPA<ITEMID_CHILD, CTContainer_PolicyUnOwned<ITEMID_CHILD>> *pdpaFilters,
	IFilterCondition *pfc,
	ITEMIDLIST_ABSOLUTE **ppidl)
{
	*ppidl = nullptr;

	IShellFolder3 *psf;
	HRESULT hr = SHBindToObject(nullptr, pidlParent, nullptr, IID_PPV_ARGS(&psf));
	if (SUCCEEDED(hr))
	{
		if (pfc)
		{
			ITEMID_CHILD *pidlChild;
			hr = psf->ModifyFilteredIDList(pdpaFilters->GetPtr(0), pfc, &pidlChild);
			if (SUCCEEDED(hr))
			{
				hr = SHILCombine(pidlParent, pidlChild, ppidl);
				ILFree(pidlChild);
			}
		}
		else
		{
			if (pdpaFilters->GetPtrCount() == 1)
			{
				IShellFolder3 *psfFilter;
				hr = SHBindToObject(psf, pdpaFilters->GetPtr(0), nullptr, IID_PPV_ARGS(&psfFilter));
				if (SUCCEEDED(hr))
				{
					PROPERTYKEY key;
					hr = psfFilter->GetStackedKey(&key);
					if (SUCCEEDED(hr))
					{
						hr = psf->CreateStackedIDList(key, ppidl);
					}

					psfFilter->Release();
				}
			}
			else
			{
				hr = SHILClone(pidlParent, (ITEMIDLIST_RELATIVE **)ppidl);
			}
		}

		psf->Release();
	}

	return hr;
}

HRESULT _ILCombineAndFree(ITEMIDLIST_RELATIVE *pidlFirst, ITEMIDLIST_RELATIVE *pidlSecond, ITEMIDLIST_RELATIVE **ppidlCombined)
{
	HRESULT hr;

	if (pidlSecond && pidlFirst)
	{
		hr = SHILCombine((const ITEMIDLIST_ABSOLUTE *)pidlFirst, pidlSecond, (ITEMIDLIST_ABSOLUTE **)ppidlCombined);
		ILFree(pidlFirst);
		ILFree(pidlSecond);
	}
	else
	{
		hr = S_OK;

		if (pidlFirst)
		{
			*ppidlCombined = pidlFirst;
		}
		else if (pidlSecond)
		{
			*ppidlCombined = pidlSecond;
		}
		else
		{
			hr = E_INVALIDARG;
		}
	}

	return hr;
}

STDAPI SHILAppend(ITEMIDLIST_RELATIVE *pidl, ITEMIDLIST_RELATIVE **ppidlCombined)
{
	return _ILCombineAndFree(*ppidlCombined, pidl, ppidlCombined);
}

HRESULT ModifyOrRemoveFilterHelper(
	ITEMIDLIST_ABSOLUTE *pidlParent,
	IFilterCondition *pfc,
	CDPA<ITEMID_CHILD, CTContainer_PolicyUnOwned<ITEMID_CHILD>> *pdpaFilters,
	ITEMIDLIST_ABSOLUTE **ppidlFiltered)
{
	*ppidlFiltered = nullptr;

	ITEMIDLIST_ABSOLUTE *pidlFilter;
	HRESULT hr = RebuildMatchedFilterIDList(pidlParent, pdpaFilters, pfc, &pidlFilter);
	if (SUCCEEDED(hr))
	{
		for (int i = 1; i < pdpaFilters->GetPtrCount() && SUCCEEDED(hr); ++i)
		{
			IShellFolder3 *psf;
			hr = SHBindToObject(nullptr, pidlFilter, nullptr, IID_PPV_ARGS(&psf));
			if (SUCCEEDED(hr))
			{
				ITEMIDLIST_RELATIVE *pidlNewChild;
				hr = psf->ReparentFilteredIDList(pdpaFilters->GetPtr(i), &pidlNewChild);
				if (SUCCEEDED(hr))
				{
					hr = SHILAppend(pidlNewChild, (ITEMIDLIST_RELATIVE **)&pidlFilter);
				}

				psf->Release();
			}
		}

		if (SUCCEEDED(hr))
		{
			ITEMIDLIST_ABSOLUTE *temp = pidlFilter;
			pidlFilter = nullptr;
			*ppidlFiltered = temp;
		}

		ILFree(pidlFilter);
	}

	return hr;
}

template<typename T>
int CALLBACK DPA_ILFreeCB(T *self, void *pData)
{
	CoTaskMemFree(self);
	return 1;
}

STDAPI ModifyOrRemoveItemFilterByIndex(IShellItem *psi, UINT uiMatchFromLast, IFilterCondition *pfc, ITEMIDLIST_ABSOLUTE **ppidlFiltered)
{
	*ppidlFiltered = nullptr;

	CDPA<ITEMID_CHILD, CTContainer_PolicyUnOwned<ITEMID_CHILD>> dpaFilters;
	HRESULT hr = dpaFilters.Create(uiMatchFromLast + 1) ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		ITEMIDLIST_ABSOLUTE *pidlParent;
		hr = SplitAtMatchedFilter(psi, nullptr, uiMatchFromLast, &dpaFilters, &pidlParent);
		if (SUCCEEDED(hr))
		{
			hr = ModifyOrRemoveFilterHelper(pidlParent, pfc, &dpaFilters, ppidlFiltered);
			ILFree(pidlParent);
		}

		dpaFilters.DestroyCallback(DPA_ILFreeCB, nullptr);
	}

	return hr;
}

enum RESTATEMENT_OPTIONS
{
	RO_NONE = 0x0,
	RO_LANGUAGE_INDEPENDENT = 0x1,
	RO_LOWERCASE_KEYWORDS = 0x2,
};

DEFINE_ENUM_FLAG_OPERATORS(RESTATEMENT_OPTIONS);

class DECLSPEC_UUID("934d4698-6a59-48f8-9f29-9fb30670320e")
StructuredQueryHelper;

MIDL_INTERFACE("b3ddac23-10ba-4407-98de-f5fef5db4629")
IStructuredQueryHelper : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE ParseStructuredQuery(const WCHAR *, DWORD, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDefaultKeyword(REFPROPERTYKEY, RESTATEMENT_OPTIONS, WCHAR **) = 0;
	virtual HRESULT STDMETHODCALLTYPE LogSqlQuery(const WCHAR *, const WCHAR *) = 0;
	virtual HRESULT STDMETHODCALLTYPE LogSqlSuccess(DWORD) = 0;
	virtual HRESULT STDMETHODCALLTYPE LogSqlFailure(HRESULT) = 0;
	virtual HRESULT STDMETHODCALLTYPE CreateQueryParser(WORD, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE ParseWithQueryParser(IQueryParser *, const WCHAR *, DWORD, REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetLanguageResourcePool(REFIID, void **) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetSearchApiQueryHelper(const WCHAR *, REFIID, void **) = 0;
};

STDAPI GetStructuredQueryHelper(REFIID riid, void **ppv)
{
	return CoCreateInstance(__uuidof(StructuredQueryHelper), nullptr, CLSCTX_INPROC_SERVER, riid, ppv);
}

STDAPI TextToCondition(const WCHAR *pszQuery, ICondition **ppc)
{
	*ppc = nullptr;
	// SHTracePerf(&ShellTraceId_StructuredQuery_ParseText_Start);

	IStructuredQueryHelper *psqh;
	HRESULT hr = GetStructuredQueryHelper(IID_PPV_ARGS(&psqh));
	if (SUCCEEDED(hr))
	{
		hr = psqh->ParseStructuredQuery(pszQuery, MAKELONG(GetKeyboardLayout(0), 0), IID_PPV_ARGS(ppc));
		psqh->Release();
	}

	// SHTracePerf(&ShellTraceId_StructuredQuery_ParseText_Stop);
	return hr;
}

DEFINE_PROPERTYKEY(PKEY_WordWheel, 0x1E3EE840, 0xBC2B, 0x476C, 0x82, 0x37, 0x2A, 0xCD, 0x1A, 0x83, 0x9B, 0x22, 5);

STDAPI TextToFilterEx(const WCHAR *pszName, REFPROPERTYKEY propkey, const WCHAR *pszQuery, IFilterCondition **ppfc)
{
	*ppfc = nullptr;

	ICondition *pc;
	HRESULT hr = TextToCondition(pszQuery, &pc);
	if (SUCCEEDED(hr))
	{
		FC_FLAGS fcFlags = FCT_FREEFORMED;
		if (IsEqualPropertyKey(propkey, PKEY_WordWheel))
		{
			fcFlags |= FCT_WORDWHEEL;
		}

		hr = SHCreateFilter(pszName, nullptr, propkey, fcFlags | FCT_FORMATFORDISPLAY, pc, IID_PPV_ARGS(ppfc));
		pc->Release();
	}

	return hr;
}

STDAPI TextToFilter(const WCHAR *pszQuery, IFilterCondition **ppv)
{
	return TextToFilterEx(pszQuery, PKEY_WordWheel, pszQuery, ppv);
}

STDAPI SHFullIDListFromFolderAndRelativeItem(IShellFolder *psf, const ITEMIDLIST_RELATIVE *pidl, ITEMIDLIST_ABSOLUTE **ppidl)
{
	*ppidl = nullptr;

	ITEMIDLIST_ABSOLUTE *pidlFolder;
	HRESULT hr = SHGetIDListFromObject(psf, &pidlFolder);
	if (SUCCEEDED(hr))
	{
		hr = SHILCombine(pidlFolder, pidl, ppidl);
		ILFree(pidlFolder);
	}

	return hr;
}

STDAPI SHCreateFilteredIDList(
	IShellFolder *psf,
	FILTERIDLISTTYPE filterType,
	IFilterCondition **rgConditions,
	int cConditions,
	ITEMIDLIST_ABSOLUTE **ppidlOutFiltered)
{
	*ppidlOutFiltered = nullptr;

	IShellFolder3 *psf3;
	HRESULT hr = psf->QueryInterface(IID_PPV_ARGS(&psf3));
	if (SUCCEEDED(hr))
	{
		hr = cConditions ? S_OK : E_INVALIDARG;
		ITEMID_CHILD *pidl = nullptr;

		for (int i = 0; SUCCEEDED(hr) && i < cConditions; ++i)
		{
			if (pidl)
			{
				IUnknown_SafeReleaseAndNullPtr(&psf3);
				hr = SHBindToObject(psf, pidl, nullptr, IID_PPV_ARGS(&psf3));
			}
			if (SUCCEEDED(hr))
			{
				ILFree(pidl);
				hr = psf3->CreateFilteredIDList(rgConditions[i], filterType, nullptr, &pidl);
			}
		}

		if (SUCCEEDED(hr))
		{
			hr = SHFullIDListFromFolderAndRelativeItem(psf3, pidl, ppidlOutFiltered);
		}

		ILFree(pidl);
		IUnknown_SafeReleaseAndNullPtr(&psf3);
	}

	return hr;
}

#pragma endregion

HRESULT CSearchOpenView::_AddTextFilterToLocation(
	LPCITEMIDLIST pidl,
	IFilterCondition *pfc,
	LPITEMIDLIST *ppidl)
{
	IShellFolder *ppv = 0;
	HRESULT hr = SHBindToObject(0, pidl, 0, IID_PPV_ARGS(&ppv));
	if (hr >= 0)
	{
		hr = SHCreateFilteredIDList(ppv, FIT_FILTER, &pfc, 1, ppidl);
	}

	if (ppv)
		ppv->Release();
	return hr;
}

HRESULT CSearchOpenView::_UpdateSearchTextFilter()
{
	// edi
	ITEMIDLIST **v3; // eax
	LPWSTR v4; // eax
	LPITEMIDLIST *v5; // eax
	ITEMIDLIST *v7; // eax
	LPITEMIDLIST v8; // eax
	IShellItem *v9; // eax
	LPITEMIDLIST v11; // [esp+10h] [ebp-2Ch] BYREF
	UINT v12; // [esp+14h] [ebp-28h] BYREF
	IFilterCondition *pfc; // [esp+18h] [ebp-24h] MAPDST SPLIT BYREF
	LPITEMIDLIST ppv; // [esp+1Ch] [ebp-20h] BYREF
	IShellItem *v16; // [esp+1Ch] [ebp-20h] SPLIT BYREF
	//CPPEH_RECORD ms_exc; // [esp+24h] [ebp-18h]

	HRESULT hr = 0;
	ppv = 0;

	LPITEMIDLIST pidl = 0;
	//CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>::CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>(&pidl);

	//if (!this->_psiFolder
	//	&& CcshellAssertFailedW(L"d:\\win7m2\\shell\\explorer\\desktop2\\srchview.cpp", 1098, L"_psiFolder", 0))
	//{
	//	AttachUserModeDebugger();
	//	do
	//	{
	//		__debugbreak();
	//		ms_exc.registration.TryLevel = -2;
	//	} while (dword_10B43C0);
	//}

	if (FindIDListbyFilterFlags(this->_psiFolder, FCT_WORDWHEEL, &v12, 0) >= 0)
	{
		//v3 = (ITEMIDLIST **)CTSmartObj<_ITEMIDLIST_ABSOLUTE *, CTSmartPtr_PolicyComplete<CTContainer_PolicyCoTaskMem>>::operator&(&pidl);
		hr = ModifyOrRemoveItemFilterByIndex(this->_psiFolder, v12, 0, &pidl);
	}
	if (hr >= 0)
	{
		v4 = this->field_A8;
		if (!v4 || !*v4)
		{
		LABEL_19:
			if (hr >= 0 && pidl)
			{
				v16 = 0;
				hr = SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&v16));
				if (hr >= 0)
				{
					this->_psiFolder->Release();
					v9 = v16;
					v16 = 0;
					this->_psiFolder = v9;
				}
				if (v16)
					v16->Release();
			}
			goto LABEL_25;
		}
		if (!pidl)
		{
			//v5 = (LPITEMIDLIST *)CTSmartObj<_ITEMIDLIST_ABSOLUTE *, CTSmartPtr_PolicyComplete<CTContainer_PolicyCoTaskMem>>::operator&(&pidl);
			hr = SHGetIDListFromObject(this->_psiFolder, &pidl);
		}
		if (hr >= 0)
		{
			pfc = nullptr;
			hr = TextToFilter(/*this->_pqp, &this->field_58,*/ this->field_A8, &pfc);
			if (hr >= 0)
			{
				//CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>::CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>(&ppv);
				//v7 = (ITEMIDLIST *)CTSmartObj<_ITEMIDLIST_ABSOLUTE *, CTSmartPtr_PolicyComplete<CTContainer_PolicyCoTaskMem>>::operator&(&ppv);
				hr = _AddTextFilterToLocation(pidl, pfc, &ppv);
				if (hr >= 0)
				{
					v8 = ppv;
					ppv = 0;
					v11 = v8;
					pidl = v11;
					//CTSmartObj<_ITEMIDLIST_ABSOLUTE *, CTSmartPtr_PolicyComplete<CTContainer_PolicyCoTaskMem>>::Attach(
					//	(void **)&pidl,
					//	&v11);
				}
				CoTaskMemFree(ppv);
				//CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>::~CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>(&ppv);
			}
			if (pfc)
			{
				pfc->Release();
			}
			goto LABEL_19;
		}
	}
LABEL_25:
	CoTaskMemFree(pidl);
	//CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>::~CCoTaskMemPtr<_ITEMIDLIST_ABSOLUTE>((void **)&pidl);
	return hr;
}

#define HYBRID_CODE

HRESULT CSearchOpenView::_UpdateSearchText(LPCWSTR psz)
{
	_SwitchToMode(VIEWMODE_DEFAULT, 0);

	CoTaskMemFree(field_A8);
	HRESULT hr = SHStrDup(psz, &field_A8);
	if (SUCCEEDED(hr))
	{
		if (field_C4)
		{
			int v17 = 0;
			if (_pFolderView && _psiFolder && _viewMode != VIEWMODE_PATHCOMPLETE)
				_CancelNavigation();
			else
				v17 = 1;

			// Skipped telemetry StartPane_FindItem_Start

			IUnknown_SafeReleaseAndNullPtr(&_psiFolder);
			hr = _InitRegularAutoListItem(&this->_psiFolder);
			if (SUCCEEDED(hr))
			{
				hr = _UpdateSearchTextFilter();
				if (SUCCEEDED(hr))
				{
					_peb->SetOptions(EBO_NOTRAVELLOG | EBO_ALWAYSNAVIGATE);
					IExplorerBrowser* peb = _peb;
					_fIsBrowsing = v17;
					int v10 = peb->BrowseToObject(_psiFolder, SBSP_WRITENOHISTORY);
					_fIsBrowsing = 0;
					return v10;
				}
			}
		}
		else if (!_pFolderView || _viewMode == VIEWMODE_PATHCOMPLETE)
		{
			IShellItem* psiAutoList = nullptr;
			// Skipped telemetry StartPane_FindItem_Start
			hr = _InitRegularAutoListItem(&psiAutoList);
			if (SUCCEEDED(hr))
			{
				_peb->SetOptions(EBO_NOTRAVELLOG | EBO_ALWAYSNAVIGATE);
				_fIsBrowsing = TRUE;
				hr = _peb->BrowseToObject(psiAutoList, SBSP_WRITENOHISTORY);
				_fIsBrowsing = FALSE;
			}
			if (psiAutoList)
			{
				psiAutoList->Release();
			}
		}
		else
		{
			_CancelNavigation();

			IFilterView* pfv;
			hr = _pFolderView->QueryInterface(IID_PPV_ARGS(&pfv));
			if (SUCCEEDED(hr))
			{
				hr = _FilterView(pfv);
				pfv->Release();
			}
		}
	}
	return hr;
}

FOLDERFLAGS CSearchOpenView::_GetFolderFlags()
{
	FOLDERFLAGS ret = FWF_ALLOWRTLREADING | FWF_SUBSETGROUPS
		| FWF_NOBROWSERVIEWSTATE | FWF_NOCOLUMNHEADER | FWF_FULLROWSELECT
		| FWF_NOWEBVIEW | FWF_SINGLECLICKACTIVATE | FWF_NOCLIENTEDGE | FWF_SINGLESEL;
	if (_hTheme)
	{
		ret |= FWF_TRANSPARENT;
	}
	return ret;
}

void CSearchOpenView::_CancelNavigation()
{
	if (!field_B8 && _pFolderView)
	{
		IOleCommandTarget *pct;
		if (SUCCEEDED(_pFolderView->QueryInterface(IID_PPV_ARGS(&pct))))
		{
			pct->Exec(nullptr, 23, 0, nullptr, nullptr);
			pct->Release();
		}
	}
}

void CSearchOpenView::_ConnectShellView(IShellView* psv)
{
	ASSERT(_pDispatchView == NULL); // 796
	if (SUCCEEDED(psv->GetItemObject(0, IID_PPV_ARGS(&_pDispatchView))))
	{
		ConnectToConnectionPoint(static_cast<IDispatch*>(this), DIID_DShellFolderViewEvents, TRUE, _pDispatchView, &field_74, nullptr);
	}
}

void CSearchOpenView::_DisconnectShellView()
{
	if (_pDispatchView)
	{
		ConnectToConnectionPoint(static_cast<IDispatch *>(this), DIID_DShellFolderViewEvents, FALSE, _pDispatchView, &field_74, nullptr);
		IUnknown_SafeReleaseAndNullPtr(&_pDispatchView);
	}
}

void CSearchOpenView::_DoKeyBoardContextMenu()
{
	int iCurSel = _GetCurSel();
	IHitTestView* phtv;
	if (iCurSel >= 0 && _pFolderView && SUCCEEDED(_pFolderView->QueryInterface(IID_PPV_ARGS(&phtv))))
	{
		RECT rc;
		if (phtv->GetItemRect(iCurSel, &rc) >= 0)
		{
			POINT pt;
			pt.x = (rc.right + rc.left) / 2;
			pt.y = (rc.bottom + rc.top) / 2;

			IShellView* psv;
			if (_pFolderView->QueryInterface(IID_PPV_ARGS(&psv)) >= 0)
			{
				IContextMenu* pcm;
				if (psv->GetItemObject(1, IID_PPV_ARGS(&pcm)) >= 0)
				{
					IContextMenuSite* v8;
					if (_pFolderView->QueryInterface(IID_PPV_ARGS(&v8)) >= 0)
					{
						HWND hwnd;
						if (psv->GetWindow(&hwnd) >= 0)
						{
							ClientToScreen(hwnd, &pt);
							v8->DoContextMenuPopup(pcm, 0, pt);
						}
						v8->Release();
					}
					pcm->Release();
				}
				psv->Release();
			}
		}
		phtv->Release();
	}
}

// Thanks to ep_taskbar by @amrsatrio for these functions

#pragma region Hybrid 7601 code

const WCHAR *const off_107FDF4[] =
{
	L"System.StructuredQueryType.SortKeyDescription",
	L"System.StructuredQueryType.AnyBitsSet",
	L"System.StructuredQueryType.AllBitsSet"
};

BOOL SemanticTypeIsNamedEntity(const WCHAR *pszSemanticType)
{
	BOOL bRet = TRUE;
	for (const WCHAR *it : off_107FDF4) // @MOD Use foreach
	{
		if (CompareStringW(LOCALE_INVARIANT, 0, pszSemanticType, -1, it, -1) == CSTR_EQUAL)
		{
			bRet = FALSE;
			break;
		}
	}
	return bRet;
}

STDAPI WritePropVariantForNamedEntity(
	REFPROPVARIANT propvarIn,
	const WCHAR *pszKeyValueSep,
	const WCHAR *pszNumberEnd,
	WCHAR *pszDest,
	size_t cchDest,
	WCHAR **ppszDestEnd,
	size_t *pcchRemaining)
{
	HRESULT hr;
	WCHAR *pszDestEnd = pszDest;
	size_t cchRemaining = cchDest;

	if (cchDest)
	{
		switch (propvarIn.vt)
		{
			case VT_I1:
			case VT_I2:
			case VT_I4:
			case VT_I8:
			{
				LONGLONG llRet;
				hr = PropVariantToInt64(propvarIn, &llRet);
				if (SUCCEEDED(hr))
				{
					hr = StringCchPrintfExW(pszDest, cchDest, &pszDestEnd, &cchRemaining, 0, L"%s%I64d%s", pszKeyValueSep, llRet, pszNumberEnd);
				}
				break;
			}
			case VT_UI1:
			case VT_UI2:
			case VT_UI4:
			case VT_UI8:
			{
				ULONGLONG ullRet;
				hr = PropVariantToUInt64(propvarIn, &ullRet);
				if (SUCCEEDED(hr))
				{
					hr = StringCchPrintfExW(pszDest, cchDest, &pszDestEnd, &cchRemaining, 0, L"%s%I64u%s", pszKeyValueSep, ullRet, pszNumberEnd);
				}
				break;
			}
			case VT_R8:
			{
				hr = StringCchPrintfExW(pszDest, cchDest, &pszDestEnd, &cchRemaining, 0, L"%s%g%s", pszKeyValueSep, propvarIn.dblVal, pszNumberEnd);
				break;
			}
			case VT_BOOL:
			{
				hr = StringCchPrintfExW(pszDest, cchDest, &pszDestEnd, &cchRemaining, 0, L"%s%s", pszKeyValueSep, propvarIn.boolVal ? L"TRUE" : L"FALSE");
				break;
			}
			case VT_LPWSTR:
			{
				hr = StringCchPrintfExW(pszDest, cchDest, &pszDestEnd, &cchRemaining, 0, L"%s", propvarIn.pwszVal);
				break;
			}
			default:
			{
				hr = E_FAIL;
				break;
			}
		}

		if (FAILED(hr))
		{
			*pszDest = 0;
			pszDestEnd = pszDest;
			cchRemaining = cchDest;
		}
	}
	else
	{
		hr = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
	}

	if (ppszDestEnd)
		*ppszDestEnd = pszDestEnd;
	if (pcchRemaining)
		*pcchRemaining = cchRemaining;

	return hr;
}

STDAPI CreateFilterConditionValueEx(
	const WCHAR *a1,
	const WCHAR *a2,
	REFPROPERTYKEY propkey,
	CONDITION_OPERATION cop,
	REFPROPVARIANT propvarIn,
	const WCHAR *pszSemanticType,
	IFilterCondition **ppfcOut)
{
	*ppfcOut = nullptr;

	IConditionFactory2 *pcf2;
	HRESULT hr = SHCreateConditionFactory(IID_PPV_ARGS(&pcf2));
	if (SUCCEEDED(hr))
	{
		WCHAR szPropValue[1024];
		ICondition *pcondResolved = nullptr;
		if (pszSemanticType
			&& SemanticTypeIsNamedEntity(pszSemanticType)
			&& SUCCEEDED(WritePropVariantForNamedEntity(propvarIn, L"=", L"", szPropValue, ARRAYSIZE(szPropValue), nullptr, nullptr)))
		{
			PROPVARIANT propvar;
			propvar.vt = VT_LPWSTR;
			propvar.pwszVal = szPropValue;

			ICondition *pcond;
			hr = pcf2->CreateLeaf(propkey, cop, propvar, pszSemanticType, L"", nullptr, nullptr, nullptr, CONDITION_CREATION_DEFAULT, IID_PPV_ARGS(&pcond));
			if (SUCCEEDED(hr))
			{
				SYSTEMTIME SystemTime;
				GetLocalTime(&SystemTime);
				hr = pcf2->ResolveCondition(pcond, SQRO_DONT_RESOLVE_DATETIME, &SystemTime, IID_PPV_ARGS(&pcondResolved));
				pcond->Release();
			}
		}
		else
		{
			hr = pcf2->CreateLeaf(propkey, cop, propvarIn, pszSemanticType, L"", nullptr, nullptr, nullptr, CONDITION_CREATION_VECTOR_LEAF, IID_PPV_ARGS(&pcondResolved));
		}

		if (SUCCEEDED(hr))
		{
			hr = SHCreateFilter(a1, a2, propkey, FCT_DEFAULT, pcondResolved, IID_PPV_ARGS(ppfcOut));
			pcondResolved->Release();
		}

		pcf2->Release();
	}

	return hr;
}

STDAPI CreateFilterConditionValue(
	const WCHAR *a1,
	const WCHAR *a2,
	REFPROPERTYKEY propkey,
	CONDITION_OPERATION cop,
	REFPROPVARIANT propvarIn,
	IFilterCondition **ppfcOut)
{
	return CreateFilterConditionValueEx(a1, a2, propkey, cop, propvarIn, nullptr, ppfcOut);
}

#pragma endregion

void CSearchOpenView::_FilterPathCompleteView(LPCWSTR pszPath)
{
	if (_pFolderView && _ppci && pszPath)
	{
		IFilterView* pfv;
		if (SUCCEEDED(_pFolderView->QueryInterface(IID_PPV_ARGS(&pfv))))
		{
			PROPVARIANT pvar;
			if (SUCCEEDED(InitPropVariantFromString(pszPath, &pvar)))
			{
				IFilterCondition* pfc = nullptr;
				if (SUCCEEDED(CreateFilterConditionValue(pszPath, nullptr, PKEY_ItemNameDisplay, COP_VALUE_STARTSWITH, pvar, &pfc)))
				{
					pfv->FilterByCondition(pfc);
					pfc->Release();
				}
				PropVariantClear(&pvar);
			}
			pfv->Release();
		}
	}
}

void CSearchOpenView::_InstrumentActivation(int iItem)
{
	// EXEX-Vista(allison): TODO.
}

#define HYBRID_CODE

void CSearchOpenView::_PathCompleteUpdate(CPathCompleteInfo* ppciNew)
{
	BOOL v3 = 0;
	if (ppciNew && _ppci)
	{
		v3 = StrCmpI(_ppci->_psz2, ppciNew->_psz2) == 0;
	}
	if (_ppci)
	{
		delete _ppci;
	}

	_ppci = ppciNew;
	_SwitchToMode(VIEWMODE_PATHCOMPLETE, 0);

	if (v3)
	{
		_FilterPathCompleteView(_ppci->_psz1);
	}
	else
	{
		ITEMIDLIST_ABSOLUTE* pidl;
		if (SUCCEEDED(_InitPathCompletePidlAutoList(_ppci->_psz2, &pidl)))
		{
			_peb->SetOptions(EBO_NOTRAVELLOG | EBO_ALWAYSNAVIGATE);
			_peb->BrowseToIDList(pidl, SBSP_WRITENOHISTORY);
			ILFree(pidl);
		}
	}
}

void CSearchOpenView::_ReleaseExplorerBrowser()
{
	if (_peb)
	{
		_peb->Unadvise(_dwCookie);
		_peb->Destroy();
		IUnknown_SetSite(_peb, NULL);
	}
	IUnknown_SafeReleaseAndNullPtr(&_peb);
}

void RevokeFromGIT(DWORD dwCookie)
{
	if (dwCookie)
	{
		IGlobalInterfaceTable *pgit;
		if (SUCCEEDED(CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pgit))))
		{
			pgit->RevokeInterfaceFromGlobal(dwCookie);
			pgit->Release();
		}
	}
}

void CSearchOpenView::_RevokeQuerySink()
{
	if (_dwCookieSink)
	{
		RevokeFromGIT(_dwCookieSink);
		_dwCookieSink = 0;
	}
}

void CSearchOpenView::_SizeExplorerBrowser(int cx, int cy)
{
	RECT rc;
	rc.left = _margins.cxLeftWidth;
	rc.top = _margins.cyTopHeight;
	rc.right = cx - _margins.cxRightWidth;
	rc.bottom = cy - _margins.cyBottomHeight;
	_peb->SetRect(NULL, rc);

	IColumnManager *pcm;
	if (_pFolderView && SUCCEEDED(_pFolderView->QueryInterface(IID_PPV_ARGS(&pcm))))
	{
		CM_COLUMNINFO cmci = {0};
		cmci.cbSize = sizeof(cmci);
		cmci.dwState = CM_STATE_NONE;
		cmci.dwMask = CM_MASK_WIDTH;
		cmci.uWidth = (UINT)RECTWIDTH(rc);
		if (field_B0 != 0) // EXEX-Vista(allison): TODO: Check why field_B0 is never true
		{
			cmci.uWidth -= GetSystemMetrics(SM_CXVSCROLL);
		}

		pcm->SetColumnInfo(PKEY_ItemNameDisplay, &cmci);
		pcm->Release();
	}
}

const PROPERTYKEY PKEY_Null = { { 0u, 0u, 0u, { 0u, 0u, 0u, 0u, 0u, 0u, 0u, 0u } }, 0u };

void CSearchOpenView::_SwitchToMode(VIEWMODE viewModeNew, int a3)
{
	VARIANT vt;
	vt.lVal = -1;
	vt.vt = VT_I4;
	IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 317, 0, &vt, NULL);
	if (_pFolderView && (viewModeNew != _viewMode || a3))
	{
		vt.lVal = viewModeNew;
		vt.vt = VT_I4;
		IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 328, 0, &vt, NULL);
		if (viewModeNew == VIEWMODE_PATHCOMPLETE)
		{
			SORTCOLUMN sc;
			sc.propkey = PKEY_ItemNameDisplay;
			sc.direction = SORT_ASCENDING;
			_pFolderView->SetSortColumns(&sc, 1);
			_pFolderView->SetGroupBy(PKEY_Null, 1);
			_LimitViewResults(_pFolderView, 0);
		}
		else if (_viewMode == VIEWMODE_PATHCOMPLETE)
		{
			if (_ppci)
			{
				delete _ppci;
				_ppci = NULL;
			}
		}

		_pFolderView->SetGroupSubsetCount(0);
		if (!a3)
		{
			field_7C = 1;
		}
	}
	_viewMode = viewModeNew;
}

// Thanks to ep_taskbar by @amrsatrio for these functions
#pragma region Resource String Copy Helpers

HRESULT ResourceStringCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR *ppsz, UINT cch)
{
	const WCHAR *pch;
	WORD len;
	HRESULT hr = ResourceStringFindAndSizeEx(hModule, uId, wLanguage, &pch, &len);
	if (SUCCEEDED(hr))
	{
		if (len < cch)
		{
			memcpy(ppsz, pch, sizeof(WCHAR) * len);
			ppsz[len] = 0;
		}
		else
		{
			hr = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
		}
	}
	return hr;
}

HRESULT ResourceStringCchCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR *ppsz, UINT cch)
{
	return ResourceStringCopyEx(hModule, uId, wLanguage, ppsz, cch);
}

#pragma endregion

void CSearchOpenView::_UpdateIndexState(int iState)
{
	if (iState)
		field_AC = 1;

	WCHAR szText[255];
	ResourceStringCchCopyEx(g_hinstCabinet, iState == 0 ? 7027 : 7028, LANG_NEUTRAL, szText, ARRAYSIZE(szText));
	_peb->SetEmptyText(szText);
}

void CSearchOpenView::_UpdateOpenBoxText()
{
	WCHAR szParsingName[MAX_PATH];
	if (_viewMode == VIEWMODE_PATHCOMPLETE && _pFolderView
		&& SUCCEEDED(_GetSelectedItemParsingName(szParsingName, ARRAYSIZE(szParsingName))))
	{
		VARIANT vt;
		vt.vt = VT_BYREF;
		vt.byref = szParsingName;
		IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenBox, &SID_SM_DV2ControlHost, 315, 0, &vt, NULL);
	}
}

void CSearchOpenView::_UpdateScrolling()
{
	IFolderView2 *pFolderView; // eax
	IFolderView2 *v3; // eax
	int *v4; // eax
	RECT v5; // [esp+8h] [ebp-38h] BYREF
	RECT v6; // [esp+18h] [ebp-28h] BYREF
	RECT rc; // [esp+28h] [ebp-18h] BYREF
	IHitTestView *v8; // [esp+38h] [ebp-8h] BYREF
	int iItem; // [esp+3Ch] [ebp-4h] BYREF

	iItem = -1;
	pFolderView = _pFolderView;
	if (pFolderView)
	{
		pFolderView->GetVisibleItem(-1, 1, &iItem);
		if (iItem > 0)
		{
			v3 = this->_pFolderView;
			if (v3)
			{
				if (v3->QueryInterface(IID_PPV_ARGS(&v8)) >= 0)
				{
					GetClientRect(this->_hwnd, &rc);
					if (v8->GetItemRect(iItem, &v5) < 0 || v8->GetItemRect(0, &v6) < 0)
					{
						goto LABEL_13;
					}
					v4 = &this->field_AC;
					if (rc.bottom >= v5.bottom + this->_margins.cyBottomHeight - v6.top)
					{
						if (*v4)
						{
							*v4 = 0;
							goto LABEL_12;
						}
					}
					else if (!*v4)
					{
						*v4 = 1;
					LABEL_12:
						_SizeExplorerBrowser(rc.right, rc.bottom);
					}
				LABEL_13:
					v8->Release();
				}
			}
		}
	}
}

DEFINE_PROPERTYKEY(PKEY_StartMenu_Query, 0x4BD13B3D, 0xE68B, 0x44EC, 0x89, 0xEE, 0x76, 0x11, 0x78, 0x9D, 0x40, 0x70, 102);

void CSearchOpenView::_UpdateTopMatch(UTM_REASON reason)
{
	const WCHAR *v3; // eax
	DWORD v4; // eax
	const WCHAR *v5; // eax
	const WCHAR *v6; // eax
	DWORD v7; // eax
	int v8; // ecx
	LONG v9; // ebx
	IUnknown *v10; // [esp-20h] [ebp-5Ch]
	IUnknown *punkSite; // [esp-20h] [ebp-5Ch]
	VARIANT vt; // [esp+8h] [ebp-34h] BYREF
	VARIANT varIn; // [esp+18h] [ebp-24h] BYREF
	int v14; // [esp+28h] [ebp-14h]
	int v15; // [esp+2Ch] [ebp-10h] BYREF
	IShellFolder2 *v16; // [esp+30h] [ebp-Ch] BYREF
	LPITEMIDLIST pidl; // [esp+34h] [ebp-8h] BYREF
	LONG v18; // [esp+38h] [ebp-4h]
	HRESULT hr; // [esp+44h] [ebp+8h]

	if (this->_viewMode)
		return;
	v18 = this->field_98;
	v14 = this->field_88;
	if (!CSearchOpenView::_GetItemCount())
	{
		v7 = this->field_98;
		v18 = v7 != 3 && v7;
		v8 = this->field_84;
		if (v8 || reason == UTM_REASON_1)
		{
			this->field_88 = 1;
			if (v8)
			{
				vt.lVal = 1;
				vt.vt = 3;
				IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 328, 0, &vt, 0);
			}
		}
		v14 = 0;
		goto LABEL_48;
	}

	this->_pFolderView->GetVisibleItem(-1, 0, &v15);
	if (v15 >= 0 && this->_pFolderView->Item(v15, &pidl) >= 0)
	{
		hr = this->_pFolderView->GetFolder(IID_PPV_ARGS(&v16));
		if (hr >= 0)
		{
			if (this->field_A8)
			{
				hr = v16->GetDetailsEx(pidl, &PKEY_StartMenu_Query, &varIn);
				if (hr >= 0)
				{
					v3 = VariantToStringWithDefault(varIn, L"");
					if (!StrCmpW(this->field_A8, v3))
					{
						if (!this->field_88)
						{
							//if (EventEnabled(g_SHPerfRegHandle, &Explorer_StartPane_OpenBox_TopMatchReady_Info))
							//	TemplateEventDescriptor(g_SHPerfRegHandle, &Explorer_StartPane_OpenBox_TopMatchReady_Info);
							//if (EventEnabled(g_SHPerfRegHandle, &ShellTraceId_PerfTrack_StartPane_FindItem_Stop))
							//	TemplateEventDescriptor(g_SHPerfRegHandle, &ShellTraceId_PerfTrack_StartPane_FindItem_Stop);
						}
						v18 = -1;
						this->field_88 = 1;
					}
					VariantClear(&varIn);
				}
			}
			if (this->field_88)
			{
				v4 = this->field_98;
				if (v4 < 2 || v4 == 4)
				{
					hr = v16->GetDetailsEx(pidl, &PKEY_StartMenu_Group, &varIn);
					if (hr >= 0)
					{
						v5 = VariantToStringWithDefault(varIn, L"");
						if (StrCmpICW(L"programs", v5))
						{
							v6 = VariantToStringWithDefault(varIn, L"");
							if (StrCmpICW(L"internet", v6))
							{
								v18 = this->field_98 != 4 ? this->field_98 : 0;
							}
						}
						VariantClear(&varIn);
					}
				}
				else if (v4 == 3)
				{
					v18 = 0;
				}
				else if (v4 == 2)
				{
					v18 = 1;
				}
			}
			v16->Release();
		}
		if (v18 == -1)
		{
			if (!v14)
				hr = this->_pFolderView->SelectAndPositionItems(1, (LPCITEMIDLIST *)&pidl, 0, 1);
			if (hr >= 0 && this->field_88 && this->field_8C)
				_ActivateItem(v15);
		}
		ILFree(pidl);
	}
	if (this->field_88)
	{
		vt.vt = 3;
		vt.lVal = 0;
		IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 328, 0, &vt, 0);
	LABEL_48:
		if (this->field_88)
		{
			v9 = v18;
			varIn.vt = 3;
			varIn.lVal = v18;
			if (!v14)
			{
				IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 317, 0, &varIn, 0);
				VariantClear(&varIn);
			}
			if (this->field_8C)
			{
				if (v9 != -1)
				{
					IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 301, 0, 0, 0);
				}
			}
		}
	}
}

BOOL SearchView_RegisterClass()
{
	WNDCLASSEX wc;
	ZeroMemory(&wc, sizeof(wc));


	wc.cbSize = sizeof(wc);
	wc.style = CS_GLOBALCLASS;
	wc.lpfnWndProc = CSearchOpenView::s_WndProc;
	wc.hInstance = g_hinstCabinet;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.lpszClassName = WC_SRCHVIEW;
	return RegisterClassEx(&wc);
}

#endif