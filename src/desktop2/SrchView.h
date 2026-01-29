#ifndef _SRCHVIEW_H_
#define _SRCHVIEW_H_

#define COMPILE_SRCHVIEW

#ifdef COMPILE_SRCHVIEW

#include "desktop2.h"

#include "HostUtil.h"
#include "COWSite.h"
#include <structuredquery.h>

#define WC_SRCHVIEW  TEXT("Desktop Search Open View")

#pragma region Utilities

// Special thanks to ep_taskbar by @amrsatrio for the following interfaces/structs/enums,
// the only thing changed is the styling to match the rest of the codebase
MIDL_INTERFACE("5D1E7C84-1596-49A2-A89F-4E3A431B0920")
INavigationOptions : IUnknown
{
	STDMETHOD(CanNavigateToIDList)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
};

MIDL_INTERFACE("26B79130-4C9F-4424-AEFB-52CC63F4D3C6")
IContextMenuModifier : IUnknown
{
	STDMETHOD(GetContextMenu)(IContextMenu *, IContextMenu **) PURE;
};

MIDL_INTERFACE("A8B505A9-10A6-45D1-90AA-77424C8CB6F3")
IAccessibleProvider : IUnknown
{
	STDMETHOD(CreateAccessibleObject)(HWND, LONG, REFIID, void **) PURE;
};

typedef enum BROWSER_VIEW_FLAGS
{
	BVF_DEFAULT = 0x0,
	BVF_NOLINKOVERLAY = 0x1,
	BVF_FORCETHUMBNAILDISPLAY = 0x2,
	BVF_NOEXPANDOBUTTONS = 0x4,
	BVF_NOHITHIGHLIGHTING = 0x10,
	BVF_NOTRYHARDER = 0x20,
	BVF_STARTMENUMODE = 0x40,
	BVF_NOSCROLLBARS = 0x80,
	BVF_SUPPORTTILEVIEWINFO = 0x100,
	BVF_NOINTERMEDIATEWINDOW = 0x200,
	BVF_PREVENTWINDOWACTIVATION = 0x400,
	BVF_NOWAITONENUMERATION = 0x800,
	BVF_ITEMCOLLECTIONPROVIDER = 0x1000,
} BROWSER_VIEW_FLAGS;

DEFINE_ENUM_FLAG_OPERATORS(BROWSER_VIEW_FLAGS);

MIDL_INTERFACE("DD1E21CC-E2C7-402C-BF05-10328D3F6BAD")
IBrowserSettings : IUnknown
{
	STDMETHOD(GetEnumerationTimeout)(DWORD *) PURE;
	STDMETHOD(GetViewFlags)(BROWSER_VIEW_FLAGS *) PURE;
	STDMETHOD(GetSearchBoxTimerDelay)(UINT *) PURE;
};

enum FC_FILTERNAME
{
	FCFN_DISPLAY = 0,
	FCFN_INFOLDER = 1,
};

enum FC_FLAGS
{
	FCT_DEFAULT = 0x0,
	FCT_WORDWHEEL = 0x1,
	FCT_FORMATFORDISPLAY = 0x2,
	FCT_FREEFORMED = 0x4,
};

DEFINE_ENUM_FLAG_OPERATORS(FC_FLAGS);

MIDL_INTERFACE("FCA2857D-1760-4AD3-8C63-C9B602FCBAEA")
IFilterCondition : IPersistStream
{
	STDMETHOD(GetFilterName)(FC_FILTERNAME, WCHAR **) PURE;
	STDMETHOD(GetTypeFlags)(FC_FLAGS *) PURE;
	STDMETHOD(GetPropertyKey)(PROPERTYKEY *) PURE;
	STDMETHOD(GetCondition)(ICondition **) PURE;
};

// _GUID_bf78cc76_73e3_4c61_8822_f5f651a9c6d4 = Vista guid
MIDL_INTERFACE("A5DA0C32-4F6C-4AAA-A6BC-A7C1FE27A059")
IFilterView : IUnknown
{
	STDMETHOD(FilterByCondition)(IFilterCondition *) PURE;
	STDMETHOD(FilterByText)(LPCWSTR, LPCWSTR) PURE;
	STDMETHOD(StackByProperty)(REFPROPERTYKEY) PURE;
	STDMETHOD(FilterContent)() PURE;
};

typedef enum tagSM_QUERY_STATUS
{
	SMQS_PROGRAMS_QUERY_DONE = 0x1,
} SM_QUERY_STATUS;

MIDL_INTERFACE("7D61FB0F-8ECB-4BF2-A196-B32C2545CDC4")
IStartMenuQuerySink : IUnknown
{
	STDMETHOD(SetStatusForQuery)(LPCWSTR, SM_QUERY_STATUS) PURE;
};


MIDL_INTERFACE("B99D542E-CABE-4AD8-ADE5-AEBCCDDB1509")
IStartMenuQueryCache : IUnknown
{
	STDMETHOD(InitializeCache)() PURE;
	STDMETHOD(ReleaseCache)() PURE;
};

#pragma endregion

class CPathCompleteInfo
{
public:
	CPathCompleteInfo()
		: _psz1(NULL)
		, _psz2(NULL)
	{
	}

	~CPathCompleteInfo()
	{
		_psz2 = NULL;
		CoTaskMemFree(_psz2);
		_psz1 = NULL;
		CoTaskMemFree(_psz1);
	}

	LPWSTR _psz1;
	LPWSTR _psz2;
};

#include "RunTask.h"
#include "shstr.h"

class CPathCompletionTask : public CRunnableTask
{
public:
	CPathCompletionTask(IUnknown *punk, IExplorerBrowser *peb, HWND hwnd, WCHAR *pszPath)
		: CRunnableTask(RTF_DEFAULT)
		, _punk(punk)
		, _peb(peb)
		, _hwnd(hwnd)
		, _pszPath(pszPath)
	{
		IUnknown_Set(&_punk, punk);
		peb->AddRef();
		PathRemoveBlanks(pszPath);
	}

	~CPathCompletionTask()
	{
		CoTaskMemFree(_pszPath);
		IUnknown_SafeReleaseAndNullPtr(&_peb);
	}

	// *** IRunnableTask ***
	STDMETHODIMP InternalResumeRT()
	{
		const WCHAR *pszPath; // eax
		int v3; // eax
		LPWSTR v4; // eax
		LPWSTR v5; // edi
		const WCHAR *v6; // eax
		LPWSTR pszStr; // ebx
		struct IACList2Vtbl *lpVtbl; // ecx
		LPWSTR *v9; // eax
		LPWSTR *v10; // eax
		LPWSTR *v12; // eax
		struct IPersistIDList *v14; // [esp+8h] [ebp-A8h] BYREF
		int v15; // [esp+Ch] [ebp-A4h]
		LPITEMIDLIST pidl; // [esp+10h] [ebp-A0h] BYREF
		HRESULT hr; // [esp+14h] [ebp-9Ch]
		struct IACList2 *ppv; // [esp+18h] [ebp-98h] BYREF
		struct CPathCompleteInfo *ppci; // [esp+1Ch] [ebp-94h] MAPDST
		ShStrW v20; // [esp+20h] [ebp-90h] BYREF

		hr = CoCreateInstance(CLSID_ACListISF, 0, 1u, IID_PPV_ARGS(&ppv));
		if (hr >= 0)
		{
			ppv->SetOptions(0x2F);

			CPathCompleteInfo *ppci = new CPathCompleteInfo();
			if (!ppci)
				hr = E_OUTOFMEMORY;
			if (hr < 0)
				goto LABEL_42;
			pszPath = this->_pszPath;
			pidl = 0;
			if (pszPath && *pszPath)
			{
				v3 = lstrlenW(pszPath);
				v4 = CharPrevW(this->_pszPath, &this->_pszPath[v3]);
				v15 = 1;
				while (1)
				{
					v5 = v4;
					while (1)
					{
						v6 = this->_pszPath;
						if (v5 != v6)
						{
							if (*v5 == 47)
								goto LABEL_18;
							if (*v5 != 92)
								break;
						}
						if (*v5 == 47 || *v5 == 92)
						{
						LABEL_18:
							//ShStrW::ShStrW(&v20);
							++v5;
							if (v20.SetStr(this->_pszPath) >= 0)
							{
								pszStr = v20.GetInplaceStr();
								if ((((char *)v5 - (char *)this->_pszPath) & 0xFFFFFFFE) != 4 || StrCmpN(v20.GetStr(), L"\\\\", 2))
								{
									pszStr[v5 - this->_pszPath] = 0;
								}
								else
								{
									pszStr[1] = 0;
									v5 -= 2;
								}
								if ((CRunnableTask::ShouldContinue() & 0x80000000) == 0)
								{
									if (ppv->Expand(v20.GetStr()) < 0)
									{
										v5 = CharPrevW(this->_pszPath, v5 - 1);
									}
									else
									{
										v15 = 0;
										if (ppv->QueryInterface(
											IID_IPersistIDList,
											(void **)&v14) >= 0)
										{
											if (v14->GetIDList(&pidl) >= 0)
											{
												if (SHStrDup(v5, &ppci->_psz1) >= 0)
												{
													if (SHStrDup(pszStr, &ppci->_psz2) >= 0 && (CRunnableTask::ShouldContinue() & 0x80000000) == 0)
													{
														PostMessageW(_hwnd, 0x401u, (WPARAM)ppci, (LPARAM)pidl);
														ppci = 0;
													}
												}
											}
											v14->Release();
										}
									}
								}
							}
							v20.Reset();
						}
						if (!v15 || v5 <= this->_pszPath)
							goto LABEL_36;
					}
					v4 = CharPrevW(v6, v5);
				}
			}
		LABEL_36:
			if ((CRunnableTask::ShouldContinue() & 0x80000000) == 0)
			{
				if (!ppci)
				{
				LABEL_42:
					ppv->Release();
					return hr;
				}
				if (SHStrDupW(this->_pszPath, &ppci->_psz2) >= 0)
				{
					PostMessageW(this->_hwnd, 0x401u, (WPARAM)ppci, 0);
					ppci = 0;
				}
			}
			if (ppci)
			{
				delete ppci;
				goto LABEL_42;
			}
		}
		return hr;
	}

	WCHAR *_pszPath;
	IUnknown *_punk;
	IExplorerBrowser *_peb;
	HWND _hwnd;
};

class CSearchOpenView
	: public CUnknown
	, public IServiceProvider
	, public IOleCommandTarget
	, public CObjectWithSite
	, public IExplorerBrowserEvents
	, public IDispatch
	, public INavigationOptions
	, public ICommDlgBrowser3
	, public IContextMenuModifier
	, public IAccessibleProvider
	, public IBrowserSettings
{
public:
	CSearchOpenView();
	~CSearchOpenView();

	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup,
		ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pCmdText);
	STDMETHODIMP Exec(const GUID *pguidCmdGroup,
		DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut);

	// *** IObjectWithSite ***
	STDMETHODIMP SetSite(IUnknown *punkSite);

	// *** IExplorerBrowserEvents ***
	STDMETHODIMP OnNavigationPending(PCIDLIST_ABSOLUTE pidlFolder);
	STDMETHODIMP OnViewCreated(IShellView *psv);
	STDMETHODIMP OnNavigationComplete(PCIDLIST_ABSOLUTE pidlFolder);
	STDMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE pidlFolder);

	// *** IDispatch ***
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	// *** INavigationOptions ***
	STDMETHODIMP CanNavigateToIDList(PCIDLIST_ABSOLUTE pidl1, PCIDLIST_ABSOLUTE pidl2);

	// *** ICommDlgBrowser ***
	STDMETHODIMP OnDefaultCommand(IShellView *ppshv);
	STDMETHODIMP OnStateChange(IShellView *ppshv, ULONG uChange);
	STDMETHODIMP IncludeObject(IShellView *ppshv, PCUITEMID_CHILD pidl);

	// *** ICommDlgBrowser2 ***
	STDMETHODIMP Notify(IShellView *ppshv, DWORD dwNotifyType);
	STDMETHODIMP GetDefaultMenuText(IShellView *ppshv, LPWSTR pszText, int cchMax);
	STDMETHODIMP GetViewFlags(DWORD *pdwFlags);

	// *** ICommDlgBrowser3 ***
	STDMETHODIMP OnColumnClicked(IShellView *ppshv, int iColumn);
	STDMETHODIMP GetCurrentFilter(LPWSTR pszFileSpec, int cchFileSpec);
	STDMETHODIMP OnPreViewCreated(IShellView *ppshv);

	// *** IContextMenuModifier ***
	STDMETHODIMP GetContextMenu(IContextMenu *pcmIn, IContextMenu **ppcmOut);

	// *** IAccessibleProvider ***
	STDMETHODIMP CreateAccessibleObject(HWND hwnd, LONG idObject, REFIID riid, void **ppv);

	// *** IBrowserSettings ***
	STDMETHODIMP GetEnumerationTimeout(DWORD *pdwTimeout);
	STDMETHODIMP GetViewFlags(BROWSER_VIEW_FLAGS *pbvf);
	STDMETHODIMP GetSearchBoxTimerDelay(UINT *puiDelay) { return S_OK; }

	HRESULT Initialize(HWND hwnd);
	HRESULT AddPathCompletionTask(LPCWSTR pszPath);

private:
	static DWORD WINAPI s_ExecuteCommandLine(LPVOID pv);
	static DWORD WINAPI s_ExecuteIDList(LPVOID pv);

	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm);
	LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);
	LRESULT _OnSize(LPWINDOWPOS lpwp);
	int _OnEraseBkGnd(HWND hwnd, HDC hdc);
	int _GetCurSel();
	int _GetItemCount();

	HRESULT _ActivateItem(int iItem);
	HRESULT _AddCheckIndexerStaterTask();
	HRESULT _AddCondition(IObjectArray* poa, REFPROPERTYKEY a3, CONDITION_OPERATION a4, LPCWSTR a5);
	HRESULT _CreateConditions(LPCWSTR pcszURL, ICondition **ppCondition);
	HRESULT _CreateExplorerBrowser(HWND hwnd);
	HRESULT _FilterView(IFilterView *pfv);
	HRESULT _GetItemKeyWord(int itemIndex, LPWSTR pszDest, UINT cchDest);
	HRESULT _GetQueryFactoryWithSink(REFIID riid, void **ppv);
	HRESULT _GetSelectedItemParsingName(LPWSTR pszName, UINT cchName);
	HRESULT _InitPathCompletePidlAutoList(LPCWSTR pszPath, PIDLIST_ABSOLUTE *ppidl);
	HRESULT _InitPidlAutoList(PIDLIST_ABSOLUTE *ppidl);

	HRESULT _InitRegularAutoListItem(IShellItem **ppsi);

	HRESULT _LimitViewResults(IFolderView *pfv, int nMaxItems);
	HRESULT _Parse(IObjectArray *poa, REFPROPERTYKEY pkey, LPCWSTR psz);
	HRESULT _RecreateBrowserObject();
	HRESULT _RegisterQuerySink(IStartMenuQuerySink *psmqs);
	HRESULT _UpdateMRU(LPCWSTR psz);
	HRESULT _UpdateSearchText(LPCWSTR psz);
	HRESULT _UpdateSearchTextFilter();
	HRESULT _AddTextFilterToLocation(LPCITEMIDLIST pidl, IFilterCondition *pfc, LPITEMIDLIST *ppidl);

	FOLDERFLAGS _GetFolderFlags();
	void _CancelNavigation();
	void _ConnectShellView(IShellView *psv);
	void _DisconnectShellView();
	void _DoKeyBoardContextMenu();
	void _FilterPathCompleteView(LPCWSTR pszPath);
	void _InstrumentActivation(int iItem);
	void _PathCompleteUpdate(CPathCompleteInfo *ppciNew, PIDLIST_ABSOLUTE pidl);
	void _ReleaseExplorerBrowser();
	void _RevokeQuerySink();
	void _SizeExplorerBrowser(int cx, int cy);

	typedef enum VIEWMODE
	{
		VIEWMODE_DEFAULT = 0,
		VIEWMODE_PATHCOMPLETE = 1,
	} VIEWMODE;

	void _SwitchToMode(VIEWMODE viewModeNew, int a3);
	void _UpdateIndexState(int iState);
	void _UpdateOpenBoxText();
	void _UpdateScrolling();

	typedef enum tagUTM_REASON
	{
		UTM_REASON_0 = 0,
		UTM_REASON_1 = 1,
	} UTM_REASON;

	void _UpdateTopMatch(UTM_REASON reason);

private:
	IExplorerBrowser* _peb;
	IShellItemArray* _psiaStartMenuProvider;
	IShellItemArray* _psiaStartMenuAutoComplete;
	IFolderView2* _pFolderView;
	IShellTaskScheduler* _psched;
	IDispatch* _pDispatchView;
	IQueryParser* _pqp;
	IConditionFactory2* _pcf;
	IStartMenuQueryCache* _psmqc;
	int field_54;
	DWORD _dwCookie;
	HTHEME _hTheme;
	MARGINS _margins;
	HWND _hwnd;
	DWORD field_74;
	VIEWMODE _viewMode;
	int field_7C;
	CPathCompleteInfo* _ppci;
	int field_84;
	int field_88;
	int field_8C;
	int field_90;
	int field_94;
	DWORD field_98;
	DWORD dword9C;
	LPWSTR _pszGroupPrograms;
	LPWSTR _pszGroupInternet;
	LPWSTR field_A8;
	int field_AC;
	int field_B0;
	int field_B4;
	int field_B8;
	DWORD _dwCookieSink;
	IStartMenuQuerySink* _psmqs;
	int field_C0;
	int field_C4;
	IShellItem* _psiFolder;

	friend BOOL SearchView_RegisterClass();
};

#endif // COMPILE_SRCHVIEW

#endif