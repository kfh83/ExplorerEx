#include "pch.h"

#include "NSCHost.h"

#include "cabinet.h"

#include "HostUtil.h"
#include "ShUndoc.h"
#include "ShGuidP.h"

HRESULT CNSCHost::QueryInterface(REFIID riid, void **ppvObj)
{
	static const QITAB qit[] =
	{
		QITABENT(CNSCHost, INameSpaceTreeControlEvents),
		QITABENT(CNSCHost, INameSpaceTreeAccessible),
		QITABENT(CNSCHost, INameSpaceTreeControlCustomDraw),
		QITABENT(CNSCHost, INameSpaceTreeControlDropHandler),
		QITABENT(CNSCHost, IServiceProvider),
		QITABENT(CNSCHost, IOleCommandTarget),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObj);
}

ULONG CNSCHost::AddRef()
{
	return InterlockedIncrement(&_cRef);
}

ULONG CNSCHost::Release()
{
	ASSERT(0 != _cRef); // 153
	ULONG cRef = InterlockedDecrement(&_cRef);
	if (cRef == 0)
	{
		delete this;
	}
	return cRef;
}

HRESULT CNSCHost::OnItemClick(IShellItem *psi, NSTCEHITTEST nstceHitTest, NSTCECLICKTYPE nstceClickType)
{
	HRESULT hr = S_FALSE;
	if ((nstceClickType & NSTCECT_BUTTON) == NSTCECT_LBUTTON && (nstceHitTest & 0x6E) != 0)
	{
		hr = _Invoke(psi, FALSE);
	}
	return hr;
}

HRESULT CNSCHost::OnPropertyItemCommit(IShellItem *psi)
{
	return S_OK;
}

HRESULT CNSCHost::OnItemStateChanging(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState)
{
	return S_OK;
}

HRESULT CNSCHost::OnItemStateChanged(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState)
{
	return S_OK;
}

HRESULT CNSCHost::OnSelectionChanged(IShellItemArray *psiaSelection)
{
	return S_OK;
}

HRESULT CNSCHost::OnKeyboardInput(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if ((uMsg == WM_KEYDOWN || uMsg == WM_SYSKEYDOWN)
		&& (wParam == VK_RETURN || wParam == VK_SPACE))
	{
		IShellItem *psi;
		if (SUCCEEDED(_GetSelectedItem(&psi)))
		{
			_Invoke(psi, FALSE);
			psi->Release();
		}
	}
	return S_OK;
}

HRESULT CNSCHost::OnBeforeExpand(IShellItem *psi)
{
	field_5C = 1;
	// SHTracePerf(&ShellTraceId_Explorer_StartPane_AllProgram_Folder_Open_Start);
	return S_OK;
}

HRESULT CNSCHost::OnAfterExpand(IShellItem *psi)
{
	// SHTracePerf(&ShellTraceId_Explorer_StartPane_AllProgram_Folder_Open_Stop);
	return S_OK;
}

HRESULT CNSCHost::OnBeginLabelEdit(IShellItem *psi)
{
	_NotifyCaptureInput(TRUE);
	return S_OK;
}

HRESULT CNSCHost::OnEndLabelEdit(IShellItem *psi)
{
	_NotifyCaptureInput(FALSE);
	return S_OK;
}

HRESULT CNSCHost::OnGetToolTip(IShellItem *psi, LPWSTR pszTip, int cchTip)
{
	ASSERT(pszTip); // 654

	*pszTip = 0;

	SFGAOF attr = SFGAO_FOLDER;
	if (psi->GetAttributes(SFGAO_FOLDER, &attr) < 0 || (attr & SFGAO_FOLDER) != 0)
		return S_OK;

	IQueryInfo *pqi;
	HRESULT hr = psi->BindToHandler(NULL, BHID_SFUIObject, IID_PPV_ARGS(&pqi));
	if (SUCCEEDED(hr))
	{
		LPWSTR pszInfoTip;
		hr = pqi->GetInfoTip(QITIPF_LINKNOTARGET, &pszInfoTip);
		if (SUCCEEDED(hr) && pszInfoTip)
		{
			StringCchCopyW(pszTip, cchTip, pszInfoTip);
			CoTaskMemFree(pszInfoTip);
		}
		pqi->Release();
	}
	return hr;
}

HRESULT CNSCHost::OnBeforeItemDelete(IShellItem *psi)
{
	return S_OK;
}

HRESULT CNSCHost::OnItemAdded(IShellItem *psi, BOOL fIsRoot)
{
	return S_OK;
}

HRESULT CNSCHost::OnItemDeleted(IShellItem *psi, BOOL fIsRoot)
{
	return S_OK;
}

HRESULT CNSCHost::OnBeforeContextMenu(IShellItem *psi, REFIID riid, void **ppv)
{
	*ppv = NULL;
	return E_NOTIMPL;
}

HRESULT CNSCHost::OnAfterContextMenu(IShellItem *psi, IContextMenu *pcmIn, REFIID riid, void **ppv)
{
	*ppv = NULL;
	return E_NOTIMPL;
}

HRESULT CNSCHost::OnBeforeStateImageChange(IShellItem *psi)
{
	return E_NOTIMPL;
}

HRESULT CNSCHost::OnGetDefaultIconIndex(IShellItem *psi, int *piDefaultIcon, int *piOpenIcon)
{
	return E_NOTIMPL;
}

HRESULT CNSCHost::OnGetDefaultAccessibilityAction(IShellItem *psi, BSTR *pbstrDefaultAction)
{
	*pbstrDefaultAction = 0;

	WCHAR szRole[MAX_PATH];
	if (GetRoleText(0xFFFFFCB6, szRole, ARRAYSIZE(szRole)))
		*pbstrDefaultAction = SysAllocString(szRole);
	return *pbstrDefaultAction != 0 ? S_OK : E_OUTOFMEMORY;
}

HRESULT CNSCHost::OnDoDefaultAccessibilityAction(IShellItem *psi)
{
	return _Invoke(psi, TRUE);
}

HRESULT CNSCHost::OnGetAccessibilityRole(IShellItem *psi, VARIANT *pvarRole)
{
	SFGAOF attr = SFGAO_FOLDER;
	if (SUCCEEDED(psi->GetAttributes(SFGAO_FOLDER, &attr)) && (attr & SFGAO_FOLDER) != 0)
		pvarRole->lVal = ROLE_SYSTEM_OUTLINEITEM;
	else
		pvarRole->lVal = ROLE_SYSTEM_MENUITEM;

	pvarRole->vt = VT_I4;
	return S_OK;
}

HRESULT CNSCHost::PrePaint(HDC hdc, RECT *prc, LRESULT *plres)
{
	*plres = CDRF_NOTIFYITEMDRAW;
	return S_OK;
}

HRESULT CNSCHost::PostPaint(HDC hdc, RECT *prc)
{
	return S_OK;
}

HRESULT CNSCHost::ItemPrePaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem, COLORREF *pclrText, COLORREF *pclrTextBk, LRESULT *plres)
{
	if ((pnstccdItem->uItemState & (CDIS_SELECTED | CDIS_HOT)))
	{
		if (!_hTheme)
		{
			*pclrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
			*pclrTextBk = GetSysColor(COLOR_HIGHLIGHT);
		}
	}
	else
	{
		if (IUnknown_QueryServiceExec(_punkSite, SID_SM_MFU, &SID_SM_DV2ControlHost, 311, 0, NULL, NULL) == S_OK
			&& _IsNewItem(pnstccdItem->psi) == S_OK)
		{
			*pclrText = _clrText;
			*pclrTextBk = _clrTextBk;
		}

		if (SHRestricted(REST_GREYMSIADS) && _IsItemMSIAds(pnstccdItem->psi) == S_OK)
		{
			*pclrText = GetSysColor(COLOR_BTNSHADOW);
		}
	}
	return S_OK;
}

HRESULT CNSCHost::ItemPostPaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem)
{
	return S_OK;
}

HRESULT CNSCHost::OnDragEnter(IShellItem *psiOver, IShellItemArray *psiaData, BOOL fOutsideSource, DWORD grfKeyState, DWORD *pdwEffect)
{
	VARIANT vt;
	vt.vt = VT_BOOL;
	vt.boolVal = VARIANT_TRUE;
	IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 312, 0, &vt, NULL);

	field_64 = fOutsideSource;
	if (field_64 && (grfKeyState & (MK_SHIFT | MK_CONTROL | MK_XBUTTON1)) == 0)
	{
		*pdwEffect = DROPEFFECT_LINK;
	}
	return S_OK;
}

HRESULT CNSCHost::OnDragOver(IShellItem *psiOver, IShellItemArray *psiaData, DWORD grfKeyState, DWORD *pdwEffect)
{
	if (field_64 && (grfKeyState & (MK_SHIFT | MK_CONTROL | MK_XBUTTON1)) == 0)
	{
		*pdwEffect = DROPEFFECT_LINK;
	}
	return S_OK;
}

HRESULT CNSCHost::OnDragPosition(IShellItem *psiOver, IShellItemArray *psiaData, int iNewPosition, int iOldPosition)
{
	HRESULT hr = S_OK;
	if (iNewPosition == -1)
		return E_FAIL;
	return hr;
}

HRESULT CNSCHost::OnDrop(IShellItem *psiOver, IShellItemArray *psiaData, int iPosition, DWORD grfKeyState, DWORD *pdwEffect)
{
	VARIANT vt;
	vt.vt = VT_BOOL;
	vt.boolVal = VARIANT_FALSE;
	IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 312, 0, &vt, NULL);

	field_64 = 0;
	return S_OK;
}

HRESULT CNSCHost::OnDropPosition(IShellItem *psiOver, IShellItemArray *psiaData, int iNewPosition, int iOldPosition)
{
	return S_OK;
}

HRESULT CNSCHost::OnDragLeave(IShellItem *psiOver)
{
	VARIANT vt;
	vt.vt = VT_BOOL;
	vt.boolVal = VARIANT_FALSE;
	IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 312, 0, &vt, NULL);

	field_64 = 0;
	return S_OK;
}

HRESULT CNSCHost::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
	HRESULT hr = E_FAIL;
	if (IsEqualGUID(guidService, SID_SM_NSCHOST))
	{
		return QueryInterface(riid, ppvObject);
	}
	return hr;
}

HRESULT CNSCHost::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

HRESULT CNSCHost::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
	HRESULT hr = E_INVALIDARG;

	if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
	{
		if (nCmdID == 302)
		{
			if (nCmdexecopt == 1 && (!pvarargIn || pvarargIn->boolVal != VARIANT_TRUE))
			{
				IShellItem *psiFirstVisible = NULL;
				BOOL bNeedEnsureVisible = TRUE;
				if (_pns)
				{
					if (SUCCEEDED(_pns->GetNextItem(0, NSTCGNI_FIRSTVISIBLE, &psiFirstVisible)))
					{
						if (IUnknown_QueryServiceExec(_punkSite, SID_SM_MFU, &SID_SM_DV2ControlHost, 311, 0, NULL, NULL) == S_OK)
						{
							IShellItem *psiCurrent = psiFirstVisible;
							psiCurrent->AddRef();

							IShellItem *psiNext;
							while (bNeedEnsureVisible && _pns->GetNextItem(psiCurrent, NSTCGNI_NEXT, &psiNext) == S_OK)
							{
								IUnknown_Set((IUnknown **)&psiCurrent, psiNext);

								hr = _IsNewItem(psiCurrent);
								if (hr == S_OK)
								{
									_pns->EnsureItemVisible(psiCurrent);
									bNeedEnsureVisible = FALSE;
								}
								psiNext->Release();
							}

							psiCurrent->Release();
						}

						SetFocus(_hwnd);
						if (bNeedEnsureVisible)
						{
							_pns->EnsureItemVisible(psiFirstVisible);
						}

						psiFirstVisible->Release();
					}
				}
			}
		}
		else if (nCmdID == 318)
		{
			ASSERT(pvarargOut->vt == VT_BOOL); // 1047
			pvarargOut->iVal = (field_5C != 0) - 1;
			return S_OK;
		}
	}
	return hr;
}

CLIPFORMAT g_cfPreferredEffect;

void WINAPI InitClipboardFormats()
{
	if (!g_cfPreferredEffect)
	{
		g_cfPreferredEffect = RegisterClipboardFormat(TEXT("Preferred DropEffect"));
	}
}

CNSCHost::CNSCHost()
	: _cRef(1)
{
	InitClipboardFormats();
}


CNSCHost::~CNSCHost()
{
	ASSERT(!_pns); // 240
	ASSERT(!_pweh); // 241
}

LRESULT CNSCHost::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CNSCHost *self = reinterpret_cast<CNSCHost *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	if (uMsg > 0x46)
	{
		switch (uMsg)
		{
			case WM_NOTIFY:
			{
				HRESULT hr = self->_OnNotify(hwnd, uMsg, wParam, lParam);
				if (SUCCEEDED(hr))
				{
					return hr;
				}
				return self->_OnWinEvent(hwnd, uMsg, wParam, lParam);	
			}
			case WM_NCCREATE:
				return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);
			case WM_NCDESTROY:
				return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
			case WM_PALETTECHANGED:
				return self->_OnWinEvent(hwnd, uMsg, wParam, lParam);
		}
		if (uMsg == WM_APP)
		{
			return self->_CollapseAll();
		}
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
	switch (uMsg)
	{
		case WM_WINDOWPOSCHANGING:
			return self->_OnSize(hwnd, uMsg, wParam, lParam);
		case WM_CREATE:
			return self->_OnCreate(hwnd, uMsg, wParam, lParam);
		case WM_DESTROY:
			return self->_OnDestroy(hwnd, uMsg, wParam, lParam);
		case WM_ERASEBKGND:
			return self->_OnEraseBkGnd(hwnd, uMsg, wParam, lParam);
	}
	if (uMsg == WM_SYSCOLORCHANGE || uMsg == WM_SETTINGCHANGE)
	{
		return self->_OnWinEvent(hwnd, uMsg, wParam, lParam);
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CNSCHost::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPCREATESTRUCT lpcs = (LPCREATESTRUCT)lParam;
	SMPANEDATA *psmpd = (SMPANEDATA *)lpcs->lpCreateParams;

	IUnknown_Set(&psmpd->punk, SAFECAST(this, IServiceProvider*));
	_hTheme = psmpd->hTheme;
	if (_hTheme)
	{
		GetThemeColor(_hTheme, SPP_NSCHOST, 0, TMT_INFOTEXT, &_clrText);
		GetThemeColor(_hTheme, SPP_NSCHOST, 0, TMT_INFOBK, &_clrTextBk);

		GetThemeMargins(_hTheme, NULL, SPP_NSCHOST, 0, TMT_CONTENTMARGINS, NULL, &_margins);
	}
	else
	{
		_clrText = GetSysColor(COLOR_INFOTEXT);
		_clrTextBk = GetSysColor(COLOR_INFOBK);
	}

	HRESULT hr = _InitializeNSC(hwnd);
	if (_hTheme)
	{
		SIZE siz = { 0, 0 };
		HDC hdc = GetDC(_hwnd);
		if (hdc)
		{
			GetThemePartSize(_hTheme, hdc, SPP_PROGLISTSEPARATOR, 0, NULL, TS_DRAW, &siz);
			ReleaseDC(_hwnd, hdc);
		}
		field_68 = siz.cy;
	}
	else
	{
		field_68 = SHGetSystemMetricsScaled(SM_CYEDGE);
	}
	return hr;
}

LRESULT CNSCHost::_OnDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (_pns)
	{
		_pns->TreeUnadvise(_dwCookie);
	}

	IUnknown_SetSite(_pns, NULL);

	IUnknown_SafeReleaseAndNullPtr(&_pns);
	IUnknown_SafeReleaseAndNullPtr(&_pweh);
	IUnknown_SafeReleaseAndNullPtr(&_psif);
	IUnknown_SafeReleaseAndNullPtr(&_psf);
	ILFree(_pidl);
	return 0;
}

LRESULT CNSCHost::_OnEraseBkGnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	RECT rc;
	GetClientRect(hwnd, &rc);
	if (_hTheme)
	{
		if (IsCompositionActive())
		{
			SHFillRectClr((HDC)wParam, &rc, 0);
		}
		DrawThemeBackground(_hTheme, (HDC)wParam, SPP_NSCHOST, 0, &rc, NULL);
	}
	else
	{
		DWORD SysColor = GetSysColor(COLOR_MENU);
		SHFillRectClr((HDC)wParam, &rc, SysColor);
	}

	rc.left += SHGetSystemMetricsScaled(SM_CXEDGE) + _margins.cxLeftWidth;
	rc.right += SHGetSystemMetricsScaled(SM_CXEDGE) - _margins.cxRightWidth;
	rc.top = rc.bottom - field_68;
	if (_hTheme)
	{
		DrawThemeBackground(_hTheme, (HDC)wParam, SPP_PROGLISTSEPARATOR, 0, &rc, NULL);
	}
	else
	{
		DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_TOPLEFT);
	}
	return 0;
}

LRESULT CNSCHost::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CNSCHost *psch = new CNSCHost();
	if (psch)
	{
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)psch);
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
	return 0;
}

LRESULT CNSCHost::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SetWindowLongPtr(hwnd, GWLP_USERDATA, NULL);
	LRESULT lRes = DefWindowProc(hwnd, uMsg, wParam, lParam);
	if (this)
	{
		this->Release();
	}
	return lRes;
}

struct SMNMMENUBAND
{
	NMHDR nm;
	IMenuBand *pmb;
};

LRESULT CNSCHost::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HRESULT hr = E_FAIL;

	NMHDR *pnmh = (NMHDR *)lParam;
	if (pnmh)
	{
		switch (pnmh->code)
		{
		case SMN_APPLYREGION:
			return HandleApplyRegion(hwnd, _hTheme, (PSMNMAPPLYREGION)lParam, SPP_NSCHOST, 0);
		case SMN_GETMINSIZE:
			return _OnSMNGetMinSize((PSMNGETMINSIZE)lParam);
		case 215:
			return _OnSMNFindItemWorker((PSMNDIALOGMESSAGE)lParam);
		case 222:
			PostMessage(hwnd, WM_APP, 0, 0);
			return 0;
		case 223:
			return SUCCEEDED(SetSite(((SMNMMENUBAND *)pnmh)->pmb));
		}
	}
	return hr;
}

LRESULT CNSCHost::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
	return 0; // EXEX-Vista(allison): TODO.
}

LRESULT CNSCHost::_OnSMNGetMinSize(PSMNGETMINSIZE psmngms)
{
	LRESULT lres = SendMessage(this->_hwnd, 0x1110, 0, 0);
	psmngms->siz.cy = this->_margins.cyTopHeight + this->_margins.cyBottomHeight + lres * SendMessage(this->_hwnd, 0x111C, 0, 0);
	return 0;
}

LRESULT CNSCHost::_OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HRESULT hr = S_OK;

	LPWINDOWPOS pwp = reinterpret_cast<LPWINDOWPOS>(lParam);
	if (_pns && (pwp->flags & SWP_NOSIZE) == 0)
	{
		int iHeight = pwp->cy - _margins.cyBottomHeight - _margins.cyTopHeight;
		int iWidth = pwp->cx - _margins.cxRightWidth - _margins.cxLeftWidth;

		IVisualProperties *pvp;
		hr = _pns->QueryInterface(IID_PPV_ARGS(&pvp));
		if (SUCCEEDED(hr))
		{
			int iItemHeight = -1;
			hr = pvp->GetItemHeight(&iItemHeight);
			if (SUCCEEDED(hr) && iItemHeight > 1 && iHeight % iItemHeight)
			{
				iHeight -= iHeight % iItemHeight;
			}
			pvp->Release();
		}
		SetWindowPos(_hwnd, NULL, _margins.cxLeftWidth, _margins.cyTopHeight, iWidth, iHeight, SWP_NOZORDER);
	}
	return hr;
}

LRESULT CNSCHost::_OnWinEvent(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lres = 0;
	if (_pweh)
	{
		_pweh->OnWinEvent(hwnd, uMsg, wParam, lParam, &lres);
	}
	return lres;
}

BOOL CNSCHost::_AreChangesRestricted()
{
	return IsRestrictedOrUserSetting(HKEY_CURRENT_USER, REST_NOCHANGESTARMENU, L"Advanced", L"Start_EnableDragDrop", 0);
}

// Thanks to ep_taskbar by @amrsatrio
HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder **psfOut, ITEMIDLIST_ABSOLUTE **pidlOut)
{
	if (psfOut)
		*psfOut = NULL;

	*pidlOut = NULL;

	WCHAR szPath[47] = L"shell:::";
	StringFromGUID2(rclsid, &szPath[8], 39);

	ITEMIDLIST_ABSOLUTE *pidl;
	HRESULT hr = SHILCreateFromPath(szPath, &pidl, NULL);
	if (SUCCEEDED(hr))
	{
		if (psfOut)
		{
			hr = SHBindToObject(NULL, pidl, NULL, IID_PPV_ARGS(psfOut));
		}

		if (SUCCEEDED(hr))
		{
			*pidlOut = pidl;
			pidl = NULL;
		}

		ILFree(pidl);
	}

	return hr;
}

#include "cocreateinstancehook.h"

HRESULT CNSCHost::_InitializeNSC(HWND hwnd)
{
	HRESULT hr = CoCreateInstance(CLSID_NamespaceTreeControl, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pns));
	if (SUCCEEDED(hr))
	{
		IUnknown_SetSite(_pns, static_cast<INameSpaceTreeControlEvents*>(this));

		_pns->QueryInterface(IID_PPV_ARGS(&_pweh));

		hr = _pns->TreeAdvise(static_cast<INameSpaceTreeControlEvents*>(this), &_dwCookie);
		if (SUCCEEDED(hr))
		{
			NSTCSTYLE v5;
			if (_AreChangesRestricted())
				v5 = 0x585808;
			else
				v5 |= NSTCS_DISABLEDRAGDROP;

			if (_SHRegGetBoolValueFromHKCUHKLM(REGSTR_PATH_STARTPANE_SETTINGS, TEXT("Start_SortByName"), TRUE))
			{
				v5 |= NSTCS_NOORDERSTREAM;
			}

			hr = _pns->Initialize(hwnd, NULL, v5);
			if (SUCCEEDED(hr))
			{
				LPITEMIDLIST pidl;
				hr = BindToGetFolderAndPidl(CLSID_ProgramsFolderAndFastItems, NULL, &pidl);
				if (SUCCEEDED(hr))
				{
					IShellItem *psi;
					hr = SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi));
					if (SUCCEEDED(hr))
					{
						CoCreateInstance(CLSID_PersonalStartMenu, 0, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_psif));
						if (this->_psif)
						{
							IUnknown_SetSite(_psif, static_cast<IServiceProvider *>(this));
						}

						hr = _pns->AppendRoot(psi, 96, 3, _psif);
						if (SUCCEEDED(hr))
						{
							hr = IUnknown_GetWindow(_pns, &this->_hwnd);
						}
						psi->Release();
					}
					ILFree((LPITEMIDLIST)pidl);
				}

				LPCWSTR pszTheme = IsCompositionActive() ? L"StartMenuHoverComposited" : L"StartMenuHover";
				_pns->SetTheme(pszTheme);

#if 1
				WCHAR szFallback[MAX_PATH];
				ExpandEnvironmentStringsW(L"%PROGRAMDATA%\\Microsoft\\Windows\\Start Menu\\Programs", szFallback, MAX_PATH);
				IShellItem *psi = nullptr;
				if (SUCCEEDED(SHCreateItemFromParsingName(szFallback, nullptr, IID_PPV_ARGS(&psi))))
				{
					_pns->AppendRoot(psi, 96, 3, nullptr);
					psi->Release();
				}
#endif
			}
		}
	}
	return hr;
}

HRESULT CNSCHost::_CollapseAll()
{
	if (_pns)
	{
		_pns->CollapseAll();
	}
	return 1;
}

HRESULT CNSCHost::_GetSelectedItem(IShellItem **ppsi)
{
	IShellItemArray *psia;
	HRESULT hr = _pns->GetSelectedItems(&psia);
	if (SUCCEEDED(hr))
	{
		hr = psia->GetItemAt(0, ppsi);
		psia->Release();
	}
	return hr;
}

LRESULT _SendNotify(HWND hwndFrom, UINT code, OPTIONAL NMHDR *pnm);

void CNSCHost::_NotifyCaptureInput(BOOL fBlock)
{
	SMNMBOOL nmb;
	nmb.f = fBlock;
	_SendNotify(GetParent(_hwnd), SMN_BLOCKMENUMODE, &nmb.hdr);
}

HRESULT CNSCHost::_Invoke(IShellItem *psi, BOOL fKeyboard)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CNSCHost::_IsItemMSIAds(IShellItem *psi)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

HRESULT CNSCHost::_IsNewItem(IShellItem *psi)
{
	return E_NOTIMPL; // EXEX-Vista(allison): TODO.
}

BOOL NSCHost_RegisterClass()
{
	WNDCLASS wc;
	ZeroMemory(&wc, sizeof(wc));

	wc.style = CS_GLOBALCLASS;
	wc.cbWndExtra = sizeof(CNSCHost *);
	wc.lpfnWndProc = CNSCHost::s_WndProc;
	wc.hInstance = g_hinstCabinet;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.lpszClassName = WC_NSCHOST;
	return RegisterClass(&wc);
	//return SHRegisterClass(&wc);
}